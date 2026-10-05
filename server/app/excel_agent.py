from __future__ import annotations

import copy
import json
import logging
import re
import time
import uuid
from collections.abc import Awaitable, Callable
from functools import lru_cache
from typing import Any

from pydantic import ValidationError

from .agent_loop_protection import (
    LoopConfig,
    LoopState,
    detect_action_conflict,
    detect_path_loop,
    get_warning_prompt,
    safe_degradation,
    should_stop,
    update_loop_state,
)
from .action_capability import (
    ActionCapability,
    CapabilityMatrix,
    HostCapabilityMatrix,
    intersect_action_capabilities,
    is_action_allowed_by_keys,
    parse_host_id,
    resolve_host_action_bound,
)
from .capabilities import capabilities, capability_float, capability_int, capability_optional_int
from .column_mapping import (
    BlankPolicyNeededError,
    DEFAULT_MAX_MAPPING_ACTIONS,
    NeedsClarificationError,
    UnmappedCategoryError,
    build_column_mapping_actions,
    validate_column_mapping_shape,
)
from .llm import ModelConnection, OpenAICompatibleClient
from .models import (
    AnalysisPlan,
    AnswerResponse,
    AssistantResponse,
    BlankPolicy,
    ColumnMappingClarification,
    ColumnMappingPair,
    ColumnMappingRule,
    IntentClarification,
    IntentClarificationResponse,
    LocalPlanRuleResponse,
    PlanRequest,
    PlanResponse,
    ResultContext,
)

MAX_MAPPING_SPLIT_ACTIONS = capability_int("columnMapping", "maxSplitActions")


def _max_mapping_split_actions() -> int:
    """单次列映射预览允许拆成的最大写入段数。

    配置读不到时退回 `DEFAULT_MAX_MAPPING_ACTIONS`，而不是"不限" ——
    不设上限会让"保留原值"把一列拆成几百个动作，验收与撤销都不可控。
    """
    return MAX_MAPPING_SPLIT_ACTIONS or DEFAULT_MAX_MAPPING_ACTIONS


MAX_READ_ROWS = capability_int("agent", "maxReadRows")
MAX_READ_COLUMNS = capability_int("agent", "maxReadColumns")
MAX_OUTPUT_TOKENS = capability_optional_int("agent", "maxOutputTokens")
REPAIR_MAX_OUTPUT_TOKENS = capability_optional_int("agent", "repairMaxOutputTokens")

logger = logging.getLogger(__name__)


def _default_loop_config() -> LoopConfig:
    """从共享能力配置构造生产环境 Agent 循环限额。"""
    return LoopConfig(
        max_turns=capability_int("agent", "maxTurns"),
        max_total_tokens=capability_optional_int("agent", "maxTotalTokens"),
        max_wall_time_seconds=capability_float(
            "taskRecovery", "planningHardWallTimeSeconds"
        ),
        max_consecutive_errors=capability_int(
            "agent", "maxConsecutiveErrors"
        ),
        max_cache_hits_per_query=capability_int(
            "agent", "maxCacheHitsPerQuery"
        ),
        max_tool_path_loop_length=capability_int(
            "agent", "maxToolPathLoopLength"
        ),
        warning_turn=capability_int("agent", "warningTurn"),
    )

AGENT_SYSTEM_PROMPT = """
你是 Excel Bro 的规划 Agent。你的工作是理解目标、按需查看工作簿，再提交一个
可预览、可验证的 Excel 操作计划；你不能直接声称已经修改工作簿。

工作原则：
- 先理解用户想得到的结果，再决定是否需要调用工作簿工具。
- 初始用户消息中的工作簿上下文已经包含已授权工作表的准确表头。所需字段与表头
  完全匹配，且动作按表头寻址（如 computeValues、splitGroupAggregate）时，直接
  提交计划，不要再调用 get_workbook_context、find_fields 或 read_range。只有字段
  不在表头中、名称存在歧义，或确实需要定位单元格地址/判断取值时才使用只读工具。
- 不根据样本猜测字段地址；涉及单元格范围、公式引用，且上下文没有足够地址信息时，
  先调用 find_fields 或 read_range。
- 尊重“写公式”和“写计算结果”的区别。用户要求公式时使用 writeFormulas，
  不能用服务端算出的静态值代替。
- 根据用户表达组合通用动作，不依赖固定句式。
- 删除、清空、覆盖、移动数据必须在 warnings 中写清影响。
- 写类动作（removeDuplicates / clearRange / deleteRange / copyRange）支持用
  filters 做“按值选行”：filters 的 field 是列的表头名（不是列号），operator
  通常为 equals，value 是用户点名的取值。filters 只对命中行生效，未命中行
  （含其重复）原样保留、顺序不变。
- 当用户说“把 XX 的数据去重”“删除 XX 的行”“清空 XX 的行”且 XX 是数据中某列的
  取值（而非表头名或工作表名）时，属于按值选行，应填 filters，不要反问范围，
  也不要对整表去重/删除/清空。先 read_range 看数据判断 XX 是“值”还是“字段/范围”。
- 去重带 filters 时，只对命中 filters 的行之间按 columns 判重，每个命中值保留
  第一次出现的行；未命中行全部保留。
- 工作簿快照可能截断；不能把截断样本说成完整数据。
- 当用户要求按某字段拆成多个工作表，并在每张结果表中分组计数、求和或计算
  组内占比时，优先使用 splitGroupAggregate。该动作会在 Excel 端读取源工作表
  的完整使用区域，不受快照 200 行限制；不要根据截断样本生成大量 writeTable。
- splitGroupAggregate 的 splitBy 是输出工作表的拆分字段，groupBy 是每张结果表
  内的分组字段。countRows 表示记录行数；为指标填写 ratioOutputName 时，占比的
  分母是同一 groupBy 分组在全部非空 splitBy 值中的该指标总量。
- splitBy 字段会由执行器自动写入每张结果表，绝不能为了让该字段显示在结果中
  又把它放进 groupBy。groupBy 只填写除 splitBy 之外、用于确定表内每一行的字段。
- 输出工作表名称来自哪个字段，该字段就是 splitBy；每张结果表内每行由哪些字段
  确定，这些字段就是 groupBy。splitGroupAggregate 在执行时读取完整使用区域，
  不要调用 read_range 抽样计数，也不要先计算各组结果。
- 源工作表可能在实际表头前有标题行；splitGroupAggregate 会自动识别同时包含
  所需字段的表头行。existingSheetPolicy 默认使用 rename，避免覆盖已有工作表。
- 有默认值的列表字段没有内容时应省略或提交空列表，不要提交 null。
- 你只规划白名单 Excel 动作，不执行 VBA、宏、脚本或外部程序。

写入方式的优先级（从高到低，只有前一档表达不了才降级）：
1. writeFormulas —— 公式引用单元格，源数据一变 Excel 自己重算。表达得了就用它。
2. computeValues / splitGroupAggregate —— 执行时求值的动作：动作体里**只带查询
   配置，不带结果**，Excel 端在执行那一刻读取源工作表的实时数据算出来。需要
   筛选、分组、聚合、排序、取前 N 这类查询时用 computeValues。
3. writeValues / writeTable —— 只在写入的内容确实是**固定值**时使用：用户点名的
   常量、表头文字、说明文字。

不要把算出来的结果塞进 writeValues.values。那是一张一次性快照：源数据改一格
结果就过期了，而计划本身无法重算，用户只能重新提问。同样的需求用
computeValues 表达，改完数据直接重跑就是新结果。

按对应关系把一列翻译成另一列（分类、归类、打标签、编码转名称）时，调用
map_column_values，**不要**自己逐格抄写再提交 writeValues。你只给映射规则和
区域，目标值由服务端算出：逐格抄写会漏行，也会把没列出的来源值自行归类，
而这两种错误在单元格层面都看不出来。用户没说空白单元格写什么，或来源值未
被规则覆盖时，先问清楚，不要替用户决定。

map_column_values 的匹配顺序是**先整格、后拆分**：

1. 整格精确匹配优先：来源单元格整体是一个已知值（无论它看起来是否含分隔符），
   就直接用这条映射，不拆分。"甲 乙"整格就是映射表里的键时，结果取它。
2. 整格没命中、且规则声明了分隔符时，只按**已声明的**分隔符拆分（多个候选时
   按在格中出现位置最靠前者切分），逐项去首尾空格后**按完整名称精确匹配**，
   命中项之间取规则指定的**最高级**（`valueRank` 里 rank 最大者）。
   注意 `rank` 表达的是**优先级**，不是等级名的字面数字：用户说更严重/更优先的
   等级必须给更大的 rank。等级名里带数字时（命名习惯常与严重程度反向），
   千万不要照抄成 rank —— 填反了不会报错，只会安静地取到较低的那个等级。
3. 拆出来的项在映射表里找不到 → 归因**确定**是"类别没覆盖"：让用户
   **补充类别映射**即可，不要说成"分隔形式没声明"去混淆。
4. 整格没命中、又看得出格子里像有别的分隔写法（`undeclaredSeparators` 里
   登记过的那种）→ 归因**不确定**：只能说"可能需要补充分隔符、或补充映射"，
   **不要**替用户判定到底是哪一种，更**不要**自己拿一个没声明过的写法去拆。

也就是说：`delimiters` 与 `valueRank` 必须**成对**给出。只声明分隔符、不说拆分后
取哪一个，本工具会判定规则不完整并返回澄清，而不按映射表顺序硬排一个顺序
（那个顺序是给模型的声明顺序，不代表用户说的优先级）。

- 用户没讲清楚"格子里的多个值怎么处理"时，先 submit_clarification 问清楚
  （是否按分隔符拆分、按哪个分隔符、拆分后取哪一个），拿到确切答复后再调用；
  不要用未经确认的假设直接产出映射计划，也不要自行决定拆分后写哪一个
  （例如全部写出、取其一、取某种"最高级"）。
- 澄清要**一次说清**：把未命中的项合并成一条问题抛出，不要逐行、逐格地反复问。
- "某个具体来源值 → 某个具体目标标签"这类显式对应项照常工作，不要因为这条
  提示就把已经明确的单值映射也一并拿去澄清。

- 若请求包含“本地数据工具结果”，该结果由 Excel 端对用户授权范围进行确定性
  计算，可以直接引用它的 headers、rows、calculation 和完整性标记来**回答问题**
  （submit_answer），不必再调 read_range 重复读取，也不要自行改变计算口径。
  但要**写回工作表**时，不要把这些 rows 当静态值塞进 writeValues —— 改用
  computeValues 复述同一套查询条件，让 Excel 在执行时重算。

缺少用户决定时必须调用 submit_clarification，不要把问题写进 submit_answer。
fields 为每项必须补充的信息提供一个独立输入框（id、label、placeholder）；
有明确可选方案时提供 options，否则 options 用空数组。只问影响当前结果的关键
缺口，不重复询问用户已说明的内容，不为尚未观察到的假想异常阻塞全部任务。
submit_answer 仅用于无需用户补充信息的最终回答；澄清不会写入或执行任何计划。

完成分析后优先调用 submit_plan、submit_clarification 或 submit_answer。若模型不支持工具调用，也可以
直接返回符合 AssistantResponse JSON Schema 的 JSON。
""".strip()


class AgentBudgetPause(RuntimeError):
    """A complete planning boundary was saved and now awaits user budget approval."""

    def __init__(self, checkpoint: dict[str, Any], approved_seconds: float) -> None:
        super().__init__("本轮处理时间已用完，任务现场已保存。")
        self.checkpoint = checkpoint
        self.approved_seconds = approved_seconds


def _column_number(name: str) -> int:
    result = 0
    for character in name.upper():
        result = result * 26 + ord(character) - 64
    return result


def _column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _used_origin(address: str | None) -> tuple[int, int]:
    if not address:
        return 1, 1
    cell = address.rsplit("!", 1)[-1].split(":", 1)[0]
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", cell)
    if not match:
        return 1, 1
    return _column_number(match.group(1)), int(match.group(2))


def _sheet(request: PlanRequest, name: str):
    normalized = name.strip().casefold()
    for sheet in request.workbook.worksheets:
        if sheet.name.casefold() == normalized:
            return sheet
        if sheet.sourceSheet and sheet.sourceSheet.casefold() == normalized:
            return sheet
    raise ValueError(f"未找到工作表「{name}」")


def _workbook_context(request: PlanRequest) -> dict[str, Any]:
    return {
        "name": request.workbook.name,
        "activeWorksheet": request.workbook.activeWorksheet,
        "selectedRange": request.workbook.selectedRange,
        "worksheets": [
            {
                "name": sheet.name,
                "sourceFile": sheet.sourceFile,
                "sourceSheet": sheet.sourceSheet,
                "usedRange": sheet.usedRange,
                "rowCount": sheet.rowCount,
                "columnCount": sheet.columnCount,
                "headers": sheet.headers,
                "truncated": sheet.truncated,
            }
            for sheet in request.workbook.worksheets
        ],
    }


def _read_range(request: PlanRequest, sheet_name: str, address: str) -> dict[str, Any]:
    sheet = _sheet(request, sheet_name)
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?",
        address.strip(),
    )
    if not match:
        raise ValueError("range 必须是 A1 或 A1:B10 格式")
    min_column = _column_number(match.group(1))
    min_row = int(match.group(2))
    max_column = _column_number(match.group(3) or match.group(1))
    max_row = int(match.group(4) or match.group(2))
    if max_column < min_column or max_row < min_row:
        raise ValueError("range 的结束位置不能早于开始位置")
    if (
        max_row - min_row + 1 > MAX_READ_ROWS
        or max_column - min_column + 1 > MAX_READ_COLUMNS
    ):
        raise ValueError(
            f"单次最多读取 {MAX_READ_ROWS} 行、{MAX_READ_COLUMNS} 列"
        )

    origin_column, origin_row = _used_origin(sheet.usedRange)
    values = [sheet.headers, *sheet.dataRows]
    result: list[list[object]] = []
    for row in range(min_row, max_row + 1):
        output_row: list[object] = []
        for column in range(min_column, max_column + 1):
            row_index = row - origin_row
            column_index = column - origin_column
            value = None
            if (
                0 <= row_index < len(values)
                and 0 <= column_index < len(values[row_index])
            ):
                value = values[row_index][column_index]
            output_row.append(value)
        result.append(output_row)
    return {
        "sheet": sheet.name,
        "range": address.upper().replace("$", ""),
        "values": result,
        "snapshotTruncated": sheet.truncated,
    }


def _snapshot_region(
    request: PlanRequest, sheet_name: str, address: str
) -> tuple[Any, list[list[Any]]]:
    """从快照读取区域值，**越界即拒绝，不补 None**。

    `_read_range` 对快照外的格子补 `None`：那对只读查询是合适的降级，模型
    看到空值会继续追问。但映射结果是要**落盘**的 —— 补出来的 `None` 会被
    映射成空白标签，写进实际上没有数据的行，用户看到凭空多出来的一列标签。

    行边界按**实际捕获量**（`len(headers) + len(dataRows)`）判定而不是
    `rowCount`：`rowCount` 是宿主报告的总行数，快照截断时它仍然很大。
    列边界按**声明宽度** `columnCount` 判定：`dataRows` 的行可能因尾部空值
    被裁短，此时区域内缺失的项是真实空格，折成 `None` 是正确的。

    行/列 × 早于原点/超出范围，**四个方向都要拒**。共享夹具
    `config/mapping_slice_cases.json` 与客户端 `sliceMappingSource` 跑同一批
    用例，此前两侧各漏一个方向（本函数不拦左侧列早于原点、客户端不拦右侧列
    溢出），各自静默折成空值。
    """
    sheet = _sheet(request, sheet_name)
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?",
        address.strip(),
    )
    if not match:
        raise ValueError("range 必须是 A1 或 A1:B10 格式")
    min_column = _column_number(match.group(1))
    min_row = int(match.group(2))
    max_column = _column_number(match.group(3) or match.group(1))
    max_row = int(match.group(4) or match.group(2))
    if max_column < min_column or max_row < min_row:
        raise ValueError("range 的结束位置不能早于开始位置")
    if (
        max_row - min_row + 1 > MAX_READ_ROWS
        or max_column - min_column + 1 > MAX_READ_COLUMNS
    ):
        raise ValueError(f"单次最多读取 {MAX_READ_ROWS} 行、{MAX_READ_COLUMNS} 列")

    origin_column, origin_row = _used_origin(sheet.usedRange)
    rows = [sheet.headers, *sheet.dataRows]
    if min_row < origin_row:
        raise ValueError(
            f"区域 {address} 起始于第 {min_row} 行，"
            f"早于「{sheet.name}」已捕获区域的起点第 {origin_row} 行。"
        )
    last_captured_row = origin_row + len(rows) - 1
    if max_row > last_captured_row:
        truncation_note = (
            "（该表快照已截断，请缩小区域或分批处理）" if sheet.truncated else ""
        )
        raise ValueError(
            f"「{sheet.name}」的快照只到第 {last_captured_row} 行，"
            f"区域 {address} 要到第 {max_row} 行，超出已捕获范围{truncation_note}。"
            "映射结果会写入工作表，不能对没读到的行补空值。"
        )
    if min_column < origin_column:
        raise ValueError(
            f"区域 {address} 起始于第 {_column_name(min_column)} 列，"
            f"早于「{sheet.name}」已捕获区域的起点第 "
            f"{_column_name(origin_column)} 列。"
        )
    last_captured_column = origin_column + max(sheet.columnCount, 1) - 1
    if max_column > last_captured_column:
        raise ValueError(
            f"「{sheet.name}」已用区域只到第 {_column_name(last_captured_column)} 列，"
            f"区域 {address} 要到第 {_column_name(max_column)} 列，超出该范围。"
        )

    values: list[list[Any]] = []
    for row in range(min_row, max_row + 1):
        source_row = rows[row - origin_row]
        values.append(
            [
                source_row[column - origin_column]
                if 0 <= column - origin_column < len(source_row)
                else None
                for column in range(min_column, max_column + 1)
            ]
        )
    return sheet, values


def _round_trip_display(value: str) -> str:
    """把 JSON 式转义还原成用户能读的原文（`\\n` → 换行、`\\t` → 制表符）。

    为什么需要这一步：快照里的字符串来自 Excel/WPS 的 JSON 通道，换行会被
    转义成字面量 `\\n`。映射**匹配**用的必须是原样字符串（否则拆分与精确
    匹配都会错位），但报给用户看的必须是真换行 —— 否则澄清卡里会出现
    `"甲\n乙"` 这种给人读的 JSON，用户看不出那几个字符是什么。
    """
    if "\\" not in value:
        return value
    try:
        # `unicode_escape` 会把非 ASCII 也一并转走，所以先按 latin-1 往返
        # 保住原有码点，再只让转义序列生效。
        return value.encode("latin-1", "backslashreplace").decode("unicode_escape")
    except (UnicodeDecodeError, UnicodeEncodeError):
        # 还原失败不是错误路径：原文照旧展示即可，绝不因为展示而改变匹配结果。
        return value


def _mapping_from_pairs(pairs: Any) -> dict[str, str]:
    """把 `[{from,to}]` 折成映射表。

    用成对列表而不是任意键对象：后者在各家 strict function schema 下的支持
    参差不齐。重复的 `from` **报错而不是后者覆盖前者** —— 同一来源值给了两个
    目标标签是需求本身有歧义，静默取一个等于替用户做了没有依据的选择。
    """
    if not isinstance(pairs, list) or not pairs:
        raise ValueError("mapping 至少要有一条「来源值 → 目标标签」。")
    mapping: dict[str, str] = {}
    for entry in pairs:
        if not isinstance(entry, dict):
            raise ValueError("mapping 的每一项都要是 {from, to} 对象。")
        source = entry.get("from")
        target = entry.get("to")
        if not isinstance(source, str) or not isinstance(target, str) or not target:
            raise ValueError("mapping 每一项的 from 与 to 都必须是非空字符串。")
        if source in mapping and mapping[source] != target:
            raise ValueError(
                f"mapping 对来源值「{source}」给了两个目标标签"
                f"（「{mapping[source]}」与「{target}」），请确认后只保留一个。"
            )
        mapping[source] = target
    return mapping


def _string_list(value: Any) -> list[str]:
    """把模型给的字符串列表归一：非列表/含非字符串项一律报错，不做宽容转换。

    静默丢掉一个类型不对的项（比如把 `[1, "、"]` 里的数字忽略）会让规则
    与模型的原意不一致，而这种不一致在单元格层面看不出来。
    """
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError("分隔符参数必须是字符串列表。")
    items: list[str] = []
    for item in value:
        if not isinstance(item, str) or not item:
            raise ValueError("分隔符参数里每一项都必须是非空字符串。")
        items.append(item)
    return items


def _rank_entries(value: Any) -> list[dict[str, Any]]:
    """把 `[{value, rank}]` 归一。`rank` 只要求是数字，不约束取值范围。

    等级**由规则数据给出**，服务端不认识任何具体等级名，也不该给数值
    附加含义（越大越优先是 core 的约定，规则只需一致地提供它）。
    """
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError("等级次序参数必须是 {value, rank} 列表。")
    entries: list[dict[str, Any]] = []
    for entry in value:
        if not isinstance(entry, dict):
            raise ValueError("等级次序的每一项都要是 {value, rank} 对象。")
        item = entry.get("value")
        rank = entry.get("rank")
        if not isinstance(item, str) or not item:
            raise ValueError("等级次序每一项的 value 都必须是非空字符串。")
        if isinstance(rank, bool) or not isinstance(rank, (int, float)):
            raise ValueError("等级次序每一项的 rank 都必须是数字。")
        entries.append({"value": item, "rank": rank})
    return entries


def _blank_policy_argument(
    raw: Any, blank_label: str
) -> dict[str, Any] | None:
    """空白策略参数归一。

    `blankPolicy` 优先；没给时才按旧字段 `blankLabel` 折 —— 非空串等价于
    `{kind:"value"}`，语义与旧版逐格一致。**两者都没有时返回 None（未定）**，
    不是错误。
    """
    if isinstance(raw, dict):
        return BlankPolicy.model_validate(raw).model_dump(mode="json")
    if blank_label:
        return {"kind": "value", "value": blank_label}
    return None


def _blank_policy_summary(policy: dict[str, Any] | None) -> str:
    """给 `assumptions` 用的一句话。**未定时不写任何"空白按 X 处理"。**"""
    if not policy:
        return "源列空白的处理方式尚未确定（本次未读到空白，或等用户选择）。"
    kind = policy.get("kind")
    if kind == "keep":
        return "源列空白：保留目标原值，那个位置不写入。"
    if kind == "clear":
        return "源列空白：清空目标内容。"
    if kind == "value":
        rendered = policy.get("value", "")
        if rendered == "":
            return "源列空白：显式写入空字符串。"
        return f"源列空白写入「{rendered}」。"
    return "源列空白的处理方式尚未确定。"


def _clarification_rule_response(
    *,
    model_arguments: dict[str, Any],
    source_sheet_name: str,
    target_sheet_name: str,
    source_range: str,
    target_range: str,
    mapping: dict[str, str],
    blank_label: str | None,
    reason: str,
    message: str,
    unresolved_count: int,
    resolved_by: list[str],
    missing_categories: list[str] | None = None,
    blank_policy: dict[str, Any] | None = None,
    blank_cell_count: int = 0,
    delimiters: list[str] | None = None,
    undeclared_separators: list[str] | None = None,
    value_rank: list[dict[str, Any]] | None = None,
    display_values: bool = False,
) -> LocalPlanRuleResponse:
    """"规则不全"时的响应：下发规则 + `clarification`，**不下发任何计划**。

    不返回可执行计划是刻意的：归因已经确定"现在算不出正确答案"，
    返回一份建立在猜测之上的矩阵，等于把猜测伪装成结果 —— 用户会以为
    它是对的，而分歧已经静默流进下游统计。

    `message` 里**不嵌入来源格的原始内容**：那属于客户端本地数据，
    走到这里就破坏了"源值不上传"。界面要展示原文，从本端读。

    `display_values=True` 只用在**已经生成本地提示文案**的场合：
    此时 `missing_categories` 只是写入 `clarification`、给客户端界面
    列"还要补哪些类别"用的，不参与任何匹配，所以转义序列可以还原成
    可读原文。未覆盖项**参与匹配时**一律保持原样字符串。
    """
    rendered_missing = [
        _round_trip_display(item) if display_values else item
        for item in (missing_categories or [])
    ]
    return LocalPlanRuleResponse(
        provider="model",
        rule=ColumnMappingRule(
            sourceSheet=source_sheet_name,
            targetSheet=target_sheet_name,
            sourceRange=source_range,
            targetRange=target_range,
            mapping=[
                ColumnMappingPair(source=source, target=target)
                for source, target in mapping.items()
            ],
            blankLabel=blank_label,
            blankPolicy=blank_policy,
            delimiters=delimiters or [],
            valueRank=value_rank or [],
            undeclaredSeparators=undeclared_separators or [],
            clarification=ColumnMappingClarification(
                reason=reason,
                message=message,
                missingCategories=rendered_missing,
                unresolvedCellCount=unresolved_count,
                blankCellCount=blank_cell_count,
                resolvedBy=resolved_by,
            ),
        ),
        summary=message,
        warnings=["本次没有生成计划，工作表未被改动。"],
        assumptions=[],
    )


def _map_column_plan(
    request: PlanRequest, arguments: dict[str, Any]
) -> PlanResponse | LocalPlanRuleResponse:
    """按模型给定的映射参数产出计划，**值由确定性代码算，模型不经手**。

    分工：模型只决定「哪一列、按哪张映射表、空白叫什么」。模型逐格抄写映射
    结果时会漏行、会自行归类未列出的值，而这类错误在单元格层面看不出来。

    谁来算，取决于**源值在谁手上**，这由 `planningReadPolicy` 表达：

    - `"ifNeeded"`：快照带值（Office）或服务端能直接读文件（文件夹通道）→
      服务端算，返回可执行的 `PlanResponse`。
    - `"none"`：客户端声明「我自己读、值不上传」（WPS）→ 服务端手上没有源值，
      只下发规则，客户端用 `packages/core` 的同一份实现算，返回
      `LocalPlanRuleResponse`。

    分叉条件刻意是**读策略而不是宿主**：服务端不认识宿主，也不该认识。
    任何送 `"none"` 的客户端都走同一条路。

    产出的是 `writeValues`，与系统提示里「不要把算出来的结果塞进
    writeValues」看似冲突，实则不同：那条针对的是**聚合查询结果**（源数据
    一变就过期，应改用 computeValues 在执行时重算）；这里写入的是用户给定
    的分类标签表的查表结果，映射表本身就是固定输入。但它确实仍是一张快照，
    所以计划里明确写出这一点。
    """
    source_sheet_name = str(arguments.get("sourceSheet", ""))
    # 先只解析表名：`"none"` 时没有源值可读，但「表不存在」仍应立刻拦下。
    source_sheet = _sheet(request, source_sheet_name)
    target_sheet_name = str(arguments.get("targetSheet", "")).strip()
    # 未指定目标表时按源表处理；重叠由 build_column_mapping_actions 拦下。
    target_sheet = (
        _sheet(request, target_sheet_name).name
        if target_sheet_name
        else source_sheet.name
    )
    source_range = str(arguments.get("sourceRange", "")).strip().upper().replace("$", "")
    target_range = str(arguments.get("targetRange", "")).strip().upper().replace("$", "")
    mapping = _mapping_from_pairs(arguments.get("mapping"))
    blank_label = str(arguments.get("blankLabel") or "").strip()
    # **不再**因为没给空白标签就直接拒：那是"还没读源值就先问空白怎么办"的根源。
    # 策略未定时，只有真的读到空白格才升级为一次澄清（见下方 BlankPolicyNeededError）。
    blank_policy = _blank_policy_argument(arguments.get("blankPolicy"), blank_label)
    # 多值语义的三组参数都是**规则数据**，服务端只做类型归一与去空，
    # 不在这里判断"该用哪个分隔符、哪个等级更高"——那是规则的语义。
    # 服务端不认识任何具体分隔符字面量，也不认识任何等级名。
    delimiters = _string_list(arguments.get("delimiters"))
    undeclared_separators = _string_list(arguments.get("undeclaredSeparators"))
    value_rank = _rank_entries(arguments.get("valueRank"))
    # 声明了分隔符却没有任何等级次序时，拆分后"取哪一个"没有确定答案。
    # 这是**规则不全**，不是计算能兜的：交给模型回去问清楚，
    # 而不是在这里按 mapping 顺序硬排一个 —— 那个顺序是模型给的，
    # 不代表用户说的优先级。
    #
    # 这条检查必须排在 `planningReadPolicy` 分叉**之前**：两个分支用的是
    # 同一份 core 实现，规则不全时**两侧都算不出来**。若只在 `ifNeeded`
    # 分支拦，`"none"` 分支会把残缺规则原样下发，客户端拿到后按 mapping
    # 顺序硬排 —— 同一个错误在一条通道上被拦、在另一条上被静默执行。
    if delimiters and not value_rank:
        return _clarification_rule_response(
            model_arguments=arguments,
            source_sheet_name=source_sheet.name,
            target_sheet_name=target_sheet,
            source_range=source_range,
            target_range=target_range,
            mapping=mapping,
            blank_label=blank_label,
            blank_policy=blank_policy,
            reason="unresolvedSeparatorOrMapping",
            message=(
                "这条规则声明了分隔符，但没有说明拆分后应当取哪一个，"
                "无法确定多值单元格的结果。请先向用户确认："
                "同一格里的多个值按哪个分隔符拆分、拆分后取哪一个"
                "（或按什么顺序判断哪一个优先）。"
            ),
            unresolved_count=0,
            resolved_by=["specifyOtherDelimiter", "addCategories"],
            delimiters=delimiters,
            undeclared_separators=undeclared_separators,
        )

    assumptions = [
        f"映射规则共 {len(mapping)} 条，全部来自本次需求，未自行补充。",
        _blank_policy_summary(blank_policy),
    ]
    warnings = [
        f"写入的是按当前源数据查表得到的固定值；"
        f"「{source_sheet.name}」{source_range} 之后再改动，"
        f"「{target_sheet}」{target_range} 不会自动跟着变。",
        f"如果「{target_sheet}」{target_range} 已有内容，执行后会被覆盖。",
    ]

    if request.planningReadPolicy == "none":
        # 没有源值，但该拦的仍要拦：范围解析、形状不符、同表重叠三类错误
        # 不需要值就能判定。留到客户端才发现等于把错误推到模型离开循环之后，
        # 那时它已经无法修正参数了。
        rows, columns = validate_column_mapping_shape(
            source_sheet=source_sheet.name,
            target_sheet=target_sheet,
            source_range=source_range,
            target_range=target_range,
            blank_label=blank_label,
        )
        return LocalPlanRuleResponse(
            provider="model",
            rule=ColumnMappingRule(
                sourceSheet=source_sheet.name,
                targetSheet=target_sheet,
                sourceRange=source_range,
                targetRange=target_range,
                # 从校验过的映射表回建，而不是原样转发模型给的数组：
                # 重复键已在 _mapping_from_pairs 里判过，这里输出的是去重结果。
                mapping=[
                    ColumnMappingPair(source=source, target=target)
                    for source, target in mapping.items()
                ],
                blankLabel=blank_label or None,
                blankPolicy=blank_policy,
                delimiters=delimiters,
                valueRank=value_rank,
                undeclaredSeparators=undeclared_separators,
            ),
            summary=(
                f"按 {len(mapping)} 条映射规则，把「{source_sheet.name}」"
                f"{source_range}（{rows}×{columns}）写入「{target_sheet}」"
                f"{target_range}。"
                f"{_blank_policy_summary(blank_policy)}"
                "源值不上传，由本端读取并计算。"
            ),
            assumptions=assumptions,
            # 客户端还未读取源值，也未决定能否安全试算公式；此时不能宣称一定写固定值。
            warnings=warnings[1:],
        )

    source_sheet, source_values = _snapshot_region(
        request, source_sheet_name, str(arguments.get("sourceRange", ""))
    )
    try:
        actions = build_column_mapping_actions(
            source_sheet=source_sheet.name,
            target_sheet=target_sheet,
            source_range=source_range,
            target_range=target_range,
        source_values=source_values,
        mapping=mapping,
        blank_label=blank_label,
        blank_policy=blank_policy,
        delimiters=delimiters,
        value_rank=value_rank,
        undeclared_separators=undeclared_separators,
        max_split_actions=_max_mapping_split_actions(),
    )
    except BlankPolicyNeededError as error:
        # **读到源值之后**才确认"确实有空白"，且策略未定 → 结构化澄清。
        # 四个选项由客户端按 `reason` 推导，服务端不替用户挑默认值。
        return _clarification_rule_response(
            model_arguments=arguments,
            source_sheet_name=source_sheet.name,
            target_sheet_name=target_sheet,
            source_range=source_range,
            target_range=target_range,
            mapping=mapping,
            blank_label=None,
            blank_policy=None,
            blank_cell_count=error.blank_count,
            reason="blankPolicyNeeded",
            message=(
                f"源区域读出来有 {error.blank_count} 个空白格，"
                "但你还没说这些位置的目标格应当怎么处理。请选择一种做法："
                "保留目标原值（这些位置不写入）/ 清空目标内容 / 写入指定值 / "
                "显式写入空字符串。"
            ),
            unresolved_count=0,
            resolved_by=[
                "keepTarget",
                "clearTarget",
                "writeBlankValue",
                "writeEmptyString",
            ],
            delimiters=delimiters,
            undeclared_separators=undeclared_separators,
            value_rank=value_rank,
        )
    except NeedsClarificationError as error:
        # 归因**不确定**（可能是分隔符没声明全，也可能是缺类别）：
        # 不返回计划，把归因与候选动作交给客户端，让用户选。
        # 原文留在服务端这边也不下发 —— 客户端有自己的源值，能就地展示。
        return _clarification_rule_response(
            model_arguments=arguments,
            source_sheet_name=source_sheet.name,
            target_sheet_name=target_sheet,
            source_range=source_range,
            target_range=target_range,
            mapping=mapping,
            blank_label=blank_label,
            reason="unresolvedSeparatorOrMapping",
            message=error.args[0] if error.args else "有单元格无法归类，需要先澄清。",
            unresolved_count=len(error.cells),
            resolved_by=[
                "specifyOtherDelimiter",
                "addCategories",
                "treatWholeCellAsCategory",
            ],
            delimiters=delimiters,
            undeclared_separators=undeclared_separators,
            value_rank=value_rank,
            display_values=True,
        )
    except UnmappedCategoryError as error:
        # 归因**确定**是类别缺口（已按声明的分隔符拆开）。给确定的那一条
        # 补救动作，不把"补充分隔符"这类无关选项塞进去误导用户。
        missing = [_round_trip_display(str(item)) for item in error.values]
        return _clarification_rule_response(
            model_arguments=arguments,
            source_sheet_name=source_sheet.name,
            target_sheet_name=target_sheet,
            source_range=source_range,
            target_range=target_range,
            mapping=mapping,
            blank_label=blank_label,
            reason="matchedDelimiterMissingCategories",
            message=(
                "以下类别不在映射表内，已按规则声明的分隔符拆分后仍未找到对应标签："
                f"{'、'.join(missing)}。请补充这些类别的映射后重试。"
            ),
            unresolved_count=0,
            resolved_by=["addCategories"],
            missing_categories=missing,
            delimiters=delimiters,
            undeclared_separators=undeclared_separators,
            value_rank=value_rank,
            display_values=True,
        )
    plan = AnalysisPlan.model_validate(
        {
            "id": f"map-column-{uuid.uuid4().hex[:10]}",
            "title": f"按映射表填写 {target_sheet}!{target_range}",
            "summary": (
                f"读取「{source_sheet.name}」{source_range} 的 "
                f"{len(source_values)} 行，按 {len(mapping)} 条映射规则"
                f"写入「{target_sheet}」{target_range}。"
                f"{_blank_policy_summary(blank_policy)}"
            ),
            "assumptions": assumptions,
            "warnings": warnings,
            "actions": actions,
        }
    )
    # 该终局工具也产出计划，同样要过规划出口的能力收窄 ——
    # 否则它会成为绕过 `submit_plan` 检查的第二条路径。
    _enforce_action_capabilities(request, plan)
    return PlanResponse(provider="model", plan=plan)


def _find_fields(request: PlanRequest, query: str) -> dict[str, Any]:
    normalized = re.sub(r"[\W_]+", "", query).casefold()
    matches: list[dict[str, Any]] = []
    for sheet in request.workbook.worksheets:
        origin_column, origin_row = _used_origin(sheet.usedRange)
        for index, header in enumerate(sheet.headers):
            header_text = re.sub(r"[\W_]+", "", str(header)).casefold()
            if not header_text or (
                normalized not in header_text and header_text not in normalized
            ):
                continue
            column = _column_name(origin_column + index)
            populated_rows = [
                row_index + origin_row + 1
                for row_index, row in enumerate(sheet.dataRows)
                if index < len(row) and row[index] not in (None, "")
            ]
            data_start = min(populated_rows) if populated_rows else origin_row + 1
            data_end = max(populated_rows) if populated_rows else data_start
            full_data_end = max(origin_row + 1, origin_row + sheet.rowCount - 1)
            matches.append(
                {
                    "sheet": sheet.name,
                    "header": header,
                    "headerCell": f"{column}{origin_row}",
                    "dataRange": f"{column}{origin_row + 1}:{column}{full_data_end}",
                    "observedNonEmptyRange": (
                        f"{column}{data_start}:{column}{data_end}"
                    ),
                    "nonEmptyCountInSnapshot": len(populated_rows),
                    "snapshotTruncated": sheet.truncated,
                }
            )
    return {"query": query, "matches": matches}


def _inline_json_schema(schema: dict[str, Any]) -> dict[str, Any]:
    """Expand Pydantic $defs references for model APIs with stricter validators."""
    definitions = schema.get("$defs", {})

    def expand(node: Any) -> Any:
        if isinstance(node, list):
            return [expand(item) for item in node]
        if not isinstance(node, dict):
            return node

        reference = node.get("$ref")
        if isinstance(reference, str) and reference.startswith("#/$defs/"):
            definition_name = reference.removeprefix("#/$defs/")
            target = definitions.get(definition_name)
            if target is None:
                raise ValueError(f"JSON Schema 引用了未知定义：{reference}")
            resolved = expand(target)
            siblings = {
                key: expand(value)
                for key, value in node.items()
                if key != "$ref"
            }
            return {**resolved, **siblings}

        return {
            key: expand(value)
            for key, value in node.items()
            if key != "$defs"
        }

    return expand(schema)


def _strip_metadata(node: Any) -> Any:
    """Remove Pydantic-generated title/description/default annotations.

    模型需要的是字段结构和取值约束，不是自动生成的字段名标题。剥离后可显著
    降低每轮 Agent 请求携带的 schema token，不改变任何字段或校验规则。

    default 一并剥离：它描述的是**服务端在字段缺失时的行为**，不是模型必须
    满足的条件。剥离后模型能生成的合法计划集合一个都没少 —— maxLength /
    minLength / enum / anyOf 这些真约束全部保留。这条界线要守住：削注解可以，
    削约束不行，后者会让模型看不见真实边界从而反复提交非法计划。

    注意：只能剥离作为 JSON Schema 元数据的标量注解，不能误删 properties
    内部同名的业务字段（例如 AnalysisPlan 自身就有 title 字段）。因此进入
    properties / $defs 等“键即字段名”的容器时不再当作元数据处理。
    """
    if isinstance(node, list):
        return [_strip_metadata(item) for item in node]
    if not isinstance(node, dict):
        return node
    result: dict[str, Any] = {}
    for key, value in node.items():
        if key in ("title", "description", "default") and not isinstance(
            value, (dict, list)
        ):
            # 标量的 title/description/default 是 schema 注解，安全剥离。
            continue
        if key in ("properties", "$defs", "definitions") and isinstance(value, dict):
            # 这些容器的键是字段名，值继续递归但键本身保留。
            result[key] = {
                field: _strip_metadata(field_schema)
                for field, field_schema in value.items()
            }
        else:
            result[key] = _strip_metadata(value)
    return result


def _capability_matrix() -> CapabilityMatrix:
    """从 `config/capabilities.json` 的 `hostCapabilities` 构造可信上界矩阵。

    配置缺失或形态不对时返回**空矩阵**：空矩阵 → 空上界 → 空交集 →
    没有动作可规划。这是保守失败，不是静默放开。
    """
    raw = capabilities().get("hostCapabilities")
    if not isinstance(raw, dict):
        return CapabilityMatrix()
    hosts: dict[str, HostCapabilityMatrix] = {}
    for host_id, entry in raw.items():
        if not isinstance(entry, dict):
            continue
        versions: dict[str, dict[str, ActionCapability]] = {}
        for prefix, actions in (entry.get("versions") or {}).items():
            if not isinstance(actions, dict):
                continue
            versions[str(prefix)] = {
                str(action): ActionCapability.model_validate(capability)
                for action, capability in actions.items()
                if isinstance(capability, dict)
            }
        hosts[str(host_id)] = HostCapabilityMatrix(
            enabled=bool(entry.get("enabled", True)),
            versions=versions,
        )
    return CapabilityMatrix(hosts=hosts)


def _allowed_action_types(request: PlanRequest) -> set[str]:
    """服务端上界 ∩ 客户端自报（方案 §4.3）。

    请求**没有** `hostId` 时返回空集 —— 那表示调用方没走能力协商这条路，
    不是"没声明就放开"。现有的 `/api/plan` 兼容入口与文件夹通道都会命中这一支，
    它们的动作边界由各自的后续校验负责，不在这里被意外放宽。
    """
    if not parse_host_id(request.hostId):
        return set()
    bound = resolve_host_action_bound(
        _capability_matrix(), request.hostId, request.hostVersion
    )
    return intersect_action_capabilities(bound, request.selfReportedActions)


def _enforce_action_capabilities(request: PlanRequest, plan: AnalysisPlan) -> None:
    """规划出口的能力收窄检查。

    这是第二道防线（客户端收窄 → 服务端定上界 → 执行预检兜底）。

    三种情形的处理必须分开，混为一谈等于开一个口子：

    | 请求 | 处理 | 理由 |
    | --- | --- | --- |
    | **未声明** `hostId` | 不检查，保持原行为 | 兼容入口（`/api/plan`、文件夹通道）没走能力协商，动作边界由各通道后续校验负责；在这里顺手拒绝会误伤现有路径 |
    | 声明了宿主标识 | 按下述规则强制 | 调用方明确要求走能力协商 |
    | 声明了但**无法识别** | **拒绝** | 未知宿主的上界是空集，不能因为"这个 id 我不认识"就跳过整道闸 —— 那正是构造请求绕过检查的入口 |

    最后一行的判定用"字段有没有值"而不是"能不能解析出宿主"来分流，
    这两个条件在自报了无法识别标识时结果相反，而后者才是安全的那个。
    """
    declared = bool((request.hostId or "").strip())
    if not declared:
        return
    allowed = _allowed_action_types(request)
    for action in plan.actions:
        action_type = str(getattr(action, "type", ""))
        variant = action.applyTo if action_type == "clearRange" else None
        if is_action_allowed_by_keys(allowed, action_type, variant):
            continue
        key = f"{action_type}:{variant}" if variant else action_type
        raise ValueError(
            f"动作「{key}」不在当前宿主与版本已验收的能力范围内，已拒绝整份计划。"
            "未验收的动作不进入规划，也不会被执行；请改用已开放的动作，"
            "或在完成该动作的实机验收后再启用。"
        )


def _plan_tool_schema() -> dict[str, Any]:
    """Compact submit_plan schema.

    在完整内联 schema 基础上做两处安全裁剪：
    1. 移除 acceptanceCriteria 字段——它由 AnalysisPlan 的
       add_deterministic_acceptance_criteria 验证器自动推断；模型即便仍主动
       提交，Pydantic 的 default_factory 也照常接受，因此不必在工具里向模型
       公布这段最大的子 schema。
    2. 剥离 Pydantic 自动生成的 title 元数据。
    动作(actions)的字段级结构完整保留，模型仍能看到全部动作类型及其字段。
    """
    schema = _inline_json_schema(AnalysisPlan.model_json_schema())
    properties = schema.get("properties")
    if isinstance(properties, dict):
        properties.pop("acceptanceCriteria", None)
    required = schema.get("required")
    if isinstance(required, list) and "acceptanceCriteria" in required:
        required.remove("acceptanceCriteria")
    return _strip_metadata(schema)


def _readonly_tools() -> list[dict[str, Any]]:
    return [
        {
            "type": "function",
            "function": {
                "name": "get_workbook_context",
                "description": "查看工作簿、工作表、表头、活动表和当前选区。",
                "parameters": {"type": "object", "properties": {}},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "find_fields",
                "description": "按业务字段名称查找表头单元格和实际非空数据范围。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "用户提到的实际字段名称。",
                        }
                    },
                    "required": ["query"],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "read_range",
                "description": "从已捕获的工作簿快照读取一个明确区域。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "sheet": {"type": "string"},
                        "range": {"type": "string"},
                    },
                    "required": ["sheet", "range"],
                    "additionalProperties": False,
                },
            },
        },
    ]


def _terminal_tools() -> list[dict[str, Any]]:
    plan_schema = _plan_tool_schema()
    result_context_schema = _strip_metadata(
        _inline_json_schema(ResultContext.model_json_schema())
    )
    return [
        {
            "type": "function",
            "function": {
                "name": "submit_plan",
                "description": (
                    "提交 Excel 操作计划。写入优先级：writeFormulas > "
                    "computeValues（执行时按实时数据求值）> writeValues（仅固定值）。"
                    "splitGroupAggregate 的 splitBy 会自动成为结果列，groupBy 不得"
                    "重复包含 splitBy。无内容的默认列表字段请省略或传空列表，不要"
                    "传 null。"
                ),
                "parameters": {
                    "type": "object",
                    "properties": {"plan": plan_schema},
                    "required": ["plan"],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "map_column_values",
                "description": (
                    "按用户给定的对应关系把一列的值翻译成另一列的标签（分类、"
                    "归类、打标签、编码转名称等）。你只提供映射规则与区域，"
                    "**目标值由确定性代码逐格算出**，不要自己抄写结果，也不要改用 "
                    "writeValues 手填。"
                    "来源格默认按**整格**匹配。只有当同一格确实可能含多个值、"
                    "且用户已经讲清楚怎么处理时，才填 delimiters 与 valueRank："
                    "delimiters 是用户确认过的分隔符写法（按出现位置切分），"
                    "valueRank 说明拆分后取哪一个（rank 大者优先）。"
                    "两者必须**同时**给出，只给 delimiters 表达不了'取哪一个'，"
                    "本条规则会被判为不完整并要求先澄清。"
                    "undeclaredSeparators 用来登记'用户提到过、但还没确认能用来"
                    "拆分'的写法；拆分后仍有未覆盖的类别、或整格看不出分隔方式时，"
                    "本条工具会返回澄清而不是猜测的结果。"
                    "遇到说不清的地方先用 submit_clarification 问，不要自行假设。"
                ),
                "parameters": {
                    "type": "object",
                    "properties": {
                        "sourceSheet": {"type": "string", "description": "来源工作表名。"},
                        "sourceRange": {
                            "type": "string",
                            "description": "来源单列区域，如 D2:D200，不含表头。",
                        },
                        "targetSheet": {
                            "type": "string",
                            "description": "目标工作表名；与来源同表时可省略。",
                        },
                        "targetRange": {
                            "type": "string",
                            "description": "目标区域，行数须与来源一致，如 E2:E200。",
                        },
                        "mapping": {
                            "type": "array",
                            "description": (
                                "用户给出的全部对应关系，不要自行补充。每项把一个"
                                "来源值（整格或拆分后的单项）映射到一个目标标签。"
                            ),
                            "items": {
                                "type": "object",
                                "properties": {
                                    "from": {"type": "string"},
                                    "to": {"type": "string"},
                                },
                                "required": ["from", "to"],
                                "additionalProperties": False,
                            },
                        },
                        "blankLabel": {
                            "type": "string",
                            "description": (
                                "（旧字段，优先用 blankPolicy）来源为空白时写入的标签。"
                                "**不要为了填它而先去问用户**：源列到底有没有空白，"
                                "工具读到源值之后才知道。没有空白时它用不上；"
                                "真有空白且你不知道该写什么，就省略它，工具会返回"
                                "一张结构化的空白处理澄清卡让用户自己选。"
                            ),
                        },
                        "blankPolicy": {
                            "type": "object",
                            "description": (
                                "来源为空白时怎么处理。四种取值："
                                '{"kind":"keep"} = 保留目标原值（那个位置不写入）；'
                                '{"kind":"clear"} = 清空目标内容；'
                                '{"kind":"value","value":"…"} = 写入指定值'
                                "（value 可以是空串，表示显式写空字符串）；"
                                '{"kind":"undecided"} = 还没定。'
                                "**用户没提空白就整个省略这个字段**，不要替他猜；"
                                "源列真有空白且策略未定时，工具会返回澄清卡让他选，"
                                "不会静默写任何默认值。"
                            ),
                            "properties": {
                                "kind": {
                                    "type": "string",
                                    "enum": ["keep", "clear", "value", "undecided"]
                                },
                                "value": {"type": "string"}
                            },
                            "required": ["kind"],
                            "additionalProperties": False,
                        },
                        "delimiters": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": (
                                "用户确认过的分隔符写法，按出现位置取最靠前的那个"
                                "切分。**必须与 valueRank 一起给**。不要自行编造"
                                "分隔符；用户没提过就省略或用 submit_clarification 问。"
                            ),
                        },
                        "valueRank": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "value": {"type": "string"},
                                    "rank": {
                                        "type": "number",
                                        "description": (
                                            "优先级数值，**越大越优先**。"
                                            "用户说更严重/更高级/更优先的那个等级，"
                                            "rank 必须给得更大。"
                                        ),
                                    },
                                },
                                "required": ["value", "rank"],
                                "additionalProperties": False,
                            },
                            "description": (
                                "拆分后取哪一个：rank 越大的越优先。等级名称与"
                                "数值都由用户给出，不要自己发明顺序。"
                                "**方向极易填反，务必按用户的优先级而不是按等级名**："
                                "等级名里若带数字，那个数字的大小与优先级方向**可能相反**"
                                "（名字里的 0 未必是 rank 0），不要拿它当 rank。"
                                "例：用户说「甲比乙严重」，正确写法是 "
                                '[{"value":"甲","rank":1},{"value":"乙","rank":0}]；'
                                "写成反过来就会取到乙。"
                            ),
                        },
                        "undeclaredSeparators": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": (
                                "用户提到过、但尚未确认能否用来拆分的分隔写法。"
                                "登记它们能让'这格到底算不算多值'变成一次明确的"
                                "澄清，而不是由实现去猜。"
                            ),
                        },
                    },
                    # `blankLabel` / `blankPolicy` **都不在必填里**：
                    # 把它们列为必填，会逼模型在读到源值之前先去问用户
                    # "空白怎么办" —— 而多数源列根本没有空白。
                    "required": [
                        "sourceSheet",
                        "sourceRange",
                        "targetRange",
                        "mapping",
                    ],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "submit_clarification",
                "description": "缺少用户决定时提交可填写的澄清卡，不是普通文字回答。",
                "parameters": {
                    "type": "object",
                    "properties": {"clarification": _strip_metadata(_inline_json_schema(IntentClarification.model_json_schema()))},
                    "required": ["clarification"],
                    "additionalProperties": False,
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "submit_answer",
                "description": "提交无需修改 Excel 的直接回答。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "message": {"type": "string"},
                        "resultContext": {
                            "anyOf": [result_context_schema, {"type": "null"}]
                        },
                    },
                    "required": ["message"],
                    "additionalProperties": False,
                },
            },
        },
    ]


# 完整工具集：需要模型自行查数据时使用。
AGENT_TOOLS = [*_readonly_tools(), *_terminal_tools()]
# 数据已由前端确定性算好并随 dataResults 带入时，去掉三个只读工具，
# 让模型直接用现成结果提交回答或计划。减少每轮携带的 token，也避免模型
# 多绕几轮重复读取，从而缩短真实耗时、降低撞上超时中断的概率。
AGENT_TOOLS_WITH_DATA = _terminal_tools()
PLAN_ONLY_TOOLS = [
    tool
    for tool in AGENT_TOOLS_WITH_DATA
    if tool.get("function", {}).get("name") == "submit_plan"
]


@lru_cache(maxsize=1)
def _action_schemas() -> dict[str, dict[str, Any]]:
    items = _plan_tool_schema()["properties"]["actions"]["items"]
    return {variant["properties"]["type"]["const"]: variant for variant in items["oneOf"]}


def _select_plan_actions(request: PlanRequest, arguments: dict[str, Any]) -> list[str]:
    selected = arguments.get("actionTypes")
    available = set(_action_schemas())
    if request.hostId:
        available &= _allowed_action_types(request)
    if (not isinstance(selected, list) or not selected or
            any(not isinstance(item, str) or item not in available for item in selected)):
        raise ValueError("必须选择当前宿主允许的动作类型")
    return list(dict.fromkeys(selected))


def _planning_tools(request: PlanRequest, template: list[dict[str, Any]], selected: list[str]) -> list[dict[str, Any]]:
    """Model selects action types, not business fields. Backend plan validation stays complete."""
    variants = _action_schemas()
    available = sorted(set(variants) & _allowed_action_types(request)) if request.hostId else sorted(variants)
    result = []
    for tool in template:
        if tool["function"]["name"] != "submit_plan":
            result.append(tool)
        elif selected or template is PLAN_ONLY_TOOLS:
            projected = copy.deepcopy(tool)
            projected["function"]["parameters"]["properties"]["plan"]["properties"]["actions"]["items"] = {
                "anyOf": [variants[name] for name in selected] if selected else list(variants.values())
            }
            result.append(projected)
    if template is not PLAN_ONLY_TOOLS and available:
        result.append({"type": "function", "function": {
            "name": "select_plan_actions",
            "description": "仅在需要通用操作计划时选择动作类型，下一轮获得这些动作的完整参数。列映射优先直接用 map_column_values，不要先读源数据或选 writeValues。选择类型不执行、不授权任何写入。",
            "parameters": {"type": "object", "properties": {"actionTypes": {
                "type": "array", "minItems": 1, "maxItems": len(available),
                "items": {"type": "string", "enum": available}
            }}, "required": ["actionTypes"], "additionalProperties": False}
        }})
    return result


def _normalize_submitted_plan(value: Any) -> Any:
    """修正模型计划中不改变语义、且可确定消除的协议冗余。

    splitGroupAggregate 的 splitBy 本身就是输出列。模型偶尔会为了“保留该字段”
    又把它放进 groupBy；这不会增加任何信息，却会让合法计划在模型层来回重试。
    这里只移除与 splitBy 完全同名的 groupBy 项，其他无效配置仍交给 Pydantic
    拒绝，避免本地层猜测业务语义。
    """
    if not isinstance(value, dict):
        return value

    normalized = copy.deepcopy(value)
    # 模型常把“可省略且默认由服务端创建的集合”输出为 null。对所有
    # AnalysisPlan 顶层 default_factory 字段统一按“省略”处理，不绑定具体字段名，
    # 也不改变任何业务语义；必填字段和没有默认工厂的字段仍照常严格校验。
    for field_name, field_info in AnalysisPlan.model_fields.items():
        if (
            normalized.get(field_name) is None
            and field_info.default_factory is not None
        ):
            normalized.pop(field_name, None)

    actions = normalized.get("actions")
    if not isinstance(actions, list):
        return normalized

    removed_fields: list[str] = []
    for action in actions:
        if not isinstance(action, dict) or action.get("type") != "splitGroupAggregate":
            continue
        split_by = action.get("splitBy")
        group_by = action.get("groupBy")
        if not isinstance(split_by, str) or not isinstance(group_by, list):
            continue
        normalized_split = split_by.strip()
        filtered_group_by = [
            field
            for field in group_by
            if not (
                isinstance(field, str)
                and field.strip() == normalized_split
            )
        ]
        if len(filtered_group_by) == len(group_by):
            continue
        action["groupBy"] = filtered_group_by
        if normalized_split and normalized_split not in removed_fields:
            removed_fields.append(normalized_split)

    if not removed_fields:
        return normalized

    warnings = normalized.get("warnings")
    if warnings is None:
        warnings = []
        normalized["warnings"] = warnings
    if isinstance(warnings, list) and len(warnings) < 20:
        fields = "、".join(f"「{field}」" for field in removed_fields)
        warning = (
            f"已从 groupBy 移除与 splitBy 重复的字段 {fields}；"
            "拆分字段会自动保留在结果列中。"
        )
        if warning not in warnings:
            warnings.append(warning)
    return normalized


def _parse_content(content: str) -> AssistantResponse:
    text = content.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("模型没有提交计划或回答")
    payload = json.loads(text[start : end + 1])
    if payload.get("kind") == "answer":
        return AnswerResponse.model_validate({**payload, "provider": "model"})
    if payload.get("kind") == "clarification":
        return IntentClarificationResponse.model_validate({**payload, "provider": "model"})
    if payload.get("kind") == "plan":
        return PlanResponse.model_validate(
            {
                **payload,
                "provider": "model",
                "plan": _normalize_submitted_plan(payload.get("plan")),
            }
        )
    raise ValueError("模型返回了未知响应类型")


def _execute_tool(
    request: PlanRequest, name: str, arguments: dict[str, Any]
) -> AssistantResponse | dict[str, Any]:
    if name == "select_plan_actions":
        return {"selectedActionTypes": _select_plan_actions(request, arguments), "next": "submit_plan"}
    if request.planningReadPolicy == "none" and name in READONLY_TOOLS:
        raise ValueError("本轮源值仅在客户端读取，不允许调用规划只读工具")
    if name == "get_workbook_context":
        return _workbook_context(request)
    if name == "find_fields":
        return _find_fields(request, str(arguments.get("query", "")))
    if name == "read_range":
        return _read_range(
            request,
            str(arguments.get("sheet", "")),
            str(arguments.get("range", "")),
        )
    if name == "map_column_values":
        return _map_column_plan(request, arguments)
    if name == "submit_plan":
        plan = AnalysisPlan.model_validate(
            _normalize_submitted_plan(arguments.get("plan"))
        )
        # 规划出口的能力收窄：未验收的动作在生成计划时就拒绝，
        # 而不是让用户看到一个生成好的计划之后在执行预检被拒（方案 §4.3）。
        _enforce_action_capabilities(request, plan)
        return PlanResponse(provider="model", plan=plan)
    if name == "submit_answer":
        return AnswerResponse(
            provider="model",
            message=str(arguments.get("message", "")),
            resultContext=arguments.get("resultContext"),
        )
    if name == "submit_clarification":
        return IntentClarificationResponse(
            provider="model", turnId=request.turnId,
            clarification=IntentClarification.model_validate(arguments.get("clarification")),
        )
    raise ValueError(f"未知工具：{name}")


READONLY_TOOLS = frozenset(
    {"get_workbook_context", "find_fields", "read_range"}
)


def _tool_cache_key(name: str, arguments: dict[str, Any]) -> str:
    return f"{name}:{json.dumps(arguments, ensure_ascii=False, sort_keys=True)}"


def _tool_step(name: str, arguments: dict[str, Any]) -> str | None:
    """把只读工具调用翻译成用户可读的步骤文案。"""
    if name == "get_workbook_context":
        return "正在查看工作簿结构"
    if name == "find_fields":
        query = str(arguments.get("query", "")).strip()
        return f"正在查找字段「{query}」" if query else "正在查找字段"
    if name == "read_range":
        sheet = str(arguments.get("sheet", "")).strip()
        cell_range = str(arguments.get("range", "")).strip()
        if sheet and cell_range:
            return f"正在读取 {sheet}!{cell_range}"
        return "正在读取数据范围"
    return None


def _safe_progress_text(value: object, limit: int = 1000) -> str:
    """把可公开的确定性摘要限制在进度协议允许的长度内。"""
    text = re.sub(r"\s+", " ", str(value)).strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)] + "…"


def _progress_fact(
    label: str, value: object, *, style: str = "text"
) -> dict[str, str] | None:
    safe_label = _safe_progress_text(label, 40)
    safe_value = _safe_progress_text(value, 500)
    if not safe_label or not safe_value:
        return None
    return {
        "label": safe_label,
        "value": safe_value,
        "style": "code" if style == "code" else "text",
    }


def _tool_argument_facts(
    name: str, arguments: dict[str, Any]
) -> list[dict[str, str]]:
    facts: list[dict[str, str] | None] = []
    if name == "find_fields":
        facts.append(_progress_fact("查询", arguments.get("query", "")))
    elif name == "read_range":
        facts.extend(
            [
                _progress_fact("工作表", arguments.get("sheet", "")),
                _progress_fact("范围", arguments.get("range", ""), style="code"),
            ]
        )
    elif name == "map_column_values":
        mapping = arguments.get("mapping")
        facts.extend(
            [
                _progress_fact(
                    "来源",
                    f"{arguments.get('sourceSheet', '')}!{arguments.get('sourceRange', '')}",
                    style="code",
                ),
                _progress_fact("目标", arguments.get("targetRange", ""), style="code"),
                _progress_fact(
                    "映射规则", len(mapping) if isinstance(mapping, list) else 0
                ),
                _progress_fact("空白写作", arguments.get("blankLabel", "")),
            ]
        )
    elif name == "submit_plan":
        plan = arguments.get("plan")
        if isinstance(plan, dict):
            actions = plan.get("actions")
            action_items = actions if isinstance(actions, list) else []
            facts.append(_progress_fact("动作数", len(action_items)))
    return [fact for fact in facts if fact is not None]


def _readonly_tool_facts(
    name: str,
    arguments: dict[str, Any],
    result: object,
    *,
    cached: bool,
) -> list[dict[str, str]]:
    payload = result if isinstance(result, dict) else {}
    facts: list[dict[str, str] | None] = [
        _progress_fact("读取方式", "复用本轮快照结果" if cached else "读取请求快照（非宿主实时读取）")
    ]
    if name == "get_workbook_context":
        worksheets = payload.get("worksheets")
        worksheet_items = worksheets if isinstance(worksheets, list) else []
        field_count = sum(
            len(item.get("headers", []))
            for item in worksheet_items
            if isinstance(item, dict) and isinstance(item.get("headers"), list)
        )
        facts.extend(
            [
                _progress_fact("工作表数", len(worksheet_items)),
                _progress_fact("字段数", field_count),
            ]
        )
    elif name == "find_fields":
        matches = payload.get("matches")
        match_items = matches if isinstance(matches, list) else []
        facts.extend(
            [
                _progress_fact("查询", arguments.get("query", "")),
                _progress_fact("匹配数", len(match_items)),
            ]
        )
    elif name == "read_range":
        values = payload.get("values")
        rows = values if isinstance(values, list) else []
        column_count = max(
            (len(row) for row in rows if isinstance(row, list)), default=0
        )
        facts.extend(
            [
                _progress_fact(
                    "工作表", payload.get("sheet", arguments.get("sheet", ""))
                ),
                _progress_fact(
                    "范围",
                    payload.get("range", arguments.get("range", "")),
                    style="code",
                ),
                _progress_fact("结果", f"{len(rows)} 行 × {column_count} 列"),
                _progress_fact(
                    "快照",
                    "已截断"
                    if payload.get("snapshotTruncated") is True
                    else "完整返回",
                ),
            ]
        )
    return [fact for fact in facts if fact is not None]


def _plan_progress_facts(plan: AnalysisPlan) -> list[dict[str, str]]:
    action_types = [action.type for action in plan.actions]
    facts: list[dict[str, str] | None] = [
        _progress_fact("动作数", len(action_types)),
        _progress_fact("动作", "、".join(action_types), style="code"),
    ]
    return [fact for fact in facts if fact is not None]


def _readonly_tool_completion(
    name: str,
    arguments: dict[str, Any],
    result: object,
    *,
    cached: bool,
) -> tuple[str, str, dict[str, Any], dict[str, str] | None]:
    """只从工具元数据生成用户可见摘要，绝不回传 read_range.values。"""
    suffix = "（复用本轮已读取结果）" if cached else ""
    payload = result if isinstance(result, dict) else {}
    target: dict[str, str] | None = None
    if name == "get_workbook_context":
        worksheets = payload.get("worksheets")
        worksheet_items = worksheets if isinstance(worksheets, list) else []
        field_count = sum(
            len(item.get("headers", []))
            for item in worksheet_items
            if isinstance(item, dict) and isinstance(item.get("headers"), list)
        )
        truncated_count = sum(
            1
            for item in worksheet_items
            if isinstance(item, dict) and item.get("truncated") is True
        )
        summary = (
            f"已读取 {len(worksheet_items)} 张已授权工作表的结构，共 {field_count} 个字段"
        )
        if truncated_count:
            summary += f"；{truncated_count} 张表的数据快照已截断"
        summary += suffix
        return (
            "已查看工作簿结构",
            summary,
            {
                "kind": "evidence",
                "summary": summary,
                "basis": "工作簿结构工具返回的工作表数、字段数与快照状态",
            },
            None,
        )
    if name == "find_fields":
        matches = payload.get("matches")
        match_items = matches if isinstance(matches, list) else []
        query = _safe_progress_text(arguments.get("query", ""), 120)
        locations = [
            f"{item.get('sheet')}!{item.get('headerCell')}"
            for item in match_items[:5]
            if isinstance(item, dict) and item.get("sheet") and item.get("headerCell")
        ]
        summary = f"字段「{query}」找到 {len(match_items)} 处匹配{suffix}"
        basis = "、".join(locations) if locations else "没有发现匹配位置"
        return (
            "已完成字段定位",
            summary,
            {"kind": "evidence", "summary": summary, "basis": basis},
            None,
        )
    if name == "read_range":
        values = payload.get("values")
        rows = values if isinstance(values, list) else []
        row_count = len(rows)
        column_count = max(
            (len(row) for row in rows if isinstance(row, list)), default=0
        )
        sheet = _safe_progress_text(payload.get("sheet", arguments.get("sheet", "")), 500)
        cell_range = _safe_progress_text(
            payload.get("range", arguments.get("range", "")), 80
        )
        truncated = payload.get("snapshotTruncated") is True
        summary = f"已读取 {sheet}!{cell_range} 的 {row_count} 行 × {column_count} 列"
        if truncated:
            summary += "；结果来自截断快照"
        summary += suffix
        target = {"sheet": sheet, "range": cell_range}
        return (
            "已读取数据范围",
            summary,
            {
                "kind": "evidence",
                "summary": summary,
                "basis": "只展示范围和尺寸，不展示单元格原始内容",
            },
            target,
        )
    summary = f"受控工具 {name} 已完成{suffix}"
    return (
        "工具调用已完成",
        summary,
        {"kind": "evidence", "summary": summary},
        None,
    )


def _append_warning_to_last_user_message(
    messages: list[dict[str, Any]], warning: str
) -> None:
    """把收尾提示追加到最后一条用户消息，兼容文本和多模态内容。"""
    for message in reversed(messages):
        if message.get("role") != "user":
            continue
        content = message.get("content")
        if isinstance(content, str):
            message["content"] = content + warning
            return
        if isinstance(content, list):
            for item in reversed(content):
                if isinstance(item, dict) and item.get("type") == "text":
                    item["text"] = str(item.get("text", "")) + warning
                    return
            content.append({"type": "text", "text": warning})
        return


def _inject_tool_warning(tool_content: str, warning: str) -> str:
    """向 JSON 工具结果附加警告，并保留同一结果中已有的警告。"""
    try:
        result = json.loads(tool_content)
    except (TypeError, json.JSONDecodeError):
        result = {"result": tool_content}
    if not isinstance(result, dict):
        result = {"result": result}

    existing_warning = result.get("_warning")
    if isinstance(existing_warning, str) and existing_warning:
        if warning not in existing_warning:
            result["_warning"] = f"{existing_warning}\n{warning}"
    else:
        result["_warning"] = warning
    return json.dumps(result, ensure_ascii=False)


def _contains_snapshot_truncation(value: Any) -> bool:
    """递归识别工具结果中的快照截断标记。"""
    if isinstance(value, dict):
        if value.get("snapshotTruncated") is True or value.get("truncated") is True:
            return True
        return any(_contains_snapshot_truncation(item) for item in value.values())
    if isinstance(value, list):
        return any(_contains_snapshot_truncation(item) for item in value)
    return False


async def run_excel_agent(
    request: PlanRequest,
    *,
    connection: ModelConnection,
    timeout: float,
    tool_cache: dict[str, Any] | None = None,
    on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    on_checkpoint: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    resume_checkpoint: dict[str, Any] | None = None,
    approved_wall_time_seconds: float | None = None,
    config: LoopConfig | None = None,
) -> AssistantResponse:
    """运行受预算和循环检测保护的 Excel 规划 Agent。"""
    cache = tool_cache
    loop_config = config or _default_loop_config()
    if resume_checkpoint is not None:
        if (
            resume_checkpoint.get("schemaVersion") not in {1, 2}
            or resume_checkpoint.get("kind") != "planning"
        ):
            raise ValueError("规划检查点版本或类型不受支持")
        raw_loop_state = resume_checkpoint.get("loopState")
        if not isinstance(raw_loop_state, dict):
            raise ValueError("规划检查点缺少循环状态")
        state = LoopState.from_checkpoint(raw_loop_state)
    else:
        state = LoopState()
    elapsed_before_run = state.elapsed_time
    start_time = time.monotonic()

    async def emit(
        step_id: str,
        title: str,
        *,
        status: str = "progress",
        summary: str | None = None,
        detail: str | None = None,
        trace: dict[str, Any] | None = None,
        target: dict[str, str] | None = None,
        tool_name: str | None = None,
        facts: list[dict[str, str]] | None = None,
    ) -> None:
        if on_event is None:
            return
        event: dict[str, Any] = {
            "phase": "planning",
            "stepId": step_id,
            "status": status,
            "title": title,
        }
        if summary is not None:
            event["summary"] = _safe_progress_text(summary)
        if detail is not None:
            event["detail"] = _safe_progress_text(detail, 2000)
        if trace is not None:
            event["trace"] = trace
        if target is not None:
            event["target"] = target
        if tool_name is not None:
            event["toolName"] = _safe_progress_text(tool_name, 100)
        if facts:
            event["facts"] = facts[:12]
        await on_event(event)
    user_text = (
        f"用户需求：\n{request.prompt}\n\n"
        f"初始工作簿上下文：\n"
        f"{json.dumps(_workbook_context(request), ensure_ascii=False)}\n\n"
        f"上一轮结构化结果：\n"
        f"{request.lastResult.model_dump_json() if request.lastResult and not (request.planningReadPolicy == 'none' and not request.dataResults) else 'null'}\n\n"
        f"本地数据工具结果：\n"
        f"{json.dumps([result.model_dump() for result in request.dataResults], ensure_ascii=False)}\n\n"
        + (
            "本轮只发送表结构，源值留在客户端。列映射直接用 map_column_values 提交规则，"
            "不得调用工作簿只读工具。"
            "由客户端读取和计算；未知实际来源值与空白由本地工具发现后澄清，不要预先猜测。"
            "通用动作先 select_plan_actions 获取所需参数，再 submit_plan。信息不足用 submit_clarification。"
            if request.planningReadPolicy == "none"
            else "请按需调用工具，不要猜测数据地址。"
        )
    )
    user_content: str | list[dict[str, Any]] = user_text
    if request.images:
        user_content = [{"type": "text", "text": user_text}]
        user_content.extend(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{image.mediaType};base64,{image.data}"
                },
            }
            for image in request.images
        )

    if resume_checkpoint is not None:
        raw_messages = resume_checkpoint.get("messages")
        if not isinstance(raw_messages, list) or not all(
            isinstance(item, dict) for item in raw_messages
        ):
            raise ValueError("规划检查点缺少完整消息现场")
        messages = [dict(item) for item in raw_messages]
        if cache is not None and isinstance(resume_checkpoint.get("toolCache"), dict):
            cache.clear()
            cache.update(dict(resume_checkpoint["toolCache"]))
    else:
        messages = [
            {"role": "system", "content": AGENT_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": user_content,
            },
        ]

    await emit("planning.understand", "正在理解需求", status="started")
    # 前端已带回确定性查询结果，或用户已经从结构化选项确认了最终口径时，
    # 去掉只读工具，直接进入提交。运行时聚合动作会在执行阶段读取完整数据。
    initial_tools = (
        AGENT_TOOLS_WITH_DATA
        if request.dataResults or request.planningReadPolicy == "none"
        else AGENT_TOOLS
    )
    tool_mode = (
        str(resume_checkpoint.get("toolMode"))
        if resume_checkpoint is not None
        else ("terminal" if initial_tools is AGENT_TOOLS_WITH_DATA else "full")
    )
    tools = {
        "full": AGENT_TOOLS,
        "terminal": AGENT_TOOLS_WITH_DATA,
        "plan_only": PLAN_ONLY_TOOLS,
    }.get(tool_mode)
    if tools is None:
        raise ValueError("规划检查点包含未知工具集合")
    repairing_submission = bool(
        resume_checkpoint.get("repairingSubmission")
        if resume_checkpoint is not None
        else False
    )
    selected_action_types = (
        _select_plan_actions(request, {"actionTypes": resume_checkpoint["selectedActionTypes"]})
        if resume_checkpoint and resume_checkpoint.get("selectedActionTypes") else []
    )
    next_turn_index = max(
        int(resume_checkpoint.get("nextTurnIndex", 0) or 0)
        if resume_checkpoint is not None
        else 0,
        0,
    )
    pending_assistant_message = (
        dict(resume_checkpoint["pendingAssistantMessage"])
        if resume_checkpoint is not None
        and isinstance(resume_checkpoint.get("pendingAssistantMessage"), dict)
        else None
    )
    next_tool_index = max(
        int(resume_checkpoint.get("nextToolIndex", 0) or 0)
        if resume_checkpoint is not None
        else 0,
        0,
    )
    restored_tool_batch_state = (
        dict(resume_checkpoint["toolBatchState"])
        if resume_checkpoint is not None
        and isinstance(resume_checkpoint.get("toolBatchState"), dict)
        else {}
    )

    def update_elapsed() -> None:
        state.elapsed_time = elapsed_before_run + (time.monotonic() - start_time)

    def current_tool_mode() -> str:
        if tools is PLAN_ONLY_TOOLS:
            return "plan_only"
        if tools is AGENT_TOOLS_WITH_DATA:
            return "terminal"
        return "full"

    async def save_boundary(
        stage: str,
        *,
        upcoming_turn: int,
        pending_message: dict[str, Any] | None = None,
        upcoming_tool: int = 0,
        tool_batch_state: dict[str, bool] | None = None,
    ) -> dict[str, Any]:
        update_elapsed()
        checkpoint = {
            "schemaVersion": 2,
            "kind": "planning",
            "stage": stage,
            "messages": messages,
            "toolCache": dict(cache or {}),
            "loopState": state.to_checkpoint(),
            "toolMode": current_tool_mode(),
            "repairingSubmission": repairing_submission,
            "selectedActionTypes": selected_action_types,
            "nextTurnIndex": upcoming_turn,
            "pendingAssistantMessage": pending_message,
            "nextToolIndex": upcoming_tool,
            "toolBatchState": dict(tool_batch_state or {}),
        }
        if on_checkpoint is not None:
            await on_checkpoint(checkpoint)
        return checkpoint

    async def pause_if_needed(
        stage: str,
        *,
        upcoming_turn: int,
        pending_message: dict[str, Any] | None = None,
        upcoming_tool: int = 0,
        tool_batch_state: dict[str, bool] | None = None,
    ) -> None:
        # Every resumable boundary is committed, not only the boundary that
        # happens to cross the soft time budget.  A process can exit before the
        # reminder is due, and recovery still needs the latest complete model
        # response or tool batch instead of falling back to a fresh run.
        checkpoint = await save_boundary(
            stage,
            upcoming_turn=upcoming_turn,
            pending_message=pending_message,
            upcoming_tool=upcoming_tool,
            tool_batch_state=tool_batch_state,
        )
        if approved_wall_time_seconds is None:
            return
        update_elapsed()
        if state.elapsed_time < approved_wall_time_seconds:
            return
        await emit(
            "planning.budget_pause",
            "已安全保存当前进度",
            status="warning",
            summary="本轮处理时间已用完；继续前不会启动新的模型或工具调用。",
        )
        raise AgentBudgetPause(checkpoint, approved_wall_time_seconds)

    async with OpenAICompatibleClient(connection, timeout=timeout) as client:
        for turn_index in range(next_turn_index, loop_config.max_turns):
            update_elapsed()
            state.turn_index = turn_index
            stop_flag, stop_reason = should_stop(state, loop_config)
            if stop_flag:
                assert stop_reason is not None
                logger.warning(
                    "Agent loop stopped: reason=%s, turns=%s, tokens=%s, "
                    "errors=%s, loop=%s",
                    stop_reason,
                    state.turn_index,
                    state.total_tokens,
                    state.consecutive_errors,
                    state.detected_path_loop,
                )
                await emit(
                    "planning.guard",
                    "规划已安全停止",
                    status="warning",
                    summary="达到规划循环保护条件，已停止继续尝试。",
                    trace={
                        "kind": "correction",
                        "summary": "达到规划循环保护条件，已停止继续尝试。",
                        "basis": _safe_progress_text(stop_reason, 1000),
                    },
                )
                return safe_degradation(state, stop_reason)

            if turn_index == loop_config.warning_turn:
                warning_prompt = get_warning_prompt(state)
                if warning_prompt:
                    _append_warning_to_last_user_message(
                        messages, warning_prompt
                    )

            if turn_index == 0 and resume_checkpoint is None:
                authorized_count = len(request.workbook.worksheets)
                understanding_summary = (
                    f"已整理用户目标，并将规划范围限制在 {authorized_count} 张已授权工作表内。"
                )
                await emit(
                    "planning.understand",
                    "已理解需求与授权范围",
                    status="completed",
                    summary=understanding_summary,
                    trace={
                        "kind": "understanding",
                        "summary": understanding_summary,
                        "basis": "用户当前请求与已授权工作表元数据",
                    },
                )
            model_step_id = f"planning.model.{turn_index + 1}"
            resuming_tool_batch = False
            if pending_assistant_message is not None:
                message = pending_assistant_message
                pending_assistant_message = None
                resuming_tool_batch = next_tool_index > 0
                await emit(
                    model_step_id,
                    "正在继续已保存的工具步骤",
                    status="completed",
                    summary="已复用暂停前完整返回的模型响应。",
                )
            else:
                await pause_if_needed(
                    "before_model", upcoming_turn=turn_index
                )
                await emit(
                    model_step_id,
                    "正在生成操作方案"
                    if turn_index == 0
                    else f"正在修正操作方案（第 {turn_index + 1} 轮）",
                    status="started",
                )
                payload, usage = await client.chat_completions(
                    messages=messages,
                    tools=_planning_tools(request, tools, selected_action_types),
                    tool_choice="auto",
                    max_tokens=(
                        REPAIR_MAX_OUTPUT_TOKENS
                        if repairing_submission
                        else MAX_OUTPUT_TOKENS
                    ),
                )
                state.total_tokens += usage.get("total_tokens", 0)
                choice = payload["choices"][0]
                message = dict(choice["message"])
                finish_reason = choice.get("finish_reason")
                logger.info("Planning response: round=%s mode=%s finish=%s input=%s output=%s total=%s reasoning=%s output_cap=%s",
                            turn_index + 1, current_tool_mode(),
                            finish_reason if finish_reason in {"stop", "tool_calls", "length", "content_filter"} else "other",
                            usage.get("prompt_tokens", 0), usage.get("completion_tokens", 0), usage.get("total_tokens", 0),
                            usage.get("reasoning_tokens", "unavailable"),
                            (REPAIR_MAX_OUTPUT_TOKENS if repairing_submission else MAX_OUTPUT_TOKENS) or "provider_default")
                if finish_reason in {"length", "content_filter"}:
                    # Do not execute even syntactically valid calls from an incomplete response.
                    message = {"role": "assistant", "content": "", "_incomplete": finish_reason}
            assistant_message = dict(message)
            assistant_message.setdefault("role", "assistant")
            await pause_if_needed(
                "model_response_saved",
                upcoming_turn=turn_index,
                pending_message=assistant_message,
                upcoming_tool=next_tool_index if resuming_tool_batch else 0,
                tool_batch_state=(
                    restored_tool_batch_state if resuming_tool_batch else None
                ),
            )
            tool_calls = assistant_message.get("tool_calls") or []
            if assistant_message.get("_incomplete") == "length":
                # A truncated response is not a schema error. Retrying it with a
                # smaller repair allowance cannot fix the provider/context limit.
                # This also handles an incomplete response restored from a checkpoint.
                update_loop_state(state, None, None, error_occurred=False)
                await emit(model_step_id, "模型响应被截断", status="failed",
                           summary="未执行截断响应中的任何工具；请检查模型或网关输出与上下文额度。")
                return safe_degradation(state, "output_length")
            if not tool_calls:
                content = assistant_message.get("content") or ""
                try:
                    if assistant_message.get("_incomplete"):
                        raise ValueError("模型输出达到长度限制" if assistant_message["_incomplete"] == "length" else "模型输出被上游拦截")
                    parsed = _parse_content(content)
                    if isinstance(parsed, PlanResponse):
                        _enforce_action_capabilities(request, parsed.plan)
                        action_types = "、".join(
                            action.type for action in parsed.plan.actions
                        )
                        await emit(
                            model_step_id,
                            "操作计划已生成",
                            status="completed",
                            summary=parsed.plan.summary,
                            trace={
                                "kind": "decision",
                                "summary": parsed.plan.summary,
                                "basis": f"共 {len(parsed.plan.actions)} 步：{action_types}",
                            },
                        )
                    elif isinstance(parsed, IntentClarificationResponse):
                        await emit(
                            model_step_id, "等待补充信息", status="completed",
                            summary=parsed.clarification.question,
                        )
                    else:
                        await emit(
                            model_step_id,
                            "回答已生成",
                            status="completed",
                            summary="已根据受控上下文生成回答。",
                        )
                    return parsed
                except (ValueError, ValidationError, json.JSONDecodeError) as error:
                    repairing_submission = True
                    if tools is AGENT_TOOLS:
                        tools = AGENT_TOOLS_WITH_DATA
                    state.last_error_message = str(error)[:200]
                    update_loop_state(
                        state,
                        None,
                        None,
                        error_occurred=True,
                    )
                    # 无效自由文本可能接近输出上限；错误原因已单独反馈，历史中只
                    # 保留最小占位，避免修复轮再次携带整段无效输出。
                    messages.append(
                        {
                            "role": "assistant",
                            "content": "未提交有效的结构化计划或回答。",
                        }
                    )
                    messages.append(
                        {
                            "role": "user",
                            "content": (
                                f"你的提交没有通过结构校验：{error}。"
                                "请检查工具返回内容，并使用 submit_plan、"
                                "submit_clarification 或 submit_answer 重新提交。"
                            ),
                        }
                    )
                    correction_summary = "模型提交内容未通过结构校验，正在要求重新提交。"
                    await emit(
                        model_step_id,
                        "提交格式未通过，正在修正",
                        status="warning",
                        summary=correction_summary,
                        trace={
                            "kind": "correction",
                            "summary": correction_summary,
                            "basis": _safe_progress_text(error, 1000),
                        },
                    )
                    continue

            if not resuming_tool_batch and (
                not messages or messages[-1] != assistant_message
            ):
                messages.append(assistant_message)
            requested_tool_names = [
                str((tool_call.get("function") or {}).get("name", "")).strip()
                for tool_call in tool_calls
            ]
            requested_tool_names = [
                name for name in requested_tool_names if name
            ]
            await emit(
                model_step_id,
                "已确定下一步受控操作",
                status="completed",
                summary=(
                    f"模型请求执行 {len(tool_calls)} 个白名单工具调用"
                    + (
                        f"：{'、'.join(requested_tool_names)}。"
                        if requested_tool_names
                        else "。"
                    )
                ),
                facts=[
                    fact
                    for fact in [
                        _progress_fact(
                            "工具",
                            "、".join(requested_tool_names),
                            style="code",
                        )
                    ]
                    if fact is not None
                ],
            )
            successful_readonly_batch = bool(
                restored_tool_batch_state.get("successfulReadonly")
                if resuming_tool_batch
                else False
            )
            readonly_batch_had_error = bool(
                restored_tool_batch_state.get("readonlyHadError")
                if resuming_tool_batch
                else False
            )
            tool_start_index = next_tool_index if resuming_tool_batch else 0
            if tool_start_index > len(tool_calls):
                raise ValueError("规划检查点的工具序号超出完整模型响应范围")
            for tool_index, tool_call in enumerate(tool_calls[tool_start_index:], start=tool_start_index):
                function = tool_call.get("function") or {}
                name = str(function.get("name", ""))
                arguments: dict[str, Any] = {}
                tool_step_id = (
                    f"planning.tool.{turn_index + 1}.{tool_index + 1}"
                )
                try:
                    parsed_arguments = json.loads(
                        function.get("arguments") or "{}"
                    )
                    if not isinstance(parsed_arguments, dict):
                        raise ValueError("工具参数必须是 JSON 对象")
                    arguments = parsed_arguments
                    step_text = _tool_step(name, arguments)
                    if step_text is None:
                        if name == "map_column_values":
                            step_text = "正在按映射表计算目标列"
                        elif name == "submit_plan":
                            step_text = "正在校验操作计划"
                        elif name == "submit_answer":
                            step_text = "正在校验回答"
                        elif name == "submit_clarification":
                            step_text = "正在准备补充信息表单"
                        else:
                            step_text = "正在调用受控工具"
                    await emit(
                        tool_step_id,
                        step_text,
                        status="started",
                        tool_name=name or None,
                        facts=_tool_argument_facts(name, arguments),
                    )
                    cache_key: str | None = None
                    if cache is not None and name in READONLY_TOOLS:
                        cache_key = _tool_cache_key(name, arguments)
                    if cache_key is not None and cache_key in cache:
                        cached_value = cache[cache_key]
                        if isinstance(cached_value, str):
                            tool_content = cached_value
                            try:
                                truncation_value = json.loads(cached_value)
                            except json.JSONDecodeError:
                                truncation_value = cached_value
                        else:
                            truncation_value = cached_value
                            tool_content = json.dumps(
                                cached_value, ensure_ascii=False
                            )
                        summary_value = truncation_value
                        used_cache = True
                        if _contains_snapshot_truncation(truncation_value):
                            state.encountered_truncation = True
                        conflict_warning = detect_action_conflict(
                            state,
                            cache_key,
                            loop_config.max_cache_hits_per_query,
                        )
                        if conflict_warning:
                            tool_content = _inject_tool_warning(
                                tool_content, conflict_warning
                            )
                    else:
                        result = _execute_tool(request, name, arguments)
                        if name == "select_plan_actions" and isinstance(result, dict):
                            selected_action_types = list(result["selectedActionTypes"])
                        if isinstance(
                            result,
                            (PlanResponse, AnswerResponse, LocalPlanRuleResponse, IntentClarificationResponse),
                        ):
                            if isinstance(result, LocalPlanRuleResponse):
                                await emit(
                                    tool_step_id,
                                    "映射规则已通过结构校验",
                                    status="completed",
                                    summary=result.summary,
                                    tool_name=name,
                                    facts=[
                                        fact
                                        for fact in [
                                            _progress_fact(
                                                "计算位置", "由本端读取源值并计算"
                                            )
                                        ]
                                        if fact is not None
                                    ],
                                    trace={
                                        "kind": "decision",
                                        "summary": result.summary,
                                        "basis": (
                                            f"共 {len(result.rule.mapping)} 条映射规则，"
                                            "源值不上传"
                                        ),
                                    },
                                )
                            elif isinstance(result, PlanResponse):
                                action_types = "、".join(
                                    action.type for action in result.plan.actions
                                )
                                await emit(
                                    tool_step_id,
                                    "操作计划已通过结构校验",
                                    status="completed",
                                    summary=result.plan.summary,
                                    tool_name=name,
                                    facts=_plan_progress_facts(result.plan),
                                    trace={
                                        "kind": "decision",
                                        "summary": result.plan.summary,
                                        "basis": (
                                            f"共 {len(result.plan.actions)} 步：{action_types}"
                                        ),
                                    },
                                )
                            elif isinstance(result, IntentClarificationResponse):
                                await emit(
                                    tool_step_id, "等待补充信息", status="completed",
                                    summary=result.clarification.question, tool_name=name,
                                )
                            else:
                                await emit(
                                    tool_step_id,
                                    "回答已通过结构校验",
                                    status="completed",
                                    summary="已根据受控上下文生成回答。",
                                    tool_name=name,
                                    facts=[
                                        fact
                                        for fact in [
                                            _progress_fact("结果", "回答协议校验通过")
                                        ]
                                        if fact is not None
                                    ],
                                )
                            return result
                        summary_value = result
                        used_cache = False
                        if _contains_snapshot_truncation(result):
                            state.encountered_truncation = True
                        tool_content = json.dumps(result, ensure_ascii=False)
                        if cache_key is not None:
                            cache[cache_key] = tool_content

                    update_loop_state(
                        state,
                        name,
                        cache_key,
                        error_occurred=False,
                    )
                    if name in READONLY_TOOLS:
                        successful_readonly_batch = True
                    path_warning = detect_path_loop(
                        state, loop_config.max_tool_path_loop_length
                    )
                    if path_warning:
                        tool_content = _inject_tool_warning(
                            tool_content, path_warning
                        )
                        state.last_warning = path_warning
                    if name in READONLY_TOOLS:
                        (
                            completed_title,
                            completed_summary,
                            completed_trace,
                            completed_target,
                        ) = _readonly_tool_completion(
                            name,
                            arguments,
                            summary_value,
                            cached=used_cache,
                        )
                        await emit(
                            tool_step_id,
                            completed_title,
                            status="completed",
                            summary=completed_summary,
                            trace=completed_trace,
                            target=completed_target,
                            tool_name=name,
                            facts=_readonly_tool_facts(
                                name,
                                arguments,
                                summary_value,
                                cached=used_cache,
                            ),
                        )
                except (ValueError, ValidationError, json.JSONDecodeError) as error:
                    if name in READONLY_TOOLS:
                        readonly_batch_had_error = True
                    if name == "submit_plan":
                        failed_actions = (arguments.get("plan") or {}).get("actions", []) if isinstance(arguments.get("plan"), dict) else []
                        candidates = [action.get("type") for action in failed_actions if isinstance(action, dict)]
                        try:
                            selected_action_types = _select_plan_actions(request, {"actionTypes": candidates})
                        except ValueError:
                            pass
                        # 失败的工具参数可能接近输出上限。保留完整原文会让下一轮
                        # 再次携带数千 token；工具错误已经包含校验原因，因此把历史
                        # 调用压缩为合法的最小占位，并只暴露 submit_plan 供定向修复。
                        function["arguments"] = json.dumps(
                            {"plan": None}, ensure_ascii=False
                        )
                        tools = PLAN_ONLY_TOOLS
                        repairing_submission = True
                        failed_plan = arguments.get("plan")
                        state.last_failed_plan = (
                            dict(failed_plan)
                            if isinstance(failed_plan, dict)
                            else None
                        )
                    state.last_error_message = str(error)[:200]
                    update_loop_state(
                        state,
                        name,
                        None,
                        error_occurred=True,
                    )
                    tool_content = json.dumps(
                        {"error": str(error)}, ensure_ascii=False
                    )
                    correction_summary = (
                        "操作计划未通过结构校验，正在修正。"
                        if name == "submit_plan"
                        else f"受控工具 {name or '调用'} 未成功，正在调整下一步。"
                    )
                    await emit(
                        tool_step_id,
                        "检查未通过，正在修正",
                        status="warning",
                        summary=correction_summary,
                        tool_name=name or None,
                        facts=[
                            fact
                            for fact in [
                                _progress_fact("结果", "未通过结构校验")
                            ]
                            if fact is not None
                        ],
                        trace={
                            "kind": "correction",
                            "summary": correction_summary,
                            "basis": _safe_progress_text(error, 1000),
                        },
                    )
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_call.get("id", ""),
                        "name": name,
                        "content": tool_content,
                    }
                )
                await pause_if_needed(
                    "tool_completed",
                    upcoming_turn=turn_index,
                    pending_message=assistant_message,
                    upcoming_tool=tool_index + 1,
                    tool_batch_state={
                        "successfulReadonly": successful_readonly_batch,
                        "readonlyHadError": readonly_batch_had_error,
                    },
                )
            if (
                successful_readonly_batch
                and not readonly_batch_had_error
                and tools is AGENT_TOOLS
            ):
                tools = AGENT_TOOLS_WITH_DATA
                _append_warning_to_last_user_message(
                    messages,
                    "工作簿只读探索已完成。下一轮必须基于现有工作簿上下文和"
                    "工具结果直接调用 submit_plan、submit_clarification 或 submit_answer；不要继续读取数据。",
                )
            await pause_if_needed(
                "tool_batch_completed", upcoming_turn=turn_index + 1
            )
            next_tool_index = 0
            restored_tool_batch_state = {}
    state.turn_index = loop_config.max_turns
    update_elapsed()
    logger.warning(
        "Agent loop exhausted: turns=%s, tokens=%s, errors=%s, loop=%s",
        state.turn_index,
        state.total_tokens,
        state.consecutive_errors,
        state.detected_path_loop,
    )
    await emit(
        "planning.guard",
        "规划已安全停止",
        status="warning",
        summary="已达到最大规划轮次，停止继续尝试。",
        trace={
            "kind": "correction",
            "summary": "已达到最大规划轮次，停止继续尝试。",
        },
    )
    return safe_degradation(state, "max_turns")
