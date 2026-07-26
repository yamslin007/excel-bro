from __future__ import annotations

from pathlib import Path


PRODUCTION_ROOTS = (
    Path("server/app"),
    Path("apps/excel-addin/src"),
)
EXCLUDED_FILES = {"demo.ts"}
BUSINESS_TERMS = (
    "影院",
    "影片",
    "电影",
    "影厅",
    "销售额",
    "客户",
    "学生",
    "姓名",
    "得分",
)


def test_production_code_has_no_domain_specific_examples() -> None:
    violations: list[str] = []
    for root in PRODUCTION_ROOTS:
        for path in root.rglob("*"):
            if (
                not path.is_file()
                or path.name in EXCLUDED_FILES
                or ".test." in path.name
                or path.suffix not in {".py", ".ts", ".tsx"}
            ):
                continue
            text = path.read_text(encoding="utf-8")
            for term in BUSINESS_TERMS:
                if term in text:
                    violations.append(f"{path}:{term}")

    assert not violations, (
        "生产代码不应包含行业字段或行业示例；请把它们放进测试或演示数据："
        + "、".join(violations)
    )
