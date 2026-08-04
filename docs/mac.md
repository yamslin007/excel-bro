# macOS 打包指南

记录 Excel Bro 的 macOS 安装包（.pkg）构建流程、CI 配置，以及排障过程中踩过的坑。

## 概览

macOS 安装包通过 GitHub Actions 工作流 `.github/workflows/build-mac.yml` 构建，核心逻辑在 `packaging/build_mac.sh`。产物为 `dist/Excel-Bro-Setup-<version>-<arch>.pkg`。

仅构建 Apple Silicon（arm64）：

| 架构 | Runner | 说明 |
| --- | --- | --- |
| `arm64` | `macos-14` | Apple Silicon |

> 已移除 Intel（x86_64 / macos-13）构建：Intel Mac 用户占比极低，且 `macos-13` runner 排队严重，长时间卡在 `Waiting for a runner`。如后续确需支持 Intel，再把 x86_64 矩阵条目加回。

触发方式：

- `workflow_dispatch`（手动触发）
- 推送 `v*` tag（并自动附加到 Release）

> 注意：**推送到 main 不会自动触发构建**，需手动 `gh workflow run build-mac.yml --ref main` 或打 tag。

## 构建流程（build_mac.sh）

1. 校验 `apps/excel-addin/manifest.xml`（版本号、IconUrl、HighResolutionIconUrl）。
2. `npm ci` + `npm run build:addin` 构建前端加载项。
3. `generate_certificate.py` 生成本地 localhost 证书。
4. PyInstaller 以 `--onedir` 打包 `packaging/runtime.py` 及 `server/`、`config/`。
5. 组装安装目录 `Applications/Excel Bro/`（后端、frontend、manifest、证书、卸载脚本）。
6. `pkgbuild` + `productbuild` 生成最终 `.pkg`。

## 本地构建

```bash
python -m pip install --upgrade pip
pip install -r server/requirements.txt pyinstaller cryptography
bash packaging/build_mac.sh
# 产物：dist/Excel-Bro-Setup-<version>-<arch>.pkg
```

## 常用运维命令

```bash
# 手动触发构建
gh workflow run build-mac.yml --ref main

# 查看最近运行
gh run list --workflow=build-mac.yml --limit 5

# 查看某次运行的 job 状态
gh run view <run-id> --json jobs -q '.jobs[] | "\(.name)\t\(.status)\t\(.conclusion)\t\(.databaseId)"'

# 拉取单个 job 日志（run 未整体结束时用 API 按 job 拉）
gh api repos/<owner>/<repo>/actions/jobs/<job-id>/logs --allow-escape-sequences \
  | sed 's/\x1b\[[0-9;]*m//g' | tail -80

# 取消无用的排队/运行
gh run cancel <run-id> ...
```

## 排障记录

### 1. manifest 命名空间导致校验崩溃

**现象**：`Build pkg` 阶段两个 job 都失败，exit code 1。

**根因**：`manifest.xml` 使用了默认 XML 命名空间（`xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"`）。旧校验脚本用无命名空间的 `findtext("Version")` 查找，返回 `None`，随后 `None.split()` 崩溃。

**修复**：改为命名空间感知的查找（`findtext("office:Version", namespaces=...)`）。见提交 `8702002`。

### 2. 缺少 cryptography 依赖

**现象**：manifest 校验、npm/vite 构建都通过后，卡在生成证书这一步：

```
File ".../packaging/generate_certificate.py", line 8, in <module>
    from cryptography import x509
ModuleNotFoundError: No module named 'cryptography'
```

**根因**：`generate_certificate.py` 依赖 `cryptography`，但它既不在 `server/requirements.txt`，CI 也未单独安装。

**修复**：在 CI 的 pip install 行与 `pyinstaller` 并列加上 `cryptography`。见提交 `6a2ecef`。

### 3. macOS runner 排队

**现象**：日志停在 `Waiting for a runner to pick up this job...`，`Requested labels: macos-13`。

**说明**：这是 GitHub 在等空闲的 macOS runner，**不是代码错误**。Intel 的 `macos-13` 尤其紧俏，等待几分钟到十几分钟属正常。

## 其他备注

- 日志里的 `Node.js 20 is deprecated` 只是弃用告警，不影响构建。
- 重新触发前，建议先取消跑在旧提交上的排队运行，避免占用 runner 且产生误导性的失败结果。
