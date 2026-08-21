# 文件夹 + 工作簿 + 工作表全量刷新方案

## 需求
用户改了文件名、工作簿名、工作表名后，点一个刷新按钮就能拿到最新状态。

## 涉及的三个层级

### 1. 文件夹层级（文件列表）
- **位置**：后端 session 缓存
- **刷新逻辑**：重新扫描文件夹，更新 `folderCatalog.files`

### 2. 工作簿层级（工作簿名）
- **位置**：`workbook.name`（从 Office.context.document.url 解析）
- **刷新逻辑**：重新调用 `captureWorkbook` 或 `captureWorkbookStructure`

### 3. 工作表层级（表名列表）
- **位置**：`workbook.worksheets`
- **刷新逻辑**：随 `captureWorkbook` 自动更新

---

## 实施方案

### 方案A：两个刷新按钮（推荐）

#### 1. 文件夹刷新按钮
**位置**：文件夹选择区域（"更换文件夹"按钮旁边）

**功能**：
- 重新扫描文件夹
- 更新文件列表（文件重命名会反映）
- 保持 sessionId 不变
- 不影响当前打开的工作簿

**后端**：增加 `refresh_folder(session_id)` 函数
**前端**：增加 `refreshCurrentFolder()` 函数

---

#### 2. 工作簿刷新按钮
**位置**：工作簿/工作表选择区域（"自动"/"手动"/"文件夹"切换按钮附近）

**功能**：
- 重新捕获当前工作簿结构
- 更新工作簿名
- 更新工作表列表（表重命名、新增、删除会反映）
- 刷新选中的工作表（取交集，已删除的表会自动去掉）

**前端逻辑**：
```typescript
async function refreshWorkbook() {
  if (busy) return;
  setStatus("loading");
  
  try {
    // 重新捕获工作簿
    const refreshed = await captureWorkbook();
    
    // 更新 workbook state
    setWorkbook(refreshed);
    
    // 更新工作表选择（保留仍存在的表）
    if (sourceMode === "workbook") {
      applyWorkbookSnapshotSelection(refreshed);
    }
    
    // 如果有文件夹 session，也刷新一下（可选）
    if (folderCatalog) {
      const refreshedFolder = await refreshFolder(folderCatalog.sessionId);
      applyFolderCatalog(refreshedFolder);
    }
  } catch (error) {
    alert(`刷新失败：${error instanceof Error ? error.message : String(error)}`);
  } finally {
    setStatus("idle");
  }
}
```

**UI 位置**：
```tsx
{/* 在工作表选择区域增加刷新按钮 */}
<div className="scope-mode-picker">
  <button 
    className={sourceMode === "workbook" && workbookScopeMode === "auto" ? "active" : ""}
    onClick={chooseAutomaticScope}
  >
    自动
  </button>
  {/* ... 其他按钮 ... */}
  
  <button
    className="refresh-workbook-button"
    onClick={() => void refreshWorkbook()}
    disabled={busy || !workbook}
    title="刷新工作簿结构（工作簿名、工作表列表）"
  >
    🔄
  </button>
</div>
```

---

### 方案B：一个万能刷新按钮（更简单）

**位置**：全局工具栏（顶部或右上角）

**功能**：
- 同时刷新文件夹 + 工作簿
- 一键搞定所有层级

**前端逻辑**：
```typescript
async function refreshAll() {
  if (busy) return;
  setStatus("loading");
  
  try {
    // 1. 刷新当前工作簿
    const refreshedWorkbook = await captureWorkbook();
    setWorkbook(refreshedWorkbook);
    
    if (sourceMode === "workbook") {
      applyWorkbookSnapshotSelection(refreshedWorkbook);
    }
    
    // 2. 如果有文件夹 session，也刷新
    if (folderCatalog) {
      const refreshedFolder = await refreshFolder(folderCatalog.sessionId);
      applyFolderCatalog(refreshedFolder);
    }
  } catch (error) {
    alert(`刷新失败：${error instanceof Error ? error.message : String(error)}`);
  } finally {
    setStatus("idle");
  }
}
```

**UI 位置**：
```tsx
{/* 在页面顶部工具栏增加全局刷新按钮 */}
<div className="app-toolbar">
  <button
    className="refresh-all-button"
    onClick={() => void refreshAll()}
    disabled={busy}
    title="刷新文件夹和工作簿信息"
  >
    🔄 刷新
  </button>
  {/* ... 其他全局按钮 ... */}
</div>
```

---

## 推荐方案：方案A（两个刷新按钮）

**理由**：
1. **职责清晰**：文件夹的事归文件夹按钮，工作簿的事归工作簿按钮
2. **性能更好**：只刷需要的部分（用户改了表名就只刷工作簿，不用扫文件夹）
3. **错误隔离**：文件夹刷新失败不影响工作簿刷新

---

## 具体实施步骤

### 步骤 1：后端增加文件夹刷新接口

**文件**：`server/app/folder_workbooks.py`
```python
def refresh_folder(session_id: str) -> FolderCatalog:
    """刷新文件夹会话：重新扫描文件列表，保持 sessionId 不变"""
    _prune_sessions()
    
    try:
        old_session = _sessions[session_id]
    except KeyError:
        raise ValueError(f"会话 {session_id} 不存在或已过期")
    
    root = old_session.root
    all_candidates = sorted(
        (path for path in root.rglob("*") if _is_supported_file(path)),
        key=lambda path: str(path.relative_to(root)).casefold(),
    )
    candidates = all_candidates[:FILE_LIMIT]
    
    session_files: dict[str, Path] = {}
    files: list[FolderFileInfo] = []
    
    for path in candidates:
        file_id = uuid.uuid5(
            uuid.NAMESPACE_URL,
            str(path.relative_to(root)).replace("\\", "/").casefold(),
        ).hex
        session_files[file_id] = path
        relative_path = str(path.relative_to(root))
        try:
            workbook = load_workbook(
                path,
                read_only=True,
                data_only=True,
                keep_vba=path.suffix.lower() == ".xlsm",
            )
            worksheets = [
                FolderWorksheetInfo(
                    name=sheet.title,
                    rowCount=sheet.max_row,
                    columnCount=sheet.max_column,
                )
                for sheet in workbook.worksheets
            ]
            workbook.close()
            files.append(
                FolderFileInfo(
                    id=file_id,
                    name=path.name,
                    relativePath=relative_path,
                    worksheets=worksheets,
                )
            )
        except Exception as error:
            files.append(
                FolderFileInfo(
                    id=file_id,
                    name=path.name,
                    relativePath=relative_path,
                    error=f"无法读取：{error}",
                )
            )
    
    old_session.session_files = session_files
    old_session.last_access = time.monotonic()
    
    return FolderCatalog(
        sessionId=session_id,
        folderName=root.name,
        folderPath=str(root),
        files=files,
        totalFiles=len(all_candidates),
        truncated=len(all_candidates) > FILE_LIMIT,
        expiresAt=datetime.fromtimestamp(
            datetime.now().timestamp()
            + capability_int("folder", "sessionTtlSeconds")
        ).astimezone().isoformat(),
    )
```

**文件**：`server/app/main.py`
```python
class FolderRefreshRequest(BaseModel):
    sessionId: str

@app.post("/api/folders/refresh", response_model=FolderCatalog)
async def refresh_folder_endpoint(request: FolderRefreshRequest) -> FolderCatalog:
    from .folder_workbooks import refresh_folder
    return await run_in_threadpool(refresh_folder, request.sessionId)
```

---

### 步骤 2：前端增加 API 调用

**文件**：`apps/excel-addin/src/api.ts`
```typescript
export async function refreshFolder(sessionId: string): Promise<FolderCatalog> {
  const response = await fetch(`${API_BASE}/api/folders/refresh`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sessionId }),
  });
  if (!response.ok) {
    const error = await response.text();
    throw new Error(error || "刷新文件夹失败");
  }
  return response.json();
}
```

---

### 步骤 3：前端增加刷新逻辑和 UI

**文件**：`apps/excel-addin/src/App.tsx`

```typescript
// 1. 文件夹刷新函数
async function refreshCurrentFolder() {
  if (!folderCatalog || busy) return;
  
  setStatus("loading");
  try {
    const refreshed = await refreshFolder(folderCatalog.sessionId);
    applyFolderCatalog(refreshed);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("不存在或已过期")) {
      setFolderCatalog(null);
      alert("文件夹会话已过期，请重新选择文件夹");
    } else {
      alert(`刷新文件夹失败：${message}`);
    }
  } finally {
    setStatus("idle");
  }
}

// 2. 工作簿刷新函数
async function refreshWorkbook() {
  if (busy) return;
  
  setStatus("loading");
  try {
    const refreshed = await captureWorkbook();
    setWorkbook(refreshed);
    
    if (sourceMode === "workbook") {
      applyWorkbookSnapshotSelection(refreshed);
    }
  } catch (error) {
    alert(`刷新工作簿失败：${error instanceof Error ? error.message : String(error)}`);
  } finally {
    setStatus("idle");
  }
}

// 3. UI 修改 - 文件夹刷新按钮（在 browseFolder 按钮旁边）
<div className="folder-scope-picker">
  <div className="folder-toolbar">
    <button
      className="browse-folder-button"
      disabled={busy}
      onClick={() => void browseFolder()}
    >
      {folderCatalog ? `更换文件夹 · ${folderCatalog.folderName}` : "选择文件夹"}
    </button>
    
    {folderCatalog && (
      <button
        className="icon-button"
        disabled={busy}
        onClick={() => void refreshCurrentFolder()}
        title="刷新文件夹（更新文件列表）"
      >
        🔄
      </button>
    )}
  </div>
  {/* ... 文件列表 ... */}
</div>

// 4. UI 修改 - 工作簿刷新按钮（在范围选择器旁边）
<div className="scope-controls">
  <div className="scope-mode-picker">
    <button /* ... 自动 ... */></button>
    <button /* ... 手动 ... */></button>
    <button /* ... 文件夹 ... */></button>
  </div>
  
  <button
    className="icon-button"
    disabled={busy || !workbook}
    onClick={() => void refreshWorkbook()}
    title="刷新工作簿（更新工作簿名和工作表列表）"
  >
    🔄
  </button>
</div>
```

---

### 步骤 4：样式调整

**文件**：`apps/excel-addin/src/styles.css`
```css
.folder-toolbar {
  display: flex;
  gap: 8px;
  align-items: center;
  margin-bottom: 12px;
}

.scope-controls {
  display: flex;
  gap: 8px;
  align-items: center;
}

.icon-button {
  padding: 6px 10px;
  background: #f5f5f5;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  cursor: pointer;
  font-size: 16px;
  line-height: 1;
}

.icon-button:hover:not(:disabled) {
  background: #e8e8e8;
}

.icon-button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
```

---

## 测试验证

### 测试用例 1：刷新文件夹
1. 选择一个文件夹
2. 在资源管理器中重命名一个 Excel 文件
3. 点击文件夹刷新按钮（🔄）
4. **预期**：文件列表显示新文件名

### 测试用例 2：刷新工作表
1. 在 Excel 中重命名一个工作表
2. 点击工作簿刷新按钮（🔄）
3. **预期**：工作表列表显示新表名

### 测试用例 3：新增工作表
1. 在 Excel 中新增一个工作表
2. 点击工作簿刷新按钮（🔄）
3. **预期**：新表出现在列表中

### 测试用例 4：删除工作表
1. 选中某个工作表
2. 在 Excel 中删除该表
3. 点击工作簿刷新按钮（🔄）
4. **预期**：该表从列表中消失，选中状态自动清除

---

## 给 Codex 的指令

请按照以下顺序实施"文件夹 + 工作簿全量刷新"功能：

1. 后端：在 `server/app/folder_workbooks.py` 增加 `refresh_folder(session_id)` 函数
2. 后端：在 `server/app/main.py` 增加 `/api/folders/refresh` 端点
3. 前端：在 `apps/excel-addin/src/api.ts` 增加 `refreshFolder(sessionId)` 函数
4. 前端：在 `apps/excel-addin/src/App.tsx` 增加：
   - `refreshCurrentFolder()` 函数（文件夹刷新）
   - `refreshWorkbook()` 函数（工作簿刷新）
   - 两个刷新按钮的 UI（参考上面的代码）
5. 前端：在 `apps/excel-addin/src/styles.css` 增加按钮样式

注意事项：
- 文件夹刷新保持 sessionId 不变
- 工作簿刷新调用现有的 `captureWorkbook()` 即可
- session 过期时提示用户重新选择文件夹
- 按钮在 busy 时禁用
