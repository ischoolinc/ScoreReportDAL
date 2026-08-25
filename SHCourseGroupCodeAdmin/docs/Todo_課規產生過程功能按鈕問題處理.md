## 目標

修正 `frmCreateGPlanMain108` 在執行「產生」後，到 BackgroundWorker 重新讀取資料並回填主畫面完成之前，功能按鈕被過早重新啟用的問題。

修改後需確保：

- 產生進行中不可再次按「產生」。
- BackgroundWorker 重新讀取期間不可再次按「重新讀取」。
- 重新讀取期間不可進入「查詢與設定」或單筆「設定」。
- 不可在資料尚未重新載入完成時操作 DataGrid。
- 控制項只在 `_bgWorker_RunWorkerCompleted()` 完成後重新啟用。
- 不改動既有課程規劃產生、更新、來源課規、Level、RowIndex 等資料邏輯。

修改完成後，將本次修改內容與測試結果記錄在：

`產生課程規畫調整0825.md`

---

## 一、主要修改檔案

主要修改：

- `frmCreateGPlanMain108.cs`

本次原則上不要修改：

- `GPlanInfo108.cs`
- `frmCreateGPlanItemSetup108.cs`
- `frmCreateGPlanQueryAndSetup108.cs`
- `DataAccess.cs`

除非實作過程發現與控制項狀態直接相關，否則不要擴大修改範圍。

---

## 二、目前問題

目前 `btnCreate_Click()` 一開始會：

```csharp
ControlEnable(false);
```

產生完成後，如果有 INSERT 或 UPDATE：

```csharp
ControlEnable(false);
_bgWorker.RunWorkerAsync();
```

啟動 BackgroundWorker 重新讀取。

但 `RunWorkerAsync()` 是非同步呼叫，會立即返回。

目前 `btnCreate_Click()` 最後仍會繼續執行：

```csharp
ControlEnable(true);
```

因此實際流程會變成：

```text
按「產生」
    ↓
ControlEnable(false)
    ↓
INSERT / UPDATE
    ↓
ControlEnable(false)
    ↓
_bgWorker.RunWorkerAsync()
    ↓
BackgroundWorker 正在重新讀取
    ↓
btnCreate_Click 繼續往下
    ↓
ControlEnable(true)
    ↓
其他功能提前可操作
```

這會造成重新讀取期間使用者仍可操作：

- 產生
- 查詢與設定
- 重新讀取
- chkStartLevel
- DataGrid
- DataGrid 內的「設定」

尤其再次按：

```csharp
_bgWorker.RunWorkerAsync();
```

可能造成 BackgroundWorker busy 相關例外。

---

## 三、修正原則

當：

```text
產生成功
→ 啟動 BackgroundWorker 重新讀取
```

之後，主畫面功能必須維持 disabled。

只有：

```csharp
_bgWorker_RunWorkerCompleted(...)
```

真正完成：

- 資料讀取
- DataGrid 清除
- DataGrid 重建
- 狀態統計

之後，才執行：

```csharp
ControlEnable(true);
```

---

## 四、修改 btnCreate_Click

找到：

```csharp
if (insertSQLList.Count > 0 || updateSQLList.Count > 0)
{
    try
    {
        FISCA.Features.Invoke("GraduationPlanSyncAllBackground");
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }

    ControlEnable(false);
    _bgWorker.RunWorkerAsync();
}
```

目前後面還會執行：

```csharp
ControlEnable(true);
```

請修正成：

```csharp
if (insertSQLList.Count > 0 || updateSQLList.Count > 0)
{
    try
    {
        FISCA.Features.Invoke("GraduationPlanSyncAllBackground");
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }

    ControlEnable(false);
    _bgWorker.RunWorkerAsync();

    return;
}
```

目的：

```text
BackgroundWorker 啟動後
→ 立即離開 btnCreate_Click
→ 不執行方法最後的 ControlEnable(true)
```

---

## 五、保留無資料更新時的解鎖

以下情況仍需正常重新開啟控制項：

### 情況 A

沒有新增／更新資料：

```csharp
if (insertDataList.Count == 0 && updateDataList.Count == 0)
{
    MsgBox.Show("沒有新增或更新資料。");
    ControlEnable(true);
    return;
}
```

此邏輯保留。

### 情況 B

產生流程沒有啟動 BackgroundWorker

若因條件沒有進入：

```csharp
_bgWorker.RunWorkerAsync();
```

則可以在方法結束前：

```csharp
ControlEnable(true);
```

不要造成按鈕永久 disabled。

---

## 六、控制項重新啟用唯一正常位置

確認：

```csharp
_bgWorker_RunWorkerCompleted(...)
```

最後已經有：

```csharp
GPlanDataCount();
ControlEnable(true);
```

此處應作為重新讀取完成後的主要解鎖位置。

正確流程：

```text
按「產生」
    ↓
ControlEnable(false)
    ↓
INSERT / UPDATE
    ↓
RunWorkerAsync
    ↓
return
    ↓
所有主要功能保持 disabled
    ↓
_bgWorker_DoWork
    ↓
重新讀取課程代碼總表與課程規劃表
    ↓
_bgWorker_RunWorkerCompleted
    ↓
重新填入 dgData
    ↓
GPlanDataCount()
    ↓
ControlEnable(true)
```

---

## 七、確認 ControlEnable 控制範圍

目前：

```csharp
private void ControlEnable(bool value)
{
    dgData.Enabled =
        btnCreate.Enabled =
        btnQueryAndSet.Enabled =
        chkStartLevel.Enabled =
        btnReload.Enabled = value;
}
```

必須至少繼續控制：

- `dgData`
- `btnCreate`
- `btnQueryAndSet`
- `chkStartLevel`
- `btnReload`

確保重新讀取期間以上控制項不可操作。

---

## 八、btnCancel 是否要一起停用

請檢查目前：

```text
btnCancel
```

是否未包含在 `ControlEnable()`。

建議本次一起評估：

### 建議方案

在產生與重新讀取期間也停用：

```csharp
btnCancel.Enabled = value;
```

例如：

```csharp
private void ControlEnable(bool value)
{
    dgData.Enabled =
        btnCreate.Enabled =
        btnQueryAndSet.Enabled =
        chkStartLevel.Enabled =
        btnReload.Enabled =
        btnCancel.Enabled = value;
}
```

理由：

若 BackgroundWorker 執行期間使用者直接關閉 Form，可能造成：

- Form 已關閉但 BackgroundWorker Completed 還要操作 UI。
- 非必要的生命週期風險。

若專案原本設計刻意允許重新讀取期間關閉 Form，請不要任意改變產品行為，但需在完成紀錄中說明原因。

---

## 九、防止 BackgroundWorker 重複啟動

不要新增複雜架構，但可增加簡單防呆。

在需要呼叫：

```csharp
_bgWorker.RunWorkerAsync();
```

前確認：

```csharp
if (!_bgWorker.IsBusy)
{
    _bgWorker.RunWorkerAsync();
}
```

適用位置至少檢查：

- `frmCreateGPlanMain108_Load`
- `btnReload_Click`
- `btnCreate_Click` 產生完成後重新讀取

注意：

主要問題仍應透過 UI disabled 解決。

`IsBusy` 是第二層安全防護，不要取代正常按鈕狀態控制。

---

## 十、btnReload_Click 防呆

目前：

```csharp
private void btnReload_Click(object sender, EventArgs e)
{
    _CheckSubjectLevel = chkStartLevel.Checked;
    ControlEnable(false);
    _bgWorker.RunWorkerAsync();
}
```

建議調整成安全形式，例如：

```csharp
private void btnReload_Click(object sender, EventArgs e)
{
    if (_bgWorker.IsBusy)
        return;

    _CheckSubjectLevel = chkStartLevel.Checked;
    ControlEnable(false);
    _bgWorker.RunWorkerAsync();
}
```

不要在 BackgroundWorker busy 時再次執行。

---

## 十一、產生完成後重新讀取期間 UI 狀態

在：

```text
UPDATE / INSERT 已完成
```

但：

```text
BackgroundWorker 尚未完成
```

期間預期：

| 控制項 | 狀態 |
|---|---|
| `dgData` | Disabled |
| `btnCreate` | Disabled |
| `btnQueryAndSet` | Disabled |
| `chkStartLevel` | Disabled |
| `btnReload` | Disabled |
| DataGrid 設定功能 | 無法操作 |
| `btnCancel` | 依本次決策，建議 Disabled |

---

## 十二、重新讀取完成後 UI 狀態

只有 `_bgWorker_RunWorkerCompleted()` 完成後：

```csharp
ControlEnable(true);
```

預期：

| 控制項 | 狀態 |
|---|---|
| `dgData` | Enabled |
| `btnCreate` | Enabled |
| `btnQueryAndSet` | Enabled |
| `chkStartLevel` | Enabled |
| `btnReload` | Enabled |
| DataGrid 設定功能 | 可操作 |
| `btnCancel` | Enabled |

---

## 十三、不要修改資料處理邏輯

本次禁止修改：

- `ParseMOEXml()`
- `ParseRefGPContentXml()`
- `ParseSourceGPlan()`
- `CheckData()`
- `HasSourceGPlan()`
- `needUpdateSourceGPlan`
- `needUpdateEntryYear`
- Subject `ProcessStatus`
- 新增／更新／刪除／略過／重置規則
- Level
- startLevel
- RowIndex
- 使用者自訂科目
- `CourseGroupSetting`
- SQL UPDATE / INSERT 內容

只處理 UI enabled/disabled 與 BackgroundWorker 重複啟動安全問題。

---

## 十四、例外處理注意

確認產生流程發生 Exception 時，控制項不會永久 disabled。

目前最外層：

```csharp
catch (Exception ex)
{
    MsgBox.Show("儲存發生錯誤：" + ex.Message);
}
```

後續若沒有啟動 BackgroundWorker，應恢復：

```csharp
ControlEnable(true);
```

但如果 BackgroundWorker 已經成功啟動：

- 不可提前 `ControlEnable(true)`。
- 由 `_bgWorker_RunWorkerCompleted()` 負責解鎖。

---

## 十五、測試案例

### Test 1：正常產生

流程：

```text
按「產生」
→ 執行 UPDATE / INSERT
→ 啟動重新讀取
```

重新讀取期間測試：

- 產生不可點。
- 查詢與設定不可點。
- 重新讀取不可點。
- Checkbox 不可點。
- DataGrid 不可操作。
- 設定不可點。

重新讀取完成後：

- 全部恢復可操作。

---

### Test 2：重新讀取期間快速點擊

在 BackgroundWorker 執行期間嘗試：

```text
連續點「重新讀取」
```

預期：

- 不會再次啟動 BackgroundWorker。
- 不發生 `BackgroundWorker is currently busy`。
- 不發生 `InvalidOperationException`。

---

### Test 3：重新讀取期間再次按產生

預期：

- 按鈕 disabled。
- 不會執行第二次 UPDATE / INSERT。

---

### Test 4：重新讀取期間開查詢與設定

預期：

- `btnQueryAndSet` disabled。
- 無法進入批次查詢設定。

---

### Test 5：沒有新增或更新資料

按產生後：

```text
沒有新增或更新資料。
```

預期：

- 顯示 MessageBox。
- 控制項恢復 enabled。
- 不啟動 BackgroundWorker。

---

### Test 6：資料寫入發生 Exception

預期：

- 顯示錯誤。
- 若 BackgroundWorker 未啟動，控制項恢復 enabled。
- 不造成畫面永久鎖住。

---

### Test 7：BackgroundWorker 完成

確認：

```csharp
_bgWorker_RunWorkerCompleted
```

成功：

- 清除舊 dgData。
- 重新加入資料。
- 重新統計。
- 最後 `ControlEnable(true)`。

---

### Test 8：Cancel 行為

若本次將 `btnCancel` 納入 `ControlEnable()`：

重新讀取期間：

```text
btnCancel = Disabled
```

完成後：

```text
btnCancel = Enabled
```

確認 Form 不會在 BackgroundWorker 執行途中被使用者直接關閉。

---

## 十六、完成紀錄

修改完成後建立／更新：

`產生課程規畫調整0825.md`

至少記錄：

1. 修改檔案。
2. 原本 `btnCreate_Click()` 為何會提前 `ControlEnable(true)`。
3. `RunWorkerAsync()` 為非同步呼叫的影響。
4. 如何透過 `return` 避免產生後過早解鎖。
5. `_bgWorker_RunWorkerCompleted()` 作為重新啟用控制項的位置。
6. 是否加入 `_bgWorker.IsBusy` 防呆。
7. `btnReload_Click` 是否增加 busy 防護。
8. `btnCancel` 是否納入 `ControlEnable()`，以及決策原因。
9. 是否確認重新讀取期間無法再次產生。
10. 是否確認重新讀取期間無法進入查詢與設定。
11. 是否確認重新讀取期間 DataGrid 無法操作。
12. 測試案例與結果。
13. 實作中發現的其他風險。

---

## Final Acceptance Criteria

完成前必須確認：

- 按「產生」後立即停用主要功能。
- UPDATE / INSERT 完成後啟動 BackgroundWorker。
- `RunWorkerAsync()` 後不會立即執行 `ControlEnable(true)`。
- 重新讀取期間 `btnCreate` 不可點。
- 重新讀取期間 `btnQueryAndSet` 不可點。
- 重新讀取期間 `btnReload` 不可點。
- 重新讀取期間 `chkStartLevel` 不可操作。
- 重新讀取期間 `dgData` 不可操作。
- 不會因快速操作造成 BackgroundWorker 重複啟動。
- `_bgWorker_RunWorkerCompleted()` 完成後才重新啟用功能。
- 沒有更新資料或發生錯誤時，不會造成 UI 永久 disabled。
- 不影響現有課程規劃產生、來源課規、Level、RowIndex 等資料流程。
- 修改與測試結果已記錄在 `產生課程規畫調整0825.md`。
