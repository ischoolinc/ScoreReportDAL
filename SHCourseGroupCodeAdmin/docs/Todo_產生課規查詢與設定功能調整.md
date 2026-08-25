
## 目標

調整 `frmCreateGPlanQueryAndSetup108`「查詢與設定」功能，在科目清單中新增唯讀欄位 `來源課規`。

此欄位只顯示目前記憶體中 `chkSubjectInfo.GPlanXml` 內 Subject 的：

```xml
來源課規="..."
```

不重新查詢資料庫、不重新判斷來源課規、不修改來源課規 Attribute。

修改完成後，將本次調整內容與測試結果記錄在：

`產生課程規畫調整0825.md`

---

## 1. 主要修改檔案

主要修改：

- `frmCreateGPlanQueryAndSetup108.cs`

必要時檢查：

- `frmCreateGPlanMain108.cs`
- `GPlanInfo108.cs`
- `chkSubjectInfo.cs`

不要大範圍重構其他課程規劃產生流程。

---

## 2. 保留目前資料流

主畫面目前已先完成：

```text
GPlanInfo108List()
→ GetMoeGroupCodeNameDict()
→ ParseMOEXml()
→ ParseRefGPContentXml()
→ ParseSourceGPlan(...)
→ CheckData()
→ 建立 chkSubjectInfoList
```

進入 `btnQueryAndSet` 時：

```csharp
fgq.SetGPlanInfos(_GPlanInfo108List);
```

傳入的 `GPlanInfo108` / `chkSubjectInfo` 已包含記憶體中的來源課規結果。

因此 `frmCreateGPlanQueryAndSetup108` 不要再次：

- Query `graduation_plan`
- 建立 `moe_group_code -> name` 索引
- 比對課程代碼前 16 碼
- 執行 `ParseSourceGPlan()`

此畫面只顯示既有結果。

---

## 3. 新增「來源課規」欄位

在 `LoadDataGridViewColumns()` 新增唯讀欄位。

建議位置：

```text
群科班
差異狀態
處理方式
領域
分項類別
科目名稱
來源課規
校訂部定
必選修
...
```

建議程式：

```csharp
DataGridViewTextBoxColumn tbSourceGPlan =
    new DataGridViewTextBoxColumn();

tbSourceGPlan.Name = "來源課規";
tbSourceGPlan.HeaderText = "來源課規";
tbSourceGPlan.Width = 180;
tbSourceGPlan.ReadOnly = true;
```

加入順序：

```csharp
dgData.Columns.Add(tbSubjectName);
dgData.Columns.Add(tbSourceGPlan);
dgData.Columns.Add(tbRequiredBy);
```

要求：

- 欄位必須唯讀。
- 不允許直接編輯來源課規。
- 不修改「處理方式」欄位原有行為。

---

## 4. 新增來源課規顯示 Helper

在 `frmCreateGPlanQueryAndSetup108.cs` 增加：

```csharp
private string GetSourceGPlan(chkSubjectInfo subj)
```

資料來源只使用：

```csharp
subj.GPlanXml
```

不要從 `MOEXml` 讀取。

建議：

```csharp
private string GetSourceGPlan(chkSubjectInfo subj)
{
    if (subj == null || subj.GPlanXml == null)
        return "";

    List<string> sourceNames = new List<string>();

    foreach (XElement elm in subj.GPlanXml)
    {
        if (elm == null)
            continue;

        XAttribute attr = elm.Attribute("來源課規");

        if (attr == null)
            continue;

        string value = attr.Value.Trim();

        if (string.IsNullOrEmpty(value))
            continue;

        if (!sourceNames.Contains(value))
            sourceNames.Add(value);
    }

    return string.Join(",", sourceNames.ToArray());
}
```

顯示規則：

- `subj == null`：空白。
- `GPlanXml == null`：空白。
- 沒有 `來源課規`：空白。
- 有值：顯示來源課規名稱。
- 同一科目多學期來源相同：只顯示一次。
- 若異常出現多個不同來源：去重後以逗號串接。
- 不可拋 Exception。

---

## 5. LoadData 顯示來源課規

目前：

```csharp
dgData.Rows[rowIdx].Cells["科目名稱"].Value = subj.SubjectName;
dgData.Rows[rowIdx].Cells["校訂部定"].Value = subj.RequiredBy;
```

調整：

```csharp
dgData.Rows[rowIdx].Cells["科目名稱"].Value = subj.SubjectName;
dgData.Rows[rowIdx].Cells["來源課規"].Value = GetSourceGPlan(subj);
dgData.Rows[rowIdx].Cells["校訂部定"].Value = subj.RequiredBy;
```

範例：

```text
群科班   科目名稱   來源課規                  處理方式
------------------------------------------------------
普通科   國語文                               略過
普通科   數學       113資訊科課程規劃表       略過
普通科   物理       113自然組課程規劃表       略過
```

---

## 6. UI 不重新處理來源課規

本次禁止在 `frmCreateGPlanQueryAndSetup108` 新增：

- `GetMoeGroupCodeNameDict()`
- `ParseSourceGPlan()`
- `Substring(0,16)` 來源判斷
- `UPDATE graduation_plan`

此畫面只負責：

```text
讀取記憶體資料
→ 顯示來源課規
```

---

## 7. 不可影響 ProcessStatus

新增「來源課規」欄位後，不得因來源課規有值而在 UI 中重新修改：

- 新增
- 更新
- 刪除
- 略過
- 重置

來源 Subject 是否「略過」仍由：

```text
GPlanInfo108.CheckData()
HasSourceGPlan()
```

決定。

UI 只顯示結果。

---

## 8. 不可修改來源課規 XML

「查詢與設定」畫面只顯示 `來源課規`。

禁止：

- `SetAttributeValue("來源課規", ...)`
- Remove `來源課規`
- 修改來源課規名稱
- 清除來源課規
- 從畫面直接寫回 XML

Attribute 的建立、修改、移除仍由 `ParseSourceGPlan(...)` 負責。

---

## 9. 不要直接寫資料庫

正確流程保持：

```text
主畫面重新讀取
→ 來源課規動態比對
→ 記憶體 XML
→ btnQueryAndSet
→ 查詢與設定畫面顯示來源課規
→ 按儲存
→ 回主畫面
→ 使用者按「產生 / 建立」
→ 才正式 UPDATE graduation_plan.content
```

`frmCreateGPlanQueryAndSetup108` 不可以直接 UPDATE DB。

---

## 10. 保持原本物件參考方式

目前：

```csharp
public void SetGPlanInfos(List<GPlanInfo108> data)
{
    _GPlanInfoList = data;
}
```

以及：

```csharp
dgData.Rows[rowIdx].Tag = subj;
```

都是直接使用原物件。

不要因本次功能改成 Clone 或 new：

- `GPlanInfo108`
- `chkSubjectInfo`

避免破壞主畫面與批次設定共用同一份記憶體資料的流程。

---

## 11. btnSave 原有流程保持不變

`btnSave_Click` 目前主要負責：

```text
DataGridView
→ 讀取處理方式
→ 回存 chkSubjectInfo.ProcessStatus
→ 重整 GPlanInfo108.chkSubjectInfoList
→ 回傳 _GPlanInfoDict
```

新增「來源課規」後：

- 不需要在 btnSave 回寫來源課規。
- 不需要把來源課規另外存進 property。
- 不需要重新建立 XML。
- 不需要 DB Update。

---

## 12. 確認 btnSave 不會遺失來源課規

因為 `來源課規` 存在：

```csharp
subj.GPlanXml
```

而 Row.Tag 仍是原本的：

```csharp
chkSubjectInfo
```

所以 btnSave 後確認：

```xml
來源課規="..."
```

仍存在於 `subj.GPlanXml`。

不得因重整 `chkSubjectInfoList` 而遺失。

---

## 13. ParseStatus 必要驗證

`btnQueryAndSet` 回主畫面後會呼叫：

```csharp
dataDict[data.GDCCode].ParseStatus();
```

請確認 `GPlanInfo108.ParseStatus()` 能保留來源課規 Attribute 異動造成的更新狀態：

```csharp
if (needUpdateSourceGPlan)
    Status = "更新";
```

並保留原有需要的：

```csharp
if (needUpdateEntryYear)
    Status = "更新";
```

如果目前已經有：

- 保留，不要移除。

如果目前尚未有：

- 在不改變其他 Status 邏輯的前提下補上。
- 原因是來源課規 Attribute 若在記憶體中有新增／修正／移除，經過 `btnQueryAndSet` 後仍必須保持整份課規為「更新」，主畫面按「產生」時才能寫回 DB。

不要新增新的 Status 類型。

---

## 14. 測試案例

### Test 1：一般科目

沒有：

```xml
來源課規
```

預期「來源課規」欄位顯示空白。

### Test 2：有來源課規

```xml
<Subject
    SubjectName="數學"
    來源課規="113資訊科課程規劃表"
/>
```

預期顯示：

```text
113資訊科課程規劃表
```

### Test 3：同一科目多學期來源相同

若 `subj.GPlanXml` 有多個 semester，且來源相同，預期只顯示一次。

### Test 4：異常多來源

若同一 `chkSubjectInfo.GPlanXml` 意外有：

```text
A課規
B課規
```

預期：

```text
A課規,B課規
```

且不拋錯、不修改 XML。

### Test 5：來源 Subject 處理方式

來源 Subject 原本：

```text
ProcessStatus = 略過
```

進入「查詢與設定」後仍應為：

```text
略過
```

不得因新增欄位改變。

### Test 6：儲存查詢與設定

流程：

```text
開啟 btnQueryAndSet
→ 查看來源課規
→ 按儲存
→ 回主畫面
```

確認：

- `subj.GPlanXml` 的 `來源課規` 仍存在。
- 沒有直接 UPDATE DB。
- ProcessStatus 正常保存。

### Test 7：本次剛動態補來源課規

流程：

```text
重新讀取
→ ParseSourceGPlan()
→ needUpdateSourceGPlan = true
→ 主畫面 Status = 更新
→ 開 btnQueryAndSet
→ 顯示來源課規
→ 按儲存
→ 回主畫面
```

預期：

```text
Status 仍為更新
```

不能變成：

```text
無變動
```

### Test 8：正式產生

完成 Test 7 後按主畫面：

```text
產生 / 建立
```

預期：

- `graduation_plan.content` 正式寫入來源課規 Attribute。
- 來源 Subject 仍為「略過」。
- 不被 MOE XML 覆蓋。
- 不被刪除。

---

## 15. Scope restrictions

本次不要：

- 改來源課規前 16 碼比對規則。
- 改 `ParseSourceGPlan()` 主邏輯。
- 改 `HasSourceGPlan()` 主邏輯。
- 改 Subject Level。
- 改 RowIndex。
- 改新增／更新／刪除／重置規則。
- 改使用者自訂科目。
- 改 `CourseGroupSetting`。
- 在查詢與設定畫面 Query DB。
- 在查詢與設定畫面直接寫 DB。
- 讓「來源課規」可編輯。
- 大範圍重構既有程式。

---

## 16. 完成紀錄

修改完成後建立／更新：

`產生課程規畫調整0825.md`

至少記錄：

1. 修改檔案。
2. `frmCreateGPlanQueryAndSetup108` 新增「來源課規」欄位。
3. 欄位位置與 ReadOnly 設定。
4. `GetSourceGPlan()` 讀取方式。
5. 資料來源為 `subj.GPlanXml`。
6. 無來源課規時顯示空白。
7. 多學期相同來源去重方式。
8. btnSave 是否確認不會清掉 `來源課規` Attribute。
9. 是否確認 `btnQueryAndSet` 不會直接寫 DB。
10. 是否確認回主畫面後 `needUpdateSourceGPlan` 的更新狀態沒有被 `ParseStatus()` 清掉。
11. 測試案例與結果。
12. 實作中發現的問題或風險。

---

## Final Acceptance Criteria

完成前確認：

- 「查詢與設定」畫面已新增唯讀「來源課規」欄位。
- 欄位資料直接來自 `chkSubjectInfo.GPlanXml`。
- 沒有來源課規時顯示空白。
- 有來源課規時顯示課規名稱。
- 同一科目多學期相同來源不重複顯示。
- UI 不重新查詢 DB。
- UI 不重新比對課程代碼前 16 碼。
- UI 不修改 `來源課規` Attribute。
- UI 不直接 UPDATE DB。
- btnSave 不會遺失 `subj.GPlanXml`。
- 來源 Subject 原有 `ProcessStatus = 略過` 保持不變。
- `btnQueryAndSet` 回主畫面後，來源課規 Attribute 有異動的課規仍保持 `Status = 更新`。
- 主畫面按「產生 / 建立」後仍可正常寫回 `graduation_plan.content`。
- 不影響既有新增、更新、刪除、略過、重置、Level、RowIndex 等流程。
- 修改與測試結果已記錄在 `產生課程規畫調整0825.md`。
