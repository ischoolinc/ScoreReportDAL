## Goal

調整 108 課程規劃產生流程，使用 `graduation_plan.moe_group_code` 與 `graduation_plan.name` 建立記憶體索引，並比對既有課程規劃表內每一個 Subject 的 `課程代碼` 前 16 碼。

當 Subject 的課程代碼前 16 碼與目前課規的 `moe_group_code` 不同時，查找對應來源課規並維護 Subject XML 的：

```xml
來源課規="來源課程規劃表名稱"
```

同時在 `frmCreateGPlanItemSetup108` 新增唯讀欄位「來源課規」。

調整完成後，將修改內容與測試結果記錄在：

`產生課程規劃0824.md`

---

## Related files

請優先檢查／修改：

- `frmCreateGPlanMain108.cs`
- `GPlanInfo108.cs`
- `frmCreateGPlanItemSetup108.cs`
- `DataAccess.cs` 或目前負責查詢 `graduation_plan` 的 DAO

注意：不要改動其他無關邏輯。

必須保留既有：

- 新增
- 更新
- 刪除
- 略過
- 重置
- `HasSourceGPlan()`
- Subject Level 計算
- RowIndex 重算
- 使用者自訂科目
- `CourseGroupSetting`

---

## 1. 建立 `moe_group_code -> name` 索引

新增 DataAccess 查詢，讀取：

```text
graduation_plan.id
graduation_plan.moe_group_code
graduation_plan.name
```

建議 SQL：

```sql
SELECT
    id,
    moe_group_code,
    name
FROM graduation_plan
WHERE moe_group_code IS NOT NULL
  AND TRIM(moe_group_code) <> ''
ORDER BY id DESC;
```

建立：

```text
Dictionary<string, string>
Key   = moe_group_code
Value = name
```

例如：

```text
1131234567890123 -> 113普通科課程規劃表
1139876543210123 -> 113資訊科課程規劃表
```

要求：

- 一次讀取後在記憶體使用。
- 不可以每個 Subject 都重新 Query DB。
- 如果 `moe_group_code` 重複，避免 duplicate key exception。
- 可依 `id DESC` 保留最新一筆。

---

## 2. 在 `GPlanInfo108` 增加來源課規比對

新增方法，例如：

```csharp
ParseSourceGPlan(...)
```

資料來源：

```csharp
RefGPContentXml
```

逐筆處理：

```xml
<Subject ... />
```

取得：

```text
課程代碼
```

安全條件：

- `課程代碼` 不存在：略過。
- `課程代碼` 空白：略過。
- `課程代碼.Length < 16`：略過。
- 不可以因為格式異常造成 Exception。

正常情況：

```csharp
string subjectGroupCode = courseCode.Substring(0, 16);
```

再與目前：

```csharp
GDCCode
```

比較。

---

## 3. 來源課規判斷規則

### Case A：Subject 前 16 碼 == 目前 GDCCode

代表 Subject 屬於目前課規。

要求：

- `來源課規` 應為空白／不存在。
- 如果原本錯誤保留 `來源課規`，移除該 Attribute。
- 只有實際 XML 發生變更時才標記需要更新。

例如：

```xml
<Subject
    SubjectName="數學"
    課程代碼="1131234567890123ABCDEFG"
    ... />
```

---

### Case B：Subject 前 16 碼 != 目前 GDCCode，而且索引查得到

例如：

```text
目前 GDCCode：
1141234567890123

Subject 課程代碼：
1139876543210123ABCDEFG
```

索引：

```text
1139876543210123 -> 113資訊科課程規劃表
```

結果：

```xml
<Subject
    SubjectName="數學"
    課程代碼="1139876543210123ABCDEFG"
    來源課規="113資訊科課程規劃表"
    ... />
```

要求：

- Attribute 值使用來源課規的 `name`。
- 不要只寫 `moe_group_code`。
- 如果現有值已正確，不需標記變更。
- 如果現有值錯誤，更新成正確名稱。

---

### Case C：Subject 前 16 碼 != 目前 GDCCode，但索引查不到

要求：

- 不猜來源。
- 不建立錯誤 `來源課規`。
- 不拋錯。
- 不用 MessageBox 阻斷使用者。
- 不要用未知值覆蓋原有有效資料。

---

## 4. 增加來源課規異動旗標

在 `GPlanInfo108` 新增，例如：

```csharp
public bool needUpdateSourceGPlan = false;
```

只有以下情況設定 `true`：

- 新增 `來源課規`
- 修正 `來源課規`
- 移除錯誤／過期的 `來源課規`

如果值本來就正確，不要設定。

---

## 5. 執行順序

來源課規比對必須發生在：

```text
ParseMOEXml()
    ↓
ParseRefGPContentXml()
    ↓
來源課規動態比對
    ↓
SetCheckSubjectLevel(...)
    ↓
CheckData()
```

原因：

既有 `CheckData()` 已使用：

```csharp
HasSourceGPlan(...)
```

有 `來源課規` 的 Subject 必須繼續：

```text
ProcessStatus = "略過"
```

不要移除或改寫這項規則。

---

## 6. 保持來源課規 Subject 強制略過

既有行為：

```text
Subject 有 來源課規
    ↓
HasSourceGPlan() == true
    ↓
ProcessStatus = "略過"
```

必須保留。

尤其原本可能被判斷：

- 更新
- 刪除

的 Subject，只要確認來自其他課規，就必須保持「略過」。

目的：

> 從其他課程規劃表合併進來的 Subject，不應被目前課程代碼總表判定成更新或刪除。

---

## 7. 讓來源課規 Attribute 的變更可以被主畫面產生

因為來源 Subject 會：

```text
ProcessStatus = "略過"
```

整份課規可能原本會被算成：

```text
Status = "無變動"
```

因此主畫面狀態判斷需補：

```csharp
if (data.needUpdateSourceGPlan)
{
    data.Status = "更新";
}
```

目的：

即使所有來源 Subject 都是「略過」，只要 XML 中 `來源課規` Attribute 有新增／修正／移除，仍要讓這份課規進入主畫面的 UPDATE 流程。

不要新增新的 Status 類型。

---

## 8. 資料庫寫入時機

非常重要：

以下流程都只能修改記憶體，不可以直接寫 DB：

- BackgroundWorker 讀取
- 動態比對來源課規
- `frmCreateGPlanItemSetup108_Load`
- 開啟設定畫面
- 設定畫面按 Save / OK

正確流程：

```text
重新讀取
    ↓
動態比對來源課規
    ↓
只修改記憶體 XML
    ↓
設定畫面顯示
    ↓
回主畫面
    ↓
使用者按「產生 / 建立」
    ↓
既有 btnCreate_Click
    ↓
UPDATE graduation_plan.content
```

只有主畫面「產生 / 建立」才正式寫回資料庫。

---

## 9. 確認略過 Subject 仍保留來源課規 Attribute

主畫面 `btnCreate_Click` 中：

```text
ProcessStatus = "略過"
```

應持續使用：

```csharp
subj.GPlanXml
```

組回最終 `GraduationPlan` XML。

必須確認：

```xml
來源課規="..."
```

不會因重組 XML 遺失。

不要將「略過」Subject 改成使用 `MOEXml`，避免：

- 遺失來源課規
- 遺失既有課規設定
- 破壞合併進來的 Subject

最終寫入欄位：

```text
graduation_plan.content
```

---

## 10. `frmCreateGPlanItemSetup108` 新增「來源課規」欄位

在：

```text
frmCreateGPlanItemSetup108.cs
```

增加唯讀欄位：

```text
來源課規
```

建議欄位順序：

```text
科目名稱
來源課規
校訂部定
```

建議設定：

```csharp
Name = "來源課規";
HeaderText = "來源課規";
ReadOnly = true;
Width = 180;
```

此欄位只顯示，不可以人工編輯。

---

## 11. 設定畫面顯示來源課規

LoadData 時從：

```csharp
subj.GPlanXml
```

讀取 Subject：

```xml
來源課規
```

顯示規則：

- Attribute 有值：顯示內容。
- Attribute 無值：顯示空白。
- 同一 `chkSubjectInfo` 多個 XML Subject 都是相同來源：只顯示一次。
- 如果意外出現不同來源名稱，可去重後顯示。
- 不可拋錯。

範例：

```text
科目名稱          來源課規
------------------------------------
國語文
數學              113普通科課程規劃表
物理              113自然組課程規劃表
```

UI 層不要再次查 DB 或重新判斷來源，只負責顯示已經在資料層判斷好的結果。

---

## 12. 不要改動既有 Subject 差異判斷

不要修改既有：

- 領域
- 分項類別
- 報部科目名稱
- 校訂部定
- 必選修
- 開課方式
- 授課學期學分
- 學分數
- 不需評分
- 不計學分
- Level
- startLevel
- 開課學期數

這次只新增「來源課規辨識與顯示」。

---

## 13. 不要影響 Subject Level 與 RowIndex

不要修改：

```text
Utility.CalculateSubjectLevel(...)
```

不要改主畫面最後：

```text
RowIndex
```

重算邏輯。

來源課規 Subject 即使：

```text
ProcessStatus = 略過
```

仍必須保留在最終 XML，並照現有流程排序與重算 RowIndex。

---

## 14. 防呆

必須安全處理：

- `RefGPContentXml == null`
- `GDCCode` null / empty
- `課程代碼` Attribute 不存在
- 課程代碼短於 16 碼
- `來源課規` Attribute 不存在
- 找不到來源 `moe_group_code`
- `moe_group_code` 有重複資料

不可因以上資料造成 NullReferenceException 或 ArgumentOutOfRangeException。

---

## 15. 測試案例

### Test 1：同一課規

```text
目前 moe_group_code：
AAAAAAAAAAAAAAAA

Subject 課程代碼：
AAAAAAAAAAAAAAAAXXXX
```

預期：

```text
來源課規 = 空白
```

如果原本沒有來源課規，不應產生不必要 UPDATE。

---

### Test 2：不同課規且找到來源

```text
目前 moe_group_code：
AAAAAAAAAAAAAAAA

Subject 課程代碼：
BBBBBBBBBBBBBBBBXXXX
```

索引：

```text
BBBBBBBBBBBBBBBB -> B課程規劃表
```

預期記憶體 XML：

```xml
來源課規="B課程規劃表"
```

預期：

```text
Subject ProcessStatus = 略過
GPlan Status = 更新
```

在按主畫面「產生」之前：

```text
資料庫不變
```

按「產生」後：

```text
graduation_plan.content
```

應包含：

```xml
來源課規="B課程規劃表"
```

---

### Test 3：不同課規但找不到來源

預期：

- 不拋錯。
- 不猜名稱。
- 不建立錯誤資料。
- 不破壞現有值。

---

### Test 4：來源課規原本已正確

現有：

```xml
來源課規="B課程規劃表"
```

索引結果也是：

```text
B課程規劃表
```

預期：

```text
needUpdateSourceGPlan = false
```

除非其他 Subject 有異動。

---

### Test 5：來源課規名稱錯誤

現有：

```xml
來源課規="舊課程規劃表"
```

正確索引：

```text
B課程規劃表
```

預期：

```xml
來源課規="B課程規劃表"
```

並：

```text
needUpdateSourceGPlan = true
```

---

### Test 6：同一課規卻殘留來源課規

目前：

```text
moe_group_code = AAAAAAAAAAAAAAAA
```

Subject：

```text
課程代碼 = AAAAAAAAAAAAAAAAXXXX
來源課規 = B課程規劃表
```

預期：

- 移除 `來源課規`。
- `needUpdateSourceGPlan = true`。
- 只有按主畫面「產生」後才真正移除 DB 內 Attribute。

---

### Test 7：設定畫面

開啟：

```text
frmCreateGPlanItemSetup108
```

預期：

- 有「來源課規」欄位。
- 正常 Subject 顯示空白。
- 來源 Subject 顯示來源課規名稱。
- 欄位唯讀。

---

### Test 8：只查看、不產生

流程：

```text
重新讀取
→ 開設定畫面
→ 查看來源課規
→ 關閉
```

但沒有按主畫面「產生」。

預期：

```text
資料庫完全不變
```

---

### Test 9：正式產生

來源課規在記憶體有異動後：

```text
按主畫面「產生」
```

預期：

- 執行既有 UPDATE。
- `graduation_plan.content` 正式寫入來源課規 Attribute。
- 來源 Subject 仍然保持略過。
- 不被 MOE XML 覆蓋或刪除。

---

## 16. Scope restrictions

不要：

- 在動態比對階段直接 UPDATE DB。
- 在 `frmCreateGPlanItemSetup108` 寫 DB。
- 改掉 `HasSourceGPlan()` 的原本用途。
- 將來源 Subject 改成「更新」。
- 修改其他 Subject merge 規則。
- 修改 Subject Level 計算。
- 修改 RowIndex 邏輯。
- 大範圍重構無關程式。
- 新增資料庫欄位存放來源課規。

`來源課規` 必須維持為：

```text
graduation_plan.content
```

內 Subject 的 XML Attribute。

---

## 17. 完成紀錄

修改完成後建立／更新：

```text
產生課程規劃0824.md
```

內容至少記錄：

1. 修改檔案。
2. `moe_group_code -> name` 索引做法。
3. Subject 課程代碼前 16 碼比對規則。
4. `來源課規` 新增／修正／移除條件。
5. `needUpdateSourceGPlan` 用途。
6. 為何來源 Subject 維持 `ProcessStatus = 略過`。
7. 為何重新讀取／設定畫面不直接寫 DB。
8. 主畫面「產生」如何將記憶體 XML 回寫 `graduation_plan.content`。
9. `frmCreateGPlanItemSetup108` 新增「來源課規」欄位。
10. 實際測試案例與結果。
11. 實作中發現的例外資料或風險。

---

## Final acceptance criteria

完成前確認：

- 已建立 `moe_group_code -> name` 記憶體索引。
- Subject 可安全取得課程代碼前 16 碼。
- 可判斷是否與目前課規 `moe_group_code` 相同。
- 可正確新增／修正／移除 `來源課規` Attribute。
- 找不到來源課規時不猜、不拋錯。
- `HasSourceGPlan()` 原有保護邏輯保持有效。
- 來源課規 Subject 維持 `ProcessStatus = 略過`。
- `來源課規` 有異動時，整份課規會進入主畫面 UPDATE 流程。
- 設定畫面只顯示記憶體資料，不寫 DB。
- 只有主畫面按「產生 / 建立」才寫回 `graduation_plan.content`。
- `frmCreateGPlanItemSetup108` 已新增唯讀「來源課規」欄位。
- 不影響既有新增、更新、刪除、重置、Subject Level、RowIndex、使用者自訂科目、CourseGroupSetting。
- 修改與測試結果已記錄在 `產生課程規劃0824.md`。
