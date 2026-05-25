
## 目標

調整畢業預警報表中「科目清單」的輸出邏輯。

目前程式邏輯：

- 學分數 = 0 時，原則上不輸出。
- 但「非課規科目」即使學分數 = 0，仍會保留輸出。

本次要調整為：

- 只要學分數 = 0，一律不輸出。
- 即使是「非課規科目」，學分數 = 0 也不輸出。
- 不影響學分數不是 0 的「非課規科目」輸出邏輯。

---

## 修改檔案

frmMain.cs

---

## 修改位置

請搜尋以下註解或判斷式：

```csharp
// 2026-05-06 調整：非課規科目即使學分數 = 0 仍需保留輸出
if (xmlRuleS.GetAttribute("學分數") == "0" && !isNonGraduationPlanSubject)
    continue;
````

此段位於：

```csharp
BgwGrandCheckReport_DoWork
```

內部處理：

```csharp
foreach (XmlElement xmlRuleS in xmlRule.SelectNodes("科目"))
```

的流程中。

---

## 修改方式

將原本：

```csharp
// 原本邏輯：學分數 = 0 不列入報表科目清單
// 2026-05-06 調整：非課規科目即使學分數 = 0 仍需保留輸出
if (xmlRuleS.GetAttribute("學分數") == "0" && !isNonGraduationPlanSubject)
    continue;
```

改成：

```csharp
// 學分數 = 0 不列入報表科目清單
// 2026-05-25 調整：非課規科目即使學分數 = 0 也不輸出
if (xmlRuleS.GetAttribute("學分數") == "0")
    continue;
```

---

## 注意事項

1. 只移除 `&& !isNonGraduationPlanSubject` 判斷。
2. 不要刪除 `isNonGraduationPlanSubject` 變數。
3. 不要刪除後面加入狀態的邏輯：

```csharp
if (isNonGraduationPlanSubject)
    statusList.Add("非課規科目");
```

4. 調整後，只有「學分數不是 0」的非課規科目，才可以繼續輸出為「非課規科目」。
5. 不要影響「可重修」、「可補修」、「不採計」原本輸出邏輯。
6. 不要調整其他報表、SQL、畫面載入邏輯。

---

## 驗證項目

請使用測試資料確認以下情境：

### 情境 1：一般科目，學分數 = 0

預期結果：

* 不輸出到畢業預警通知單科目清單。

### 情境 2：非課規科目，學分數 = 0

預期結果：

* 不輸出到畢業預警通知單科目清單。

### 情境 3：非課規科目，學分數 > 0

預期結果：

* 仍可輸出到畢業預警通知單科目清單。
* 狀態欄位顯示「非課規科目」。

### 情境 4：可重修 / 可補修 / 不採計，學分數 > 0

預期結果：

* 原本輸出邏輯不變。

---

## 修改完成記錄

請新增或更新以下檔案：

畢業預警調整0525.md

內容請記錄：

````md
# 畢業預警調整0525

## 調整目標

調整畢業預警報表科目清單中 0 學分科目的輸出判斷。

## 調整內容

原本「非課規科目」即使學分數 = 0 仍會保留輸出。

本次調整為：

- 只要學分數 = 0，一律不輸出。
- 非課規科目也不例外。
- 學分數不是 0 的非課規科目，仍維持原本輸出邏輯。

## 修改檔案

- frmMain.cs

## 修改位置

- BgwGrandCheckReport_DoWork
- foreach (XmlElement xmlRuleS in xmlRule.SelectNodes("科目"))

## 修改前

```csharp
if (xmlRuleS.GetAttribute("學分數") == "0" && !isNonGraduationPlanSubject)
    continue;
````

## 修改後

```csharp
if (xmlRuleS.GetAttribute("學分數") == "0")
    continue;
```

## 影響範圍

* 影響畢業預警通知單中科目清單輸出。
* 不影響畢業審查計算。
* 不影響可重修、可補修、不採計的判斷。
* 不影響學分數不是 0 的非課規科目輸出。

## 驗證結果

* 一般 0 學分科目不輸出。
* 非課規 0 學分科目不輸出。
* 非課規且學分數大於 0 的科目仍可輸出。
* 可重修、可補修、不採計原本邏輯不變。

```
```

這次調整點就是 `frmMain.cs` 裡 `BgwGrandCheckReport_DoWork` 的 0 學分過濾條件。原本有排除非課規科目，現在改成學分數為 0 就直接 `continue`。
