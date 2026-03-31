# Todo.md

目標：在 `Student6thSemesterCorseCodeRank.cs` 產生 PDF 後，使用 **Aspose.PDF 8.8.0** 將第六學期成績資料寫入 PDF 自訂 Metadata 欄位 `GradeData`，並同步產出同名 `.json` 備份檔。
注意：**不要更動原本 Word / PDF 產出邏輯與畫面操作流程**，只在既有流程後補上 Metadata 寫入。
修改完成後，記錄在：`第六學期成績單加註Metadata調整0331.md`

---

## 一、修改檔案

* `Student6thSemesterCorseCodeRank.cs`

---

## 二、修改原則

1. 保留原本：

   * Word 輸出 `.docx`
   * PDF 輸出 `.PDF`
   * 開啟輸出資料夾
   * 班級分資料夾邏輯
   * 報表欄位與套印流程
2. 僅新增：

   * PDF 產出後寫入 `GradeData` Metadata
   * 產出同名 `.json`
   * 必要的資料整理與驗證方法
3. 不要破壞原本背景執行流程 `BackgroundWorker`
4. 若 Metadata 寫入失敗，要能顯示明確錯誤訊息
5. 若單一學生 PDF 已成功產出，但 Metadata 寫入失敗，不要影響其他學生繼續產出

---

## 三、先補 class-level 欄位

### 1. 將結果清單提升為 class-level

目前 `ResultList` 是 `BgWorkerReport_DoWork` 內區域變數，後續 PDF 寫入 Metadata 時會用不到。
請新增 class-level 欄位：

```csharp
private List<rptStudSemsScoreCodeChkInfo> _ResultList = new List<rptStudSemsScoreCodeChkInfo>();
private Dictionary<string, int> _CourseStudentCountDic = new Dictionary<string, int>();
```

### 2. 在 `BgWorkerReport_DoWork` 內改用 class-level

原本：

```csharp
List<rptStudSemsScoreCodeChkInfo> ResultList = new List<rptStudSemsScoreCodeChkInfo>();
Dictionary<string, int> courseStudentCountDic = new Dictionary<string, int>();
```

改為：

* 清空 `_ResultList`
* 清空 `_CourseStudentCountDic`
* 後續合併結果、計算修課人數都寫入這兩個 class-level 欄位

例如：

```csharp
_ResultList.Clear();
_CourseStudentCountDic.Clear();
```

並將原本：

```csharp
ResultList.AddRange(...)
```

改成：

```csharp
_ResultList.AddRange(...)
```

將原本：

```csharp
courseStudentCountDic[data.CourseCode] = ...
```

改成：

```csharp
_CourseStudentCountDic[data.CourseCode] = ...
```

### 3. 原本填 MailMerge 資料時，若有用到 `ResultList`、`courseStudentCountDic`

全部改成：

* `_ResultList`
* `_CourseStudentCountDic`

---

## 四、新增 using

在檔案上方新增：

```csharp
using Aspose.Pdf.Facades;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
```

注意：

* 若專案尚未引用 `Aspose.Pdf`
* 需確認專案已有 `Aspose.PDF 8.8.0`
* `System.Web.Extensions` 若未引用，要補上，因為 `JavaScriptSerializer` 會用到

---

## 五、新增 Metadata 專用資料類別

在 `Student6thSemesterCorseCodeRank` 類別內，新增一個內部 class：

```csharp
private class GradeMetaItem
{
    public string 科目名稱 { get; set; }
    public int 單科學分數 { get; set; }
    public int 單科成績 { get; set; }
    public int 修課人數 { get; set; }
    public decimal 單科成績排名百分比 { get; set; }
    public string 課程類別代碼 { get; set; }
    public string 領域名稱代碼 { get; set; }
}
```

---

## 六、新增驗證方法

新增方法：`ValidateGradeMetaItem(GradeMetaItem item)`

驗證規則：

1. `科目名稱`

   * 不可空白
   * 長度 1~200
2. `單科學分數`

   * 整數
   * 範圍 1~9
3. `單科成績`

   * 整數
   * 範圍 0~100
4. `修課人數`

   * 整數
   * 範圍 1~9999
5. `單科成績排名百分比`

   * 範圍 0~100
   * 最多小數第 2 位
6. `課程類別代碼`

   * 必須符合 `^[1-9A-F]$`
   * 只能是 1~9 或大寫 A~F
7. `領域名稱代碼`

   * 固定 2 碼
   * 不可空白

若不合法：

* 直接 `throw new Exception("...")`
* 訊息要能指出哪個欄位錯

---

## 七、新增輔助方法

### 1. 取得修課人數

新增：

```csharp
private int GetCourseStudentCount(string courseCode)
```

邏輯：

* 若 `courseCode` 空白，回傳 0
* 若 `_CourseStudentCountDic` 有資料，回傳對應值
* 否則回傳 0

---

### 2. 取得課程類別代碼

新增：

```csharp
private string GetCourseCategoryCode(rptStudSemsScoreCodeChkInfo data)
```

先用「課程代碼字串解析」方式處理。
請依你們目前第六學期成績單規則，從 `data.CourseCode` 取出 **課程類別代碼**。

注意：

* 文件要求是 1 碼
* 只能是 `1~9` 或 `A~F`
* 若無法正確取得，先丟出例外，不要默默帶空值

請先在程式內加註解，說明「課程類別代碼取法需依現有課程代碼規則確認」

例如：

```csharp
// TODO: 依課程代碼正式規則取得課程類別代碼
```

---

### 3. 取得領域名稱代碼

新增：

```csharp
private string GetDomainCode(rptStudSemsScoreCodeChkInfo data)
```

先用「課程代碼字串解析」方式處理。
請依你們目前第六學期成績單規則，從 `data.CourseCode` 取出 **領域名稱代碼**。

注意：

* 文件要求固定 2 碼
* 若無法正確取得，先丟出例外，不要默默帶空值

請先在程式內加註解，說明「領域名稱代碼取法需依現有課程代碼規則確認」

例如：

```csharp
// TODO: 依課程代碼正式規則取得領域名稱代碼
```

---

### 4. 建立單一學生的 GradeData JSON 內容

新增：

```csharp
private string BuildGradeDataJson(string studentID)
```

邏輯：

1. 從 `_ResultList` 篩出該學生資料
2. 只保留：

   * 有 `CourseCode`
   * 不是 `修課中`
   * 有 `Score`
3. 依原本第六學期報表顯示邏輯，套用不計學分 / 不需評分篩選規則
   注意：**不要另發明新規則**，要與目前報表畫面輸出邏輯一致
4. 每筆轉成 `GradeMetaItem`
5. `單科成績排名百分比`

   * 使用數值
   * 不要加 `%`
   * 例如原本報表顯示 `33%`，Metadata 寫成 `33`
6. `修課人數`

   * 用 `_CourseStudentCountDic`
7. `單科成績`

   * 必須是整數
8. 每筆資料都先呼叫 `ValidateGradeMetaItem(item)`
9. 至少要有 1 筆

   * 若 0 筆，丟出例外
10. 用 `JavaScriptSerializer` 序列化成 JSON Array 字串後回傳

注意：

* Metadata JSON 欄位名要使用繁體中文
* 不可加入額外欄位

---

### 5. 寫入 PDF Metadata

新增：

```csharp
private void WriteGradeDataToPdf(string pdfPath, string studentID)
```

邏輯：

1. 呼叫 `BuildGradeDataJson(studentID)` 取得 JSON
2. 使用 `PdfFileInfo`
3. `BindPdf(pdfPath)`
4. `SetMetaInfo("GradeData", json)`
5. `SaveNewInfoWithXmp(pdfPath)`
6. `Close()`
7. 再同步輸出同名 `.json`

   * 檔名與 PDF 主檔名相同，只改副檔名 `.json`
   * UTF-8 編碼
8. 若發生錯誤，丟出例外或包裝清楚訊息：

   * 例如：`第六學期成績單 Metadata 寫入失敗：...`

---

## 八、在 PDF 產生後插入 Metadata 寫入

找到 `BgWorkerReport_RunWorkerCompleted` 內 PDF 產生區塊：

原本：

```csharp
document.Save(pathPDF, SaveFormat.Pdf);
```

改成：

```csharp
document.Save(pathPDF, SaveFormat.Pdf);
WriteGradeDataToPdf(pathPDF, sid);
```

注意：

* 插在 `Save PDF` 成功之後
* 不要插在 Word 儲存區塊
* 不要改掉原本另存新檔流程

---

## 九、另存新檔流程也要補 Metadata

在 PDF `SaveFileDialog` 成功後，原本有：

```csharp
document.Save(sd.FileName, Aspose.Words.SaveFormat.Pdf);
```

在這行後面也補：

```csharp
WriteGradeDataToPdf(sd.FileName, sid);
```

避免使用者改存其他路徑時，PDF 有存出來但沒寫 Metadata

---

## 十、避免影響其他學生輸出

在單一學生 PDF 區塊中，若 `WriteGradeDataToPdf(...)` 失敗：

* 要顯示清楚錯誤
* 但不要讓整批流程整個中斷
* 建議單一學生失敗時記錄訊息，繼續下一位學生

可考慮：

* 蒐集 `errorMsgList`
* 最後一次顯示全部失敗學生與錯誤原因

若目前程式風格不適合集中顯示，也至少要 `MsgBox.Show(...)` 說明是哪位學生失敗

---

## 十一、JSON 產製規則要與文件一致

請確認輸出的 `GradeData` JSON 符合下列規格：

每筆固定 7 欄：

* `科目名稱`
* `單科學分數`
* `單科成績`
* `修課人數`
* `單科成績排名百分比`
* `課程類別代碼`
* `領域名稱代碼`

格式要求：

* 最外層是 `[]`
* `單科學分數`、`單科成績`、`修課人數` 為整數
* `單科成績排名百分比` 為 number，可含小數，最多 2 位
* `課程類別代碼` 僅允許 1~9 / A~F
* `領域名稱代碼` 固定 2 碼
* 欄位名稱不可改英文
* 不可加入額外欄位
* JSON 內容與同名 `.json` 檔內容必須一致

---

## 十二、需自行確認的重要點

### 1. 課程類別代碼與領域名稱代碼取法

這是本次最重要的實作細節。
請依目前系統 `CourseCode` 規則確認：

* 課程類別代碼取哪一碼
* 領域名稱代碼取哪兩碼

若規則未確認，不要硬寫死錯誤位置。

---

### 2. 修課中資料是否要排除

因文件欄位 `單科成績` 要求整數 0~100，`修課中` 不是合法值。
因此 Metadata 應排除 `IsStudying == true` 的資料。

---

### 3. 排名百分比格式

畫面上目前是 `Rank + "%"`.
但 Metadata 必須寫數值：

* 畫面：`33%`
* Metadata：`33`

---

### 4. 不計學分 / 不需評分篩選

Metadata 內容要與實際報表顯示一致。
請重用目前輸出報表時已經存在的篩選判斷，不要另外發明一套不同條件。

---

## 十三、建議測試案例

請至少測：

### 測試 1：正常 1 位學生

* 有 3~5 筆科目
* 成功產出 PDF
* Acrobat 可看到 `GradeData`
* 同資料夾有同名 `.json`

### 測試 2：多位學生批次

* 每位都能產出 PDF 與 `.json`
* Metadata 不會互相覆蓋

### 測試 3：含修課中資料

* `修課中` 不應進入 Metadata JSON

### 測試 4：課程代碼空白

* 不應進入 Metadata JSON

### 測試 5：課程類別代碼不合法

* 應拋出錯誤，不可寫入不合法資料

### 測試 6：領域名稱代碼長度錯誤

* 應拋出錯誤

### 測試 7：同名另存

* 若 PDF 改名另存，Metadata 與 `.json` 仍要正確產出

---

## 十四、完成後自我檢查

1. 是否仍能正常產出 `.docx`
2. 是否仍能正常產出 `.PDF`
3. PDF 產出後是否都有呼叫 `WriteGradeDataToPdf`
4. 是否同步產出同名 `.json`
5. `GradeData` 是否出現在 Acrobat 的自訂 Metadata
6. JSON 是否為陣列 `[]`
7. 是否只有 7 個欄位，且欄位名為繁中
8. `單科成績排名百分比` 是否未帶 `%`
9. 是否未把 `修課中` 寫入 Metadata
10. 是否未改壞原本第六學期報表功能

---

## 十五、完成紀錄

修改完成後，新增紀錄檔：

`第六學期成績單加註Metadata調整0331.md`

內容至少包含：

1. 修改目的
2. 修改檔案
3. 新增方法名稱
4. Metadata 寫入位置
5. JSON 備份檔產出規則
6. 課程類別代碼 / 領域名稱代碼取法
7. 測試結果
8. 風險與注意事項
