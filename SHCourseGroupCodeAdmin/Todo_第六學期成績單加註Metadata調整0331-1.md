# Todo.md

目標：修正 `Student6thSemesterCorseCodeRank.cs` 在 PDF 產出後寫入 `GradeData` Metadata 時，發生
`Unknown prefix. Need to call method RegisterNamespaceURI`
的問題。
做法：**不要走 XMP prefix / namespace 寫法**，改用 **Aspose.PDF 8.8.0 的 `PdfFileInfo.SetMetaInfo("GradeData", json)` + `SaveNewInfo(...)`** 寫入一般自訂 Metadata。
注意：**不要更動原本 Word / PDF 產出流程與報表內容邏輯**，只修正 Metadata 寫入方式，並保留同名 `.json` 備份檔輸出。
修改完成後，記錄在：`第六學期成績單加註Metadata調整0331.md`

---

## 一、修改檔案

* `Student6thSemesterCorseCodeRank.cs`

---

## 二、修改重點

1. 保留原本：

   * `.docx` 產出
   * `.PDF` 產出
   * 資料夾建立
   * 依班級建立子資料夾
   * MailMerge 與報表欄位填值邏輯
2. 修正 Metadata 寫入方式：

   * 使用 `PdfFileInfo.SetMetaInfo("GradeData", json)`
   * 使用 `SaveNewInfo(pdfPath)`
   * **不要使用** `SaveNewInfoWithXmp(...)`
   * **不要使用** `Document.Metadata[...]`
   * **不要使用** 帶 prefix 的 metadata key，例如 `dc:GradeData`、`xmp:GradeData`
3. 保留同名 `.json` 備份檔輸出
4. 單一學生 Metadata 失敗時，不影響其他學生繼續輸出

---

## 三、確認 class-level 欄位

若前一版已做可沿用；若尚未做，請補上：

```csharp
private List<rptStudSemsScoreCodeChkInfo> _ResultList = new List<rptStudSemsScoreCodeChkInfo>();
private Dictionary<string, int> _CourseStudentCountDic = new Dictionary<string, int>();
```

用途：

* `_ResultList`：供 PDF 產出後建立該學生的 GradeData JSON
* `_CourseStudentCountDic`：供建立 JSON 時填入 `修課人數`

---

## 四、在 `BgWorkerReport_DoWork` 改用 class-level 結果

若目前仍是區域變數，請改成：

* 開始前先 `_ResultList.Clear();`
* 開始前先 `_CourseStudentCountDic.Clear();`

將原本：

```csharp
List<rptStudSemsScoreCodeChkInfo> ResultList = new List<rptStudSemsScoreCodeChkInfo>();
Dictionary<string, int> courseStudentCountDic = new Dictionary<string, int>();
```

改為 class-level 欄位版本。

後續所有：

* `ResultList`
* `courseStudentCountDic`

改成：

* `_ResultList`
* `_CourseStudentCountDic`

注意：

* MailMerge 內容邏輯也要同步改為使用 `_ResultList` / `_CourseStudentCountDic`
* 不要只改 Metadata 區塊，避免後面抓不到資料

---

## 五、新增 using

請確認檔案上方有：

```csharp
using Aspose.Pdf.Facades;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
```

注意：

* `JavaScriptSerializer` 需要 `System.Web.Extensions`
* 若專案沒參考，要補參考
* Aspose.PDF 需確認為 8.8.0 並可正常編譯

---

## 六、新增 Metadata 專用類別

在 `Student6thSemesterCorseCodeRank` 類別內新增：

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

## 七、新增驗證方法

新增：

```csharp
private void ValidateGradeMetaItem(GradeMetaItem item)
```

檢核規則：

1. `科目名稱`

   * 不可空白
   * 長度 1~200
2. `單科學分數`

   * 1~9
3. `單科成績`

   * 0~100
4. `修課人數`

   * 1~9999
5. `單科成績排名百分比`

   * 0~100
   * 最多小數第 2 位
6. `課程類別代碼`

   * 必須符合 `^[1-9A-F]$`
7. `領域名稱代碼`

   * 不可空白
   * 長度固定 2

不合法時：

* `throw new Exception("欄位XXX格式錯誤：...")`

---

## 八、新增輔助方法

### 1. 取得修課人數

新增：

```csharp
private int GetCourseStudentCount(string courseCode)
{
    if (string.IsNullOrWhiteSpace(courseCode))
        return 0;

    if (_CourseStudentCountDic.ContainsKey(courseCode))
        return _CourseStudentCountDic[courseCode];

    return 0;
}
```

---

### 2. 取得課程類別代碼

新增：

```csharp
private string GetCourseCategoryCode(rptStudSemsScoreCodeChkInfo data)
```

要求：

* 從 `data.CourseCode` 依正式規則解析
* 回傳 1 碼
* 僅允許 `1~9` 或 `A~F`
* 若無法取得，直接丟例外

方法內加註解：

```csharp
// TODO: 依正式課程代碼規則取得課程類別代碼
```

---

### 3. 取得領域名稱代碼

新增：

```csharp
private string GetDomainCode(rptStudSemsScoreCodeChkInfo data)
```

要求：

* 從 `data.CourseCode` 依正式規則解析
* 回傳固定 2 碼
* 若無法取得，直接丟例外

方法內加註解：

```csharp
// TODO: 依正式課程代碼規則取得領域名稱代碼
```

---

## 九、新增建立 JSON 方法

新增：

```csharp
private string BuildGradeDataJson(string studentID)
```

邏輯：

1. 從 `_ResultList` 篩出該學生資料
2. 只保留：

   * `CourseCode` 有值
   * `IsStudying == false`
   * `Score.HasValue == true`
3. 套用目前報表畫面已存在的篩選規則：

   * `不計學分`
   * `不需評分`
   * 必須與原本列印內容一致
   * 不要另外發明新規則
4. 每筆建立 `GradeMetaItem`
5. `單科成績排名百分比`

   * 寫數值，不加 `%`
6. `修課人數`

   * 用 `GetCourseStudentCount(data.CourseCode)`
7. `單科成績`

   * 必須轉成整數
8. 逐筆 `ValidateGradeMetaItem(item)`
9. 至少要有 1 筆

   * 沒資料就丟例外
10. 用 `JavaScriptSerializer` 轉成 JSON Array 回傳

注意：

* JSON 欄位名稱必須是繁中
* 不可多欄位
* 不可寫 `修課中`
* 不可寫 `%`

---

## 十、新增 Metadata 寫入方法（修正版）

新增：

```csharp
private void WriteGradeDataToPdf(string pdfPath, string studentID)
```

請用下面邏輯實作：

```csharp
private void WriteGradeDataToPdf(string pdfPath, string studentID)
{
    string json = BuildGradeDataJson(studentID);

    PdfFileInfo fileInfo = new PdfFileInfo();
    try
    {
        fileInfo.BindPdf(pdfPath);

        // 寫入自訂 Metadata 欄位
        fileInfo.SetMetaInfo("GradeData", json);

        // 修正 Unknown prefix 問題：
        // 不使用 SaveNewInfoWithXmp，避免 XMP namespace/prefix 造成錯誤
        fileInfo.SaveNewInfo(pdfPath);
    }
    catch (Exception ex)
    {
        throw new Exception("第六學期成績單 Metadata 寫入失敗：" + ex.Message, ex);
    }
    finally
    {
        fileInfo.Close();
    }

    // 同步輸出同名 JSON 備份檔
    File.WriteAllText(
        Path.ChangeExtension(pdfPath, ".json"),
        json,
        new UTF8Encoding(false)
    );
}
```

### 重要限制

1. **不要使用**

   ```csharp
   fileInfo.SaveNewInfoWithXmp(...)
   ```
2. **不要使用**

   ```csharp
   pdfDocument.Metadata["..."]
   ```
3. **不要使用帶 prefix 的 key**

   * `dc:GradeData`
   * `xmp:GradeData`
   * 任何含 `:` 的 metadata 名稱
4. Metadata key 一律固定：

   ```csharp
   "GradeData"
   ```

---

## 十一、在 PDF 成功儲存後補寫 Metadata

找到 `BgWorkerReport_RunWorkerCompleted` 內 PDF 區塊：

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

* 一定要在 PDF 成功存出後才呼叫
* 不要加在 Word 區塊
* 不要改原本命名與另存處理

---

## 十二、另存新檔流程也要補

在 PDF `SaveFileDialog` 另存成功後：

原本：

```csharp
document.Save(sd.FileName, Aspose.Words.SaveFormat.Pdf);
```

改成：

```csharp
document.Save(sd.FileName, Aspose.Words.SaveFormat.Pdf);
WriteGradeDataToPdf(sd.FileName, sid);
```

避免使用者選擇其他路徑時，PDF 有存出來但沒有 Metadata

---

## 十三、單一學生失敗不影響整批

在單一學生 PDF 區塊中，請將 Metadata 寫入失敗視為「該學生失敗」而不是整批失敗。

建議作法：

* 建立 `List<string> errorMsgList = new List<string>();`
* 若某學生 `WriteGradeDataToPdf(...)` 發生錯誤：

  * 將錯誤記錄到 `errorMsgList`
  * 繼續下一位學生
* 全部完成後若有錯誤，再一次顯示

錯誤訊息建議至少含：

* 學生識別（例如身份證號 / reportNameSingle）
* 錯誤原因

---

## 十四、確認不要再出現 Unknown prefix

請全面檢查這支檔案是否還有以下寫法，若有一律移除或改寫：

### 不可再使用

```csharp
SaveNewInfoWithXmp(...)
```

```csharp
Document.Metadata[...]
```

```csharp
pdfDocument.Metadata[...]
```

```csharp
SetMetaInfo("dc:GradeData", ...)
SetMetaInfo("xmp:GradeData", ...)
```

```csharp
任何含冒號 prefix 的 metadata key
```

---

## 十五、測試重點

### 測試 1：正常單一學生

* 產出 PDF 成功
* 同資料夾有同名 `.json`
* Acrobat 可看到 `GradeData`

### 測試 2：多位學生批次

* 每位學生各自產出 PDF / JSON
* Metadata 不互相覆蓋
* 若其中一位失敗，其餘仍完成

### 測試 3：不再出現錯誤

* 不再出現：
  `Unknown prefix. Need to call method RegisterNamespaceURI`

### 測試 4：修課中資料

* 不應寫進 Metadata

### 測試 5：課程代碼空白

* 不應寫進 Metadata

### 測試 6：另存新檔

* 使用另存新檔時，Metadata 與 `.json` 仍正確產出

---

## 十六、自我檢查

1. `.docx` 是否仍正常輸出
2. `.PDF` 是否仍正常輸出
3. `GradeData` 是否用 `SetMetaInfo("GradeData", json)` 寫入
4. 是否改成 `SaveNewInfo(...)`
5. 是否已移除 `SaveNewInfoWithXmp(...)`
6. 是否沒有任何 prefix metadata key
7. 是否同步輸出同名 `.json`
8. 是否未更動原本報表邏輯
9. 是否單一學生錯誤不會中斷整批

---

## 十七、完成紀錄

修改完成後，新增紀錄檔：

`第六學期成績單加註Metadata調整0331.md`

內容至少包含：

1. 問題現象

   * `Unknown prefix. Need to call method RegisterNamespaceURI`
2. 原因

   * 舊版 Aspose.PDF 寫入 Metadata 時碰到 XMP namespace/prefix 問題
3. 修正方式

   * 改用 `SetMetaInfo("GradeData", json)` + `SaveNewInfo(...)`
4. 修改檔案
5. 新增方法
6. 測試結果
7. 仍需確認事項

   * 課程類別代碼取法
   * 領域名稱代碼取法
