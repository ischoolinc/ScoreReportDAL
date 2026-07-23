## Objective

Refactor the sixth-semester report export process so that the following operations are handled separately:

1. Build GradeData JSON content.
2. Write GradeData into PDF Metadata.
3. Export the standalone `.json` file.

Fix the current issue where `chkExportJSON.Checked == false` may still display an error message labeled as:

```text
第六學期成績單 Metadata 寫入失敗
```

The current checkbox controls only the standalone JSON file. PDF Metadata generation must still execute regardless of the checkbox state. However, failures must be reported according to the actual failing operation.

## Current Problem

The current method performs all operations together:

```csharp
private void WriteGradeDataToPdf(string pdfPath, string studentID)
```

The current flow is approximately:

```text
BuildGradeDataJson
    ↓
Write PDF Metadata
    ↓
Optionally write standalone JSON file
```

All exceptions are currently caught by the outer method and displayed as:

```text
第六學期成績單 Metadata 寫入失敗
```

This is misleading because the actual failure may occur while:

* Building GradeData JSON
* Validating GradeData
* Parsing credits or scores
* Extracting course codes
* Writing PDF Metadata
* Writing the standalone JSON file

The error message must clearly identify which step failed.

## Required Behavior

### When `chkExportJSON.Checked == true`

The report process must:

1. Generate the Word file.
2. Generate the PDF file.
3. Build GradeData JSON content.
4. Write GradeData into PDF Metadata.
5. Generate the standalone `.json` file.

### When `chkExportJSON.Checked == false`

The report process must:

1. Generate the Word file.
2. Generate the PDF file.
3. Build GradeData JSON content.
4. Write GradeData into PDF Metadata.
5. Skip only the standalone `.json` file.

Do not skip GradeData creation or PDF Metadata generation when JSON export is unchecked.

## Implementation Requirements

### 1. Keep the captured checkbox state

Keep the form-level field:

```csharp
private bool _ExportJSON = true;
```

Keep capturing the selected value before starting the report:

```csharp
private void btnPrint_Click(object sender, EventArgs e)
{
    UserControlEnable(false);
    SchoolYear = iptSchoolYear.Value;
    Semester = iptSemester.Value;
    IsAccordingToClass = chkAccordingToClass.Checked;
    _ExportJSON = chkExportJSON.Checked;
    bgWorkerReport.RunWorkerAsync();
}
```

Do not directly use `chkExportJSON.Checked` inside export-processing methods.

### 2. Separate GradeData generation

Create a dedicated method that only builds the JSON content.

The existing method may remain:

```csharp
private string BuildGradeDataJson(string studentID)
```

Do not change its business rules unless required to fix a confirmed bug.

Wrap its invocation separately so that its failures are identified as GradeData generation errors.

Recommended method:

```csharp
private string TryBuildGradeDataJson(string studentID)
{
    try
    {
        return BuildGradeDataJson(studentID);
    }
    catch (Exception ex)
    {
        throw new Exception(
            "GradeData 資料建立失敗：" + ex.Message,
            ex);
    }
}
```

A different method name is acceptable, but the responsibilities must remain separate.

### 3. Separate PDF Metadata writing

Create a method that only writes the prepared JSON content into PDF Metadata.

Recommended structure:

```csharp
private void WriteGradeDataMetadataToPdf(string pdfPath, string json)
{
    PdfFileInfo fileInfo = new PdfFileInfo();

    try
    {
        fileInfo.BindPdf(pdfPath);
        fileInfo.SetMetaInfo("GradeData", json);
        fileInfo.SaveNewInfo(pdfPath);
    }
    catch (Exception ex)
    {
        throw new Exception(
            "PDF Metadata 寫入失敗：" + ex.Message,
            ex);
    }
    finally
    {
        fileInfo.Close();
    }
}
```

Requirements:

* Keep the Metadata field name as `GradeData`.
* Continue using `SetMetaInfo`.
* Continue using `SaveNewInfo`.
* Do not switch to `SaveNewInfoWithXmp`.
* Do not change the existing PDF Metadata format.
* Ensure `PdfFileInfo.Close()` is always called.

### 4. Separate standalone JSON file writing

Create a method that only writes the `.json` file.

Recommended structure:

```csharp
private void WriteGradeDataJsonFile(string pdfPath, string json)
{
    try
    {
        File.WriteAllText(
            Path.ChangeExtension(pdfPath, ".json"),
            json,
            new UTF8Encoding(false));
    }
    catch (Exception ex)
    {
        throw new Exception(
            "JSON 檔案輸出失敗：" + ex.Message,
            ex);
    }
}
```

Requirements:

* Keep the JSON filename identical to the PDF filename.
* Only change the file extension to `.json`.
* Continue using UTF-8 without BOM.
* This method must be called only when `_ExportJSON == true`.

### 5. Refactor the main GradeData processing method

Refactor the current `WriteGradeDataToPdf` method so it coordinates the three separate operations.

Recommended structure:

```csharp
private void WriteGradeDataToPdf(string pdfPath, string studentID)
{
    string json = TryBuildGradeDataJson(studentID);

    WriteGradeDataMetadataToPdf(pdfPath, json);

    if (_ExportJSON)
    {
        WriteGradeDataJsonFile(pdfPath, json);
    }
}
```

The method names may differ, but the execution order must remain:

1. Build GradeData JSON.
2. Write PDF Metadata.
3. Conditionally write standalone JSON.

### 6. Correct the outer error message

The current outer method uses a fixed Metadata-related error message.

Modify:

```csharp
private void TryWriteGradeDataToPdf(
    string pdfPath,
    string studentID,
    List<string> gradeMetaErrorList)
```

Do not prepend another fixed `Metadata 寫入失敗` message to every exception.

Recommended implementation:

```csharp
private void TryWriteGradeDataToPdf(
    string pdfPath,
    string studentID,
    List<string> gradeDataErrorList)
{
    try
    {
        WriteGradeDataToPdf(pdfPath, studentID);
    }
    catch (Exception ex)
    {
        string idNumber =
            StudentDocNameDict.ContainsKey(studentID)
            ? StudentDocNameDict[studentID]
            : "";

        gradeDataErrorList.Add(string.Format(
            "第六學期成績資料處理失敗（學生系統編號 {0}，身分證號 {1}）：{2}",
            studentID,
            idNumber,
            ex.Message));
    }
}
```

Expected messages must include the actual failing stage, such as:

```text
第六學期成績資料處理失敗（學生系統編號 123，身分證號 A123456789）：GradeData 資料建立失敗：單科學分數無法轉為整數……
```

```text
第六學期成績資料處理失敗（學生系統編號 123，身分證號 A123456789）：PDF Metadata 寫入失敗：……
```

```text
第六學期成績資料處理失敗（學生系統編號 123，身分證號 A123456789）：JSON 檔案輸出失敗：……
```

### 7. Rename the error list for clarity

Rename:

```csharp
gradeMetaErrorList
```

to a more accurate name, such as:

```csharp
gradeDataErrorList
```

Update all related references.

This list may contain:

* GradeData build errors
* PDF Metadata errors
* JSON file output errors

### 8. Correct the final message box title

Change the current title:

```text
第六學期成績單 Metadata 寫入失敗
```

to a broader title:

```text
第六學期成績資料處理失敗
```

Recommended code:

```csharp
if (gradeDataErrorList.Count > 0)
{
    MsgBox.Show(
        string.Join(Environment.NewLine, gradeDataErrorList),
        "第六學期成績資料處理失敗",
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning);
}
```

### 9. Verify placeholder replacement

Confirm that the displayed message does not literally contain:

```text
{0}
{1}
{2}
```

The final message must contain the real values:

* Student system ID
* Student ID number
* Actual exception message

Use `string.Format(...)` correctly or use string interpolation.

Example using interpolation:

```csharp
gradeDataErrorList.Add(
    $"第六學期成績資料處理失敗（學生系統編號 {studentID}，身分證號 {idNumber}）：{ex.Message}");
```

Do not leave unformatted placeholders in the user-facing message.

## Error Handling Rules

### GradeData build error

If `BuildGradeDataJson` fails:

* Record a GradeData build error.
* Do not attempt PDF Metadata writing.
* Do not attempt JSON file writing.
* Continue processing other students.

### PDF Metadata error

If PDF Metadata writing fails:

* Record a PDF Metadata error.
* Do not write the standalone JSON file for that student unless the existing business requirement explicitly allows it.
* Recommended behavior: stop processing GradeData output for that student.
* Continue processing other students.

### JSON file output error

If standalone JSON writing fails:

* Record a JSON file output error.
* Keep the generated PDF.
* Keep the PDF Metadata already written.
* Continue processing other students.

## Preserve Existing Logic

Do not modify unrelated behavior.

Do not change:

* Student data queries
* Semester score queries
* Ranking calculations
* Course filtering rules
* Word generation
* PDF generation
* PDF filenames
* JSON schema
* GradeData property names
* GradeData validation limits
* Course code extraction rules
* `GradeData` Metadata field name
* Report folder structure
* Existing template storage
* Existing XML or UDT structure
* Existing checkbox persistence behavior
* Existing `chkNCredit` and `chkNScore` behavior

## Test Cases

### Test 1: JSON export enabled

1. Open the report form.
2. Check `chkExportJSON`.
3. Generate a valid report.
4. Confirm Word is generated.
5. Confirm PDF is generated.
6. Confirm PDF contains `GradeData` Metadata.
7. Confirm the matching `.json` file is generated.
8. Confirm no error message is displayed.

### Test 2: JSON export disabled

1. Uncheck `chkExportJSON`.
2. Generate a valid report.
3. Confirm Word is generated.
4. Confirm PDF is generated.
5. Confirm PDF contains `GradeData` Metadata.
6. Confirm no standalone `.json` file is generated.
7. Confirm no JSON-related error message is displayed.

### Test 3: GradeData validation failure

Use or simulate invalid GradeData, such as:

* Invalid credit value
* Missing ranking
* Invalid course code
* No valid subject data

Confirm the error message starts with:

```text
GradeData 資料建立失敗
```

Confirm it is not incorrectly labeled only as a PDF Metadata failure.

### Test 4: PDF Metadata failure

Use an invalid or inaccessible PDF path, or simulate a Metadata write failure.

Confirm the message starts with:

```text
PDF Metadata 寫入失敗
```

### Test 5: JSON file output failure

Enable JSON export and use a location where the `.json` file cannot be written.

Confirm:

* The PDF remains generated.
* The PDF Metadata remains written.
* The error message starts with:

```text
JSON 檔案輸出失敗
```

### Test 6: Multiple students

Generate reports for multiple students where one student has invalid GradeData.

Confirm:

* The invalid student's error is recorded.
* Other students continue processing.
* Each error includes the correct student system ID and ID number.
* No `{0}`, `{1}`, or `{2}` placeholders remain in the displayed text.

### Test 7: SaveFileDialog fallback

Test the alternative PDF save path.

Confirm:

* Metadata is written to the manually selected PDF.
* JSON is written beside the selected PDF only when `_ExportJSON == true`.
* Errors are categorized correctly.

## Completion Record

After completing the code modification, create or update:

```text
第6學期成單調整.md
```

Record:

1. Modified files.
2. Original cause of the misleading error message.
3. Methods created or refactored.
4. How GradeData generation is separated.
5. How PDF Metadata writing is separated.
6. How standalone JSON output is separated.
7. How `chkExportJSON` affects only the standalone JSON file.
8. Updated error-list variable name.
9. Updated message box title.
10. Confirmation that `{0}`, `{1}`, and `{2}` are replaced with actual values.
11. Test cases executed.
12. Test results.
13. Build result.
14. Any unresolved issue or limitation.

## Final Validation

Before finishing:

* Build the project successfully.
* Confirm there are no compilation errors.
* Confirm the PDF is still generated when JSON export is disabled.
* Confirm PDF Metadata is still generated when JSON export is disabled.
* Confirm no `.json` file is generated when JSON export is disabled.
* Confirm error messages identify the correct processing stage.
* Confirm processing continues for other students after one student's failure.
* Confirm no unrelated program logic was changed.
