
## Objective

Add checkbox-based control for exporting the standalone JSON file generated together with each sixth-semester report.

The checkbox name is:

```csharp
chkExportJSON
```

## Requirements

### 1. Control standalone JSON file output

Locate the existing JSON file generation code.

The current JSON backup file is generated in or near the following method:

```csharp
WriteGradeDataToPdf(string pdfPath, string studentID)
```

The standalone JSON file is currently written using logic similar to:

```csharp
File.WriteAllText(
    Path.ChangeExtension(pdfPath, ".json"),
    json,
    new UTF8Encoding(false)
);
```

Modify this logic so that the standalone `.json` file is generated only when:

```csharp
chkExportJSON.Checked == true
```

Expected behavior:

* When `chkExportJSON.Checked` is `true`, generate the JSON file.
* When `chkExportJSON.Checked` is `false`, do not generate the JSON file.
* The JSON filename must remain the same as the PDF filename, with the extension changed to `.json`.
* Continue using UTF-8 encoding without BOM.

Recommended implementation:

```csharp
if (chkExportJSON.Checked)
{
    File.WriteAllText(
        Path.ChangeExtension(pdfPath, ".json"),
        json,
        new UTF8Encoding(false)
    );
}
```

### 2. Do not affect PDF Metadata generation

The checkbox must control only the standalone `.json` backup file.

Do not change or disable the existing PDF Metadata process:

```csharp
fileInfo.SetMetaInfo("GradeData", json);
fileInfo.SaveNewInfo(pdfPath);
```

Expected behavior:

* PDF Metadata `GradeData` must always be written.
* Clearing `chkExportJSON` must not stop PDF generation.
* Clearing `chkExportJSON` must not stop PDF Metadata generation.

### 3. Default checkbox state

When the report form opens, `chkExportJSON` must be checked by default.

Set the default value to:

```csharp
chkExportJSON.Checked = true;
```

Prefer setting the default in the Designer control initialization:

```csharp
this.chkExportJSON.Checked = true;
this.chkExportJSON.CheckState = System.Windows.Forms.CheckState.Checked;
```

If the project convention initializes runtime defaults in the form load event, it may instead be set in:

```csharp
Student6thSemesterCorseCodeRank_Load
```

Avoid assigning `true` during every load if a previously saved user setting is expected to override the default.

### 4. Check existing setting persistence

Inspect whether this form already saves and restores checkbox or report-option settings.

Check:

* `Configure`
* `_Configure`
* `_Configure.Save()`
* `_Configure.Encode()`
* `_Configure.Decode()`
* Form load logic
* Existing settings for `chkNCredit`
* Existing settings for `chkNScore`
* Existing settings for `chkAccordingToClass`
* Any XML, UDT, registry, or local configuration storage used by this form

If an existing mechanism already persists report options, add `chkExportJSON` to the same mechanism.

The saved setting must support:

```csharp
chkExportJSON.Checked
```

Expected persistence behavior:

1. User changes the checkbox.
2. The value is saved using the existing settings mechanism.
3. The next time the form opens, the saved value is restored.
4. When no saved value exists, the default value is `true`.

Do not create a new and unrelated settings system if this form does not currently persist report-option checkboxes.

If `_Configure` only stores the Word template and does not store checkbox states, keep `chkExportJSON` as a default-checked runtime option and document that no existing checkbox persistence mechanism was found.

### 5. Thread safety

The report uses `BackgroundWorker`.

Do not directly read WinForms controls from the worker thread if avoidable.

Capture the checkbox value before calling:

```csharp
bgWorkerReport.RunWorkerAsync();
```

For example, add a form-level field:

```csharp
private bool _ExportJSON = true;
```

In `btnPrint_Click`, assign:

```csharp
_ExportJSON = chkExportJSON.Checked;
```

Then use the captured value when deciding whether to create the JSON file:

```csharp
if (_ExportJSON)
{
    File.WriteAllText(
        Path.ChangeExtension(pdfPath, ".json"),
        json,
        new UTF8Encoding(false)
    );
}
```

This ensures that the export process uses the value selected when the user starts the report.

### 6. Preserve existing behavior

Do not change unrelated report logic.

In particular, do not modify:

* Student selection logic
* Semester score queries
* Ranking calculation
* Course filtering
* Word generation
* PDF generation
* PDF filename rules
* JSON data structure
* `GradeData` validation rules
* PDF Metadata field name
* Existing exception handling
* Existing folder structure
* Existing XML or UDT storage structure
* Existing template storage behavior

## Test Cases

### Test 1: JSON export checked

1. Open the report form.
2. Confirm `chkExportJSON` is checked by default.
3. Generate the sixth-semester report.
4. Confirm the `.docx` file is generated.
5. Confirm the `.pdf` file is generated.
6. Confirm the same-name `.json` file is generated.
7. Confirm the PDF still contains the `GradeData` Metadata.

### Test 2: JSON export unchecked

1. Clear `chkExportJSON`.
2. Generate the report.
3. Confirm the `.docx` file is generated.
4. Confirm the `.pdf` file is generated.
5. Confirm no standalone `.json` file is generated.
6. Confirm the PDF still contains the `GradeData` Metadata.

### Test 3: Multiple students

1. Select multiple students.
2. Enable JSON export.
3. Generate the reports.
4. Confirm each PDF has a matching JSON file.
5. Disable JSON export and generate again.
6. Confirm no standalone JSON files are generated.

### Test 4: Alternative PDF save path

Force or test the `SaveFileDialog` fallback path.

Confirm:

* When JSON export is enabled, the JSON file is created beside the manually saved PDF.
* When JSON export is disabled, no JSON file is created.
* PDF Metadata is still written in both cases.

### Test 5: Setting persistence

If existing option persistence is available:

1. Clear `chkExportJSON`.
2. Close the form.
3. Reopen the form.
4. Confirm the saved unchecked state is restored.
5. Check the option again.
6. Reopen the form and confirm the checked state is restored.

If no existing checkbox persistence mechanism exists, confirm the checkbox opens as checked by default and document this finding.

## Completion Record

After completing the modification, create or update:

```text
第6學期成單調整.md
```

Record the following:

* Modified files
* Location of the original JSON generation code
* How `chkExportJSON` controls JSON output
* How the checkbox value is captured before starting the `BackgroundWorker`
* Confirmation that PDF Metadata output remains unchanged
* Whether an existing checkbox-setting persistence mechanism was found
* How the default checked state is implemented
* Test cases performed
* Test results
* Any unresolved issue or limitation

## Final Validation

Before finishing:

* Build the project.
* Confirm there are no compilation errors.
* Confirm the Designer event and control names remain correct.
* Confirm `chkExportJSON` exists and is not duplicated.
* Confirm disabling JSON output does not affect PDF Metadata.
* Confirm no unrelated code or storage structure was changed.
