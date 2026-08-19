## Goal

Modify the curriculum plan generation logic so that:

> Any `<Subject>` merged from another curriculum plan, identified by a non-empty `來源課規` attribute, must not be treated as `更新` or `刪除` by the course-code master table comparison.

Such subjects must be forced to:

```text
ProcessStatus = "略過"
```

Important:

* Do not change any unrelated logic.
* Do not prevent the final curriculum plan from recalculating `RowIndex`, `Level`, `FullName`, or `startLevel`.
* Only change the comparison/status decision for Subjects that contain a valid `來源課規` value.
* Keep the current overall architecture and XML generation flow unchanged.

After implementation and testing, record the changes in:

`產生課程規劃調整0819.md`

---

## Target Files

Primary modification:

```text
GPlanInfo108.cs
```

Main method:

```csharp
GPlanInfo108.CheckData()
```

Reference only, avoid modifying unless absolutely necessary:

```text
frmCreateGPlanMainBatch.cs
```

---

## Current Behavior

`GPlanInfo108.CheckData()` compares:

```text
MOEXml
```

with:

```text
RefGPContentXml
```

and groups Subjects mainly by:

```text
課程代碼
```

using:

```csharp
MOEDict
GPlanDict
```

The current logic determines Subject status such as:

```text
新增
更新
刪除
略過
```

### Add

When the course-code master table contains a course code but the existing curriculum plan does not:

```text
MOEDict has CourseCode
GPlanDict does not
```

the Subject is marked:

```csharp
subj.ProcessStatus = "新增";
```

### Delete

When the existing curriculum plan contains a course code but the course-code master table does not:

```text
GPlanDict has CourseCode
MOEDict does not
```

the Subject is currently marked:

```csharp
subj.ProcessStatus = "刪除";
```

### Update / Skip

When both sides contain the same course code, the program compares Subject properties.

If differences exist:

```csharp
subj.ProcessStatus = "更新";
```

Otherwise:

```csharp
subj.ProcessStatus = "略過";
```

---

# Required New Rule

Add the following priority rule:

```text
If an existing curriculum-plan Subject contains
來源課規 != null / empty / whitespace

then:

ProcessStatus = "略過"
```

This rule has higher priority than:

```text
更新
刪除
```

Expected priority:

```text
來源課規有值
    ↓
強制略過
    ↓
Do not update from MOEXml
Do not delete from final curriculum plan
```

If `來源課規` does not exist or is empty, preserve the existing behavior.

---

# 1. Add a Shared Helper Method

Add a small helper method inside `GPlanInfo108`.

Suggested implementation:

```csharp
/// <summary>
/// Check whether any Subject in the existing curriculum plan
/// was merged from another curriculum plan.
/// </summary>
private bool HasSourceGPlan(List<XElement> elements)
{
    if (elements == null)
        return false;

    foreach (XElement elm in elements)
    {
        XAttribute attr = elm.Attribute("來源課規");

        if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
            return true;
    }

    return false;
}
```

Expected behavior:

```xml
<Subject ... />
```

Result:

```text
false
```

---

```xml
<Subject 來源課規="" ... />
```

Result:

```text
false
```

---

```xml
<Subject 來源課規="   " ... />
```

Result:

```text
false
```

---

```xml
<Subject 來源課規="113普通科課規" ... />
```

Result:

```text
true
```

Do not use only:

```csharp
elm.Attribute("來源課規") != null
```

because an empty attribute must not trigger the protection rule.

---

# 2. Protect Subjects Currently Marked for Deletion

Locate the section:

```csharp
// 大表沒有，課程規劃表有 多出來
foreach (string mCo in DelList)
```

Current behavior eventually sets:

```csharp
subj.ProcessStatus = "刪除";
subj.DiffStatusList.Add("多");
subj.GPlanXml = GPlanDict[mCo];
```

Modify this logic so that the original curriculum-plan XML is assigned first:

```csharp
subj.GPlanXml = GPlanDict[mCo];
```

Then check:

```csharp
HasSourceGPlan(subj.GPlanXml)
```

Expected logic:

```csharp
subj.GPlanXml = GPlanDict[mCo];

if (HasSourceGPlan(subj.GPlanXml))
{
    subj.ProcessStatus = "略過";
}
else
{
    subj.ProcessStatus = "刪除";
    subj.DiffStatusList.Add("多");
}
```

Important:

If `來源課規` has a value:

```text
Do not add "多" to DiffStatusList.
Do not mark the Subject as 刪除.
```

The Subject must remain in the final curriculum plan.

---

# 3. Protect Subjects Currently Marked for Update

Locate the final status decision in the branch where both:

```text
MOEDict
GPlanDict
```

contain the same course code.

Current logic:

```csharp
if (subj.DiffStatusList.Count > 0)
    subj.ProcessStatus = "更新";
else
    subj.ProcessStatus = "略過";

subj.GPlanXml = GPlanDict[mCo];
subj.MOEXml = MOEDict[mCo];
chkSubjectInfoList.Add(subj);
```

Modify the order so that:

```csharp
subj.GPlanXml = GPlanDict[mCo];
subj.MOEXml = MOEDict[mCo];
```

are assigned before determining the final status.

Then apply the `來源課規` priority rule.

Expected logic:

```csharp
subj.GPlanXml = GPlanDict[mCo];
subj.MOEXml = MOEDict[mCo];

if (HasSourceGPlan(subj.GPlanXml))
{
    subj.ProcessStatus = "略過";
}
else
{
    if (subj.DiffStatusList.Count > 0)
        subj.ProcessStatus = "更新";
    else
        subj.ProcessStatus = "略過";
}

chkSubjectInfoList.Add(subj);
```

Important:

Do not remove or bypass the existing difference comparison.

The current comparisons should still execute normally.

Only override the final `ProcessStatus` when `來源課規` has a value.

---

# 4. Do Not Change Normal Add Logic

Do not modify the existing:

```csharp
subj.ProcessStatus = "新增";
```

logic for Subjects that exist in `MOEXml` but not in the existing curriculum plan.

Reason:

A newly added MOE Subject does not have an existing curriculum-plan Subject from which `來源課規` could be read.

Normal add behavior must remain unchanged.

---

# 5. CourseCode Group Rule

Current `GPlanDict` structure is:

```csharp
Dictionary<string, List<XElement>> GPlanDict
```

where the key is primarily:

```text
課程代碼
```

One course code may contain multiple Subject XML elements for different:

```text
GradeYear
Semester
```

Therefore use:

```csharp
HasSourceGPlan(GPlanDict[mCo])
```

instead of checking only:

```csharp
GPlanDict[mCo][0]
```

Required rule:

> If any Subject within the same CourseCode group contains a non-empty `來源課規`, the entire CourseCode group must be treated as `略過`.

Example:

```xml
<Subject
    課程代碼="A001"
    GradeYear="1"
    Semester="1"
    來源課規="來源課規A" />

<Subject
    課程代碼="A001"
    GradeYear="1"
    Semester="2" />
```

Expected result:

```text
CourseCode A001
→ ProcessStatus = 略過
```

Do not partially update one semester while preserving another semester under the same current `chkSubjectInfo` / CourseCode grouping architecture.

---

# 6. Do Not Change Final XML Recalculation

Do not modify the following logic in:

```text
frmCreateGPlanMainBatch.cs
```

The existing final-generation logic must continue to run for all Subjects, including Subjects with `來源課規`.

Keep the current recalculation of:

```text
RowIndex
Level
FullName
startLevel
```

unchanged.

In particular, do not add conditions such as:

```csharp
if (elm.Attribute("來源課規") == null)
```

around these recalculation sections.

---

## Expected Final Data Flow

```text
Existing Curriculum Plan Subject
        ↓
GPlanDict[CourseCode]
        ↓
CheckData()
        ↓
Check 來源課規
        │
        ├─ Has non-empty value
        │       ↓
        │  ProcessStatus = 略過
        │       ↓
        │  Preserve GPlanXml
        │
        └─ No value
                ↓
          Existing comparison logic
                ↓
        新增 / 更新 / 刪除 / 略過
```

Then:

```text
frmCreateGPlanMainBatch
        ↓
ProcessStatus = 略過
        ↓
Use subj.GPlanXml
        ↓
Build new GraduationPlan XML
        ↓
Recalculate RowIndex
        ↓
Recalculate Level
        ↓
Recalculate FullName
        ↓
Recalculate startLevel
        ↓
UPDATE graduation_plan.content
```

---

# Expected Behavior Examples

## Case 1: Existing Subject has source plan and MOE data differs

Existing curriculum plan:

```xml
<Subject
    課程代碼="A001"
    SubjectName="國文"
    來源課規="來源課規A" />
```

MOE comparison detects differences.

Old result:

```text
更新
```

New result:

```text
略過
```

The original `GPlanXml` must be used.

---

## Case 2: Existing Subject has source plan but MOE no longer contains CourseCode

Existing curriculum plan:

```xml
<Subject
    課程代碼="B001"
    SubjectName="特殊課程"
    來源課規="來源課規B" />
```

Course-code master table does not contain `B001`.

Old result:

```text
刪除
```

New result:

```text
略過
```

The Subject must remain in the final curriculum plan.

---

## Case 3: Existing Subject does not have source plan and has differences

```xml
<Subject
    課程代碼="C001"
    SubjectName="英文" />
```

Differences exist.

Expected result remains:

```text
更新
```

---

## Case 4: Existing Subject does not have source plan and MOE no longer contains it

```xml
<Subject
    課程代碼="D001"
    SubjectName="數學" />
```

Expected result remains:

```text
刪除
```

---

## Case 5: Source attribute exists but is empty

```xml
<Subject
    課程代碼="E001"
    來源課規="" />
```

Do not protect the Subject.

Continue using the original update/delete comparison rules.

---

# Do Not Change

Do not change:

* `MOEDict` construction.
* `GPlanDict` construction.
* CourseCode comparison keys.
* `AddList` logic.
* Normal `DelList` creation logic.
* Existing Subject property comparison logic.
* `_CheckSubjectLevel`.
* Subject level difference checking.
* `calSubjDiffCount()`.
* `calSubjUpdateCount()`.
* `calSubjNoChangeCount()`.
* `calSubjDelCount()`.
* `calSubjAddCount()`.
* `ParseStatus()` unless required by a proven issue.
* XML sorting.
* `RowIndex` recalculation.
* `Level` recalculation.
* `FullName` recalculation.
* `startLevel` recalculation.
* SQL UPDATE behavior.
* Existing UI behavior.
* Any unrelated business rules.

Keep the change focused and minimal.

---

# Important Regression Requirement

This modification is only intended to protect Subjects imported from another curriculum plan.

It must not alter normal behavior for standard Subjects originating from the course-code master table.

Normal Subjects must still be able to:

```text
新增
更新
刪除
略過
```

according to the existing logic.

---

# Validation Checklist

* [ ] Build the project successfully.
* [ ] Confirm no compilation errors.
* [ ] Test a normal Subject with no `來源課規`; existing behavior remains unchanged.
* [ ] Test a Subject with `來源課規` that would normally be marked `更新`; confirm it becomes `略過`.
* [ ] Test a Subject with `來源課規` that would normally be marked `刪除`; confirm it becomes `略過`.
* [ ] Confirm a `來源課規=""` Subject is not automatically protected.
* [ ] Confirm whitespace-only `來源課規` is not automatically protected.
* [ ] Test multiple Subject elements under the same CourseCode.
* [ ] Confirm if any Subject in that CourseCode group contains `來源課規`, the entire group is `略過`.
* [ ] Confirm protected Subjects are written from `subj.GPlanXml`.
* [ ] Confirm the `來源課規` attribute remains in the final XML.
* [ ] Confirm normal MOE Subjects still update correctly.
* [ ] Confirm normal obsolete Subjects without `來源課規` are still deleted correctly.
* [ ] Confirm final `RowIndex` is still recalculated.
* [ ] Confirm final `Level` is still recalculated.
* [ ] Confirm final `FullName` is still recalculated.
* [ ] Confirm final `startLevel` is still recalculated.
* [ ] Confirm no unrelated curriculum-plan generation logic changed.

---

# Completion Record

After implementation and validation, create or update:

```text
產生課程規劃調整0819.md
```

Record at least:

1. Modified files.
2. Original behavior.
3. New `來源課規` protection rule.
4. Helper method added.
5. How `更新` is overridden to `略過`.
6. How `刪除` is overridden to `略過`.
7. CourseCode group handling rule.
8. Confirmation that normal Subjects are unchanged.
9. Confirmation that `RowIndex` recalculation remains unchanged.
10. Confirmation that `Level` recalculation remains unchanged.
11. Confirmation that `FullName` recalculation remains unchanged.
12. Confirmation that `startLevel` recalculation remains unchanged.
13. Test cases performed.
14. Test results.
15. Any regression or compatibility considerations.

## Final Requirement

Keep the implementation minimal.

The exact business requirement is:

> Subjects merged from another curriculum plan, identified by a non-empty `來源課規` attribute, must not be updated or deleted by the course-code master table comparison. Force these Subjects to `ProcessStatus = "略過"`, while preserving all existing final curriculum-plan recalculation behavior.
