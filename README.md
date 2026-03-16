# ExcelReportDemo

A .NET 10 console application that queries a SQLite database and generates a formatted Excel report using **ClosedXML.Report** templates. The Excel template is also generated programmatically — no manually-authored `.xlsx` file is required.

---

## Prerequisites

| Requirement | Version |
|---|---|
| .NET SDK | 10.0+ |
| NuGet source | `https://api.nuget.org/v3/index.json` |

If NuGet is not yet configured:
```bash
dotnet nuget add source https://api.nuget.org/v3/index.json --name nuget.org
```

---

## Running

```bash
dotnet run
```

Produces `report.xlsx` in the working directory.

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `ClosedXML` | 0.105.x | Low-level Excel read/write (OpenXML wrapper) |
| `ClosedXML.Report` | 0.2.x | Template engine — binds data to named ranges and substitutes `{{tags}}` |
| `Microsoft.Data.Sqlite` | 9.x | Embedded SQLite database driver |

---

## How It Works

The program runs four steps in sequence:

```
SetupDatabase → CreateTemplate → QueryProducts → FillTemplate → report.xlsx
```

### 1. SetupDatabase
Creates `inventory.db` (SQLite file on disk) with a `Products` table and seeds it with six sample rows. Uses parameterised commands throughout to prevent SQL injection.

### 2. CreateTemplate
Builds `template.xlsx` programmatically using ClosedXML. The worksheet layout is:

| Row | Content |
|---|---|
| 1 | `{{Title}}` — merged across A:F, styled as a large heading |
| 2 | `{{GeneratedOn}}` — italic date label |
| 3 | Spacer (height 8) |
| 4 | Static column headers with dark-blue fill |
| 5 | Data template row — `{{item.Id}}`, `{{item.Name}}`, etc. |
| 6 | Service row — `<<Range>>` in A6 (removed by ClosedXML.Report after generation) |

A workbook-level **Named Range** called `Products` covers `A5:F6`.

### 3. QueryProducts
Runs a SQL query against `inventory.db`, computing `TotalValue = Price × Stock` in the database, and returns a `List<ProductRow>`.

### 4. FillTemplate
Opens `template.xlsx` via `XLTemplate`, binds three variables, then calls `Generate()`:

```csharp
template.AddVariable("Title",       "Monthly Inventory Report");
template.AddVariable("GeneratedOn", DateTime.Now.ToString("dddd, MMMM d, yyyy"));
template.AddVariable("Products",    products);   // List<ProductRow>
```

---

## ClosedXML.Report Template Rules

Understanding these rules is essential for extending the template.

### Scalar variables
Any cell containing `{{Key}}` is replaced with the value from `AddVariable("Key", value)`.

```
Cell A1:  {{Title}}   →   "Monthly Inventory Report"
Cell B2:  {{GeneratedOn}}   →   "Sunday, March 16, 2026"
```

### Collection variables (named ranges)
A named range whose name matches an `AddVariable` key causes the library to repeat the template rows — one copy per item in the collection.

**Tag syntax inside a named range:**
```
{{item.PropertyName}}
```
`item` is the implicit loop variable. `PropertyName` must match a public property on the model class **exactly** (case-sensitive).

**Required two-row structure:**

A named range covering only **one row** triggers **horizontal** expansion (data spreads across columns). Two or more rows trigger **vertical** expansion (data adds new rows downward). The second row must be a service row with `<<Range>>` in its leftmost cell. ClosedXML.Report deletes the service row after processing.

```
Row 5  │ {{item.Id}} │ {{item.Name}} │ … │   ← template row (repeated per item)
Row 6  │ <<Range>>   │               │   │   ← service row (deleted after generation)
       └─────────────── Named Range "Products" ──────────────┘
```

### Style inheritance
Number formats, alignment, fill colour, and other styles applied to the template row (row 5) are carried forward to every generated data row.

---

## Data Model

`ProductRow` properties must match the `{{item.X}}` tags in the template exactly.

```csharp
public class ProductRow
{
    public int    Id         { get; set; }
    public string Name       { get; set; }
    public string Category   { get; set; }
    public double Price      { get; set; }
    public int    Stock      { get; set; }
    public double TotalValue { get; set; }  // computed by SQL: ROUND(Price * Stock, 2)
}
```

---

## Project Structure

```
ExcelReportDemo/
├── ExcelReportDemo.csproj   # Project file — targets net10.0
├── Program.cs               # All application logic (single-file)
├── .gitignore               # Excludes bin/, obj/, *.db, *.xlsx, *.zip
└── README.md
```

**Runtime-generated files** (excluded from source control):

| File | Created by |
|---|---|
| `inventory.db` | `SetupDatabase()` |
| `template.xlsx` | `CreateTemplate()` |
| `report.xlsx` | `XLTemplate.SaveAs()` |

---

## Known Gotchas

| Symptom | Cause | Fix |
|---|---|---|
| `Unknown identifier 'X'` in cells | Tags written as `{{Name}}` instead of `{{item.Name}}` inside a named range | Prefix with `item.` |
| All data in one row, spreading across columns | Named range covers only one row — horizontal expansion is the default | Add a service row with `<<Range>>` and extend the named range to cover both rows |
| `NU1605` package downgrade error on restore | `ClosedXML` version pinned below what `ClosedXML.Report` requires | Align `ClosedXML` version to match the minimum required by `ClosedXML.Report` |
| `MissingMethodException` on `Generate()` | `ClosedXML` and `ClosedXML.Report` versions are incompatible | Use `ClosedXML 0.105.*` with `ClosedXML.Report 0.2.*` |
