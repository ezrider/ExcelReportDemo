// ClosedXML.Report + SQLite demo
// Flow:  SetupDatabase → CreateTemplate → QueryDatabase → FillTemplate → Save

using ClosedXML.Excel;
using ClosedXML.Report;
using Microsoft.Data.Sqlite;

const string DbPath       = "inventory.db";
const string TemplatePath = "template.xlsx";
const string OutputPath   = "report.xlsx";

Console.WriteLine("1. Setting up SQLite database...");
SetupDatabase(DbPath);

Console.WriteLine("2. Generating Excel template...");
CreateTemplate(TemplatePath);

Console.WriteLine("3. Querying data...");
var products = QueryProducts(DbPath);

Console.WriteLine("4. Merging data into template...");
// XLTemplate opens the .xlsx file, substitutes all {{tags}}, expands named ranges, then saves.
var template = new XLTemplate(TemplatePath);
template.AddVariable("Title",       "Monthly Inventory Report");
template.AddVariable("GeneratedOn", DateTime.Now.ToString("dddd, MMMM d, yyyy"));
// The key "Products" must exactly match the Excel Named Range defined in CreateTemplate().
// XLTemplate detects that this is an IEnumerable and repeats the template row for every item.
template.AddVariable("Products", products);
template.Generate();
template.SaveAs(OutputPath);

Console.WriteLine($"\nDone!  Open: {Path.GetFullPath(OutputPath)}");

// ─────────────────────────────────────────────────────────────────────────────
//  DATABASE
// ─────────────────────────────────────────────────────────────────────────────

static void SetupDatabase(string dbPath)
{
    using var conn = new SqliteConnection($"Data Source={dbPath}");
    conn.Open();

    Run(conn, @"
        CREATE TABLE IF NOT EXISTS Products (
            Id       INTEGER PRIMARY KEY,
            Name     TEXT    NOT NULL,
            Category TEXT    NOT NULL,
            Price    REAL    NOT NULL,
            Stock    INTEGER NOT NULL
        )");

    Run(conn, "DELETE FROM Products");

    // Seed sample rows using a parameterised command to avoid injection
    using var cmd = conn.CreateCommand();
    cmd.CommandText = "INSERT INTO Products VALUES ($id,$name,$cat,$price,$stock)";
    var pId    = cmd.Parameters.Add("$id",    SqliteType.Integer);
    var pName  = cmd.Parameters.Add("$name",  SqliteType.Text);
    var pCat   = cmd.Parameters.Add("$cat",   SqliteType.Text);
    var pPrice = cmd.Parameters.Add("$price", SqliteType.Real);
    var pStock = cmd.Parameters.Add("$stock", SqliteType.Integer);

    (int id, string name, string cat, double price, int stock)[] rows =
    [
        (1, "Widget Alpha",   "Widgets",  9.99, 150),
        (2, "Widget Beta",    "Widgets", 14.99,  80),
        (3, "Gadget X-100",   "Gadgets", 29.99,  45),
        (4, "Gadget Y-200",   "Gadgets", 49.99,  20),
        (5, "Doohickey Pro",  "Misc",     4.99, 300),
        (6, "Thingamajig",    "Misc",    19.99,  60),
    ];

    foreach (var (id, name, cat, price, stock) in rows)
    {
        pId.Value = id; pName.Value = name; pCat.Value = cat;
        pPrice.Value = price; pStock.Value = stock;
        cmd.ExecuteNonQuery();
    }
}

static void Run(SqliteConnection conn, string sql)
{
    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;
    cmd.ExecuteNonQuery();
}

static List<ProductRow> QueryProducts(string dbPath)
{
    var list = new List<ProductRow>();

    using var conn = new SqliteConnection($"Data Source={dbPath}");
    conn.Open();

    using var cmd = conn.CreateCommand();
    cmd.CommandText = @"
        SELECT Id, Name, Category, Price, Stock,
               ROUND(Price * Stock, 2) AS TotalValue
        FROM   Products
        ORDER  BY Category, Name";

    using var r = cmd.ExecuteReader();
    while (r.Read())
        list.Add(new ProductRow
        {
            Id         = r.GetInt32(0),
            Name       = r.GetString(1),
            Category   = r.GetString(2),
            Price      = r.GetDouble(3),
            Stock      = r.GetInt32(4),
            TotalValue = r.GetDouble(5),
        });

    return list;
}

// ─────────────────────────────────────────────────────────────────────────────
//  TEMPLATE BUILDER
//  ClosedXML.Report rules:
//    • {{ScalarKey}}       → replaced with the value from AddVariable("ScalarKey", …)
//    • {{PropertyName}}    → within a Named Range, resolved against each collection item
//    • Named Range name    → must equal the key passed to AddVariable for a collection
//    • Styles on row 5     → carried forward to every expanded row
// ─────────────────────────────────────────────────────────────────────────────

static void CreateTemplate(string path)
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Inventory");

    // ── Row 1: report title ──────────────────────────────────────────────────
    ws.Cell("A1").Value = "{{Title}}";
    ws.Cell("A1").Style
      .Font.SetBold()
      .Font.SetFontSize(18)
      .Font.SetFontColor(XLColor.FromHtml("#1A3A5C"));
    ws.Range("A1:F1").Merge();

    // ── Row 2: generation date ───────────────────────────────────────────────
    ws.Cell("A2").Value = "Generated:";
    ws.Cell("A2").Style.Font.SetItalic();
    ws.Cell("B2").Value = "{{GeneratedOn}}";
    ws.Cell("B2").Style.Font.SetItalic();
    ws.Range("B2:F2").Merge();

    ws.Row(3).Height = 8; // visual spacer

    // ── Row 4: column headers ────────────────────────────────────────────────
    string[] headers = ["#", "Product Name", "Category", "Unit Price", "Stock", "Total Value"];
    for (int col = 1; col <= headers.Length; col++)
    {
        var hdr = ws.Cell(4, col);
        hdr.Value = headers[col - 1];
        hdr.Style
           .Font.SetBold()
           .Font.SetFontColor(XLColor.White)
           .Fill.SetBackgroundColor(XLColor.FromHtml("#2E4057"))
           .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
           .Border.SetBottomBorder(XLBorderStyleValues.Medium)
           .Border.SetBottomBorderColor(XLColor.White);
    }

    // ── Row 5: DATA TEMPLATE ROW ─────────────────────────────────────────────
    // "item" is the loop variable ClosedXML.Report binds to each collection element.
    // Property names must match the public properties of ProductRow exactly.
    ws.Cell("A5").Value = "{{item.Id}}";
    ws.Cell("B5").Value = "{{item.Name}}";
    ws.Cell("C5").Value = "{{item.Category}}";
    ws.Cell("D5").Value = "{{item.Price}}";
    ws.Cell("E5").Value = "{{item.Stock}}";
    ws.Cell("F5").Value = "{{item.TotalValue}}";

    // Styles on the template row are inherited by every generated row.
    ws.Cell("A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    ws.Cell("D5").Style.NumberFormat.Format  = "$#,##0.00";
    ws.Cell("F5").Style.NumberFormat.Format  = "$#,##0.00";
    ws.Cell("E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    ws.Range("A5:F5").Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F7FA");

    // ── Row 6: SERVICE ROW ────────────────────────────────────────────────────
    // ClosedXML.Report requires a named range of at least 2 rows to expand
    // vertically (add rows). A single-row range triggers horizontal expansion instead.
    // The service row must contain <<Range>> in its leftmost cell; the library
    // removes the service row after generation.
    ws.Cell("A6").Value = "<<Range>>";

    // ── Register the Named Range (must cover BOTH the template and service rows) ─
    // The name "Products" MUST match the key in: template.AddVariable("Products", …)
    wb.DefinedNames.Add("Products", ws.Range("A5:F6"));

    // ── Column widths ────────────────────────────────────────────────────────
    ws.Column(1).Width =  6;
    ws.Column(2).Width = 24;
    ws.Column(3).Width = 14;
    ws.Column(4).Width = 13;
    ws.Column(5).Width = 10;
    ws.Column(6).Width = 14;

    wb.SaveAs(path);
    Console.WriteLine($"   Template saved → {Path.GetFullPath(path)}");
}

// ─────────────────────────────────────────────────────────────────────────────
//  MODEL  –  property names must match the {{tags}} in the template exactly
// ─────────────────────────────────────────────────────────────────────────────

public class ProductRow
{
    public int    Id         { get; set; }
    public string Name       { get; set; } = "";
    public string Category   { get; set; } = "";
    public double Price      { get; set; }
    public int    Stock      { get; set; }
    public double TotalValue { get; set; }
}
