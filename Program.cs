using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using Microsoft.Playwright;
using Renci.SshNet;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;

namespace DailyRevenueReportAutomation;

internal static class Program
{
    // ── Load configuration ────────────────────────────────────────
    static IConfiguration config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
    .AddEnvironmentVariables()
    .Build();

    static string TDurl = config["TDPortal:Url"];
    static string TDuser = config["TDPortal:Username"];
    static string TDpass = config["TDPortal:Password"];
    static string CIBCurl = config["CIBCPortal:Url"];
    static string CIBCuser = config["CIBCPortal:Username"];
    static string CIBCpass = config["CIBCPortal:Password"];
    static string? localCsv = null;
    private static async Task Main(string[] args)
    {
        try
        {
            //await TestTDPortalHttp();

            //Saving all daily ACE transaction to Rev Report
            //GetDailyRevenueACETransactions(args);

            //TD Portal automation to download the report CSV directly
            //RunTDPortalAutomation().GetAwaiter().GetResult();

            //CIBC Portal automation to download the report CSV directly
            //RunCIBCPortalAutomation().GetAwaiter().GetResult();

            //RunCibcDrsCalculator();
            ApplyManualAdjustments();

            //Performance Report automation Agency 2 & 4
            //RunPerformanceReportAutomation().GetAwaiter().GetResult();

            Console.WriteLine("Done.");
            //return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERROR: {ex.Message}");
            //return 1;
        }
    }

    private static void ApplyManualAdjustments()
    {
        var revReportDir = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report";
        var adjustmentsRoot = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\Manual Adjustments";
        var revManualTab = "Manual Adjustments";
        var agentTabs = new[] { "Tim Katis", "Barb Boudreau", "Ian Tougas", "Frank Ramlal", "HOUSE" };

        const int DeskCol = 4, DayStartCol = 9, DayEndCol = 41;

        var now = DateTime.Now;
        var log = new List<string>();
        void Log(string m) { Console.WriteLine(m); log.Add($"{DateTime.Now:HH:mm:ss}  {m}"); }

        // Parse a tab name to a date: handles "Aug 17", "August 17", "Jul 30", "July 30".
        DateTime? ParseTabDate(string name)
        {
            name = (name ?? "").Trim();
            foreach (var f in new[] { "MMM d", "MMMM d", "MMM dd", "MMMM dd" })
                if (DateTime.TryParseExact($"{name} {now.Year}", $"{f} yyyy",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
                    return d.Date;
            return null;
        }

        Log("═══════════════════════════════════════════");
        Log("  Manual Adjustments");
        Log($"  {now:yyyy-MM-dd HH:mm:ss}");
        Log("═══════════════════════════════════════════");

        // 1. Rev report file
        var currentMonthUpper = now.ToString("MMMM").ToUpper();
        var revReportFile = Directory.GetFiles(revReportDir, $"Rev Report {currentMonthUpper} {now.Year}*.xlsx")
            .OrderByDescending(File.GetLastWriteTime).FirstOrDefault()
            ?? throw new FileNotFoundException($"Rev Report for {currentMonthUpper} {now.Year} not found.");

        // 2. Adjustments file
        var monthDir = Path.Combine(adjustmentsRoot, now.ToString("MMMM"));
        if (!Directory.Exists(monthDir))
            throw new DirectoryNotFoundException($"Month folder not found: {monthDir}");
        var adjFile = Directory.GetFiles(monthDir, "Manual Adjustments*.xlsx")
            .OrderByDescending(File.GetLastWriteTime).FirstOrDefault()
            ?? throw new FileNotFoundException($"Manual Adjustments file not found in {monthDir}.");

        void FlushLog()
        {
            try
            {
                var logDir = Path.Combine(monthDir, "Logs");
                Directory.CreateDirectory(logDir);
                var p = Path.Combine(logDir, $"ManualAdjustments_{now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllLines(p, log);
                Console.WriteLine($"Log written: {p}");
            }
            catch (Exception ex) { Console.WriteLine($"Log write failed: {ex.Message}"); }
        }

        Log($"  Adjustments: {adjFile}");
        Log($"  Rev Report:  {revReportFile}");

        // 3. Read TODAY's tab from the adjustments file
        Log("\nReading adjustments...");
        var adjustments = new List<ManualAdjustment>();
        var adjTemp = SanitizeWorkbookCopy(adjFile);   // repair blank/whitespace sheet names so ClosedXML can load
        try
        {
            using var adjWb = new XLWorkbook(adjTemp);

            var adjWs = adjWb.Worksheets.FirstOrDefault(s => ParseTabDate(s.Name) == now.Date);
            if (adjWs == null)
            {
                Log($"  No tab dated {now:MMM d} in the adjustments file — nothing to process today.");
                FlushLog();
                return;
            }
            Log($"  Reading tab: '{adjWs.Name}'");

            var row = 3; // 1=title, 2=headers
            while (true)
            {
                var dateCell = adjWs.Cell(row, 1);
                if (dateCell.IsEmpty() || string.IsNullOrWhiteSpace(dateCell.GetString())) break;

                DateTime adjDate;
                if (!dateCell.TryGetValue(out adjDate) && !TryParseAdjDate(dateCell.GetString(), out adjDate))
                {
                    Log($"  WARNING: row {row} unparseable date '{dateCell.GetString()}' — skipped.");
                    row++; continue;
                }

                double comm = 0;
                var commCell = adjWs.Cell(row, 6);
                if (!commCell.TryGetValue(out comm))
                    double.TryParse(commCell.GetString().Replace("$", "").Replace(",", "").Trim(),
                        NumberStyles.Any, CultureInfo.InvariantCulture, out comm);

                adjustments.Add(new ManualAdjustment
                {
                    Date = adjDate,
                    Collector = adjWs.Cell(row, 2).GetString().Trim(),
                    Debt = adjWs.Cell(row, 3).GetString().Trim(),
                    From = adjWs.Cell(row, 4).GetString().Trim(),
                    ToAgent = adjWs.Cell(row, 5).GetString().Trim(),
                    Comm = comm,
                    Tab = adjWs.Cell(row, 7).GetString().Trim()
                });
                row++;
            }
        }
        finally { try { File.Delete(adjTemp); } catch { } }

        Log($"  Found {adjustments.Count} adjustments.");
        if (adjustments.Count == 0) { Log("  Nothing to post."); FlushLog(); return; }

        // 4. Write into the Rev Report
        try
        {
            using (var doc = SpreadsheetDocument.Open(revReportFile, isEditable: true))
            {
                var (manPart, manData) = GetSheet(doc, revManualTab);

                uint last = manData.Elements<Row>()
                        .Where(r => r.Elements<Cell>().Any(c =>
                            !string.IsNullOrEmpty(c.CellValue?.Text) ||
                            c.InlineString != null))
                        .Select(r => r.RowIndex?.Value ?? 0u)
                        .DefaultIfEmpty(1u)
                        .Max();
                uint writeRow = last + 1;

                // double-post guard: bail if today already appears in column B
                foreach (var r in manData.Elements<Row>())
                {
                    var bCell = r.Elements<Cell>().FirstOrDefault(c => ColIndex(c.CellReference) == 2);
                    if (GetNumber(bCell) is double oa && oa >= -657435 && oa <= 2958465
    && DateTime.FromOADate(oa).Date == now.Date)
                    {
                        var cVals = string.Join(" | ", r.Elements<Cell>()
                            .OrderBy(c => ColIndex(c.CellReference))
                            .Select(c => $"{c.CellReference}={c.CellValue?.Text}"));
                        Log($"  ABORT: row {r.RowIndex?.Value} already today. Contents: {cVals}");
                        FlushLog(); return;
                    }
                }

                foreach (var a in adjustments)
                {
                    Log($"    row: date={a.Date:MM/dd/yyyy} collector={a.Collector} debt={a.Debt} comm={a.Comm} from={a.From} to={a.ToAgent}");
                    var above = manData.Elements<Row>()
                    .FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == writeRow - 1);
                    uint? bS = above?.Elements<Cell>()
                        .FirstOrDefault(c => ColIndex(c.CellReference) == 2)?.StyleIndex?.Value;
                    uint? cS = above?.Elements<Cell>()
                        .FirstOrDefault(c => ColIndex(c.CellReference) == 3)?.StyleIndex?.Value;

                    var row = GetOrCreateRow(manData, writeRow);
                    SetDate(GetOrCreateCell(row, 2), now.Date, bS);
                    SetDate(GetOrCreateCell(row, 3), a.Date, cS);
                    SetText(GetOrCreateCell(row, 4), a.Debt);
                    SetText(GetOrCreateCell(row, 5), a.From);
                    SetText(GetOrCreateCell(row, 6), a.ToAgent);
                    SetNumber(GetOrCreateCell(row, 7), a.Comm);
                    writeRow++;
                }
                manPart.Worksheet.Save();

                // TODO: agent-tab apply goes here — see note below

                // make Excel recalc formulas and refresh pivots when the file opens
                var wb = doc.WorkbookPart.Workbook;
                wb.CalculationProperties ??= wb.AppendChild(new CalculationProperties());
                wb.CalculationProperties.FullCalculationOnLoad = true;
                foreach (var pc in doc.WorkbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
                { pc.PivotCacheDefinition.RefreshOnLoad = true; pc.PivotCacheDefinition.Save(); }
                wb.Save();
            }
        }
        catch (Exception ex)
        {
            Log($"SECTION 4 FAILED: {ex.GetType().Name}: {ex.Message}");
            Log(ex.ToString());
            FlushLog();
            throw;
        }

        FlushLog();
        Log("Done.");
    }

    // ---- helpers ----
    static int ColIndex(string cellRef)
    {
        if (string.IsNullOrEmpty(cellRef)) return 0;
        int i = 0, col = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i]))
        { col = col * 26 + (char.ToUpper(cellRef[i]) - 'A' + 1); i++; }
        return col;
    }
    static string ColName(int index)
    {
        string s = ""; while (index > 0) { index--; s = (char)('A' + index % 26) + s; index /= 26; }
        return s;
    }
    static (WorksheetPart part, SheetData data) GetSheet(SpreadsheetDocument doc, string name)
    {
        var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Sheet '{name}' not found.");
        var part = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
        return (part, part.Worksheet.GetFirstChild<SheetData>());
    }
    static Row GetOrCreateRow(SheetData data, uint rowIndex)
    {
        var row = data.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == rowIndex);
        if (row != null) return row;
        row = new Row { RowIndex = rowIndex };
        var after = data.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value > rowIndex);
        if (after != null) data.InsertBefore(row, after); else data.Append(row);
        return row;
    }
    static Cell GetOrCreateCell(Row row, int col)
    {
        string reff = ColName(col) + row.RowIndex;
        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == reff);
        if (cell != null) return cell;
        var after = row.Elements<Cell>().FirstOrDefault(c => ColIndex(c.CellReference) > col);
        cell = new Cell { CellReference = reff };
        if (after != null) row.InsertBefore(cell, after); else row.Append(cell);
        return cell;
    }
    static void SetText(Cell c, string text)
    {
        c.RemoveAllChildren();                 // drop any old <f>, <v>, <is>, cached type
        c.CellFormula = null;
        c.DataType = CellValues.InlineString;
        c.Append(new InlineString(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }));
    }
    static void SetNumber(Cell c, double v)
    {
        c.RemoveAllChildren<CellFormula>(); c.RemoveAllChildren<InlineString>();
        c.DataType = null;               // number is default; this also drops a formula's "str" type
        c.CellValue = new CellValue(v.ToString(CultureInfo.InvariantCulture));
    }
    static double? GetNumber(Cell c)     // reads a value OR a formula's cached value
        => double.TryParse(c?.CellValue?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : (double?)null;
    static void SetDate(Cell c, DateTime d, uint? styleFromAbove)
    {
        c.RemoveAllChildren<CellFormula>(); c.RemoveAllChildren<InlineString>();
        c.DataType = null;
        if (styleFromAbove.HasValue) c.StyleIndex = styleFromAbove.Value;
        c.CellValue = new CellValue(d.ToOADate().ToString(CultureInfo.InvariantCulture));
    }

    // Copies the workbook to a temp file and renames any blank/whitespace sheet
    // name so ClosedXML can load it. Caller deletes the temp file. Renames rather
    // than deletes — deleting a referenced sheet part can throw or corrupt the copy.
    private static string SanitizeWorkbookCopy(string sourcePath)
    {
        var tempPath = Path.Combine(Path.GetTempPath(),
            $"adj_{DateTime.Now:yyyyMMddHHmmss}_{Guid.NewGuid():N}.xlsx");
        File.Copy(sourcePath, tempPath, overwrite: true);

        using (var doc = SpreadsheetDocument.Open(tempPath, isEditable: true))
        {
            var sheets = doc.WorkbookPart?.Workbook?.Sheets?.Elements<Sheet>().ToList();
            var changed = false;
            foreach (var sheet in sheets ?? Enumerable.Empty<Sheet>())
            {
                if (string.IsNullOrWhiteSpace(sheet.Name))
                {
                    sheet.Name = "RECOVERED_" + (sheet.SheetId?.Value.ToString()
                                                 ?? Guid.NewGuid().ToString("N"));
                    changed = true;
                }
            }
            if (changed) doc.WorkbookPart.Workbook.Save();
        }
        return tempPath;
    }

    private static void ApplySide(
        XLWorkbook wb, string[] tabs, string name, int day, double delta,
        int deskCol, int dayStart, int dayEnd,
        string sideLabel, string tag, Action<string> logLine,
        ref int applied, ref int skipped)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            logLine($"  SKIP {sideLabel} {tag}: name is blank."); skipped++; return;
        }

        var (ws, row, count) = FindDeskAcrossTabs(wb, tabs, name, deskCol);
        if (ws == null)
        {
            logLine($"  SKIP {sideLabel} {tag}: '{name}' not matched in any tab (column D)."); skipped++; return;
        }
        if (count > 1)
            logLine($"  WARN {sideLabel} {tag}: '{name}' matched {count} rows — using {ws.Name} row {row}.");

        int col = FindDayColumn(ws, day, dayStart, dayEnd);
        if (col == -1)
        {
            logLine($"  SKIP {sideLabel} {tag}: day {day} not found in '{ws.Name}' row-1 headers."); skipped++; return;
        }

        var cell = ws.Cell(row, col);
        if (cell.HasFormula)
            logLine($"  WARN {sideLabel} {tag}: {ws.Name}!{cell.Address} is a FORMULA — being replaced with a value.");

        double cur = 0; cell.TryGetValue(out cur);
        double updated = cur + delta;
        cell.Value = updated;

        logLine($"  OK   {sideLabel} {tag}: {ws.Name} row {row} ({name}) {cell.Address}  {cur:F2} -> {updated:F2}  ({(delta < 0 ? "" : "+")}{delta:F2})");
        applied++;
    }

    private static (IXLWorksheet ws, int row, int matchCount) FindDeskAcrossTabs(
        XLWorkbook wb, string[] tabs, string search, int deskCol)
    {
        var target = NormalizeName(search);
        if (target.Length == 0) return (null, -1, 0);

        IXLWorksheet firstWs = null; int firstRow = -1, count = 0;
        foreach (var tabName in tabs)
        {
            var ws = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(tabName, StringComparison.OrdinalIgnoreCase));
            if (ws == null) continue;

            int last = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int r = 2; r <= last; r++)
            {
                var d = ws.Cell(r, deskCol).GetString().Trim();
                if (d.Length == 0) continue;

                bool match = NormalizeName(d) == target;
                if (!match && d.Contains('/'))                       // desk codes like "LD / TX / BA"
                    match = d.Split('/').Any(t => NormalizeName(t) == target);

                if (match)
                {
                    if (firstWs == null) { firstWs = ws; firstRow = r; }
                    count++;
                }
            }
        }
        return (firstWs, firstRow, count);
    }

    private static int FindDayColumn(IXLWorksheet ws, int day, int start, int end)
    {
        for (int c = start; c <= end; c++)
        {
            var cell = ws.Cell(1, c);
            if (cell.TryGetValue(out double d) && (int)d == day) return c;
            if (int.TryParse(cell.GetString().Trim(), out int di) && di == day) return c;
        }
        return -1;
    }

    private static string NormalizeName(string s) =>
        new string((s ?? "").Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();

    private static bool TryParseAdjDate(string s, out DateTime date)
    {
        s = (s ?? "").Trim();
        string[] fmts = { "M.d.yyyy", "MM.dd.yyyy", "M/d/yyyy", "MM/dd/yyyy", "M/d/yy", "MM/dd/yy", "yyyy-MM-dd" };
        return DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
            || DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private static void RemoveLegacyFormControls(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, true);
        var wbPart = doc.WorkbookPart!;

        foreach (var wsPart in wbPart.WorksheetParts)
        {
            // Remove drawing parts (form controls, shapes, text boxes)
            foreach (var drawingPart in wsPart.DrawingsPart != null
                ? new[] { wsPart.DrawingsPart }
                : Array.Empty<DrawingsPart>())
            {
                wsPart.DeletePart(drawingPart);
            }

            // Remove VML drawing parts (legacy controls)
            foreach (var vmlPart in wsPart.VmlDrawingParts.ToList())
                wsPart.DeletePart(vmlPart);

            // Remove drawing references from worksheet XML
            foreach (var el in wsPart.Worksheet.Elements()
                .Where(e => e.LocalName == "drawing"
                         || e.LocalName == "legacyDrawing"
                         || e.LocalName == "controls")
                .ToList())
                el.Remove();

            wsPart.Worksheet.Save();
        }

        wbPart.Workbook.Save();
        Console.WriteLine("  Removed legacy form controls.");
    }

    // Find agent name in column B, return row number (0 if not found)
    // Uses contains/fuzzy matching since names might have slight differences
    private static int FindAgentRow(IXLWorksheet ws, string agentName)
    {
        var normalized = agentName.ToUpper().Replace(".", " ").Replace("_", " ").Trim();

        // Search column B (skip row 1 header and row 2 total)
        for (int r = 3; r <= 200; r++)
        {
            var cellVal = ws.Cell(r, 2).GetString().Trim().ToUpper()
                .Replace(".", " ").Replace("_", " ");

            if (string.IsNullOrEmpty(cellVal))
                continue;

            // Exact match first
            if (cellVal == normalized)
                return r;

            // Contains match (handles "SHAWN DAVIS" matching "Shawn Davis")
            if (cellVal.Contains(normalized) || normalized.Contains(cellVal))
                return r;
        }

        return 0;
    }

    internal class ManualAdjustment
    {
        public DateTime Date { get; set; }
        public string Collector { get; set; } = "";
        public string From { get; set; } = "";
        public string ToAgent { get; set; } = "";
        public double Comm { get; set; }
        public string Debt { get; set; } = "";
        public string Tab { get; set; } = "";
    }

    private static void GetDailyRevenueACETransactions(string[] args)
    {
        try
        {
            // ── SFTP settings ─────────────────────────────────────────────
            var sftpHost = config["Sftp:Host"]
                ?? throw new InvalidOperationException("Sftp:Host is not configured.");
            var sftpPort = int.Parse(config["Sftp:Port"] ?? "22");
            var sftpUser = config["Sftp:Username"]
                ?? throw new InvalidOperationException("Sftp:Username is not configured.");
            var sftpPass = config["Sftp:Password"]
                ?? throw new InvalidOperationException("Sftp:Password is not configured.");
            var sftpRemoteDir = config["Sftp:RemoteDir"] ?? "/";
            var filePattern = config["Sftp:FilePattern"] ?? "RevenueReportTransactions-*.csv";

            // ── Report settings ───────────────────────────────────────────
            var reportPath = config["Report:Path"]
                ?? throw new InvalidOperationException("Report:Path is not configured.");
            var sheetName = config["Report:SheetName"] ?? "Trans";
            var numCols = int.Parse(config["Report:NumCols"] ?? "18");
            var dateColIndex = int.Parse(config["Report:DateColIndex"] ?? "7");
            var dateColName = config["Report:DateColName"] ?? "Transaction_Date";
            // ── Parse optional report month arg (YYYY-MM), default = now ──
            var (year, month) = ParseReportMonth(args);
            Console.WriteLine($"Filtering to: {new DateTime(year, month, 1):MMMM yyyy}");

            // ── 1. Connect to SFTP ────────────────────────────────────────
            Console.WriteLine("Connecting to SFTP...");
            using var sftp = new SftpClient(sftpHost, sftpPort, sftpUser, sftpPass);
            sftp.Connect();
            Console.WriteLine("Connected.");

            // ── 2. Find most recent matching file ─────────────────────────
            var remotePath = FindMostRecentFile(sftp, sftpRemoteDir, filePattern);

            // ── 3. Download to temp file ──────────────────────────────────
            localCsv = DownloadCsv(sftp, remotePath);
            sftp.Disconnect();

            // ── 4. Read CSV ───────────────────────────────────────────────
            var (headers, rows) = ReadCsv(localCsv);
            Console.WriteLine($"CSV loaded: {rows.Count} rows, {headers.Length} columns");
            Console.WriteLine($"Columns: {string.Join(", ", headers)}");

            // ── 5. Resolve date column ────────────────────────────────────
            var dateIdx = ResolveDateColumnIndex(headers, dateColName, dateColIndex);

            // ── 6. Filter to target month ─────────────────────────────────
            var filtered = FilterByMonth(rows, dateIdx, year, month);
            if (filtered.Count == 0)
            {
                Console.WriteLine("No rows to write. Exiting without modifying the report.");
            }

            StripOdbcConnection(reportPath);

            // ── 7. Write to Trans tab ─────────────────────────────────────
            WriteToTransTab(filtered, headers, dateIdx, reportPath, sheetName, numCols);
            Console.WriteLine("Done.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERROR: {ex.Message}");
        }
        finally
        {
            if (localCsv is not null && File.Exists(localCsv))
                File.Delete(localCsv);
        }
    }

    private static void RunCibcDrsCalculator()
    {
        Console.WriteLine("═══════════════════════════════════════════");
        Console.WriteLine("  CIBC DRS Calculator");
        Console.WriteLine($"  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        Console.WriteLine("═══════════════════════════════════════════");

        var othersPath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others";
        var revReportPath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report";

        var csvPath = Path.Combine(othersPath, "CIBC POST.csv");
        // Find the calculator file (might be .xls or .xlsx)
        var calcPath = Directory.GetFiles(othersPath, "CIBC DRS CALCULATOR*")
                .FirstOrDefault()
                ?? throw new FileNotFoundException("CIBC DRS CALCULATOR not found.");

        // Find current month's rev report (e.g. "Rev Report JUNE 2026 V8.xlsx")
        var currentMonth = DateTime.Now.ToString("MMMM").ToUpper();
        var currentYear = DateTime.Now.Year;
        //var revReportFile = Directory.GetFiles(revReportPath, $"Rev Report {currentMonth} {currentYear}*.xlsx")
        var revReportFile = Directory.GetFiles(revReportPath, $"TESTRevReport.xlsx")
            .OrderByDescending(f => File.GetLastWriteTime(f))
            .FirstOrDefault()
            ?? throw new FileNotFoundException($"Rev Report for {currentMonth} {currentYear} not found.");

        Console.WriteLine($"  CSV: {Path.GetFileName(csvPath)}");
        Console.WriteLine($"  Calculator: {Path.GetFileName(calcPath)}");
        Console.WriteLine($"  Rev Report: {Path.GetFileName(revReportFile)}");
        Console.WriteLine();

        // ══════════════════════════════════════════════════════
        // STEP 1: Read CIBC POST.csv and find last data row
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 1: Reading CSV...");
        var csvLines = File.ReadAllLines(csvPath);
        var lastDataRow = csvLines.Length - 1; // subtract header row
                                               // Find actual last row with data in column D (index 3)
        for (int i = csvLines.Length - 1; i >= 1; i--)
        {
            var cols = csvLines[i].Split(',');
            if (cols.Length > 3 && !string.IsNullOrWhiteSpace(cols[3]))
            {
                lastDataRow = i + 1; // 1-based row number including header
                break;
            }
        }
        Console.WriteLine($"  CSV last data row: {lastDataRow}");

        if (calcPath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
            && !calcPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("  Converting .xls to .xlsx...");
            var newPath = calcPath + "x"; // .xls → .xlsx
            ConvertXlsToXlsx(calcPath, newPath);
            calcPath = newPath;
        }

        // ══════════════════════════════════════════════════════
        // STEP 2: Open CIBC DRS CALCULATOR
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 2: Opening calculator...");
        // Pre-clean: remove the B72 formula that ClosedXML can't parse
        CleanProblematicFormulas(calcPath);
        RemoveLegacyFormControls(calcPath);

        using var calcWb = new XLWorkbook(calcPath);
        var calcWs = calcWb.Worksheet("CIBC"); // adjust sheet name if different
        Console.WriteLine($"  Sheet: {calcWs.Name}");
        Console.WriteLine($"  D9 before: '{calcWs.Cell(9, 4).Value}'");
        Console.WriteLine($"  B9 before: '{calcWs.Cell(9, 2).Value}'");

        // ══════════════════════════════════════════════════════
        // STEP 3: Update B72 formula with correct last row
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 3: Validating B72...");

        // Read B71 (TOTAL)
        double b71 = 0;
        if (calcWs.Cell(71, 2).TryGetValue(out double b71Val))
            b71 = b71Val;

        // Read D28 (last row total) from the CSV directly
        double csvTotal = 0;
        var csvCols = csvLines[lastDataRow - 1].Split(',');
        if (csvCols.Length > 3)
            double.TryParse(csvCols[3].Replace("$", "").Replace("\"", "").Trim(),
                NumberStyles.Any, CultureInfo.InvariantCulture, out csvTotal);

        var difference = b71 - csvTotal;
        calcWs.Cell(72, 2).Clear();
        calcWs.Cell(72, 2).Value = difference;

        Console.WriteLine($"  B71 (Total): {b71}");
        Console.WriteLine($"  CSV D{lastDataRow}: {csvTotal}");
        Console.WriteLine($"  B72 (Difference): {difference}");

        if (Math.Abs(difference) < 0.01)
            Console.WriteLine("  Match confirmed.");
        else
            Console.WriteLine($"  WARNING: Mismatch of {difference:C2}");

        // ══════════════════════════════════════════════════════
        // STEP 4: Find dynamic range B9:Bxx (until empty or TOTAL)
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 4: Reading COMM values...");
        var commValues = new List<double>();
        int startRow = 9;
        int endRow = startRow;

        for (int r = startRow; r <= 69; r++)
        {
            var nameCell = calcWs.Cell(r, 1).GetString().Trim();

            if (nameCell.Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                break;

            double commVal = 0;
            var bCell = calcWs.Cell(r, 2);

            if (bCell.TryGetValue(out double d))
            {
                commVal = d;
            }
            else
            {
                var text = bCell.GetString().Replace("$", "").Replace(",", "").Trim();
                double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out commVal);
            }

            commValues.Add(commVal);
            endRow = r;
        }
        Console.WriteLine($"  Found {commValues.Count} rows ({startRow} to {endRow}), including hidden");

        // ══════════════════════════════════════════════════════
        // STEP 5: Paste COMM values to D9 onwards
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 5: Pasting COMM to column D...");
        for (int i = 0; i < commValues.Count; i++)
        {
            var cell = calcWs.Cell(startRow + i, 4);
            cell.Clear();
            cell.Value = commValues[i];
        }
        Console.WriteLine($"  D9 after write: '{calcWs.Cell(9, 4).Value}'");
        try
        {
            calcWb.Save();
            Console.WriteLine("  Save succeeded.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Save FAILED: {ex.Message}");
        }
        // ══════════════════════════════════════════════════════
        // STEP 6: Open Rev Report, go to Tim Katis tab
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 6: Opening Rev Report...");
        using var revWb = new XLWorkbook(revReportFile);
        var timSheet = revWb.Worksheets.FirstOrDefault(ws =>
            ws.Name.Contains("Tim Katis", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException("'Tim Katis' tab not found in Rev Report.");

        Console.WriteLine($"Found tab: '{timSheet.Name}'");

        // ══════════════════════════════════════════════════════
        // STEP 7: Grab G117:G177 from Tim Katis
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 7: Reading G117:G177...");
        int timStartRow = 118;
        var gValues = new List<double>();
        for (int i = 0; i < commValues.Count; i++)
        {
            int r = timStartRow + i;
            double val = 0;
            var gCell = timSheet.Cell(r, 7);
            var bName = timSheet.Cell(r, 2).GetString().Trim(); // agent name
            var hidden = timSheet.Row(r).IsHidden;

            if (gCell.TryGetValue(out double d))
            {
                val = d;
            }
            else
            {
                var text = gCell.GetString().Replace("$", "").Replace(",", "").Trim();
                double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out val);
            }

            gValues.Add(val);

            // Log first few to verify alignment
            if (i < 5)
                Console.WriteLine($"    Row {r}: {bName} = {val} {(hidden ? "(hidden)" : "")}");
        }
        Console.WriteLine($"  Read {gValues.Count} values from G column");

        // ══════════════════════════════════════════════════════
        // STEP 8: Subtract G from D (D - G = result)
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 8: Calculating D - G...");
        var results = new List<double>();
        for (int i = 0; i < commValues.Count; i++)
        {
            var result = commValues[i] - gValues[i];
            results.Add(result);
        }
        Console.WriteLine($"  Calculated {results.Count} results");

        // ══════════════════════════════════════════════════════
        // STEP 9: Find today's day column in Tim Katis row 1
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 9: Finding today's column...");
        var today = DateTime.Now.Day;
        int targetCol = -1;

        // Search I1:AM1 (columns 9 to 39) for today's day number
        for (int c = 9; c <= 39; c++) // I=9, AM=39
        {
            try
            {
                var headerVal = timSheet.Cell(1, c).GetDouble();
                if ((int)headerVal == today)
                {
                    targetCol = c;
                    Console.WriteLine($"  Today is day {today}, matched column {GetColumnLetter(c - 1)} (col {c})");
                    break;
                }
            }
            catch { }
        }

        if (targetCol == -1)
        {
            Console.WriteLine($"  ERROR: Could not find day {today} in row 1 (I1:AM1)");
            return;
        }

        // ══════════════════════════════════════════════════════
        // STEP 10: Paste results to Tim Katis tab
        // ══════════════════════════════════════════════════════
        Console.WriteLine($"STEP 10: Pasting results to {GetColumnLetter(targetCol - 1)}117:{GetColumnLetter(targetCol - 1)}{timStartRow + results.Count - 1}...");
        for (int i = 0; i < results.Count; i++)
        {
            timSheet.Cell(timStartRow + i, targetCol).Value = results[i];
        }

        // ══════════════════════════════════════════════════════
        // SAVE
        // ══════════════════════════════════════════════════════
        Console.WriteLine("\nSaving...");
        calcWb.Save();
        Console.WriteLine($"  Saved: {Path.GetFileName(calcPath)}");
        revWb.Save();
        Console.WriteLine($"  Saved: {Path.GetFileName(revReportFile)}");

        Console.WriteLine("\n═══════════════════════════════════════════");
        Console.WriteLine("  CIBC DRS Calculator complete.");
        Console.WriteLine("═══════════════════════════════════════════");
    }

    private static void CleanProblematicFormulas(string filePath)
    {
        using var doc = SpreadsheetDocument.Open(filePath, true);
        var wbPart = doc.WorkbookPart!;

        // Remove ALL formulas with external references from ALL sheets
        foreach (var wsPart in wbPart.WorksheetParts)
        {
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellFormula != null)
                    {
                        var formula = cell.CellFormula.Text ?? "";
                        // Remove formulas with UNC paths, external file refs, or sheet refs with brackets
                        if (formula.Contains(@"\\") ||
                            formula.Contains("[") ||
                            formula.Contains("'\\"))
                        {
                            cell.CellFormula.Remove();
                        }
                    }
                }
            }
            wsPart.Worksheet.Save();
        }

        // Remove external link parts
        foreach (var extLink in wbPart.ExternalWorkbookParts.ToList())
            wbPart.DeletePart(extLink);

        // Remove externalReferences from workbook.xml
        var workbook = wbPart.Workbook;
        foreach (var el in workbook.Elements()
            .Where(e => e.LocalName == "externalReferences").ToList())
            el.Remove();

        // Remove defined names that reference external files
        var definedNames = workbook.DefinedNames;
        if (definedNames != null)
        {
            foreach (var dn in definedNames.Elements<DefinedName>().ToList())
            {
                var val = dn.Text ?? "";
                if (val.Contains("[") || val.Contains(@"\\") || val.Contains("'\\"))
                    dn.Remove();
            }
            if (!definedNames.HasChildren)
                definedNames.Remove();
        }

        // Kill calcChain
        if (wbPart.CalculationChainPart != null)
            wbPart.DeletePart(wbPart.CalculationChainPart);

        workbook.Save();
        Console.WriteLine("  Cleaned external references, links, and formulas.");
    }

    private static void ConvertXlsToXlsx(string xlsPath, string xlsxPath)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(xlsPath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        var dataSet = reader.AsDataSet(new ExcelDataReader.ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataReader.ExcelDataTableConfiguration
            {
                UseHeaderRow = false
            }
        });

        using var wb = new XLWorkbook();
        foreach (DataTable dt in dataSet.Tables)
        {
            var ws = wb.AddWorksheet(dt.TableName);
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var val = dt.Rows[r][c];
                    if (val != null && val != DBNull.Value)
                        ws.Cell(r + 1, c + 1).Value = val.ToString();
                }
            }
        }
        wb.SaveAs(xlsxPath);
        Console.WriteLine($"  Converted to: {Path.GetFileName(xlsxPath)}");
    }
    private static string GetColumnLetter(int colIndex)
    {
        var result = "";
        var n = colIndex;
        while (n >= 0)
        {
            result = (char)('A' + n % 26) + result;
            n = n / 26 - 1;
        }
        return result;
    }
    private static void StripOdbcConnection(string reportPath)
    {
        using var doc = SpreadsheetDocument.Open(reportPath, true);
        var wbPart = doc.WorkbookPart!;

        foreach (var wsPart in wbPart.WorksheetParts)
        {
            foreach (var qtp in wsPart.QueryTableParts.ToList())
                wsPart.DeletePart(qtp);

            foreach (var el in wsPart.Worksheet.Elements()
                .Where(e => e.LocalName == "queryTableParts").ToList())
                el.Remove();

            foreach (var tp in wsPart.TableDefinitionParts)
            {
                var tbl = tp.Table;
                if (tbl?.TableColumns != null)
                    foreach (var col in tbl.TableColumns.Elements<TableColumn>())
                        col.QueryTableFieldId = null;
                tbl.ConnectionId = null;
                tbl.Save();
            }
            wsPart.Worksheet.Save();
        }

        foreach (var conn in wbPart.Workbook.Descendants<Connection>().ToList())
            conn.Remove();

        if (wbPart.CalculationChainPart != null)
            wbPart.DeletePart(wbPart.CalculationChainPart);

        wbPart.Workbook.Save();
        Console.WriteLine("ODBC connection stripped.");
    }

    // ── HELPERS ────────────────────────────────────────────────────────────

    private static (int year, int month) ParseReportMonth(string[] args)
    {
        if (args.Length == 0)
            return (DateTime.Now.Year, DateTime.Now.Month);

        if (DateTime.TryParseExact(args[0], "yyyy-MM",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            return (dt.Year, dt.Month);

        throw new ArgumentException(
            $"Invalid month argument '{args[0]}'. Expected: YYYY-MM");
    }

    private static string FindMostRecentFile(SftpClient sftp, string dir, string pattern)
    {
        var regex = WildcardToRegex(pattern);
        var matches = sftp.ListDirectory(dir)
            .Where(f => !f.IsDirectory
                        && !f.Name.StartsWith('.')
                        && regex.IsMatch(f.Name))
            .OrderByDescending(f => f.LastWriteTime)
            .ToList();

        if (matches.Count == 0)
            throw new FileNotFoundException(
                $"No files matching '{pattern}' in {dir}");

        Console.WriteLine($"Found {matches.Count} matching file(s). Using: {matches[0].Name}");
        return $"{dir.TrimEnd('/')}/{matches[0].Name}";
    }

    private static Regex WildcardToRegex(string pattern)
    {
        var escaped = "^" + Regex.Escape(pattern)
            .Replace("\\*", ".*").Replace("\\?", ".") + "$";
        return new Regex(escaped, RegexOptions.IgnoreCase);
    }

    private static string DownloadCsv(SftpClient sftp, string remotePath)
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"revrpt_{Guid.NewGuid():N}.csv");
        Console.WriteLine($"Downloading {remotePath} ...");
        using (var fs = File.Create(tmp))
            sftp.DownloadFile(remotePath, fs);
        Console.WriteLine($"Downloaded to {tmp}");
        return tmp;
    }

    private static (string[] headers, List<string[]> rows) ReadCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        if (lines.Length == 0)
            throw new InvalidOperationException("CSV file is empty.");

        var headers = ParseCsvLine(lines[0]);
        var rows = new List<string[]>(lines.Length - 1);

        for (int i = 1; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (line.Length == 0) continue;
            rows.Add(ParseCsvLine(line));
        }

        return (headers, rows);
    }

    private static string[] ParseCsvLine(string line)
    {
        var fields = new List<string>();
        bool inQuotes = false;
        var sb = new System.Text.StringBuilder();

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    sb.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                fields.Add(sb.ToString().Trim());
                sb.Clear();
            }
            else
            {
                sb.Append(c);
            }
        }
        fields.Add(sb.ToString().Trim());
        return fields.ToArray();
    }

    private static int ResolveDateColumnIndex(string[] headers, string dateColName, int dateColIndex)
    {
        var idx = Array.FindIndex(headers, h =>
            string.Equals(h, dateColName, StringComparison.OrdinalIgnoreCase));

        if (idx >= 0) return idx;

        if (headers.Length > dateColIndex)
        {
            Console.WriteLine(
                $"'{dateColName}' header not found. Falling back to column index {dateColIndex}: '{headers[dateColIndex]}'");
            return dateColIndex;
        }

        throw new InvalidOperationException(
            $"'{dateColName}' not found and CSV has fewer than {dateColIndex + 1} columns.");
    }

    private static List<string[]> FilterByMonth(
        List<string[]> rows, int dateIdx, int year, int month)
    {
        var kept = new List<string[]>();
        int unparsed = 0;

        foreach (var row in rows)
        {
            if (dateIdx >= row.Length
                || !DateTime.TryParse(row[dateIdx], CultureInfo.InvariantCulture,
                                      DateTimeStyles.None, out var dt))
            {
                unparsed++;
                continue;
            }
            if (dt.Year == year && dt.Month == month)
                kept.Add(row);
        }

        var removed = rows.Count - kept.Count;
        Console.WriteLine(
            $"Filtered: {rows.Count} rows -> {kept.Count} rows " +
            $"({removed} removed, {unparsed} unparseable dates)");

        if (kept.Count == 0)
            Console.WriteLine(
                "WARNING: Zero rows remain after filtering. " +
                "Check if Transaction_Date is parsed correctly.");

        return kept;
    }

    private static void WriteToTransTab(
    List<string[]> rows, string[] headers, int dateIdx,
    string reportPath, string sheetName, int numCols)
    {
        if (!File.Exists(reportPath))
            throw new FileNotFoundException($"Report file not found: {reportPath}");

        using var doc = SpreadsheetDocument.Open(reportPath, true);
        var workbookPart = doc.WorkbookPart!;

        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");

        var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;

        // ══════════════════════════════════════════════════════
        // STEP 1: Nuke everything problematic
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 1: Cleaning workbook...");

        // Remove ALL queryTable parts from ALL worksheets
        foreach (var part in workbookPart.WorksheetParts)
        {
            foreach (var qtp in part.QueryTableParts.ToList())
                part.DeletePart(qtp);
            foreach (var el in part.Worksheet.Elements()
                .Where(e => e.LocalName == "queryTableParts").ToList())
                el.Remove();
            part.Worksheet.Save();
        }

        // Remove ALL connections
        foreach (var conn in workbookPart.Workbook.Descendants<Connection>().ToList())
            conn.Remove();

        // Delete calcChain
        if (workbookPart.CalculationChainPart != null)
            workbookPart.DeletePart(workbookPart.CalculationChainPart);

        // Remove ALL table definitions from this worksheet
        // Save the header info first
        var oldTablePart = wsPart.TableDefinitionParts.FirstOrDefault();
        int tableFirstCol = 1;
        int tableLastCol = 1;
        var tableColumnNames = new List<string>();

        if (oldTablePart?.Table != null)
        {
            // Parse existing table range to get column positions
            var refStr = oldTablePart.Table.Reference?.Value ?? "A1";
            // Read header names from the table definition
            if (oldTablePart.Table.TableColumns != null)
            {
                foreach (var tc in oldTablePart.Table.TableColumns.Elements<TableColumn>())
                    tableColumnNames.Add(tc.Name ?? "");
            }
        }

        // Read headers from actual worksheet row 1
        var headerRow = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == 1);
        if (headerRow == null)
            throw new InvalidOperationException("Header row not found.");

        var wsHeaders = new Dictionary<int, string>();
        foreach (var cell in headerRow.Elements<Cell>())
        {
            var colIdx = GetColIndex(cell.CellReference!);
            var val = GetCellStringValue(cell, workbookPart);
            if (!string.IsNullOrEmpty(val))
                wsHeaders[colIdx] = val.Trim();
        }

        int minCol = wsHeaders.Keys.Min();
        int maxCol = wsHeaders.Keys.Max();

        // Delete all table definition parts
        foreach (var tp in wsPart.TableDefinitionParts.ToList())
            wsPart.DeletePart(tp);

        // Remove tableParts element from worksheet XML
        foreach (var el in wsPart.Worksheet.Elements()
            .Where(e => e.LocalName == "tableParts").ToList())
            el.Remove();

        workbookPart.Workbook.Save();
        Console.WriteLine("  Cleaned: queryTables, connections, calcChain, table definition.");

        // ══════════════════════════════════════════════════════
        // STEP 2: Map CSV columns to worksheet columns
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 2: Mapping columns...");

        var csvToWsCol = new Dictionary<int, int>();
        for (int csvCol = 0; csvCol < headers.Length && csvCol < numCols; csvCol++)
        {
            var csvHeader = headers[csvCol].Trim();
            var match = wsHeaders.FirstOrDefault(kv =>
                string.Equals(kv.Value, csvHeader, StringComparison.OrdinalIgnoreCase));
            if (match.Value != null)
                csvToWsCol[csvCol] = match.Key;
        }
        Console.WriteLine($"  Mapped {csvToWsCol.Count} CSV columns.");

        // Find "Actual Primary Agent" column
        int formulaColIdx = wsHeaders
            .FirstOrDefault(kv => kv.Value.Equals("Actual Primary Agent", StringComparison.OrdinalIgnoreCase)).Key;

        // ══════════════════════════════════════════════════════
        // STEP 3: Delete from the first-business-day row onward
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 3: Locating cutoff row...");

        // Resolve the Post_Date column index in the worksheet.
        int postDateCol = wsHeaders
            .FirstOrDefault(kv => kv.Value.Equals("Post_Date", StringComparison.OrdinalIgnoreCase)).Key;
        if (postDateCol == 0)
            throw new InvalidOperationException("Post_Date column not found in worksheet header.");

        // First business day of the target month (weekend-only; add holidays if you have a source).
        var firstBiz = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        while (firstBiz.DayOfWeek == DayOfWeek.Saturday || firstBiz.DayOfWeek == DayOfWeek.Sunday)
            firstBiz = firstBiz.AddDays(1);
        Console.WriteLine($"  First business day: {firstBiz:M/d/yyyy}");

        // Walk data rows; find the LAST row whose Post_Date == firstBiz.
        // Assumes the sheet is sorted ascending by Post_Date.
        uint? cutoffRowIndex = null;
        foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex?.Value > 1))
        {
            var cell = row.Elements<Cell>()
                .FirstOrDefault(c => GetColIndex(c.CellReference!) == postDateCol);
            if (cell == null) continue;

            var raw = GetCellStringValue(cell, workbookPart);
            DateTime pd;
            // Values may be serial numbers (DATEVALUE/number cells) or text.
            if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var serial))
                pd = DateTime.FromOADate(serial);
            else if (!DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out pd))
                continue;

            if (pd.Date == firstBiz.Date)
                cutoffRowIndex = row.RowIndex!.Value;  // keep updating → ends on the LAST match
        }

        if (cutoffRowIndex is null)
        {
            // Fallback: no row matched the first business day. Decide deliberately.
            Console.WriteLine("  WARNING: no row with Post_Date = first business day. Keeping all existing rows and appending.");
        }
        else
        {
            Console.WriteLine($"  Cutoff at row {cutoffRowIndex}. Removing that row and everything after.");
            foreach (var row in sheetData.Elements<Row>()
                .Where(r => r.RowIndex?.Value >= cutoffRowIndex).ToList())
                row.Remove();
        }

        // ══════════════════════════════════════════════════════
        // STEP 4: Write new data
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 4: Writing data...");
        var dataToWrite = rows.Skip(1).ToList();
        var dateColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Post_Date", "Transaction_Date", "Setup_Date"
        };

        // Start appending after the highest surviving row index (or row 2 if only the header remains).
        uint startRow = sheetData.Elements<Row>()
            .Select(r => r.RowIndex?.Value ?? 1)
            .DefaultIfEmpty(1u)
            .Max() + 1;
        if (startRow < 2) startRow = 2;

        for (int r = 0; r < dataToWrite.Count; r++)
        {
            var src = dataToWrite[r];
            var rowNum = (uint)(startRow + (uint)r);
            var dataRow = new Row { RowIndex = rowNum };

            // Collect cells with their worksheet column index, so we can sort before appending.
            var rowCells = new List<(int col, Cell cell)>();

            // Write mapped CSV values
            foreach (var kvp in csvToWsCol)
            {
                var csvCol = kvp.Key;
                var wsCol = kvp.Value;
                if (csvCol >= src.Length) continue;
                var val = src[csvCol];
                var cell = new Cell { CellReference = GetCellRef(wsCol, (int)rowNum) };
                var colName = wsHeaders.ContainsKey(wsCol) ? wsHeaders[wsCol] : "";
                if (dateColumns.Contains(colName)
                && DateTime.TryParse(val, CultureInfo.InvariantCulture,
                         DateTimeStyles.None, out var dt))
                {
                    var dateStr = dt.ToString("M/d/yyyy");
                    cell.CellFormula = new CellFormula($"DATEVALUE(\"{dateStr}\")");
                    cell.DataType = null;
                    cell.CellValue = null;
                    cell.StyleIndex = GetOrCreateDateStyle(workbookPart);
                }
                else if (decimal.TryParse(val, CultureInfo.InvariantCulture, out var num))
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(num.ToString(CultureInfo.InvariantCulture)); // normalized (fix #2)
                }
                else
                {
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(val ?? string.Empty);
                }
                rowCells.Add((wsCol, cell));
            }

            // Add formula cell for "Actual Primary Agent"
            if (formulaColIdx > 0)
            {
                var fCell = new Cell { CellReference = GetCellRef(formulaColIdx, (int)rowNum) };
                fCell.CellFormula = new CellFormula(
                    $"IF({GetCellRef(wsHeaders.FirstOrDefault(kv => kv.Value.Equals("Primary Agent", StringComparison.OrdinalIgnoreCase)).Key, (int)rowNum)}=\"\"," +
                    $"VLOOKUP({GetCellRef(wsHeaders.FirstOrDefault(kv => kv.Value.Equals("Client Group", StringComparison.OrdinalIgnoreCase)).Key, (int)rowNum)},Salesman!$H$1:$I$19,2,FALSE)," +
                    $"{GetCellRef(wsHeaders.FirstOrDefault(kv => kv.Value.Equals("Primary Agent", StringComparison.OrdinalIgnoreCase)).Key, (int)rowNum)})");
                rowCells.Add((formulaColIdx, fCell));
            }

            // Append in ascending column order — required by the OpenXML spec.
            foreach (var (_, cell) in rowCells.OrderBy(c => c.col))
                dataRow.Append(cell);

            sheetData.Append(dataRow);
        }
        Console.WriteLine($"  Wrote {dataToWrite.Count} rows.");

        // ══════════════════════════════════════════════════════
        // STEP 5: Recreate table from scratch
        // ══════════════════════════════════════════════════════
        Console.WriteLine("STEP 5: Recreating table...");
        var newLastRow = (int)startRow + dataToWrite.Count - 1;
        var firstColLetter = new string(GetCellRef(minCol, 1).TakeWhile(char.IsLetter).ToArray());
        var lastColLetter = new string(GetCellRef(maxCol, 1).TakeWhile(char.IsLetter).ToArray());
        var tableRef = $"{firstColLetter}1:{lastColLetter}{newLastRow}";

        // Create new table definition part
        var newTablePart = wsPart.AddNewPart<TableDefinitionPart>();
        var tableId = (uint)(workbookPart.WorksheetParts
            .SelectMany(wp => wp.TableDefinitionParts)
            .Count() + 1);

        var tableColumns = new TableColumns { Count = (uint)(maxCol - minCol + 1) };
        for (int c = minCol; c <= maxCol; c++)
        {
            var colName = wsHeaders.ContainsKey(c) ? wsHeaders[c] : $"Column{c + 1}";
            tableColumns.Append(new TableColumn
            {
                Id = (uint)(c - minCol + 1),
                Name = colName
            });
        }

        newTablePart.Table = new Table
        {
            Id = tableId,
            Name = "Table_Query_from_CCDS_ODBC_AIX",
            DisplayName = "Table_Query_from_CCDS_ODBC_AIX",
            Reference = tableRef,
            TotalsRowShown = false,
            AutoFilter = new AutoFilter { Reference = tableRef },
            TableColumns = tableColumns,
            TableStyleInfo = new TableStyleInfo
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            }
        };
        newTablePart.Table.Save();

        // Add tableParts element to worksheet
        var tableParts = new TableParts { Count = 1 };
        tableParts.Append(new TablePart
        {
            Id = wsPart.GetIdOfPart(newTablePart)
        });

        // Insert tableParts at the end of the worksheet (before closing tag)
        wsPart.Worksheet.Append(tableParts);

        wsPart.Worksheet.Save();
        workbookPart.Workbook.Save();

        Console.WriteLine($"  Table recreated: {tableRef}");
        Console.WriteLine($"Done. Wrote {dataToWrite.Count} rows to '{sheetName}'.");
    }

    private static uint GetOrCreateDateStyle(WorkbookPart workbookPart)
    {
        var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
        var cellFormats = stylesheet.CellFormats!;

        // Check if a format using numFmtId 14 (short date) already exists
        uint idx = 0;
        foreach (var cf in cellFormats.Elements<CellFormat>())
        {
            if (cf.NumberFormatId?.Value == 14)
                return idx;
            idx++;
        }

        // Create one
        cellFormats.Append(new CellFormat
        {
            NumberFormatId = 14,
            ApplyNumberFormat = true
        });
        cellFormats.Count = (uint)cellFormats.Elements<CellFormat>().Count();
        stylesheet.Save();

        return (uint)(cellFormats.Elements<CellFormat>().Count() - 1);
    }

    // Helper: get string value from a cell (handles shared strings)
    private static string GetCellStringValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellValue == null) return string.Empty;

        var val = cell.CellValue.Text;

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sst != null && int.TryParse(val, out var idx))
            {
                var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? val;
            }
        }
        return val;
    }

    // Helper: extract column index from cell reference like "C5" → 2
    private static int GetColIndex(string cellRef)
    {
        int col = 0;
        foreach (var ch in cellRef)
        {
            if (char.IsLetter(ch))
                col = col * 26 + (char.ToUpper(ch) - 'A' + 1);
            else
                break;
        }
        return col - 1;
    }

    //private static void WriteToTransTab(
    //List<string[]> rows, string[] headers, int dateIdx,
    //string reportPath, string sheetName, int numCols)
    //{
    //    if (!File.Exists(reportPath))
    //        throw new FileNotFoundException($"Report file not found: {reportPath}");

    //    using var doc = SpreadsheetDocument.Open(reportPath, true);
    //    var workbookPart = doc.WorkbookPart
    //        ?? throw new InvalidOperationException("Workbook is empty.");

    //    var sheet = workbookPart.Workbook.Descendants<Sheet>()
    //        .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
    //        ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");

    //    var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    //    var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;

    //    // Preserve row 1 (table headers), remove everything else
    //    var existingRows = sheetData.Elements<Row>().ToList();
    //    Row? headerRow = existingRows.FirstOrDefault(r => r.RowIndex?.Value == 1);

    //    sheetData.RemoveAllChildren<Row>();

    //    // Put header row back
    //    if (headerRow != null)
    //        sheetData.Append(headerRow);

    //    var colsToWrite = Math.Min(numCols, headers.Length);

    //    // Write filtered data starting at row 2
    //    for (int r = 0; r < rows.Count; r++)
    //    {
    //        var src = rows[r];
    //        var dataRow = new Row { RowIndex = (uint)(r + 2) };

    //        for (int c = 0; c < colsToWrite && c < src.Length; c++)
    //        {
    //            var val = src[c];
    //            var cell = new Cell { CellReference = GetCellRef(c, r + 2) };

    //            if (c == dateIdx
    //                && DateTime.TryParse(val, CultureInfo.InvariantCulture,
    //                                     DateTimeStyles.None, out var dt))
    //            {
    //                cell.DataType = CellValues.String;
    //                cell.CellValue = new CellValue(dt.ToString("yyyy/MM/dd"));
    //            }
    //            else if (decimal.TryParse(val, CultureInfo.InvariantCulture, out var num))
    //            {
    //                cell.DataType = CellValues.Number;
    //                cell.CellValue = new CellValue(val);
    //            }
    //            else
    //            {
    //                cell.DataType = CellValues.String;
    //                cell.CellValue = new CellValue(val ?? string.Empty);
    //            }

    //            dataRow.Append(cell);
    //        }

    //        sheetData.Append(dataRow);
    //    }

    //    // Update the table reference range to match new row count
    //    foreach (var tablePart in wsPart.TableDefinitionParts)
    //    {
    //        var table = tablePart.Table;
    //        if (table?.Reference != null)
    //        {
    //            // Table range: A1 to last column + last data row
    //            var lastCol = GetCellRef(colsToWrite - 1, 1).TrimEnd('1');
    //            var newRef = $"A1:{lastCol}{rows.Count + 1}";
    //            table.Reference = newRef;

    //            // Update AutoFilter range too if present
    //            if (table.AutoFilter != null)
    //                table.AutoFilter.Reference = newRef;

    //            table.Save();
    //            Console.WriteLine($"Updated table range to {newRef}");
    //        }
    //    }

    //    wsPart.Worksheet.Save();
    //    Console.WriteLine($"Wrote {rows.Count} rows to '{sheetName}' tab in {reportPath}");
    //}

    private static string GetCellRef(int colIndex, int rowNum)
    {
        string col = "";
        int c = colIndex;
        while (c >= 0)
        {
            col = (char)('A' + c % 26) + col;
            c = c / 26 - 1;
        }
        return col + rowNum;
    }

    private static async Task RunCIBCPortalAutomation()
    {
        Console.WriteLine("Launching browser...");
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(
            new BrowserTypeLaunchOptions
            {
                Headless = false // set true for background automation
            });

        // IgnoreHTTPSErrors must be on the context, not the page
        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            IgnoreHTTPSErrors = true,
            AcceptDownloads = true
        });
        var page = await context.NewPageAsync();

        await page.GotoAsync(
            CIBCurl,
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });
        Console.WriteLine("Page loaded.");

        // ── LOGIN ─────────────────────────────
        //await page.WaitForSelectorAsync("#username");
        // Clear and type username
        // First page
        await page.WaitForSelectorAsync("#username");

        await page.ClickAsync("#username");
        await page.FillAsync("#username", "");
        await page.Keyboard.TypeAsync(CIBCuser, new KeyboardTypeOptions { Delay = 50 });

        await page.ClickAsync("#random");
        await page.FillAsync("#random", "");
        await page.Keyboard.TypeAsync(CIBCpass, new KeyboardTypeOptions { Delay = 50 });

        // Submit and catch the popup
        var popupTask = context.WaitForPageAsync();
        await page.ClickAsync("input[type='submit']");
        var popup = await popupTask;

        // Second page (popup) — same credentials
        await popup.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await popup.WaitForSelectorAsync("#username");

        await popup.ClickAsync("#username");
        await popup.FillAsync("#username", "");
        await popup.Keyboard.TypeAsync(CIBCuser, new KeyboardTypeOptions { Delay = 50 });

        await popup.ClickAsync("#random");
        await popup.FillAsync("#random", "");
        await popup.Keyboard.TypeAsync(CIBCpass, new KeyboardTypeOptions { Delay = 50 });

        await popup.PressAsync("#random", "Enter");

        Console.WriteLine("Logged in.");

        // Wait for the post-login UI (the left nav with "Reports" in your screenshot)
        await popup.WaitForSelectorAsync("#REPORTS");
        await popup.ClickAsync("#REPORTS");
        await popup.WaitForTimeoutAsync(3000);

        // Get the frame directly by name
        var mainFrame = popup.Frames.FirstOrDefault(f => f.Name == "CIBCCRM_MAIN");

        await mainFrame.WaitForSelectorAsync("select[name='report_seq']");
        await mainFrame.SelectOptionAsync("select[name='report_seq']", new SelectOptionValue { Value = "16" });
        Console.WriteLine("Report selected.");

        //await mainFrame.PressAsync("select[name='report_seq']", "Enter");

        // Generate Report opens a new window — capture it
        var reportPageTask = context.WaitForPageAsync();
        await mainFrame.ClickAsync("text=Generate Report");

        var reportPage = await reportPageTask;

        await reportPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await reportPage.WaitForTimeoutAsync(3000);

        // Find which frame has the date inputs
        IFrame dateFrame = null;
        foreach (var f in reportPage.Frames)
        {
            try
            {
                var found = await f.EvalOnSelectorAllAsync<int>("input#start_date", "els => els.length");
                if (found > 0)
                {
                    dateFrame = f;
                    Console.WriteLine($"Found date inputs in frame: name='{f.Name}'");
                    break;
                }
            }
            catch { }
        }

        // Set date range: first of current month to today
        var today = DateTime.Today;
        var firstOfMonth = new DateTime(today.Year, today.Month, 1).ToString("yyyy-MM-dd");
        var todayStr = today.ToString("yyyy-MM-dd");

        await dateFrame.EvalOnSelectorAsync("#start_date",
            $"el => {{ el.value = '{firstOfMonth}'; el.dispatchEvent(new Event('change')); }}");

        await dateFrame.EvalOnSelectorAsync("#end_date",
            $"el => {{ el.value = '{todayStr}'; el.dispatchEvent(new Event('change')); }}");

        Console.WriteLine($"Date range set: {firstOfMonth} to {todayStr}");

        // Click Render Report (inspect to confirm the selector)
        await dateFrame.ClickAsync("text=Render Report");
        Console.WriteLine("Render Report clicked.");

        // Reports can take a while to render. Wait for either a known element on
        // the rendered report OR the export/download button to appear.
        // Wait 10 seconds for the report to render
        await reportPage.WaitForTimeoutAsync(10000);
        Console.WriteLine("Wait complete. Searching for export/download options...");

        // Find the frame with the export button
        IFrame exportFrame = null;
        foreach (var f in reportPage.Frames)
        {
            try
            {
                var found = await f.EvalOnSelectorAllAsync<int>("#_export_button", "els => els.length");
                if (found > 0)
                {
                    exportFrame = f;
                    Console.WriteLine($"Found export button in frame: name='{f.Name}'");
                    break;
                }
            }
            catch { }
        }

        // Click export and capture the download
        var download = await reportPage.RunAndWaitForDownloadAsync(async () =>
        {
            await exportFrame.ClickAsync("#_export_button");
        });

        var savePath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\CIBC POST.csv";

        // Delete existing file if present
        if (File.Exists(savePath))
            File.Delete(savePath);

        await download.SaveAsAsync(savePath);
        Console.WriteLine($"Downloaded: {savePath}");

        Console.WriteLine("CIBC Automation completed.");
        await Task.Delay(5000);
        await browser.CloseAsync();
    }

    //private static async Task TestTDPortalHttp()
    //{
    //    Console.WriteLine("Testing TD Portal HTTP...\n");

    //    var handler = new HttpClientHandler
    //    {
    //        ServerCertificateCustomValidationCallback = (_, _, _, _) => true,
    //        AllowAutoRedirect = true,
    //        CookieContainer = new System.Net.CookieContainer()
    //    };

    //    using var client = new HttpClient(handler);
    //    var baseUrl = "https://10.21.178.100";
    //    var userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";
    //    client.DefaultRequestHeaders.Add("User-Agent", userAgent);

    //    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

    //    try
    //    {
    //        // ── STEP 1: GET login page (establish session cookie) ──
    //        Console.WriteLine("STEP 1: GET login page...");
    //        var getResponse = await client.GetAsync(TDurl, cts.Token);
    //        Console.WriteLine($"  Status: {(int)getResponse.StatusCode}");

    //        var cookies = handler.CookieContainer.GetCookies(new Uri(baseUrl));
    //        Console.WriteLine($"  Session cookie: {cookies[0]?.Name} = {cookies[0]?.Value}");

    //        // ── STEP 2: POST browser verification (mimic requestVerifyBrowser JS) ──
    //        Console.WriteLine("\nSTEP 2: POST browser_verification.php...");

    //        // This is what bowser.parse() produces for our user agent
    //        var browserInfo = new
    //        {
    //            request_dtl = new
    //            {
    //                name = "Chrome",
    //                version = "131.0.0.0",
    //                engine = "Blink",
    //                user_agent = userAgent,
    //                os = "Windows",
    //                os_version = "10"
    //            },
    //            user_agent = userAgent
    //        };

    //        var jsonContent = new StringContent(
    //            System.Text.Json.JsonSerializer.Serialize(browserInfo),
    //            System.Text.Encoding.UTF8,
    //            "application/json"  // bowser sends as JSON, not form data
    //        );

    //        var verifyRequest = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/browser_verification.php")
    //        {
    //            Content = jsonContent
    //        };
    //        verifyRequest.Headers.Add("Referer", TDurl);
    //        verifyRequest.Headers.Add("Origin", baseUrl);
    //        verifyRequest.Headers.Add("X-Requested-With", "XMLHttpRequest"); // it's an XHR call

    //        var verifyResponse = await client.SendAsync(verifyRequest, cts.Token);
    //        var verifyBody = await verifyResponse.Content.ReadAsStringAsync();
    //        Console.WriteLine($"  Status: {(int)verifyResponse.StatusCode}");
    //        Console.WriteLine($"  Response: {verifyBody}");

    //        // Check cookies after verification
    //        cookies = handler.CookieContainer.GetCookies(new Uri(baseUrl));
    //        Console.WriteLine($"  Cookies: {cookies.Count}");
    //        foreach (System.Net.Cookie c in cookies)
    //            Console.WriteLine($"    {c.Name} = {c.Value}");

    //        // ── STEP 3: POST login with mode=attempt ──
    //        Console.WriteLine("\nSTEP 3: POST login...");
    //        var postData = new FormUrlEncodedContent(new[]
    //        {
    //        new KeyValuePair<string, string>("username", TDuser),
    //        new KeyValuePair<string, string>("random", TDpass),
    //        new KeyValuePair<string, string>("mode", "attempt"),
    //    });

    //        var loginRequest = new HttpRequestMessage(HttpMethod.Post, TDurl)
    //        {
    //            Content = postData
    //        };
    //        loginRequest.Headers.Add("Referer", TDurl);
    //        loginRequest.Headers.Add("Origin", baseUrl);

    //        var loginResponse = await client.SendAsync(loginRequest, cts.Token);
    //        var loginBody = await loginResponse.Content.ReadAsStringAsync();

    //        Console.WriteLine($"  Status: {(int)loginResponse.StatusCode}");
    //        Console.WriteLine($"  Body length: {loginBody.Length}");

    //        // Check for redirect (success) vs same page (failure)
    //        if (loginResponse.Headers.Location != null)
    //            Console.WriteLine($"  Redirect: {loginResponse.Headers.Location}");

    //        // Check for error messages
    //        var hasError = loginBody.Contains("invalid", StringComparison.OrdinalIgnoreCase)
    //            || loginBody.Contains("incorrect", StringComparison.OrdinalIgnoreCase);
    //        var hasLoginForm = loginBody.Contains("id=\"username\"", StringComparison.OrdinalIgnoreCase);

    //        Console.WriteLine($"  Has error message: {hasError}");
    //        Console.WriteLine($"  Still on login page: {hasLoginForm}");
    //        Console.WriteLine($"  First 1000 chars:");
    //        Console.WriteLine(loginBody[..Math.Min(1000, loginBody.Length)]);

    //        // Updated cookies
    //        cookies = handler.CookieContainer.GetCookies(new Uri(baseUrl));
    //        Console.WriteLine($"\n  Cookies after login: {cookies.Count}");
    //        foreach (System.Net.Cookie c in cookies)
    //            Console.WriteLine($"    {c.Name} = {c.Value}");

    //        // ── STEP 4: If login succeeded, try main page ──
    //        if (!hasLoginForm)
    //        {
    //            Console.WriteLine("\nSTEP 4: Accessing main page...");
    //            var mainResponse = await client.GetAsync($"{baseUrl}/main.phtml", cts.Token);
    //            var mainBody = await mainResponse.Content.ReadAsStringAsync();
    //            Console.WriteLine($"  Status: {(int)mainResponse.StatusCode}");
    //            Console.WriteLine($"  Body length: {mainBody.Length}");
    //            Console.WriteLine($"  First 500 chars:");
    //            Console.WriteLine(mainBody[..Math.Min(500, mainBody.Length)]);
    //        }
    //    }
    //    catch (OperationCanceledException)
    //    {
    //        Console.WriteLine("  TIMEOUT");
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine($"  ERROR: {ex.GetType().Name}: {ex.Message}");
    //    }

    //    Console.WriteLine("\nDone.");
    //}

    private static async Task TestTDPortalHttp()
    {
        Console.WriteLine("Testing TD Portal HTTP...\n");

        var handler = new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (_, _, _, _) => true,
            AllowAutoRedirect = false,
            CookieContainer = new System.Net.CookieContainer()
        };

        using var client = new HttpClient(handler);
        var baseUrl = "https://10.21.178.100";
        var userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36";
        client.DefaultRequestHeaders.Add("User-Agent", userAgent);

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

        try
        {
            // ── STEP 1: GET login page ──
            Console.WriteLine("STEP 1: GET login page...");
            var getResponse = await client.GetAsync(TDurl, cts.Token);
            Console.WriteLine($"  Status: {(int)getResponse.StatusCode}");

            var cookies = handler.CookieContainer.GetCookies(new Uri(baseUrl));
            Console.WriteLine($"  Session: {cookies[0]?.Name} = {cookies[0]?.Value}");

            var loginHtml = await getResponse.Content.ReadAsStringAsync();
            Console.WriteLine($"  Status: {(int)getResponse.StatusCode}");

            // Analyze the form
            Console.WriteLine("\n  Analyzing login form...");


            // Find <form tag
            var formIdx = loginHtml.IndexOf("<form", StringComparison.OrdinalIgnoreCase);
            if (formIdx >= 0)
            {
                var formEnd = loginHtml.IndexOf(">", formIdx);
                Console.WriteLine($"  Form tag: {loginHtml[formIdx..(formEnd + 1)]}");
            }

            // Find ALL input fields (not just hidden)
            var allInputs = System.Text.RegularExpressions.Regex.Matches(
                loginHtml, @"<input[^>]*>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            Console.WriteLine($"  All inputs ({allInputs.Count}):");
            foreach (System.Text.RegularExpressions.Match m in allInputs)
                Console.WriteLine($"    {m.Value}");

            // Find the submit button
            var submitIdx = loginHtml.IndexOf("onsubmit", StringComparison.OrdinalIgnoreCase);
            while (submitIdx >= 0)
            {
                var lineStart = Math.Max(0, submitIdx - 100);
                var lineEnd = Math.Min(loginHtml.Length, submitIdx + 100);
                var line = loginHtml[lineStart..lineEnd];
                if (line.Contains("<input", StringComparison.OrdinalIgnoreCase)
                    || line.Contains("<button", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  Submit element: ...{line.Trim()}...");
                    break;
                }
                submitIdx = loginHtml.IndexOf("submit", submitIdx + 1, StringComparison.OrdinalIgnoreCase);
            }

            // Check for any password hashing/transformation
            foreach (var keyword in new[] { "md5", "sha", "hash", "encrypt", "encode", "btoa", "digest" })
            {
                var idx = loginHtml.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    var s = Math.Max(0, idx - 100);
                    var e = Math.Min(loginHtml.Length, idx + 200);
                    Console.WriteLine($"\n  Found '{keyword}':");
                    Console.WriteLine(loginHtml[s..e]);
                }
            }

            // ── STEP 2: POST browser verification ──
            // Mimic exactly: xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded; charset=utf-8')
            // requestString = "request_dtl=" + encodeURI(JSON.stringify(tssBrowserInfo)) + "&user_agent=" + encodeURI(JSON.stringify(userAgent))
            Console.WriteLine("\nSTEP 2: POST browser_verification.php...");

            var requestDtlJson = System.Text.Json.JsonSerializer.Serialize(new
            {
                name = "Chrome",
                version = "147.0.0.0",
                engine = "Blink",
                user_agent = userAgent,
                os = "Windows",
                os_version = "10"
            });

            // encodeURI(JSON.stringify(value)) — Uri.EscapeUriString mimics encodeURI
            var requestString = "request_dtl=" + Uri.EscapeDataString(requestDtlJson)
                              + "&user_agent=" + Uri.EscapeDataString("\"" + userAgent + "\"");

            Console.WriteLine($"  Sending: {requestString[..Math.Min(200, requestString.Length)]}...");

            var verifyRequest = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/browser_verification.php")
            {
                Content = new StringContent(
                    requestString,
                    System.Text.Encoding.UTF8,
                    "application/x-www-form-urlencoded")
            };
            verifyRequest.Headers.Add("Referer", TDurl);
            verifyRequest.Headers.Add("Origin", baseUrl);
            verifyRequest.Headers.Add("X-Requested-With", "XMLHttpRequest");

            var verifyResponse = await client.SendAsync(verifyRequest, cts.Token);
            var verifyBody = await verifyResponse.Content.ReadAsStringAsync();
            Console.WriteLine($"  Status: {(int)verifyResponse.StatusCode}");
            Console.WriteLine($"  Response: {verifyBody}");

            // ── STEP 3: POST login ──
            Console.WriteLine("\nSTEP 3: POST login...");
            var loginUrl = $"{baseUrl}/login.phtml";  // NOT index.phtml

            var postData = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("mode", "attempt"),
                new KeyValuePair<string, string>("username", TDuser),
                new KeyValuePair<string, string>("random", TDpass),
            });

            var loginRequest = new HttpRequestMessage(HttpMethod.Post, loginUrl)
            {
                Content = postData
            };
            loginRequest.Headers.Add("Referer", TDurl);
            loginRequest.Headers.Add("Origin", baseUrl);

            var loginResponse = await client.SendAsync(loginRequest, cts.Token);
            var loginBody = await loginResponse.Content.ReadAsStringAsync();

            Console.WriteLine($"  Status: {(int)loginResponse.StatusCode}");
            if (loginResponse.Headers.Location != null)
                Console.WriteLine($"  Redirect: {loginResponse.Headers.Location}");

            // Check if session cookie changed
            var step3Cookies = handler.CookieContainer.GetCookies(new Uri(baseUrl));
            foreach (System.Net.Cookie c in step3Cookies)
                Console.WriteLine($"  Cookie: {c.Name} = {c.Value}");

            Console.WriteLine($"  Status: {(int)loginResponse.StatusCode}");
            Console.WriteLine($"  Body length: {loginBody.Length}");

            var hasLoginForm = loginBody.Contains("id=\"username\"", StringComparison.OrdinalIgnoreCase)
                || loginBody.Contains("id='username'", StringComparison.OrdinalIgnoreCase);
            Console.WriteLine($"  Still on login page: {hasLoginForm}");

            // Find any error text near "invalid" or "incorrect"
            foreach (var pattern in new[] { "invalid", "incorrect", "denied", "failed", "error" })
            {
                var i = loginBody.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
                if (i >= 0)
                {
                    var ctx = loginBody[Math.Max(0, i - 30)..Math.Min(loginBody.Length, i + 80)];
                    Console.WriteLine($"  Found '{pattern}': ...{ctx}...");
                }
            }

            //Console.WriteLine($"\n  First 500 chars:");
            //Console.WriteLine(loginBody[..Math.Min(500, loginBody.Length)]);

            cookies = handler.CookieContainer.GetCookies(new Uri(baseUrl));
            Console.WriteLine($"\n  Cookies: {cookies.Count}");
            foreach (System.Net.Cookie c in cookies)
                Console.WriteLine($"    {c.Name} = {c.Value}");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("  TIMEOUT");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ERROR: {ex.GetType().Name}: {ex.Message}");
        }

        Console.WriteLine("\nDone.");
    }
    //private static async Task RunTDPortalAutomation()
    //{
    //    Console.WriteLine("Launching browser...");
    //    using var playwright = await Playwright.CreateAsync();
    //    await using var browser = await playwright.Chromium.LaunchAsync(
    //        new BrowserTypeLaunchOptions
    //        {
    //            Headless = false,
    //            Channel = "chrome",
    //            Args = new[] { "--disable-blink-features=AutomationControlled" }
    //        });

    //    var context = await browser.NewContextAsync(new BrowserNewContextOptions
    //    {
    //        IgnoreHTTPSErrors = true,
    //        AcceptDownloads = true
    //    });
    //    var page = await context.NewPageAsync();

    //    await page.AddInitScriptAsync(@"
    //    Object.defineProperty(navigator, 'webdriver', { get: () => false });
    //");

    //    await page.GotoAsync(
    //        TDurl,
    //        new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });
    //    await page.WaitForSelectorAsync("#username");
    //    await page.WaitForTimeoutAsync(3000);
    //    Console.WriteLine("Page loaded.");

    //    // ── FIRST PAGE LOGIN ──────────────────
    //    await page.ClickAsync("#username");
    //    await page.FillAsync("#username", "");
    //    await page.Keyboard.TypeAsync(
    //        TDuser,
    //        new KeyboardTypeOptions { Delay = 50 });

    //    await page.ClickAsync("#random");
    //    await page.FillAsync("#random", "");
    //    await page.Keyboard.TypeAsync(
    //        TDpass,
    //        new KeyboardTypeOptions { Delay = 50 });

    //    Console.WriteLine($"Username: {TDuser}");
    //    Console.WriteLine($"Password: {TDpass}");
    //    Console.WriteLine($"Password length: {TDpass.Length}");
    //    Console.WriteLine($"Password bytes: {string.Join(" ", System.Text.Encoding.UTF8.GetBytes(TDpass).Select(b => $"{b:X2}"))}");
    //    // Submit and try to catch popup
    //    IPage popup = null;
    //    try
    //    {
    //        var popupTask = context.WaitForPageAsync(new BrowserContextWaitForPageOptions
    //        {
    //            Timeout = 10000
    //        });
    //        await page.ClickAsync("input[type='submit']");
    //        popup = await popupTask;
    //        Console.WriteLine("Popup opened.");
    //    }
    //    catch
    //    {
    //        // No popup — login might have happened on the same page
    //        Console.WriteLine("No popup. Checking current page...");
    //        popup = page;
    //    }

    //    // ── SECOND PAGE LOGIN (if popup opened) ──
    //    await popup.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
    //    await popup.WaitForTimeoutAsync(2000);

    //    var hasLoginForm = await popup.EvalOnSelectorAllAsync<int>("#username", "els => els.length");
    //    if (hasLoginForm > 0 && popup != page)
    //    {
    //        Console.WriteLine("Second login form found. Filling...");
    //        await popup.ClickAsync("#username");
    //        await popup.FillAsync("#username", "");
    //        await popup.Keyboard.TypeAsync(
    //            Environment.GetEnvironmentVariable("PORTAL2_USER") ?? "REPLACE_ME",
    //            new KeyboardTypeOptions { Delay = 50 });

    //        await popup.ClickAsync("#random");
    //        await popup.FillAsync("#random", "");
    //        await popup.Keyboard.TypeAsync(
    //            Environment.GetEnvironmentVariable("PORTAL2_PASS") ?? "REPLACE_ME",
    //            new KeyboardTypeOptions { Delay = 50 });

    //        await popup.PressAsync("#random", "Enter");
    //        await popup.WaitForTimeoutAsync(3000);
    //    }

    //    // ── CHECK LOGIN ───────────────────────
    //    var hasError = await popup.EvalOnSelectorAllAsync<int>("div.alert", "els => els.length");
    //    if (hasError > 0)
    //    {
    //        var errorText = await popup.EvalOnSelectorAsync<string>("div.alert", "el => el.innerText");
    //        Console.WriteLine($"Login FAILED: {errorText}");
    //        Console.ReadLine();
    //        return;
    //    }
    //    Console.WriteLine("Logged in.");

    //    // ── CLICK REPORTS ─────────────────────
    //    await popup.WaitForSelectorAsync("#REPORTS");
    //    await popup.ClickAsync("#REPORTS");
    //    await popup.WaitForTimeoutAsync(3000);

    //    // ── SELECT REPORT ─────────────────────
    //    var mainFrame = popup.Frames.FirstOrDefault(f => f.Name == "CIBCCRM_MAIN");
    //    await mainFrame.WaitForSelectorAsync("select[name='report_seq']");
    //    await mainFrame.SelectOptionAsync("select[name='report_seq']",
    //        new SelectOptionValue { Value = "16" });
    //    Console.WriteLine("Report selected.");

    //    // ── GENERATE REPORT (new window) ──────
    //    var reportPageTask = context.WaitForPageAsync();
    //    await mainFrame.ClickAsync("text=Generate Report");
    //    var reportPage = await reportPageTask;

    //    await reportPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
    //    await reportPage.WaitForTimeoutAsync(3000);

    //    // ── SET DATE RANGE ────────────────────
    //    IFrame dateFrame = null;
    //    foreach (var f in reportPage.Frames)
    //    {
    //        try
    //        {
    //            var found = await f.EvalOnSelectorAllAsync<int>("input#start_date", "els => els.length");
    //            if (found > 0)
    //            {
    //                dateFrame = f;
    //                Console.WriteLine($"Found date inputs in frame: name='{f.Name}'");
    //                break;
    //            }
    //        }
    //        catch { }
    //    }

    //    var today = DateTime.Today;
    //    var firstOfMonth = new DateTime(today.Year, today.Month, 1).ToString("yyyy-MM-dd");
    //    var todayStr = today.ToString("yyyy-MM-dd");

    //    await dateFrame.EvalOnSelectorAsync("#start_date",
    //        $"el => {{ el.value = '{firstOfMonth}'; el.dispatchEvent(new Event('change')); }}");
    //    await dateFrame.EvalOnSelectorAsync("#end_date",
    //        $"el => {{ el.value = '{todayStr}'; el.dispatchEvent(new Event('change')); }}");
    //    Console.WriteLine($"Date range set: {firstOfMonth} to {todayStr}");

    //    // ── RENDER REPORT ─────────────────────
    //    await dateFrame.ClickAsync("text=Render Report");
    //    Console.WriteLine("Render Report clicked.");

    //    await reportPage.WaitForTimeoutAsync(30000);
    //    Console.WriteLine("Wait complete.");

    //    // ── EXPORT ────────────────────────────
    //    IFrame exportFrame = null;
    //    foreach (var f in reportPage.Frames)
    //    {
    //        try
    //        {
    //            var found = await f.EvalOnSelectorAllAsync<int>("#_export_button", "els => els.length");
    //            if (found > 0)
    //            {
    //                exportFrame = f;
    //                Console.WriteLine($"Found export button in frame: name='{f.Name}'");
    //                break;
    //            }
    //        }
    //        catch { }
    //    }

    //    var download = await reportPage.RunAndWaitForDownloadAsync(async () =>
    //    {
    //        await exportFrame.ClickAsync("#_export_button");
    //    });

    //    // TODO: Update save path for this portal's report
    //    var savePath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\PORTAL2 REPORT.csv";

    //    if (File.Exists(savePath))
    //        File.Delete(savePath);

    //    await download.SaveAsAsync(savePath);
    //    Console.WriteLine($"Downloaded: {savePath}");

    //    Console.WriteLine("Automation completed.");
    //    await Task.Delay(5000);
    //    await browser.CloseAsync();
    //}

    private static async Task RunPerformanceReportAutomation()
    {
        Console.WriteLine("Launching Performance Report automation...");
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(
            new BrowserTypeLaunchOptions
            {
                Headless = false
            });

        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            IgnoreHTTPSErrors = true,
            AcceptDownloads = true
        });
        var page = await context.NewPageAsync();

        await page.GotoAsync(
            CIBCurl,
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });
        await page.WaitForSelectorAsync("#username");
        Console.WriteLine("Page loaded.");

        // ── FIRST PAGE LOGIN ──────────────────
        await page.ClickAsync("#username");
        await page.FillAsync("#username", "");
        await page.Keyboard.TypeAsync(CIBCuser, new KeyboardTypeOptions { Delay = 50 });

        await page.ClickAsync("#random");
        await page.FillAsync("#random", "");
        await page.Keyboard.TypeAsync(CIBCpass, new KeyboardTypeOptions { Delay = 50 });

        var popupTask = context.WaitForPageAsync();
        await page.ClickAsync("input[type='submit']");
        var popup = await popupTask;

        // ── SECOND PAGE LOGIN ─────────────────
        await popup.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await popup.WaitForSelectorAsync("#username");

        await popup.ClickAsync("#username");
        await popup.FillAsync("#username", "");
        await popup.Keyboard.TypeAsync(CIBCuser, new KeyboardTypeOptions { Delay = 50 });

        await popup.ClickAsync("#random");
        await popup.FillAsync("#random", "");
        await popup.Keyboard.TypeAsync(CIBCpass, new KeyboardTypeOptions { Delay = 50 });

        await popup.PressAsync("#random", "Enter");
        Console.WriteLine("Logged in.");

        // ── CLICK REPORTS ─────────────────────
        await popup.WaitForSelectorAsync("#REPORTS");
        await popup.ClickAsync("#REPORTS");
        await popup.WaitForTimeoutAsync(3000);

        // ── SELECT PERFORMANCE REPORT (seq 38) ─
        var mainFrame = popup.Frames.FirstOrDefault(f => f.Name == "CIBCCRM_MAIN");
        await mainFrame.WaitForSelectorAsync("select[name='report_seq']");
        await mainFrame.SelectOptionAsync("select[name='report_seq']",
            new SelectOptionValue { Value = "38" });
        Console.WriteLine("Performance Report selected.");

        // ── GENERATE REPORT (new window) ──────
        var reportPageTask = context.WaitForPageAsync();
        await mainFrame.ClickAsync("text=Generate Report");
        var reportPage = await reportPageTask;

        await reportPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await reportPage.WaitForTimeoutAsync(3000);

        // ── FIND THE CRITERIA FRAME ───────────
        IFrame criteriaFrame = null;
        foreach (var f in reportPage.Frames)
        {
            try
            {
                var found = await f.EvalOnSelectorAllAsync<int>(
                    "input#start_date, input[name='start_date']",
                    "els => els.length");
                if (found > 0)
                {
                    criteriaFrame = f;
                    Console.WriteLine($"Found criteria in frame: name='{f.Name}'");
                    break;
                }
            }
            catch { }
        }

        // ── CALCULATE DATES ───────────────────
        // Start: last year, current month + 1, current day
        // End: today
        var today = DateTime.Today;
        var startDate = new DateTime(today.Year, today.Month, 1)
                    .AddYears(-1)
                    .AddMonths(1);
        var startStr = startDate.ToString("yyyy-MM-dd");
        var endStr = today.ToString("yyyy-MM-dd");

        Console.WriteLine($"Date range: {startStr} to {endStr}");

        // ── SET ASSIGNED DATE RANGE ───────────
        // Debug: find all date inputs to get exact selectors
        var dateInputs = await criteriaFrame.EvalOnSelectorAllAsync<string>(
            "input[type='text']",
            "els => els.map(e => `id=${e.id} name=${e.name} value=${e.value}`).join('\\n')");
        Console.WriteLine($"Date inputs found:\n{dateInputs}");

        // Set Assigned Date Range (first pair of date inputs)
        await criteriaFrame.EvalOnSelectorAsync("#start_date",
            $"el => {{ el.value = '{startStr}'; el.dispatchEvent(new Event('change')); }}");
        await criteriaFrame.EvalOnSelectorAsync("#end_date",
            $"el => {{ el.value = '{endStr}'; el.dispatchEvent(new Event('change')); }}");

        // Set Payment Date Range (second pair — adjust selectors based on debug output above)
        // These IDs are guesses — check the debug output for actual IDs
        await criteriaFrame.EvalOnSelectorAsync("#start_pdate",
            $"el => {{ el.value = '{startStr}'; el.dispatchEvent(new Event('change')); }}");
        await criteriaFrame.EvalOnSelectorAsync("#end_pdate",
            $"el => {{ el.value = '{endStr}'; el.dispatchEvent(new Event('change')); }}");

        Console.WriteLine("Dates set.");

        // ── HELPER: Run report for a specific agency ──
        async Task ExportForAgency(string agencyValue, string filename)
        {
            // Select agency in phase_to_view dropdown
            // First deselect all, then select the target agency
            await criteriaFrame.EvalOnSelectorAsync("#phase_to_view",
                $@"el => {{
                for (var i = 0; i < el.options.length; i++) {{
                    el.options[i].selected = (el.options[i].value === '{agencyValue}');
                }}
                el.dispatchEvent(new Event('change'));
            }}");
            Console.WriteLine($"Selected {agencyValue}.");

            // Click Render Report
            await criteriaFrame.ClickAsync("text=Render Report");
            Console.WriteLine("Render Report clicked. Waiting 2.5 minutes...");

            await reportPage.WaitForTimeoutAsync(150000);

            // Find export button
            IFrame exportFrame = null;
            foreach (var f in reportPage.Frames)
            {
                try
                {
                    var found = await f.EvalOnSelectorAllAsync<int>(
                        "#_export_button", "els => els.length");
                    if (found > 0)
                    {
                        exportFrame = f;
                        break;
                    }
                }
                catch { }
            }

            // Export
            var download = await reportPage.RunAndWaitForDownloadAsync(async () =>
            {
                await exportFrame.ClickAsync("#_export_button");
            });

            var savePath = $@"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\VTs\{filename}";

            if (File.Exists(savePath))
                File.Delete(savePath);

            await download.SaveAsAsync(savePath);
            Console.WriteLine($"Downloaded: {savePath}");
        }

        // ── EXPORT AGENCY2 ───────────────────
        await ExportForAgency("AGENCY2", "2.csv");

        // ── GO BACK TO CRITERIA ──────────────
        IFrame editFrame = null;
        foreach (var f in reportPage.Frames)
        {
            try
            {
                var found = await f.EvalOnSelectorAllAsync<int>(
                    "text=Edit Criteria", "els => els.length");
                if (found > 0)
                {
                    editFrame = f;
                    break;
                }
            }
            catch { }
        }
        await editFrame.ClickAsync("text=Edit Criteria");
        await reportPage.WaitForTimeoutAsync(3000);

        // Re-find the criteria frame (it may have reloaded)
        criteriaFrame = null;
        foreach (var f in reportPage.Frames)
        {
            try
            {
                var found = await f.EvalOnSelectorAllAsync<int>(
                    "#phase_to_view", "els => els.length");
                if (found > 0)
                {
                    criteriaFrame = f;
                    Console.WriteLine($"Re-found criteria frame: name='{f.Name}'");
                    break;
                }
            }
            catch { }
        }

        // ── EXPORT AGENCY4 ───────────────────
        await ExportForAgency("AGENCY4", "4.csv");

        Console.WriteLine("Performance Report automation completed.");
        await Task.Delay(5000);
        await browser.CloseAsync();
    }
}