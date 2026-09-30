using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using Microsoft.Playwright;
using Renci.SshNet;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Path = System.IO.Path;
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
    [STAThread]
    private static async Task Main(string[] args)
    {
        try
        {
            //Saving all daily ACE transaction to Rev Report
            //GetDailyRevenueACETransactions(args);

            //TD Portal automation to download the report CSV directly
            //RunTDPortalAutomation().GetAwaiter().GetResult();

            //CIBC Portal automation to download the report CSV directly
            //RunCIBCPortalAutomation().GetAwaiter().GetResult();

            //ApplyManualAdjustments();

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
    private static void RunOnSta(Action work)
    {
        Exception err = null;
        var t = new Thread(() =>
        {
            ExcelMessageFilter.Register();      // your existing class
            try { work(); }
            catch (Exception ex) { err = ex; }
            finally { ExcelMessageFilter.Revoke(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (err != null) throw new Exception(err.Message, err);
    }

    static string DescribeLock(string path)
    {
        try
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                return null;   // we got exclusive access — nothing is holding it
        }
        catch (IOException ex)
        {
            return $"LOCKED: {ex.Message}";
        }
        catch (UnauthorizedAccessException ex)
        {
            return $"PERMISSION DENIED (not a lock): {ex.Message}";
        }
    }

    private static async Task RunTDPortalAutomation()
    {
        Console.WriteLine("Launching browser...");
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(
            new BrowserTypeLaunchOptions { Headless = false });

        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            IgnoreHTTPSErrors = true,
            AcceptDownloads = true
        });
        var page = await context.NewPageAsync();

        await page.GotoAsync(TDurl, new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });
        Console.WriteLine("Page loaded.");

        // ── FIRST LOGIN ──
        await page.WaitForSelectorAsync("#username");
        await page.ClickAsync("#username");
        await page.FillAsync("#username", "");
        await page.Keyboard.TypeAsync(TDuser, new KeyboardTypeOptions { Delay = 50 });
        await page.ClickAsync("#random");
        await page.FillAsync("#random", "");
        await page.Keyboard.TypeAsync(TDpass, new KeyboardTypeOptions { Delay = 50 });

        // Capture the popup that opens on submit
        var popupTask = context.WaitForPageAsync();
        await page.ClickAsync("#ok");
        var popup = await popupTask;

        // ── SECOND LOGIN (in the popup) ──
        await popup.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await popup.WaitForSelectorAsync("#username");

        await popup.ClickAsync("#username");
        await popup.FillAsync("#username", "");
        await popup.Keyboard.TypeAsync(TDuser, new KeyboardTypeOptions { Delay = 50 });

        await popup.ClickAsync("#random");
        await popup.FillAsync("#random", "");
        await popup.Keyboard.TypeAsync(TDpass, new KeyboardTypeOptions { Delay = 50 });

        await popup.ClickAsync("#ok");
        await popup.WaitForLoadStateAsync(LoadState.DOMContentLoaded);

        // ── OPEN REPORTS ──
        await popup.WaitForSelectorAsync("#REPORTS");
        await popup.ClickAsync("#REPORTS");
        await popup.WaitForTimeoutAsync(1000);

        // Report UI loads into the TDVMDRS_MAIN frame
        var mainFrame = popup.Frames.FirstOrDefault(f => f.Name == "TDVMDRS_MAIN");

        // ── SELECT "Collection and Commission Report" (value 16) ──
        await mainFrame.SelectOptionAsync("select[name='report_seq']",
    new SelectOptionValue { Value = "16" });
        Console.WriteLine("Report selected.");

        var criteriaFrame = popup.Frame("report_criteria");

        await mainFrame.ClickAsync("text=Generate Report");
        Console.WriteLine("Generate Report clicked.");

        // Give report_check.phtml's POST and the resulting GET time to complete
        await popup.WaitForTimeoutAsync(5000);

        //Console.WriteLine($"Total pages in context: {context.Pages.Count}");
        //foreach (var p in context.Pages)
        //    Console.WriteLine($"  Page: {p.Url}");

        IFrame reportFrame = null;
        IPage reportPage = null;
        foreach (var p in context.Pages)
        {
            var f = p.Frames.FirstOrDefault(fr => fr.Url.Contains("collection_and_commission.phtml"));
            if (f != null) { reportFrame = f; reportPage = p; break; }
        }

        if (reportFrame != null)
        {
            // Confirm the widget is actually there before touching it
            var hasDateFields = await reportFrame.EvalOnSelectorAllAsync<int>("input#y_start_date", "els => els.length");
            Console.WriteLine($"Date fields found on report page: {hasDateFields}");

            if (hasDateFields > 0)
            {
                var today = DateTime.Today;
                async Task SetDateAsync(string prefix, DateTime date)
                {
                    await reportFrame.FillAsync($"#y_{prefix}", date.ToString("yyyy"));
                    await reportFrame.FillAsync($"#m_{prefix}", date.ToString("MM"));
                    await reportFrame.FillAsync($"#d_{prefix}", date.ToString("dd"));
                    await reportFrame.Locator($"#d_{prefix}").PressAsync("Tab");
                }
                await SetDateAsync("start_date", new DateTime(today.Year, today.Month, 1));
                await SetDateAsync("end_date", today);

                var hiddenStart = await reportFrame.EvalOnSelectorAsync<string>("#start_date", "el => el.value");
                var hiddenEnd = await reportFrame.EvalOnSelectorAsync<string>("#end_date", "el => el.value");
                Console.WriteLine($"Hidden fields after fill: start_date='{hiddenStart}', end_date='{hiddenEnd}'");
            }
        }
        else
        {
            Console.WriteLine("collection_and_commission.phtml not found in any page/frame after 5s wait.");
            // Dump every frame in every page so we can see what's actually there
            foreach (var p in context.Pages)
            {
                Console.WriteLine($"--- Frames in page {p.Url} ---");
                foreach (var f in p.Frames)
                    Console.WriteLine($"  '{f.Name}' | {f.Url}");
            }
        }

        // ── CLICK RENDER REPORT (applies the date filter you just set) ──
        await reportFrame.ClickAsync("#render_button");
        Console.WriteLine("Render Report clicked.");

        // This page loads pleaseWait.js/Progress.js — there's likely a busy
        // indicator while it re-renders with the new date range. Wait for
        // network activity to settle rather than a blind fixed delay.
        await reportPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await reportPage.WaitForTimeoutAsync(1000); // small buffer on top

        // ── FIND EXPORT — search reportPage's frames, not popup's ──
        IFrame exportFrame = null;
        foreach (var f in reportPage.Frames)
        {
            try
            {
                var found = await f.EvalOnSelectorAllAsync<int>("#export_link", "els => els.length");
                if (found > 0) { exportFrame = f; Console.WriteLine($"Found export button in frame: '{f.Name}'"); break; }
            }
            catch { }
        }
        // If it's not inside a sub-frame at all, check reportFrame directly as a fallback
        if (exportFrame == null)
        {
            var foundOnReportFrame = await reportFrame.EvalOnSelectorAllAsync<int>("#export_link", "els => els.length");
            if (foundOnReportFrame > 0) exportFrame = reportFrame;
        }
        if (exportFrame == null) throw new Exception("Could not find a frame containing #export_link.");

        var download = await reportPage.RunAndWaitForDownloadAsync(async () =>
        {
            await exportFrame.ClickAsync("#export_link");
        });

        var suggested = download.SuggestedFilename;
        var ext = Path.GetExtension(suggested);
        var destFileName = $"TD_{DateTime.Today:MM-dd-yyyy}{ext}";
        var destFolder = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\TD";
        var destPath = Path.Combine(destFolder, destFileName);
        await download.SaveAsAsync(destPath);
        Console.WriteLine($"Exported to: {destPath}");

        CopyReportValuesToRevenueWorkbook();
    }

    private static void CopyReportValuesToRevenueWorkbook()
    {
        // ── 1. LOCATE AND VERIFY TODAY'S DOWNLOADED FILE ──
        var sourceFolder = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\TD";
        var sourceFileName = $"TD_{DateTime.Today:MM-dd-yyyy}.csv";
        var sourcePath = Path.Combine(sourceFolder, sourceFileName);

        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Today's downloaded file was not found: {sourcePath}");

        var sourceLastWrite = File.GetLastWriteTime(sourcePath);
        if (sourceLastWrite.Date != DateTime.Today)
            throw new InvalidOperationException(
                $"File {sourcePath} exists but was last modified {sourceLastWrite:yyyy-MM-dd}, not today. Aborting — refusing to copy stale data.");

        Console.WriteLine($"Source file verified as today's: {sourcePath}");

        // ── 2. READ VALUES A3:F<lastRow> FROM THE SOURCE ──
        List<object[]> sourceRows = new List<object[]>();
        using (var reader = new StreamReader(sourcePath))
        using (var csv = new CsvHelper.CsvReader(reader, System.Globalization.CultureInfo.InvariantCulture))
        {
            int lineNum = 0;
            while (csv.Read())
            {
                lineNum++;
                if (lineNum < 2) continue; // skip rows 1-2

                var row = new object[6];
                // replace these with the ACTUAL column meanings — don't auto-detect
                row[0] = ParseDecimal(csv.GetField(0));                     // A: text/ID → string
                row[1] = ParseDecimal(csv.GetField(1));                     // B: text → string
                row[2] = ParseDecimal(csv.GetField(2));        // C: amount → double
                row[3] = ParseDecimal(csv.GetField(3));        // D: amount → double
                row[4] = ParseDecimal(csv.GetField(4));         // E: date → DateTime?
                row[5] = ParseDecimal(csv.GetField(5));                     // F: text → string
                sourceRows.Add(row);
            }
        }

        static object ParseDecimal(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;

            var cleaned = raw.Trim();
            bool negParens = cleaned.StartsWith("(") && cleaned.EndsWith(")");
            if (negParens) cleaned = cleaned.Substring(1, cleaned.Length - 2);
            cleaned = cleaned.Replace("$", "").Replace(",", "").Trim();

            if (double.TryParse(cleaned, NumberStyles.Float | NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out var d))
                return negParens ? -Math.Abs(d) : d;

            return raw; // genuinely unparseable — keep as text rather than silently dropping it
        }

        // ── 3. FIND THE DESTINATION WORKBOOK ──
        var destFolder = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report";
        var monthName = DateTime.Today.ToString("MMMM").ToUpper(); // e.g. "SEPTEMBER"
        var yearStr = DateTime.Today.Year.ToString();

        var candidates = Directory.GetFiles(destFolder, "*.xlsx")
            .Where(f => Path.GetFileName(f).ToUpper().Contains(monthName))
            .Where(f => Path.GetFileName(f).Contains(yearStr))
            .Where(f => Path.GetFileName(f).Contains("V8")) // hardcoded per your spec — see note above
            .OrderByDescending(f => File.GetLastWriteTime(f))
            .ToList();

        if (candidates.Count == 0)
            throw new FileNotFoundException(
                $"No destination workbook found in {destFolder} matching month '{monthName}', year '{yearStr}', and 'V8'.");

        var destPath = candidates.First();
        Console.WriteLine($"Destination workbook: {destPath}");

        var lockInfo = DescribeLock(destPath);
        Console.WriteLine(lockInfo ?? "File is free — no lock.");

        if (candidates.Count > 1)
            Console.WriteLine($"  WARNING: {candidates.Count} files matched — picked the most recently modified. Others: {string.Join(", ", candidates.Skip(1).Select(Path.GetFileName))}");
        var formulasBefore = CountFormulas(destPath);

        var localPath = Path.Combine(Path.GetTempPath(), $"revreport_{Guid.NewGuid():N}.xlsx");
        File.Copy(destPath, localPath, true);
        Console.WriteLine($"Working locally: {localPath}");
        // ================= PHASE 1: OpenXML — write CSV data into TD DATA =================
        try
        {
            using (var destDoc = SpreadsheetDocument.Open(destPath, true))
            {
                var workbookPart = destDoc.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name == "TD DATA");
                if (sheet == null) throw new InvalidOperationException("Worksheet 'TD DATA' not found.");

                var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var worksheet = wsPart.Worksheet;

                var totalRow = (uint)(2 + sourceRows.Count - 1);
                var cell = worksheet.Descendants<Cell>()
                             .FirstOrDefault(c => c.CellReference == $"I{totalRow}");
                Console.WriteLine($"Verify TD DATA I{totalRow} = {cell?.CellValue?.Text}  (expected {sourceRows.Last()[1]})");

                for (int i = 0; i < sourceRows.Count; i++)
                    for (int c = 0; c < 6; c++)
                        SetCellValue(worksheet, (uint)(2 + i), (uint)(8 + c), sourceRows[i][c]);

                // Excel won't recalc formulas that depend on H:M unless told to
                var calc = workbookPart.Workbook.CalculationProperties
                           ?? workbookPart.Workbook.AppendChild(new CalculationProperties());
                calc.FullCalculationOnLoad = true;
                worksheet.Save();
                workbookPart.Workbook.Save();
            }   // handle released HERE
            Console.WriteLine("PHASE 1 OK: TD DATA written, file closed.");
            WaitForFileFree(localPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("PHASE 1 FAILED:");
            Console.WriteLine(ex.ToString());
            throw;
        }

        // ================= verify the handle really dropped =================
        GC.Collect();
        GC.WaitForPendingFinalizers();
        WaitForFileFree(localPath);

        // ================= PHASE 2: Interop — Tim Katis + day snapshot =================
        try
        {
            // ---------- 3. verify no formulas were destroyed ----------
            var formulasAfter = CountFormulas(localPath);
            foreach (var kv in formulasBefore)
            {
                var after = formulasAfter.TryGetValue(kv.Key, out var a) ? a : 0;
                if (after != kv.Value)
                    Console.WriteLine($"  WARNING: sheet '{kv.Key}' formula count {kv.Value} → {after}");
            }

            // ---------- 4. Interop: TD DATA C1:C19 → Tim Katis, then snapshot F into today's column ----------
            PushToTimKatis(localPath);
            WaitForFileFree(localPath);
            WaitForFileFree(destPath);
            File.Copy(localPath, destPath, true);
            Console.WriteLine($"Copied back to {destPath}");
            File.Delete(localPath);
            Console.WriteLine("PHASE 2 OK.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("PHASE 2 FAILED:");
            Console.WriteLine(ex.ToString());
            throw;
        }
        finally
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
    static void FindTodayDayCell(dynamic ws, int today, out int headerRow, out int dayCol)
    {
        headerRow = -1; dayCol = -1;

        dynamic used = ws.UsedRange;
        object[,] vals = used.Value2;          // ONE marshal, 1-based [row, col]
        int rowOffset = (int)used.Row;
        int colOffset = (int)used.Column;
        int rows = vals.GetLength(0);
        int cols = vals.GetLength(1);

        for (int i = 1; i <= rows; i++)
        {
            var hits = new List<KeyValuePair<int, int>>();   // (col index in array, day value)
            for (int j = 1; j <= cols; j++)
            {
                object v = vals[i, j];
                if (v == null) continue;
                double d;
                if (!double.TryParse(Convert.ToString(v), out d)) continue;
                if (d < 1 || d > 31 || d != Math.Floor(d)) continue;
                hits.Add(new KeyValuePair<int, int>(j, (int)d));
            }

            int run = 1;
            for (int k = 1; k < hits.Count; k++)
            {
                run = (hits[k].Value == hits[k - 1].Value + 1) ? run + 1 : 1;
                if (run >= 4)                                  // looks like a day-header row
                {
                    foreach (var h in hits)
                    {
                        if (h.Value == today)
                        {
                            headerRow = rowOffset + i - 1;
                            dayCol = colOffset + h.Key - 1;
                            return;
                        }
                    }
                    break;                                     // right row shape, today not present — keep looking
                }
            }
        }
    }
    static Dictionary<string, int> CountFormulas(string path)
    {
        var counts = new Dictionary<string, int>();
        using (var doc = SpreadsheetDocument.Open(path, false))
        {
            var wbPart = doc.WorkbookPart;
            foreach (var sheet in wbPart.Workbook.Descendants<Sheet>())
            {
                var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
                int n = wsPart.Worksheet.Descendants<CellFormula>().Count();
                counts[sheet.Name] = n;
            }
        }
        return counts;
    }
    static void WaitForFileFree(string path, int timeoutMs = 30000)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        string last = null;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            last = DescribeLock(path);
            if (last == null)
            {
                if (sw.ElapsedMilliseconds > 0)
                    Console.WriteLine($"  File free after {sw.ElapsedMilliseconds} ms.");
                return;
            }
            System.Threading.Thread.Sleep(250);
        }
        throw new IOException($"File still locked after {timeoutMs} ms: {last}");
    }
    static void PushToTimKatis(string destPath)
    {
        const string TD_SHEET = "TD DATA";
        const string DEST_SHEET = "Tim Katis";
        const int DEST_TOTAL_ROW = 180;
        const int DEST_FIRST_ROW = 181;
        const int DEST_LAST_ROW = 198;
        const int DEST_CODE_COL = 4;   // D
        const int DEST_VALUE_COL = 5;   // E
        const int DEST_RESULT_COL = 6;   // F

        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new InvalidOperationException("Excel is not installed on this machine.");

        dynamic app = Activator.CreateInstance(excelType);
        dynamic wb = null;
        try
        {
            ExcelMessageFilter.Register();

            app.Visible = true;
            app.DisplayAlerts = false;

            wb = app.Workbooks.Open(destPath);
            dynamic tdWs = wb.Sheets[TD_SHEET];
            dynamic dstWs = wb.Sheets[DEST_SHEET];

            // ---- code → value map from TD DATA A2:A19 / C2:C19 ----
            var tdMap = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int r = 2; r <= 19; r++)
            {
                object rawCode = tdWs.Cells[r, 1].Value2;      // dynamic → object HERE
                string code = Convert.ToString(rawCode)?.Trim();
                if (string.IsNullOrWhiteSpace(code)) continue;

                object rawVal = tdWs.Cells[r, 3].Value2;
                tdMap[code] = rawVal == null ? 0d : Convert.ToDouble(rawVal);
            }

            object rawTotal = tdWs.Cells[1, 3].Value2;          // C1
            dstWs.Cells[DEST_TOTAL_ROW, DEST_VALUE_COL].Value2 =
                rawTotal == null ? 0d : Convert.ToDouble(rawTotal);

            // ---- matched BY CODE, not by position ----
            var problems = new List<string>();
            var unused = new HashSet<string>(tdMap.Keys, StringComparer.OrdinalIgnoreCase);

            for (int r = DEST_FIRST_ROW; r <= DEST_LAST_ROW; r++)
            {
                object rawCode = dstWs.Cells[r, DEST_CODE_COL].Value2;
                string code = Convert.ToString(rawCode)?.Trim();
                if (string.IsNullOrWhiteSpace(code)) continue;

                double val;                                     // explicit, not var
                if (tdMap.TryGetValue(code, out val))
                {
                    dstWs.Cells[r, DEST_VALUE_COL].Value2 = val;
                    unused.Remove(code);
                }
                else
                {
                    dstWs.Cells[r, DEST_VALUE_COL].Value2 = 0d;
                    problems.Add($"  row {r}: code '{code}' not in {TD_SHEET} → wrote 0");
                }
            }
            foreach (string left in unused)
                problems.Add($"  {TD_SHEET} '{left}' ({tdMap[left]:N2}) has no row on {DEST_SHEET} → NOT written");

            if (problems.Count > 0)
            {
                Console.WriteLine("WARNING: code mismatches between tabs:");
                problems.ForEach(Console.WriteLine);
            }

            app.CalculateFullRebuild();

            // ---- NEW ----
            int today = DateTime.Today.Day;
            int headerRow, dayCol;
            FindTodayDayCell(dstWs, today, out headerRow, out dayCol);

            if (dayCol < 0)
                throw new InvalidOperationException($"Could not locate a day-header row containing '{today}' on {DEST_SHEET}.");

            Console.WriteLine($"Day-header row {headerRow}; day {today} → column {dayCol}. Verify before trusting this.");

            // ---- snapshot F into that column ----
            for (int r = DEST_TOTAL_ROW; r <= DEST_LAST_ROW; r++)
            {
                object f = dstWs.Cells[r, DEST_RESULT_COL].Value2;
                dstWs.Cells[r, dayCol].Value2 = f == null ? 0d : Convert.ToDouble(f);
            }

            wb.Save();
            Console.WriteLine($"Tim Katis: E{DEST_TOTAL_ROW}:E{DEST_LAST_ROW} updated; F snapshot → column {dayCol} (day {today}).");
        }
        finally
        {
            if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); wb = null; }
            if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); app = null; }
            ExcelMessageFilter.Revoke();
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
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
            //using (var doc = SpreadsheetDocument.Open(revReportFile, isEditable: true))
            using (var doc = SpreadsheetDocument.Open(revReportFile, true, new OpenSettings { AutoSave = false }))
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
                //manPart.Worksheet.Save();

                // ── agent-tab apply ──
                // Pre-scan every agent tab's Desk column (D) once: normalized name -> matches across all tabs
                var deskIndex = new Dictionary<string, List<(string tab, WorksheetPart part, uint row)>>();
                foreach (var tabName in agentTabs)
                {
                    var (aPart, aData) = GetSheet(doc, tabName);
                    foreach (var r in aData.Elements<Row>())
                    {
                        if (r.RowIndex == null) continue;
                        var dCell = r.Elements<Cell>().FirstOrDefault(c => ColIndex(c.CellReference) == DeskCol);
                        var key = Normalize(GetCellText(dCell, doc));
                        if (key.Length == 0) continue;
                        if (!deskIndex.TryGetValue(key, out var lst)) deskIndex[key] = lst = new();
                        lst.Add((tabName, aPart, r.RowIndex.Value));
                    }
                }

                var touched = new HashSet<WorksheetPart>();
                int applied = 0, skipped = 0;
                Log("\nApplying to agent tabs...");
                foreach (var a in adjustments)
                {
                    int day = a.Date.Day;
                    string rec = $"[{a.Date:MM/dd} debt {a.Debt} ${a.Comm:F2} from '{a.From}' to '{a.ToAgent}']";

                    deskIndex.TryGetValue(Normalize(a.From), out var fromM);
                    deskIndex.TryGetValue(Normalize(a.ToAgent), out var toM);

                    string reason =
                        (fromM == null || fromM.Count == 0) ? $"FROM '{a.From}' not found in any Desk (D) column"
                      : (fromM.Count > 1) ? $"FROM '{a.From}' matches {fromM.Count} desks (ambiguous)"
                      : (toM == null || toM.Count == 0) ? $"TO '{a.ToAgent}' not found in any Desk (D) column"
                      : (toM.Count > 1) ? $"TO '{a.ToAgent}' matches {toM.Count} desks (ambiguous)"
                      : null;

                    if (reason == null)
                    {
                        var (fTab, fPart, fRow) = fromM[0];
                        var (tTab, tPart, tRow) = toM[0];
                        var fCol = FindDayColumn(fPart, doc, day, DayStartCol, DayEndCol);
                        var tCol = FindDayColumn(tPart, doc, day, DayStartCol, DayEndCol);
                        if (fCol == null) reason = $"day {day} column not found on '{fTab}'";
                        else if (tCol == null) reason = $"day {day} column not found on '{tTab}'";
                        else
                        {
                            var fCell = GetOrCreateCell(GetOrCreateRow(fPart.Worksheet.GetFirstChild<SheetData>(), fRow), (int)fCol.Value);
                            var tCell = GetOrCreateCell(GetOrCreateRow(tPart.Worksheet.GetFirstChild<SheetData>(), tRow), (int)tCol.Value);
                            double fOld = GetNumber(fCell) ?? 0, tOld = GetNumber(tCell) ?? 0;
                            SetNumber(fCell, fOld - a.Comm);      // deduct from FROM desk
                            SetNumber(tCell, tOld + a.Comm);      // credit TO desk
                            touched.Add(fPart); touched.Add(tPart);
                            applied++;
                            Log($"  OK   {rec}: {fTab} {ColName((int)fCol.Value)}{fRow} {fOld:F2}->{fOld - a.Comm:F2}  |  {tTab} {ColName((int)tCol.Value)}{tRow} {tOld:F2}->{tOld + a.Comm:F2}");
                        }
                    }

                    if (reason != null) { Log($"  SKIP {rec}: {reason}"); skipped++; }
                }
                Log($"\nApplied {applied}, Skipped {skipped}.");

                // ── persist everything at once (nothing was saved before this point) ──
                manPart.Worksheet.Save();
                foreach (var p in touched) p.Worksheet.Save();

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
    // ── HELPER: fill the y/m/d text boxes for a given date field prefix ──
    // helper — note the name is ToXLCellValue, NOT XLCellValue
    static XLCellValue ToXLCellValue(object o) => o switch
    {
        null => Blank.Value,
        string s => s,
        double d => d,
        bool b => b,
        DateTime dt when dt.Year is >= 1900 and <= 9999 => dt,
        DateTime dt => dt.ToString("yyyy-MM-dd"), // out-of-range date → write as text, don't throw
        _ => o.ToString()
    };
    static async Task SetDateFieldAsync(IFrame frame, string prefix, DateTime date)
    {
        // prefix is "start_date" or "end_date"
        await frame.FillAsync($"#y_{prefix}", date.ToString("yyyy"));
        await frame.FillAsync($"#m_{prefix}", date.ToString("MM"));
        await frame.FillAsync($"#d_{prefix}", date.ToString("dd"));

        // These widgets (InputDate) typically sync the hidden field on blur —
        // force it by tabbing off the last field.
        await frame.Locator($"#d_{prefix}").PressAsync("Tab");

        // Belt-and-suspenders: also set the hidden input directly and fire
        // input/change in case the widget doesn't listen for blur on the
        // individual boxes. Harmless if the blur already did it.
        await frame.EvaluateAsync(@"(id, val) => {
        const el = document.getElementById(id);
        if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }", new object[] { prefix, date.ToString("yyyy-MM-dd") });
    }
    static string Normalize(string s)   // "Richard Carlsen" and "RICHARD.CARLSEN" both -> "RICHARDCARLSEN"
    {
        if (string.IsNullOrEmpty(s)) return "";
        var sb = new StringBuilder();
        foreach (var ch in s)
            if (!char.IsWhiteSpace(ch) && ch != '.' && ch != '\u00A0')
                sb.Append(char.ToUpperInvariant(ch));
        return sb.ToString();
    }

    static string GetCellText(Cell c, SpreadsheetDocument doc)   // resolves shared strings (column D is text)
    {
        if (c == null) return "";
        if (c.DataType?.Value == CellValues.SharedString && c.CellValue != null && int.TryParse(c.CellValue.Text, out var i))
            return doc.WorkbookPart.SharedStringTablePart?.SharedStringTable?.ElementAtOrDefault(i)?.InnerText ?? "";
        if (c.DataType?.Value == CellValues.InlineString) return c.InlineString?.InnerText ?? "";
        return c.CellValue?.Text ?? "";
    }

    static uint? FindDayColumn(WorksheetPart part, SpreadsheetDocument doc, int day, int startCol, int endCol)
    {
        var header = part.Worksheet.GetFirstChild<SheetData>()?
            .Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == 1);
        if (header == null) return null;
        for (int col = startCol; col <= endCol; col++)
        {
            var cell = header.Elements<Cell>().FirstOrDefault(c => ColIndex(c.CellReference) == col);
            if (int.TryParse(GetCellText(cell, doc), out var n) && n == day) return (uint)col;
        }
        return null;
    }
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
        Console.WriteLine($"  Apartment: {Thread.CurrentThread.GetApartmentState()}");
        var othersPath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others";
        var revReportPath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report";
        var csvPath = Path.Combine(othersPath, "CIBC POST.csv");

        // The real .xls only — no conversion
        var calcPath = Directory.GetFiles(othersPath, "CIBC DRS CALCULATOR*.xls")
            .Where(f => f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            .Single();

        // SAME rev report as TD — reuse your TD finder
        var month = DateTime.Today.ToString("MMMM").ToUpper();
        var year = DateTime.Today.Year.ToString();
        var revPath = Directory.GetFiles(revReportPath, "*.xlsx")
            .Where(f => { var n = Path.GetFileName(f).ToUpper(); return n.Contains(month) && n.Contains(year) && n.Contains("V8"); })
            .OrderByDescending(File.GetLastWriteTime)
            .First();

        Console.WriteLine($"Calculator: {calcPath}\nRev Report: {revPath}");

        // ── CSV: last row with data in D, parsed properly (quoted commas) ──
        int lastDataRow = 0; double csvTotal = 0;
        using (var reader = new StreamReader(csvPath))
        using (var csv = new CsvHelper.CsvReader(reader, CultureInfo.InvariantCulture))
        {
            int line = 0;
            while (csv.Read())
            {
                line++;
                var d = csv.GetField(3);
                if (!string.IsNullOrWhiteSpace(d))
                {
                    lastDataRow = line;
                    double.TryParse(d.Replace("$", "").Replace(",", "").Trim(),
                        NumberStyles.Any, CultureInfo.InvariantCulture, out csvTotal);
                }
            }
        }
        Console.WriteLine($"CSV last data row D{lastDataRow} = {csvTotal}");

        // ── Work on local copies ──
        var calcLocal = Path.Combine(Path.GetTempPath(), $"cibccalc_{Guid.NewGuid():N}.xls");
        var revLocal = Path.Combine(Path.GetTempPath(), $"revreport_{Guid.NewGuid():N}.xlsx");
        File.Copy(calcPath, calcLocal, true);
        File.Copy(revPath, revLocal, true);

        Excel.Application xl = null;
        Excel.Workbook csvWb = null, calcWb = null, revWb = null;
        try
        {
            xl = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
                AskToUpdateLinks = false,
                ScreenUpdating = false,
                EnableEvents = false
            };

            // Open the CSV first so any external refs in the calculator resolve to today's data
            csvWb = xl.Workbooks.Open(csvPath, ReadOnly: true);
            calcWb = xl.Workbooks.Open(calcLocal, UpdateLinks: 0);

            // Log and re-point every link to CIBC POST.csv at the open network CSV
            var links = calcWb.LinkSources(Excel.XlLink.xlExcelLinks) as Array;
            if (links == null)
                Console.WriteLine("  WARNING: calculator has NO external links — B isn't reading the CSV at all.");
            else
                foreach (string link in links)
                {
                    Console.WriteLine($"  Link found: {link}");
                    if (Path.GetFileName(link).Equals("CIBC POST.csv", StringComparison.OrdinalIgnoreCase)
                        && !link.Equals(csvPath, StringComparison.OrdinalIgnoreCase))
                    {
                        calcWb.ChangeLink(link, csvPath, Excel.XlLinkType.xlLinkTypeExcelLinks);
                        Console.WriteLine($"    → re-pointed to {csvPath}");
                    }
                }

            xl.CalculateFull();

            var calc = (Excel.Worksheet)calcWb.Worksheets["CIBC"];

            // ── Read A/B from row 9 until TOTAL ──
            const int startRow = 9;
            int endRow = startRow;
            var names = new List<string>(); var comm = new List<double>();
            for (int r = startRow; r <= 70; r++)
            {
                var name = Convert.ToString(((Excel.Range)calc.Cells[r, 1]).Value2)?.Trim() ?? "";
                if (name.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) break;
                var v = ((Excel.Range)calc.Cells[r, 2]).Value2;
                names.Add(name);
                comm.Add(v is double d ? d : 0);
                endRow = r;
            }
            Console.WriteLine($"COMM rows {startRow}-{endRow} ({comm.Count})");

            // ── Paste COMM values to D (values only, one range write) ──
            var dVals = new object[comm.Count, 1];
            for (int i = 0; i < comm.Count; i++) dVals[i, 0] = comm[i];
            calc.Range[$"D{startRow}:D{endRow}"].Value2 = dVals;

            // ── B72: keep it a formula ──
            // Check this matches what the original B72 formula looked like
            calc.Range["B72"].Formula = $"=B71-'[CIBC POST.csv]CIBC POST'!$D${lastDataRow}";

            xl.Calculate();

            var diff = Convert.ToDouble(calc.Range["B72"].Value2);
            // ── Abort if calculator doesn't tie to CSV ──
            if (Math.Abs(diff) >= 0.01)
                throw new InvalidOperationException($"B72 mismatch {diff:C2} — not touching Rev Report.");

            // ══════════════════════════════════════════════
            // STEP A: B9:B72 → D9:D72 (values only)
            // ══════════════════════════════════════════════
            const int calcFirst = 9, calcLast = 69;
            int n = calcLast - calcFirst + 1;                       // 64 rows

            var bRange = calc.Range[$"B{calcFirst}:B{calcLast}"];
            var bVals = (object[,])bRange.Value2;                  // 1-based [row, col]
            calc.Range[$"D{calcFirst}:D{calcLast}"].Value2 = bVals; // values, no formulas, no clipboard
            Console.WriteLine($"  Copied B{calcFirst}:B{calcLast} → D{calcFirst}:D{calcLast}");

            // read D back as numbers
            var dNums = new double[n];
            var cNames = new string[n];
            var aVals = (object[,])calc.Range[$"A{calcFirst}:A{calcLast}"].Value2;
            for (int i = 0; i < n; i++)
            {
                dNums[i] = bVals[i + 1, 1] is double d ? d : 0;
                cNames[i] = Convert.ToString(aVals[i + 1, 1])?.Trim() ?? "";
            }

            // ══════════════════════════════════════════════
            // STEP B: Open Rev Report (calculator + CSV stay open)
            // ══════════════════════════════════════════════
            File.Copy(revPath, revLocal, true);
            revWb = xl.Workbooks.Open(revLocal, UpdateLinks: 0);
            xl.CalculateFull();

            Excel.Worksheet tim = null;
            foreach (Excel.Worksheet w in revWb.Worksheets)
                if (w.Name.IndexOf("Tim Katis", StringComparison.OrdinalIgnoreCase) >= 0) { tim = w; break; }
            if (tim == null) throw new InvalidOperationException("'Tim Katis' tab not found.");

            // ══════════════════════════════════════════════
            // STEP C: Read Tim Katis G, compute D − G
            // ══════════════════════════════════════════════
            const int timFirst = 116;                               // CONFIRM 117 vs 118
            int timLast = timFirst + n - 1;
            var gVals = (object[,])tim.Range[$"G{timFirst}:G{timLast}"].Value2;
            var tNames = (object[,])tim.Range[$"B{timFirst}:B{timLast}"].Value2;

            var result = new object[n, 1];
            for (int i = 0; i < n; i++)
            {
                double g = gVals[i + 1, 1] is double gd ? gd : 0;
                result[i, 0] = dNums[i] - g;
                Console.WriteLine($"  calc r{calcFirst + i,-3} {cNames[i],-25} | tim r{timFirst + i,-3} {Convert.ToString(tNames[i + 1, 1]),-40} | {dNums[i],12:N2} - {g,12:N2} = {dNums[i] - g,12:N2}");
            }

            // ══════════════════════════════════════════════
            // STEP D: Today's column in Tim Katis → paste results
            // ══════════════════════════════════════════════
            int targetCol = -1;
            var header = (object[,])tim.Range["I1:AM1"].Value2;
            for (int c = 1; c <= header.GetLength(1); c++)
                if (header[1, c] is double h && (h < 32 ? (int)h : DateTime.FromOADate(h).Day) == DateTime.Today.Day)
                { targetCol = 8 + c; break; }                       // I = col 9
            if (targetCol == -1) throw new InvalidOperationException($"Day {DateTime.Today.Day} not in Tim Katis I1:AM1.");

            tim.Range[tim.Cells[timFirst, targetCol], tim.Cells[timLast, targetCol]].Value2 = result;
            Console.WriteLine($"  Pasted {n} results to Tim Katis col {targetCol}, rows {timFirst}-{timLast}");

            // ══════════════════════════════════════════════
            // SAVE everything, THEN close
            // ══════════════════════════════════════════════
            calcWb.CheckCompatibility = false;
            calcWb.Save();
            revWb.Save();

            revWb.Close(false); revWb = null;
            calcWb.Close(false); calcWb = null;
            csvWb.Close(false); csvWb = null;
        }
        finally
        {
            if (calcWb != null) calcWb.Close(false);
            if (csvWb != null) csvWb.Close(false);
            if (revWb != null) revWb.Close(false);
            if (xl != null) { xl.Quit(); Marshal.ReleaseComObject(xl); }
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        WaitForFileFree(calcPath); File.Copy(calcLocal, calcPath, true);
        WaitForFileFree(revPath); File.Copy(revLocal, revPath, true);
        File.Delete(calcLocal); File.Delete(revLocal);
        Console.WriteLine("CIBC DRS Calculator complete.");
    }
    //private static void RunCibcDrsCalculator()
    //{
    //    Console.WriteLine("═══════════════════════════════════════════");
    //    Console.WriteLine("  CIBC DRS Calculator");
    //    Console.WriteLine($"  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    //    Console.WriteLine("═══════════════════════════════════════════");

    //    var othersPath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others";
    //    var revReportPath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report";

    //    var csvPath = Path.Combine(othersPath, "CIBC POST.csv");
    //    // Find the calculator file (might be .xls or .xlsx)
    //    var calcPath = Directory.GetFiles(othersPath, "CIBC DRS CALCULATOR*")
    //            .FirstOrDefault()
    //            ?? throw new FileNotFoundException("CIBC DRS CALCULATOR not found.");

    //    // Find current month's rev report (e.g. "Rev Report JUNE 2026 V8.xlsx")
    //    var currentMonth = DateTime.Now.ToString("MMMM").ToUpper();
    //    var currentYear = DateTime.Now.Year;
    //    //var revReportFile = Directory.GetFiles(revReportPath, $"Rev Report {currentMonth} {currentYear}*.xlsx")
    //    var revReportFile = Directory.GetFiles(revReportPath, $"TESTRevReport.xlsx")
    //        .OrderByDescending(f => File.GetLastWriteTime(f))
    //        .FirstOrDefault()
    //        ?? throw new FileNotFoundException($"Rev Report for {currentMonth} {currentYear} not found.");

    //    Console.WriteLine($"  CSV: {Path.GetFileName(csvPath)}");
    //    Console.WriteLine($"  Calculator: {Path.GetFileName(calcPath)}");
    //    Console.WriteLine($"  Rev Report: {Path.GetFileName(revReportFile)}");
    //    Console.WriteLine();

    //    // ══════════════════════════════════════════════════════
    //    // STEP 1: Read CIBC POST.csv and find last data row
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 1: Reading CSV...");
    //    var csvLines = File.ReadAllLines(csvPath);
    //    var lastDataRow = csvLines.Length - 1; // subtract header row
    //                                           // Find actual last row with data in column D (index 3)
    //    for (int i = csvLines.Length - 1; i >= 1; i--)
    //    {
    //        var cols = csvLines[i].Split(',');
    //        if (cols.Length > 3 && !string.IsNullOrWhiteSpace(cols[3]))
    //        {
    //            lastDataRow = i + 1; // 1-based row number including header
    //            break;
    //        }
    //    }
    //    Console.WriteLine($"  CSV last data row: {lastDataRow}");

    //    if (calcPath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
    //        && !calcPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
    //    {
    //        Console.WriteLine("  Converting .xls to .xlsx...");
    //        var newPath = calcPath + "x"; // .xls → .xlsx
    //        ConvertXlsToXlsx(calcPath, newPath);
    //        calcPath = newPath;
    //    }

    //    // ══════════════════════════════════════════════════════
    //    // STEP 2: Open CIBC DRS CALCULATOR
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 2: Opening calculator...");
    //    // Pre-clean: remove the B72 formula that ClosedXML can't parse
    //    CleanProblematicFormulas(calcPath);
    //    RemoveLegacyFormControls(calcPath);

    //    using var calcWb = new XLWorkbook(calcPath);
    //    var calcWs = calcWb.Worksheet("CIBC"); // adjust sheet name if different
    //    Console.WriteLine($"  Sheet: {calcWs.Name}");
    //    Console.WriteLine($"  D9 before: '{calcWs.Cell(9, 4).Value}'");
    //    Console.WriteLine($"  B9 before: '{calcWs.Cell(9, 2).Value}'");

    //    // ══════════════════════════════════════════════════════
    //    // STEP 3: Update B72 formula with correct last row
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 3: Validating B72...");

    //    // Read B71 (TOTAL)
    //    double b71 = 0;
    //    if (calcWs.Cell(71, 2).TryGetValue(out double b71Val))
    //        b71 = b71Val;

    //    // Read D28 (last row total) from the CSV directly
    //    double csvTotal = 0;
    //    var csvCols = csvLines[lastDataRow - 1].Split(',');
    //    if (csvCols.Length > 3)
    //        double.TryParse(csvCols[3].Replace("$", "").Replace("\"", "").Trim(),
    //            NumberStyles.Any, CultureInfo.InvariantCulture, out csvTotal);

    //    var difference = b71 - csvTotal;
    //    calcWs.Cell(72, 2).Clear();
    //    calcWs.Cell(72, 2).Value = difference;

    //    Console.WriteLine($"  B71 (Total): {b71}");
    //    Console.WriteLine($"  CSV D{lastDataRow}: {csvTotal}");
    //    Console.WriteLine($"  B72 (Difference): {difference}");

    //    if (Math.Abs(difference) < 0.01)
    //        Console.WriteLine("  Match confirmed.");
    //    else
    //        Console.WriteLine($"  WARNING: Mismatch of {difference:C2}");

    //    // ══════════════════════════════════════════════════════
    //    // STEP 4: Find dynamic range B9:Bxx (until empty or TOTAL)
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 4: Reading COMM values...");
    //    var commValues = new List<double>();
    //    int startRow = 9;
    //    int endRow = startRow;

    //    for (int r = startRow; r <= 69; r++)
    //    {
    //        var nameCell = calcWs.Cell(r, 1).GetString().Trim();

    //        if (nameCell.Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
    //            break;

    //        double commVal = 0;
    //        var bCell = calcWs.Cell(r, 2);

    //        if (bCell.TryGetValue(out double d))
    //        {
    //            commVal = d;
    //        }
    //        else
    //        {
    //            var text = bCell.GetString().Replace("$", "").Replace(",", "").Trim();
    //            double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out commVal);
    //        }

    //        commValues.Add(commVal);
    //        endRow = r;
    //    }
    //    Console.WriteLine($"  Found {commValues.Count} rows ({startRow} to {endRow}), including hidden");

    //    // ══════════════════════════════════════════════════════
    //    // STEP 5: Paste COMM values to D9 onwards
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 5: Pasting COMM to column D...");
    //    for (int i = 0; i < commValues.Count; i++)
    //    {
    //        var cell = calcWs.Cell(startRow + i, 4);
    //        cell.Clear();
    //        cell.Value = commValues[i];
    //    }
    //    Console.WriteLine($"  D9 after write: '{calcWs.Cell(9, 4).Value}'");
    //    try
    //    {
    //        calcWb.Save();
    //        Console.WriteLine("  Save succeeded.");
    //    }
    //    catch (Exception ex)
    //    {
    //        Console.WriteLine($"  Save FAILED: {ex.Message}");
    //    }
    //    // ══════════════════════════════════════════════════════
    //    // STEP 6: Open Rev Report, go to Tim Katis tab
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 6: Opening Rev Report...");
    //    using var revWb = new XLWorkbook(revReportFile);
    //    var timSheet = revWb.Worksheets.FirstOrDefault(ws =>
    //        ws.Name.Contains("Tim Katis", StringComparison.OrdinalIgnoreCase))
    //        ?? throw new InvalidOperationException("'Tim Katis' tab not found in Rev Report.");

    //    Console.WriteLine($"Found tab: '{timSheet.Name}'");

    //    // ══════════════════════════════════════════════════════
    //    // STEP 7: Grab G117:G177 from Tim Katis
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 7: Reading G117:G177...");
    //    int timStartRow = 118;
    //    var gValues = new List<double>();
    //    for (int i = 0; i < commValues.Count; i++)
    //    {
    //        int r = timStartRow + i;
    //        double val = 0;
    //        var gCell = timSheet.Cell(r, 7);
    //        var bName = timSheet.Cell(r, 2).GetString().Trim(); // agent name
    //        var hidden = timSheet.Row(r).IsHidden;

    //        if (gCell.TryGetValue(out double d))
    //        {
    //            val = d;
    //        }
    //        else
    //        {
    //            var text = gCell.GetString().Replace("$", "").Replace(",", "").Trim();
    //            double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out val);
    //        }

    //        gValues.Add(val);

    //        // Log first few to verify alignment
    //        if (i < 5)
    //            Console.WriteLine($"    Row {r}: {bName} = {val} {(hidden ? "(hidden)" : "")}");
    //    }
    //    Console.WriteLine($"  Read {gValues.Count} values from G column");

    //    // ══════════════════════════════════════════════════════
    //    // STEP 8: Subtract G from D (D - G = result)
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 8: Calculating D - G...");
    //    var results = new List<double>();
    //    for (int i = 0; i < commValues.Count; i++)
    //    {
    //        var result = commValues[i] - gValues[i];
    //        results.Add(result);
    //    }
    //    Console.WriteLine($"  Calculated {results.Count} results");

    //    // ══════════════════════════════════════════════════════
    //    // STEP 9: Find today's day column in Tim Katis row 1
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("STEP 9: Finding today's column...");
    //    var today = DateTime.Now.Day;
    //    int targetCol = -1;

    //    // Search I1:AM1 (columns 9 to 39) for today's day number
    //    for (int c = 9; c <= 39; c++) // I=9, AM=39
    //    {
    //        try
    //        {
    //            var headerVal = timSheet.Cell(1, c).GetDouble();
    //            if ((int)headerVal == today)
    //            {
    //                targetCol = c;
    //                //Console.WriteLine($"  Today is day {today}, matched column {GetColumnLetter(c - 1)} (col {c})");
    //                break;
    //            }
    //        }
    //        catch { }
    //    }

    //    if (targetCol == -1)
    //    {
    //        Console.WriteLine($"  ERROR: Could not find day {today} in row 1 (I1:AM1)");
    //        return;
    //    }

    //    // ══════════════════════════════════════════════════════
    //    // STEP 10: Paste results to Tim Katis tab
    //    // ══════════════════════════════════════════════════════
    //    //Console.WriteLine($"STEP 10: Pasting results to {GetColumnLetter(targetCol - 1)}117:{GetColumnLetter(targetCol - 1)}{timStartRow + results.Count - 1}...");
    //    for (int i = 0; i < results.Count; i++)
    //    {
    //        timSheet.Cell(timStartRow + i, targetCol).Value = results[i];
    //    }

    //    // ══════════════════════════════════════════════════════
    //    // SAVE
    //    // ══════════════════════════════════════════════════════
    //    Console.WriteLine("\nSaving...");
    //    calcWb.Save();
    //    Console.WriteLine($"  Saved: {Path.GetFileName(calcPath)}");
    //    revWb.Save();
    //    Console.WriteLine($"  Saved: {Path.GetFileName(revReportFile)}");

    //    Console.WriteLine("\n═══════════════════════════════════════════");
    //    Console.WriteLine("  CIBC DRS Calculator complete.");
    //    Console.WriteLine("═══════════════════════════════════════════");
    //}

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
    static void SetCellValue(Worksheet worksheet, uint rowIndex, uint colIndex, object value)
    {
        string cellRef = $"{GetColumnLetter(colIndex)}{rowIndex}";
        var sheetData = worksheet.GetFirstChild<SheetData>();

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            var rowAfter = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex > rowIndex);
            if (rowAfter != null) sheetData.InsertBefore(row, rowAfter); else sheetData.Append(row);
        }

        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellRef);
        if (cell == null)
        {
            // NEW cell — no template formatting to inherit. Flag this if it happens a lot;
            // it means sourceRows.Count exceeds the pre-formatted template range.
            cell = new Cell { CellReference = cellRef };
            var cellAfter = row.Elements<Cell>()
                .FirstOrDefault(c => string.Compare(c.CellReference.Value, cellRef, StringComparison.OrdinalIgnoreCase) > 0);
            if (cellAfter != null) row.InsertBefore(cell, cellAfter); else row.Append(cell);
        }

        cell.CellFormula = null; // never leave a stale formula behind a literal value

        switch (value)
        {
            case null:
                cell.CellValue = null;
                cell.DataType = null;
                break;
            case string s:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(s));
                cell.CellValue = null;
                break;
            case bool b:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(b ? "1" : "0");
                cell.InlineString = null;
                break;
            case DateTime dt:
                cell.DataType = null; // numeric; relies on the cell's EXISTING number format to render as a date
                cell.CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.InlineString = null;
                break;
            case double d:
                cell.DataType = null;
                cell.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
                cell.InlineString = null;
                break;
            default:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(value.ToString()));
                cell.CellValue = null;
                break;
        }
    }
    static string GetColumnLetter(uint colIndex)
    {
        string letter = "";
        while (colIndex > 0)
        {
            uint rem = (colIndex - 1) % 26;
            letter = (char)('A' + rem) + letter;
            colIndex = (colIndex - 1) / 26;
        }
        return letter;
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

        RunOnSta(RunCibcDrsCalculator);
    }

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