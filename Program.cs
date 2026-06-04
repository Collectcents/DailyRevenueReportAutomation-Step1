using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;
using Microsoft.Playwright;
using Renci.SshNet;
using System.Globalization;
using System.Text.RegularExpressions;

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
    private static int Main(string[] args)
    {
        string? localCsv = null;

        try
        {
            RunPerformanceReportAutomation().GetAwaiter().GetResult();

            // -- In TESTING
            //TD Portal automation to download the report CSV directly
            RunTDPortalAutomation().GetAwaiter().GetResult();
            //CIBC Portal automation to download the report CSV directly
            RunCIBCPortalAutomation().GetAwaiter().GetResult();

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
                return 1;
            }

            // ── 7. Write to Trans tab ─────────────────────────────────────
            WriteToTransTab(filtered, headers, dateIdx, reportPath, sheetName, numCols);
            Console.WriteLine("Done.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERROR: {ex.Message}");
            return 1;
        }
        finally
        {
            if (localCsv is not null && File.Exists(localCsv))
                File.Delete(localCsv);
        }
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
        var workbookPart = doc.WorkbookPart
            ?? throw new InvalidOperationException("Workbook is empty.");

        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");

        var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;

        // Preserve row 1 (table headers), remove everything else
        var existingRows = sheetData.Elements<Row>().ToList();
        Row? headerRow = existingRows.FirstOrDefault(r => r.RowIndex?.Value == 1);

        sheetData.RemoveAllChildren<Row>();

        // Put header row back
        if (headerRow != null)
            sheetData.Append(headerRow);

        var colsToWrite = Math.Min(numCols, headers.Length);

        // Write filtered data starting at row 2
        for (int r = 0; r < rows.Count; r++)
        {
            var src = rows[r];
            var dataRow = new Row { RowIndex = (uint)(r + 2) };

            for (int c = 0; c < colsToWrite && c < src.Length; c++)
            {
                var val = src[c];
                var cell = new Cell { CellReference = GetCellRef(c, r + 2) };

                if (c == dateIdx
                    && DateTime.TryParse(val, CultureInfo.InvariantCulture,
                                         DateTimeStyles.None, out var dt))
                {
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(dt.ToString("yyyy/MM/dd"));
                }
                else if (decimal.TryParse(val, CultureInfo.InvariantCulture, out var num))
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(val);
                }
                else
                {
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(val ?? string.Empty);
                }

                dataRow.Append(cell);
            }

            sheetData.Append(dataRow);
        }

        // Update the table reference range to match new row count
        foreach (var tablePart in wsPart.TableDefinitionParts)
        {
            var table = tablePart.Table;
            if (table?.Reference != null)
            {
                // Table range: A1 to last column + last data row
                var lastCol = GetCellRef(colsToWrite - 1, 1).TrimEnd('1');
                var newRef = $"A1:{lastCol}{rows.Count + 1}";
                table.Reference = newRef;

                // Update AutoFilter range too if present
                if (table.AutoFilter != null)
                    table.AutoFilter.Reference = newRef;

                table.Save();
                Console.WriteLine($"Updated table range to {newRef}");
            }
        }

        wsPart.Worksheet.Save();
        Console.WriteLine($"Wrote {rows.Count} rows to '{sheetName}' tab in {reportPath}");
    }

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
        await page.Keyboard.TypeAsync("", new KeyboardTypeOptions { Delay = 50 });

        await page.ClickAsync("#random");
        await page.FillAsync("#random", "");
        await page.Keyboard.TypeAsync("!", new KeyboardTypeOptions { Delay = 50 });

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

    private static async Task RunTDPortalAutomation()
    {
        Console.WriteLine("Launching browser...");
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(
            new BrowserTypeLaunchOptions
            {
                Headless = false,
                Channel = "chrome",
                Args = new[] { "--disable-blink-features=AutomationControlled" }
            });

        var context = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            IgnoreHTTPSErrors = true,
            AcceptDownloads = true
        });
        var page = await context.NewPageAsync();

        await page.AddInitScriptAsync(@"
        Object.defineProperty(navigator, 'webdriver', { get: () => false });
    ");

        await page.GotoAsync(
            TDurl,
            new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });
        await page.WaitForSelectorAsync("#username");
        await page.WaitForTimeoutAsync(3000);
        Console.WriteLine("Page loaded.");

        // ── FIRST PAGE LOGIN ──────────────────
        await page.ClickAsync("#username");
        await page.FillAsync("#username", "");
        await page.Keyboard.TypeAsync(
            TDuser,
            new KeyboardTypeOptions { Delay = 50 });

        await page.ClickAsync("#random");
        await page.FillAsync("#random", "");
        await page.Keyboard.TypeAsync(
            TDpass,
            new KeyboardTypeOptions { Delay = 50 });

        Console.WriteLine($"Username: {TDuser}");
        Console.WriteLine($"Password: {TDpass}");
        Console.WriteLine($"Password length: {TDpass.Length}");
        Console.WriteLine($"Password bytes: {string.Join(" ", System.Text.Encoding.UTF8.GetBytes(TDpass).Select(b => $"{b:X2}"))}");
        // Submit and try to catch popup
        IPage popup = null;
        try
        {
            var popupTask = context.WaitForPageAsync(new BrowserContextWaitForPageOptions
            {
                Timeout = 10000
            });
            await page.ClickAsync("input[type='submit']");
            popup = await popupTask;
            Console.WriteLine("Popup opened.");
        }
        catch
        {
            // No popup — login might have happened on the same page
            Console.WriteLine("No popup. Checking current page...");
            popup = page;
        }

        // ── SECOND PAGE LOGIN (if popup opened) ──
        await popup.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await popup.WaitForTimeoutAsync(2000);

        var hasLoginForm = await popup.EvalOnSelectorAllAsync<int>("#username", "els => els.length");
        if (hasLoginForm > 0 && popup != page)
        {
            Console.WriteLine("Second login form found. Filling...");
            await popup.ClickAsync("#username");
            await popup.FillAsync("#username", "");
            await popup.Keyboard.TypeAsync(
                Environment.GetEnvironmentVariable("PORTAL2_USER") ?? "REPLACE_ME",
                new KeyboardTypeOptions { Delay = 50 });

            await popup.ClickAsync("#random");
            await popup.FillAsync("#random", "");
            await popup.Keyboard.TypeAsync(
                Environment.GetEnvironmentVariable("PORTAL2_PASS") ?? "REPLACE_ME",
                new KeyboardTypeOptions { Delay = 50 });

            await popup.PressAsync("#random", "Enter");
            await popup.WaitForTimeoutAsync(3000);
        }

        // ── CHECK LOGIN ───────────────────────
        var hasError = await popup.EvalOnSelectorAllAsync<int>("div.alert", "els => els.length");
        if (hasError > 0)
        {
            var errorText = await popup.EvalOnSelectorAsync<string>("div.alert", "el => el.innerText");
            Console.WriteLine($"Login FAILED: {errorText}");
            Console.ReadLine();
            return;
        }
        Console.WriteLine("Logged in.");

        // ── CLICK REPORTS ─────────────────────
        await popup.WaitForSelectorAsync("#REPORTS");
        await popup.ClickAsync("#REPORTS");
        await popup.WaitForTimeoutAsync(3000);

        // ── SELECT REPORT ─────────────────────
        var mainFrame = popup.Frames.FirstOrDefault(f => f.Name == "CIBCCRM_MAIN");
        await mainFrame.WaitForSelectorAsync("select[name='report_seq']");
        await mainFrame.SelectOptionAsync("select[name='report_seq']",
            new SelectOptionValue { Value = "16" });
        Console.WriteLine("Report selected.");

        // ── GENERATE REPORT (new window) ──────
        var reportPageTask = context.WaitForPageAsync();
        await mainFrame.ClickAsync("text=Generate Report");
        var reportPage = await reportPageTask;

        await reportPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        await reportPage.WaitForTimeoutAsync(3000);

        // ── SET DATE RANGE ────────────────────
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

        var today = DateTime.Today;
        var firstOfMonth = new DateTime(today.Year, today.Month, 1).ToString("yyyy-MM-dd");
        var todayStr = today.ToString("yyyy-MM-dd");

        await dateFrame.EvalOnSelectorAsync("#start_date",
            $"el => {{ el.value = '{firstOfMonth}'; el.dispatchEvent(new Event('change')); }}");
        await dateFrame.EvalOnSelectorAsync("#end_date",
            $"el => {{ el.value = '{todayStr}'; el.dispatchEvent(new Event('change')); }}");
        Console.WriteLine($"Date range set: {firstOfMonth} to {todayStr}");

        // ── RENDER REPORT ─────────────────────
        await dateFrame.ClickAsync("text=Render Report");
        Console.WriteLine("Render Report clicked.");

        await reportPage.WaitForTimeoutAsync(30000);
        Console.WriteLine("Wait complete.");

        // ── EXPORT ────────────────────────────
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

        var download = await reportPage.RunAndWaitForDownloadAsync(async () =>
        {
            await exportFrame.ClickAsync("#_export_button");
        });

        // TODO: Update save path for this portal's report
        var savePath = @"\\fro-vfs-01\Shared\Reporting\SDriveDown\Rev Report\Rev Report Others\PORTAL2 REPORT.csv";

        if (File.Exists(savePath))
            File.Delete(savePath);

        await download.SaveAsAsync(savePath);
        Console.WriteLine($"Downloaded: {savePath}");

        Console.WriteLine("Automation completed.");
        await Task.Delay(5000);
        await browser.CloseAsync();
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