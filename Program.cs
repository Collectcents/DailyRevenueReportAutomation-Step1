using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;
using Renci.SshNet;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DailyRevenueReportAutomation;

internal static class Program
{
    private static int Main(string[] args)
    {
        string? localCsv = null;

        try
        {
            // ── Load configuration ────────────────────────────────────────
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                .AddEnvironmentVariables()
                .Build();

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
}