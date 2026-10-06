using System.Text;

namespace DailyRevenueReportAutomation
{
    class TeeWriter : TextWriter
    {
        private readonly TextWriter _console;
        private readonly TextWriter _file;

        public TeeWriter(TextWriter console, TextWriter file) { _console = console; _file = file; }

        public override Encoding Encoding => Encoding.UTF8;

        public override void Write(char value) { _console.Write(value); _file.Write(value); }
        public override void Write(string value) { _console.Write(value); _file.Write(value); }

        public override void WriteLine(string value)
        {
            _console.WriteLine(value);
            _file.WriteLine($"{DateTime.Now:HH:mm:ss}  {value}");   // timestamp in the file only
        }

        public override void Flush() { _console.Flush(); _file.Flush(); }
    }
}
