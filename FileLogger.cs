using System.Text;

public class FileLogger : TextWriter
{
    private string logPath;
    private static object _lock = new object();

    public FileLogger(string path)
    {
        logPath = path;
    }

    public override void WriteLine(string value)
    {
        lock (_lock)
        {
            File.AppendAllText(logPath, value + Environment.NewLine);
        }
    }

    public override Encoding Encoding
    {
        get { return Encoding.UTF8; }
    }
}

