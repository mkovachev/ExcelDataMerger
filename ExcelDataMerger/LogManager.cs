public class LogManager
{
    private string logFilePath;

    public LogManager(string destinationFolderPath)
    {
        string logFileName = "log.txt";
        logFilePath = Path.Combine(destinationFolderPath, logFileName);
    }

    public void Log(string message)
    {
        Console.WriteLine(message);
        File.AppendAllText(logFilePath, message + Environment.NewLine);
    }

    public void ClearLog()
    {
        File.WriteAllText(logFilePath, string.Empty);
    }
}
