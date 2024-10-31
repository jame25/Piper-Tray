using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace PiperTray
{
    static class Program
    {
        static Mutex mutex = new Mutex(true, "{8F6F0AC4-B9A1-45fd-A8CF-72F04E6BDE8F}");

        [STAThread]
        static void Main()
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                try
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    var app = PiperTrayApp.GetInstance();
                    app.Initialize();

                    // Set up logging only if it's enabled in settings
                    if (PiperTrayApp.IsLoggingEnabled)
                    {
                        string logPath = Path.Combine(Application.StartupPath, "system.log");
                        Console.SetOut(new FileLogger(logPath));
                    }

                    Application.Run(app);
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
            else
            {
                MessageBox.Show("Another instance of Piper Tray is already running.", "Piper Tray", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            LogUnhandledException(e.Exception, "Thread Exception");
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            LogUnhandledException((Exception)e.ExceptionObject, "Unhandled Exception");
        }

        static void LogUnhandledException(Exception ex, string source)
        {
            string logPath = Path.Combine(Application.StartupPath, "crash.log");
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {source}");
                writer.WriteLine(ex.ToString());
                writer.WriteLine();
            }
        }
    }
}