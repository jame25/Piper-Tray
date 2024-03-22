using System;
using System.Threading;
using System.Windows.Forms;
using ClipboardTTS;

namespace ClipboardTTS
{
    public static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                bool isFirstInstance;
                using (Mutex mutex = new Mutex(true, "PiperTrayMutex", out isFirstInstance))
                {
                    if (isFirstInstance)
                    {
                        Application.Run(new TrayApplicationContext());
                    }
                    else
                    {
                        MessageBox.Show("Another instance of the application is already running.", "Piper Tray", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Application.Exit();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Piper Tray", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
    }
}
