using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;
using TextCopy;
using System.Windows.Input;
using System.Runtime.InteropServices;

namespace ClipboardTTS
{
    public class HotkeyWindow : Form
    {
        private const int WM_HOTKEY = 0x0312;
        public const int HotKeyId = 1; // Make HotKeyId public

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HotKeyId)
            {
                // Hotkey pressed, execute the command
                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/C taskkill /F /IM sox.exe",
                    CreateNoWindow = true,
                    UseShellExecute = false
                }).WaitForExit();
            }

            base.WndProc(ref m);
        }
    }

    public class TrayApplicationContext : ApplicationContext
    {
        private const string TempFile = "clipboard_temp.txt";
        private const string SoxPath = "sox.exe";
        private const string PiperPath = "piper.exe";
        private string PiperArgs;
        private const string SoxArgs = "-t raw -b 16 -e signed-integer -r 22050 -c 1 - -t waveaudio pad 0 0.010";
        private const string SettingsFile = "settings.conf";

        private NotifyIcon trayIcon;
        private bool isRunning = true;
        private bool EnableLogging { get; set; }
        private bool isMonitoringEnabled = true;
        private static Mutex mutex = null;


        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const uint MOD_ALT = 0x0001;
        private const uint VK_Q = 0x51;

        private HotkeyWindow hotkeyWindow;

        private Icon idleIcon;
        private Icon activeIcon;

        public TrayApplicationContext()
        {
            const string appName = "PiperTray";
            bool createdNew;

            mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                // Another instance of the application is already running
                MessageBox.Show("Another instance of the application is already running.", "Piper Tray", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            // Check if model exists in the application directory
            string[] onnxFiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.onnx");
            if (onnxFiles.Length == 0)
            {
                MessageBox.Show("No speech model files found. Please make sure you place the .onnx file and .json is in the same directory as the application.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }

            LoadSettings();

            idleIcon = new Icon("icon_idle.ico");
            activeIcon = new Icon("icon_active.ico");

            trayIcon = new NotifyIcon();
            trayIcon.Icon = idleIcon;
            trayIcon.Text = "Piper Tray";
            trayIcon.Visible = true;

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem monitoringItem = new ToolStripMenuItem("Disable Monitoring", null, MonitoringItem_Click);
            contextMenu.Items.Add(monitoringItem);
            contextMenu.Items.Add("Stop Speech", null, StopItem_Click);
            contextMenu.Items.Add("Exit", null, ExitItem_Click);

            trayIcon.ContextMenuStrip = contextMenu;

            hotkeyWindow = new HotkeyWindow();
            RegisterHotKey(hotkeyWindow.Handle, HotkeyWindow.HotKeyId, MOD_ALT, VK_Q);

            // Start monitoring the clipboard automatically
            Thread monitoringThread = new Thread(StartMonitoring);
            monitoringThread.Start();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                UnregisterHotKey(hotkeyWindow.Handle, HotkeyWindow.HotKeyId);
                hotkeyWindow.Dispose();
                trayIcon.Dispose();
                idleIcon.Dispose();
                activeIcon.Dispose();
            }

            base.Dispose(disposing);
        }
        private void UpdateTrayIcon(bool isActive)
        {
            if (isActive)
            {
                trayIcon.Icon = activeIcon;
            }
            else
            {
                trayIcon.Icon = idleIcon;
            }
        }

        private void LoadSettings()
        {
            bool enableLogging = false; // Default value (logging disabled)

            if (File.Exists(SettingsFile))
            {
                string[] lines = File.ReadAllLines(SettingsFile);
                string model = "";
                string speed = "";

                foreach (string line in lines)
                {
                    if (line.StartsWith("model="))
                    {
                        model = line.Substring("model=".Length).Trim();
                    }
                    else if (line.StartsWith("speed="))
                    {
                        speed = line.Substring("speed=".Length).Trim();
                    }
                    else if (line.StartsWith("logging="))
                    {
                        string loggingValue = line.Substring("logging=".Length).Trim();

                        if (int.TryParse(loggingValue, out int value))
                        {
                            enableLogging = value != 0;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(model) && !string.IsNullOrEmpty(speed))
                {
                    PiperArgs = $"--model {model} --length_scale {speed} --output-raw";
                }
                else
                {
                    // Use default model and speed if either is missing in the settings file
                    PiperArgs = "--model en_US-libritts_r-medium.onnx --length_scale 1 --output-raw";
                }
            }
            else
            {
                // Use default model and speed if settings file doesn't exist
                PiperArgs = "--model en_US-libritts_r-medium.onnx --length_scale 1 --output-raw";
            }

            // Store the logging setting in a class-level variable
            EnableLogging = enableLogging;
        }

        private void MonitoringItem_Click(object sender, EventArgs e)
        {
            isMonitoringEnabled = !isMonitoringEnabled;
            ToolStripMenuItem monitoringItem = (ToolStripMenuItem)sender;
            monitoringItem.Text = isMonitoringEnabled ? "Disable Monitoring" : "Enable Monitoring";
        }

        private void StopItem_Click(object sender, EventArgs e)
        {
            isRunning = false;

            // Kill the sox.exe process
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/C taskkill /F /IM sox.exe",
                CreateNoWindow = true,
                UseShellExecute = false
            }).WaitForExit();
        }

        private void ExitItem_Click(object sender, EventArgs e)
        {
            isRunning = false;

            // Kill the sox.exe process if it is running
            Process[] soxProcesses = Process.GetProcessesByName("sox");
            foreach (Process process in soxProcesses)
            {
                process.Kill();
            }

            trayIcon.Visible = false;
            Application.Exit();
        }


        private void StartMonitoring()
        {
            long prevSize = 0;

            while (isRunning)
            {
                if (isMonitoringEnabled)
                {
                    // Get the current clipboard text
                    string clipboardText = ClipboardService.GetText();

                    // Write the clipboard text to a temporary file
                    File.WriteAllText(TempFile, clipboardText);

                    // Get the current size of the temporary file
                    long currentSize = new FileInfo(TempFile).Length;

                    // Check if the file size has changed
                    if (currentSize != prevSize)
                    {
                        prevSize = currentSize;

                        // Wait a short time to ensure the clipboard data is stable
                        Thread.Sleep(1000);

                        // Update the tray icon to indicate active state
                        UpdateTrayIcon(true);

                        // Use Piper TTS to convert the text from the temporary file to raw audio and pipe it to SoX
                        string piperCommand = $"{PiperPath} {PiperArgs} < \"{TempFile}\"";
                        string soxCommand = $"{SoxPath} {SoxArgs}";
                        Process piperProcess = new Process();
                        piperProcess.StartInfo.FileName = "cmd.exe";
                        piperProcess.StartInfo.Arguments = $"/C {piperCommand} | {soxCommand}";
                        piperProcess.StartInfo.UseShellExecute = false;
                        piperProcess.StartInfo.CreateNoWindow = true;
                        piperProcess.StartInfo.RedirectStandardError = true;

                        piperProcess.Start();
                        string errorOutput = piperProcess.StandardError.ReadToEnd();
                        piperProcess.WaitForExit();

                        // Update the tray icon to indicate idle state
                        UpdateTrayIcon(false);
                    }

                    // Add a small delay to reduce CPU usage
                    Thread.Sleep(100);
                }
            }
        }
    }
}
