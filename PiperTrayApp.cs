using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using NAudio.Wave;
using System.Threading.Tasks;
using NAudio.CoreAudioApi;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.ComponentModel;

namespace PiperTray
{
    public class PiperTrayApp : Form
    {
        private static readonly Lazy<PiperTrayApp> lazy =
        new Lazy<PiperTrayApp>(() => new PiperTrayApp());

        private enum AudioPlaybackState
        {
            Idle,
            Initializing,
            Playing,
            Stopping,
            StopRequested
        }

        private AudioPlaybackState CurrentAudioState { get; set; } = AudioPlaybackState.Idle;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);


        private class VoiceModelState
        {
            public int CurrentIndex { get; set; }
            public List<string> Models { get; set; }
            public bool IsDirty { get; set; }
        }

        private VoiceModelState voiceModelState;
        private ToolStripMenuItem exportMenuItem;

        public class CustomVoiceMenuItem : ToolStripMenuItem
        {
            public bool IsSelected { get; set; }
            public bool CheckMarkEnabled { get; set; } = false;
            public Color SelectionColor { get; set; } = Color.Green;

            protected override void OnPaint(PaintEventArgs e)
            {
                // Set CheckState to None to prevent the checkmark from showing
                this.CheckState = CheckState.Unchecked;

                base.OnPaint(e);

                if (IsSelected)
                {
                    int columnWidth = 24;
                    using (SolidBrush grayBrush = new SolidBrush(Color.LightGray))
                    {
                        e.Graphics.FillRectangle(grayBrush, 0, 0, columnWidth, this.Height);
                    }

                    using (SolidBrush greenBrush = new SolidBrush(Color.Green))
                    {
                        int squareSize = columnWidth - 4;
                        int yPosition = (this.Height - squareSize) / 2;
                        e.Graphics.FillRectangle(greenBrush, 2, yPosition, squareSize, squareSize);
                    }
                }
            }
        }

        public class PresetSettings
        {
            public string Name { get; set; }
            public string VoiceModel { get; set; }
            public int Speaker { get; set; }
            public double Speed { get; set; }
            public float SentenceSilence { get; set; }
        }

        public const int HOTKEY_ID_STOP_SPEECH = 9000;
        public const int HOTKEY_ID_MONITORING = 9001;
        public const int HOTKEY_ID_CHANGE_VOICE = 9002;
        public const int HOTKEY_ID_SPEED_INCREASE = 9003;
        public const int HOTKEY_ID_SPEED_DECREASE = 9004;

        private uint monitoringModifiers;
        private uint monitoringVk;
        private uint stopSpeechModifiers;
        private uint stopSpeechVk;
        private uint changeVoiceModifiers;
        private uint changeVoiceVk;
        private uint speedIncreaseModifiers;
        private uint speedIncreaseVk;
        private uint speedDecreaseModifiers;
        private uint speedDecreaseVk;

        private Dictionary<int, Action> hotkeyActions;

        private Dictionary<string, string> customCharacterMappings = new Dictionary<string, string>
        {
            { "ć", "ch" } // Map 'ć' to a phonetic approximation
        };

        private Dictionary<string, bool> menuVisibilitySettings = new Dictionary<string, bool>();
        private Dictionary<string, ToolStripMenuItem> menuItems = new Dictionary<string, ToolStripMenuItem>();

        private HashSet<string> ignoreWords;
        private HashSet<string> bannedWords;
        private Dictionary<string, string> replaceWords;

        private NotifyIcon trayIcon;
        private SettingsForm settingsForm;
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem toggleMonitoringMenuItem;
        private ToolStripMenuItem presetsMenuItem;
        private ToolStripMenuItem speedMenuItem;
        private ToolStripMenuItem fasterMenuItem;
        private ToolStripMenuItem slowerMenuItem;
        private ToolStripMenuItem resetSpeedMenuItem;
        private ToolStripMenuItem voiceMenuItem;
        private ToolStripMenuItem settingsMenuItem;
        private ToolStripMenuItem exitMenuItem;
        private System.Windows.Forms.Timer clipboardTimer;
        private string lastClipboardContent = "";
        private bool isMonitoring = true;
        private bool ignoreCurrentClipboard = true;
        private CancellationTokenSource playbackCancellationTokenSource;
        private Thread playbackThread;
        private string piperPath;
        public static string LogFilePath { get; private set; }
        public static bool IsLoggingEnabled { get; private set; }
        private static readonly object logLock = new object();
        private bool isProcessing = false;
        private WaveOutEvent currentWaveOut;
        private bool isLoggingEnabled = false;
        private List<string> voiceModels;
        private int currentVoiceModelIndex = 0;
        private int currentSpeaker = 0;
        private readonly double[] speedOptions = { 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0 };
        private int currentSpeedIndex = 0; // Default to 1.0x speed
        private double currentSpeed = 1.0;

        private int currentPresetIndex = -1;

        private DateTime lastScanTime = DateTime.MinValue;
        private const int ScanCooldownSeconds = 5;

        private SynchronizationContext syncContext;

        public static PiperTrayApp Instance { get { return lazy.Value; } }

        private volatile bool stopRequested = false;

        private PiperTrayApp()
        {

        }

        private void InitializeMenuItems()
        {
            toggleMonitoringMenuItem = new ToolStripMenuItem("Monitoring");
            voiceMenuItem = new ToolStripMenuItem("Voice");
            settingsMenuItem = new ToolStripMenuItem("Settings");
            exitMenuItem = new ToolStripMenuItem("Exit");

            menuItems["Monitoring"] = toggleMonitoringMenuItem;
            menuItems["Speed"] = speedMenuItem;
            menuItems["Voice"] = voiceMenuItem;
            menuItems["Presets"] = presetsMenuItem;
            menuItems["Export to WAV"] = exportMenuItem;

            // Load initial visibility settings
            LoadMenuVisibilitySettings();
            ApplyMenuVisibility();
        }

        private void LoadMenuVisibilitySettings()
        {
            var settings = ReadCurrentSettings();
            foreach (var menuKey in menuItems.Keys)
            {
                string settingKey = $"MenuVisible_{menuKey.Replace(" ", "_")}";
                if (settings.TryGetValue(settingKey, out string value))
                {
                    menuVisibilitySettings[menuKey] = bool.Parse(value);
                }
                else
                {
                    menuVisibilitySettings[menuKey] = true; // Default to visible
                }
            }
        }

        private void ApplyMenuVisibility()
        {
            foreach (var kvp in menuItems)
            {
                if (menuVisibilitySettings.TryGetValue(kvp.Key, out bool isVisible))
                {
                    kvp.Value.Visible = isVisible;
                }
            }
        }

        public void UpdateMenuVisibility(string menuItem, bool isVisible)
        {
            if (menuItems.TryGetValue(menuItem, out ToolStripMenuItem item))
            {
                menuVisibilitySettings[menuItem] = isVisible;
                item.Visible = isVisible;
            }
        }

        protected override void CreateHandle()
        {
            base.CreateHandle();
        }

        private void InitializeComponent()
        {
            try
            {
                // Initialize voiceModelState first
                voiceModelState = new VoiceModelState
                {
                    CurrentIndex = 0,
                    Models = new List<string>(),
                    IsDirty = false
                };

                // Load voice models
                LoadVoiceModels();

                CreateNotifyIcon();

                // Initialize menu items
                toggleMonitoringMenuItem = new ToolStripMenuItem("Monitoring", null, SafeEventHandler(ToggleMonitoring))
                {
                    Checked = true
                };
                Log($"[InitializeComponent] Created toggleMonitoringMenuItem with ToggleMonitoring handler");
                voiceMenuItem = new ToolStripMenuItem("Voice");
                speedMenuItem = new ToolStripMenuItem("Speed");
                exportMenuItem = new ToolStripMenuItem("Export to WAV", null, SafeEventHandler(ExportWav));
                presetsMenuItem = new ToolStripMenuItem("Presets");

                // Initialize the menuItems dictionary
                menuItems = new Dictionary<string, ToolStripMenuItem>();

                InitializeMenuItems();

                // Continue with the rest of your InitializeComponent code
                fasterMenuItem = new ToolStripMenuItem("Faster", null, (sender, e) => IncreaseSpeed());
                slowerMenuItem = new ToolStripMenuItem("Slower", null, (sender, e) => DecreaseSpeed());
                resetSpeedMenuItem = new ToolStripMenuItem("Reset", null, (sender, e) => ResetSpeed(sender, e));

                speedMenuItem.DropDownItems.Add(fasterMenuItem);
                speedMenuItem.DropDownItems.Add(slowerMenuItem);
                speedMenuItem.DropDownItems.Add(resetSpeedMenuItem);

                exitMenuItem = new ToolStripMenuItem("Exit", null, SafeEventHandler(Exit));

                PopulateContextMenu();

                this.FormBorderStyle = FormBorderStyle.None;
                this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;

                UpdateSpeedDisplay();
            }
            catch (Exception ex)
            {
                Log($"Error in InitializeComponent: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"An error occurred while initializing the application: {ex.Message}", "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        public static PiperTrayApp GetInstance()
        {
            return Instance;
        }

        public void Initialize()
        {
            try
            {
                syncContext = SynchronizationContext.Current ?? new SynchronizationContext();

                // Set up logging first
                SetLogFilePath();
                InitializeLogging();
                SetPiperPath();
                Log("Starting application initialization");

                // Initialize core components
                InitializeComponent();
                LoadAndCacheDictionaries();
                InitializeClipboardMonitoring();
                InitializeHotkeyActions();

                // Load settings and models
                LoadVoiceModels();
                var settings = ReadSettings();
                ApplyMonitoringState(settings.monitoringEnabled);
                ApplySettings(
                    settings.model,
                    settings.speed,
                    settings.logging,
                    settings.monitoringHotkeyModifiers,
                    settings.monitoringHotkeyVk,
                    settings.changeVoiceHotkeyModifiers,
                    settings.changeVoiceHotkeyVk,
                    settings.monitoringEnabled,
                    settings.speedIncreaseHotkeyModifiers,
                    settings.speedIncreaseHotkeyVk,
                    settings.speedDecreaseHotkeyModifiers,
                    settings.speedDecreaseHotkeyVk,
                    settings.speaker,
                    settings.sentenceSilence
                );

                RegisterHotkey(HOTKEY_ID_STOP_SPEECH, stopSpeechModifiers, stopSpeechVk, "Stop Speech");

                // Keep the application running
                Application.Run(this);

                Log("Application initialization completed successfully");
            }
            catch (Exception ex)
            {
                Log($"Critical initialization error: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                MessageBox.Show("Application failed to initialize properly. Check the log file for details.", "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InitializeLogging()
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            string logFilePath = Path.Combine(assemblyDirectory, "system.log");

            // Read logging setting from config
            bool isLoggingEnabled = ReadLoggingSettingFromConfig();

            // Set up static logging properties
            IsLoggingEnabled = isLoggingEnabled;
            LogFilePath = logFilePath;
        }

        private bool ReadLoggingSettingFromConfig()
        {
            string configPath = GetConfigPath();
            if (File.Exists(configPath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(configPath);
                    foreach (string line in lines)
                    {
                        string[] parts = line.Split('=');
                        if (parts.Length == 2 && parts[0].Trim().Equals("Logging", StringComparison.OrdinalIgnoreCase))
                        {
                            if (bool.TryParse(parts[1].Trim(), out bool loggingEnabled))
                            {
                                return loggingEnabled;
                            }
                            else
                            {
                                Log($"[ReadLoggingSettingFromConfig] Invalid value for Logging: '{parts[1].Trim()}'");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[ReadLoggingSettingFromConfig] Exception: {ex.Message}");
                }
            }
            else
            {
                Log("[ReadLoggingSettingFromConfig] settings.conf not found. Defaulting Logging to false.");
            }
            return false; // Default to logging disabled if setting not found or invalid
        }

        private void TestWndProc()
        {
            Log($"[TestWndProc] Sending test message to WndProc");
            SendMessage(this.Handle, 0x0400, IntPtr.Zero, IntPtr.Zero);
        }

        private void CreateNotifyIcon()
        {
            if (trayIcon != null)
            {
                return;
            }

            var icon = LoadIconFromResources();
            if (icon == null)
            {
                throw new FileNotFoundException("Tray icon resource not found");
            }

            trayIcon = new NotifyIcon()
            {
                Icon = icon,
                ContextMenuStrip = new ContextMenuStrip(),
                Visible = true,
                Text = "Piper Tray"
            };

            // Add logging for menu click events
            trayIcon.MouseClick += (s, e) => Log($"[CreateNotifyIcon] Tray icon clicked: {e.Button}");
            trayIcon.ContextMenuStrip.Opening += (s, e) => Log($"[CreateNotifyIcon] Context menu opening");
            trayIcon.ContextMenuStrip.ItemClicked += (s, e) => Log($"[CreateNotifyIcon] Menu item clicked: {e.ClickedItem.Text}");

            trayIcon.ContextMenuStrip.Opening += ContextMenuStrip_Opening;
        }

        private void LogEmbeddedResources()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            Log("Embedded resources:");
            foreach (var name in resourceNames)
            {
                Log($"- {name}");
            }
        }

        private Icon LoadIconFromResources()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            var iconResourceName = resourceNames.FirstOrDefault(name => name.EndsWith("icon.ico"));

            if (iconResourceName == null)
            {
                Log("Icon resource not found in embedded resources");
                return null;
            }

            using (var stream = assembly.GetManifestResourceStream(iconResourceName))
            {
                if (stream == null)
                {
                    Log($"Failed to load icon stream for resource: {iconResourceName}");
                    return null;
                }
                return new Icon(stream);
            }
        }

        public static Icon GetApplicationIcon()
        {
            var icon = new PiperTrayApp().LoadIconFromResources();
            if (icon == null)
            {
                throw new FileNotFoundException("Application icon resource not found");
            }
            return icon;
        }

        private void ContextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            RebuildContextMenu();
            UpdateSpeedDisplay();
            UpdateVoiceMenuCheckedState();
        }

        private void ApplyFinalMenuState()
        {
            foreach (ToolStripItem item in voiceMenuItem.DropDownItems)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    int index = voiceMenuItem.DropDownItems.IndexOf(item);
                    menuItem.Checked = (index == voiceModelState.CurrentIndex);
                }
            }
            trayIcon.ContextMenuStrip.Refresh();
        }

        private void LoadAndCacheDictionaries()
        {
            (ignoreWords, bannedWords, replaceWords) = LoadDictionaries();
            Log($"Dictionaries loaded and cached. Ignore words: {ignoreWords.Count}, Banned words: {bannedWords.Count}, Replace words: {replaceWords.Count}");
        }

        private void ApplySettings(string model, float speed, bool logging,
            uint monitoringHotkeyModifiers, uint monitoringHotkeyVk,
            uint changeVoiceHotkeyModifiers, uint changeVoiceHotkeyVk,
            bool monitoringEnabled,
            uint speedIncreaseHotkeyModifiers, uint speedIncreaseHotkeyVk,
            uint speedDecreaseHotkeyModifiers, uint speedDecreaseHotkeyVk,
            int speaker,
            float sentenceSilence)
        {
            UpdateVoiceModel(model);
            UpdateSpeedFromSettings(speed);
            isLoggingEnabled = logging;
            currentSpeaker = speaker;
            ApplyHotkeySettings(
                monitoringHotkeyModifiers,
                monitoringHotkeyVk,
                changeVoiceHotkeyModifiers,
                changeVoiceHotkeyVk,
                speedIncreaseHotkeyModifiers,
                speedIncreaseHotkeyVk,
                speedDecreaseHotkeyModifiers,
                speedDecreaseHotkeyVk
            );
            ApplyMonitoringState(monitoringEnabled);

            if (settingsForm != null)
            {
                settingsForm.UpdateSentenceSilence(sentenceSilence);
            }

            Log($"[ApplySettings] Applied settings - Model: {model}, Speed: {speed}, Logging: {logging}, Speaker: {speaker}, Sentence Silence: {sentenceSilence}");
        }


        public (bool success, uint errorCode) RegisterHotkey(int hotkeyId, uint modifiers, uint vk, string hotkeyName)
        {
            string hotkeyDescription = GetHotkeyDescription(hotkeyId, hotkeyName);
            Log($"[RegisterHotkey] Attempting to register {hotkeyDescription}. ID: {hotkeyId}, Modifiers: 0x{modifiers:X2}, VK: 0x{vk:X2}");
            bool result = RegisterHotKey(this.Handle, hotkeyId, modifiers, vk);
            uint errorCode = result ? 0 : GetLastError();
            Log($"[RegisterHotkey] {hotkeyDescription} registration result: {(result ? "Success" : $"Failed (Error: {errorCode})")}");
            return (result, errorCode);
        }

        public string GetConfigPath()
        {
            return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "settings.conf");
        }

        private string GetHotkeyDescription(int hotkeyId, string hotkeyName)
        {
            switch (hotkeyId)
            {
                case HOTKEY_ID_MONITORING: return "Monitoring hotkey";
                case HOTKEY_ID_CHANGE_VOICE: return "Change Voice hotkey";
                case HOTKEY_ID_SPEED_INCREASE: return "Speed Increase hotkey";
                case HOTKEY_ID_SPEED_DECREASE: return "Speed Decrease hotkey";
                default: return $"{hotkeyName} hotkey";
            }
        }

        private void ApplyHotkeySettings(uint monitoringHotkeyModifiers, uint monitoringHotkeyVk,
            uint changeVoiceHotkeyModifiers, uint changeVoiceHotkeyVk,
            uint speedIncreaseHotkeyModifiers, uint speedIncreaseHotkeyVk,
            uint speedDecreaseHotkeyModifiers, uint speedDecreaseHotkeyVk)
        {
            Log($"[ApplyHotkeySettings] Entering method");

            Log($"[ApplyHotkeySettings] Monitoring hotkey - Modifiers: 0x{monitoringHotkeyModifiers:X}, VK: 0x{monitoringHotkeyVk:X}");
            var result = UpdateHotkey(HOTKEY_ID_MONITORING, monitoringHotkeyModifiers, monitoringHotkeyVk);
            Log($"[ApplyHotkeySettings] Monitoring hotkey registration result - Modifiers: 0x{result.modifiers:X}, VK: 0x{result.vk:X}");

            Log($"[ApplyHotkeySettings] Change Voice hotkey - Modifiers: 0x{changeVoiceHotkeyModifiers:X}, VK: 0x{changeVoiceHotkeyVk:X}");
            result = UpdateHotkey(HOTKEY_ID_CHANGE_VOICE, changeVoiceHotkeyModifiers, changeVoiceHotkeyVk);
            Log($"[ApplyHotkeySettings] Change Voice hotkey registration result - Modifiers: 0x{result.modifiers:X}, VK: 0x{result.vk:X}");

            Log($"[ApplyHotkeySettings] Speed Increase hotkey - Modifiers: 0x{speedIncreaseHotkeyModifiers:X}, VK: 0x{speedIncreaseHotkeyVk:X}");
            result = UpdateHotkey(HOTKEY_ID_SPEED_INCREASE, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk);
            Log($"[ApplyHotkeySettings] Speed Increase hotkey registration result - Modifiers: 0x{result.modifiers:X}, VK: 0x{result.vk:X}");

            Log($"[ApplyHotkeySettings] Speed Decrease hotkey - Modifiers: 0x{speedDecreaseHotkeyModifiers:X}, VK: 0x{speedDecreaseHotkeyVk:X}");
            result = UpdateHotkey(HOTKEY_ID_SPEED_DECREASE, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk);
            Log($"[ApplyHotkeySettings] Speed Decrease hotkey registration result - Modifiers: 0x{result.modifiers:X}, VK: 0x{result.vk:X}");

            Log($"[ApplyHotkeySettings] Hotkey settings applied");
            Log($"[ApplyHotkeySettings] Exiting method");
        }



        private EventHandler SafeEventHandler(EventHandler handler)
        {
            return (sender, e) =>
            {
                try
                {
                    handler(sender, e);
                }
                catch (Exception ex)
                {
                    Log($"Error in event handler: {ex.Message}");
                    Log($"Stack trace: {ex.StackTrace}");
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }

        private void InitializeHotkeyActions()
        {
            hotkeyActions = new Dictionary<int, Action>
            {
                { HOTKEY_ID_STOP_SPEECH, StopCurrentSpeech },
                { HOTKEY_ID_MONITORING, () => ToggleMonitoring(this, EventArgs.Empty) },
                { HOTKEY_ID_CHANGE_VOICE, ChangeVoice },
                { HOTKEY_ID_SPEED_INCREASE, IncreaseSpeed },
                { HOTKEY_ID_SPEED_DECREASE, DecreaseSpeed }
            };

            Log($"[InitializeHotkeyActions] Hotkey actions initialized. Count: {hotkeyActions.Count}");
        }

        private void StopCurrentSpeech()
        {
            Log($"[StopCurrentSpeech] Stopping current speech playback");
            if (CurrentAudioState == AudioPlaybackState.Playing)
            {
                CurrentAudioState = AudioPlaybackState.StopRequested;

                if (currentWaveOut != null)
                {
                    try
                    {
                        currentWaveOut.Stop();
                        currentWaveOut.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Log($"[StopCurrentSpeech] Error stopping playback: {ex.Message}");
                    }
                    finally
                    {
                        currentWaveOut = null;
                    }
                }

                if (playbackCancellationTokenSource != null)
                {
                    playbackCancellationTokenSource.Cancel();
                    playbackCancellationTokenSource = null;
                }

                CurrentAudioState = AudioPlaybackState.Idle;
            }
        }

        private void CheckVoiceFiles()
        {
            string appDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Log($"[CheckVoiceFiles] Checking for .onnx files in: {appDirectory}");
            string[] onnxFiles = Directory.GetFiles(appDirectory, "*.onnx");
            Log($"[CheckVoiceFiles] Found {onnxFiles.Length} .onnx files");
            foreach (var file in onnxFiles)
            {
                Log($"[CheckVoiceFiles] Found file: {Path.GetFileNameWithoutExtension(file)}");
            }
            if (onnxFiles.Length == 0)
            {
                Log($"[CheckVoiceFiles] No .onnx files detected");
                MessageBox.Show("No voice files (.onnx) detected in the application directory. The application cannot function correctly.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        [DllImport("kernel32.dll")]
        static extern uint GetLastError();

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80; // WS_EX_TOOLWINDOW
                return cp;
            }
        }

        protected override void SetVisibleCore(bool value)
        {
            if (!this.IsHandleCreated)
            {
                CreateHandle();
                Log($"[SetVisibleCore] Handle created: {this.Handle}");
            }
            base.SetVisibleCore(false);
        }

        private void RegisterHotkeys()
        {
            Log($"[RegisterHotkeys] Starting hotkey registration process");
            UnregisterAllHotkeys();

            // Register all hotkeys in a single pass
            RegisterHotkey(HOTKEY_ID_MONITORING, monitoringModifiers, monitoringVk, "Monitoring");
            RegisterHotkey(HOTKEY_ID_STOP_SPEECH, stopSpeechModifiers, stopSpeechVk, "Stop Speech");
            RegisterHotkey(HOTKEY_ID_CHANGE_VOICE, changeVoiceModifiers, changeVoiceVk, "Change Voice");
            RegisterHotkey(HOTKEY_ID_SPEED_INCREASE, speedIncreaseModifiers, speedIncreaseVk, "Speed Increase");
            RegisterHotkey(HOTKEY_ID_SPEED_DECREASE, speedDecreaseModifiers, speedDecreaseVk, "Speed Decrease");

            Log($"[RegisterHotkeys] Hotkey registration completed");
        }

        public void UnregisterAllHotkeys()
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID_MONITORING);
            UnregisterHotKey(this.Handle, HOTKEY_ID_CHANGE_VOICE);
            UnregisterHotKey(this.Handle, HOTKEY_ID_SPEED_INCREASE);
            UnregisterHotKey(this.Handle, HOTKEY_ID_SPEED_DECREASE);
            Log("All hotkeys unregistered");
        }

        private int GetHotkeyId(string hotkeyName)
        {
            switch (hotkeyName)
            {
                case "Monitoring": return HOTKEY_ID_MONITORING;
                case "ChangeVoice": return HOTKEY_ID_CHANGE_VOICE;
                case "SpeedIncrease": return HOTKEY_ID_SPEED_INCREASE;
                case "SpeedDecrease": return HOTKEY_ID_SPEED_DECREASE;
                default: throw new ArgumentException($"Unknown hotkey name: {hotkeyName}");
            }
        }

        public bool TryRegisterHotkey(uint modifiers, uint vk)
        {
            if (RegisterHotKey(this.Handle, HOTKEY_ID_MONITORING, modifiers, vk))
            {
                Log($"Hotkey registered successfully. Modifiers: 0x{modifiers:X}, VK: 0x{vk:X}");
                return true;
            }
            else
            {
                uint errorCode = GetLastError();
                Log($"Failed to register hotkey. Modifiers: 0x{modifiers:X}, VK: 0x{vk:X}, Error code: {errorCode}");
                return false;
            }
        }

        public (uint modifiers, uint vk) UpdateHotkey(int hotkeyId, uint modifiers, uint vk)
        {
            Log($"[UpdateHotkey] Entering method. HotkeyId: {hotkeyId}, Modifiers: 0x{modifiers:X}, VK: 0x{vk:X}");
            Log($"[UpdateHotkey] Received values - Modifier: 0x{modifiers:X2}, Key: 0x{vk:X2}");

            // Unregister existing hotkey
            bool unregisterResult = UnregisterHotKey(this.Handle, hotkeyId);
            Log($"[UpdateHotkey] Unregister result: {unregisterResult}. Last error: {GetLastError()}");

            if (modifiers == 0 && vk == 0)
            {
                Log($"[UpdateHotkey] No hotkey combination set for ID: {hotkeyId}");
                return (0, 0);
            }

            // Attempt to register new hotkey
            bool registerResult = RegisterHotKey(this.Handle, hotkeyId, modifiers, vk);
            uint errorCode = GetLastError();

            Log($"[UpdateHotkey] Register result: {registerResult}. Error code: {errorCode}");

            if (registerResult)
            {
                Log($"[UpdateHotkey] Hotkey registered successfully. ID: {hotkeyId}, Modifiers: 0x{modifiers:X}, VK: 0x{vk:X}");

                // Update the appropriate variables based on the hotkey ID
                switch (hotkeyId)
                {
                    case HOTKEY_ID_MONITORING:
                        monitoringModifiers = modifiers;
                        monitoringVk = vk;
                        break;
                    case HOTKEY_ID_CHANGE_VOICE:
                        changeVoiceModifiers = modifiers;
                        changeVoiceVk = vk;
                        break;
                    case HOTKEY_ID_SPEED_INCREASE:
                        speedIncreaseModifiers = modifiers;
                        speedIncreaseVk = vk;
                        break;
                    case HOTKEY_ID_SPEED_DECREASE:
                        speedDecreaseModifiers = modifiers;
                        speedDecreaseVk = vk;
                        break;
                }

                SaveSettings();
                return (modifiers, vk);
            }
            else
            {
                Log($"[UpdateHotkey] Failed to register hotkey. ID: {hotkeyId}, Modifiers: 0x{modifiers:X}, VK: 0x{vk:X}, Error code: {errorCode}");
                Log($"[UpdateHotkey] Error description: {new Win32Exception((int)errorCode).Message}");
                return (0, 0);
            }
        }

        private bool IsValidHotkeyCombination(uint modifiers, uint vk)
        {
            if (modifiers == 0 || vk == 0)
            {
                Log($"Invalid hotkey combination. Modifiers: 0x{modifiers:X}, VK: 0x{vk:X}");
                return false;
            }
            return true;
        }

        private uint GetHotkeyModifiers(int hotkeyId)
        {
            switch (hotkeyId)
            {
                case HOTKEY_ID_MONITORING:
                    return monitoringModifiers;
                case HOTKEY_ID_CHANGE_VOICE:
                    return changeVoiceModifiers;
                case HOTKEY_ID_SPEED_INCREASE:
                    return speedIncreaseModifiers;
                case HOTKEY_ID_SPEED_DECREASE:
                    return speedDecreaseModifiers;
                default:
                    return 0;
            }
        }

        private uint GetHotkeyKey(int hotkeyId)
        {
            switch (hotkeyId)
            {
                case HOTKEY_ID_MONITORING:
                    return monitoringVk;
                case HOTKEY_ID_CHANGE_VOICE:
                    return changeVoiceVk;
                case HOTKEY_ID_SPEED_INCREASE:
                    return speedIncreaseVk;
                case HOTKEY_ID_SPEED_DECREASE:
                    return speedDecreaseVk;
                default:
                    return 0;
            }
        }

        public bool IsHotkeyRegistered(uint modifiers, uint vk)
        {
            // Implementation goes here
            // For example:
            return modifiers != 0 && vk != 0;
        }

        public bool IsMonitoringHotkeyRegistered()
        {
            var (model, speed, logging, monitoringHotkeyModifiers, monitoringHotkeyVk, changeVoiceHotkeyModifiers, changeVoiceHotkeyVk, monitoringEnabled, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence) = ReadSettings();
            return IsHotkeyRegistered(monitoringHotkeyModifiers, monitoringHotkeyVk);
        }

        public void LoadAndApplyHotkeySettings()
        {
            var (_, _, _, monitoringHotkeyModifiers, monitoringHotkeyVk, changeVoiceHotkeyModifiers, changeVoiceHotkeyVk, _, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence) = ReadSettings();

            UnregisterAllHotkeys();

            RegisterHotkey(HOTKEY_ID_MONITORING, monitoringHotkeyModifiers, monitoringHotkeyVk, "Monitoring");
            RegisterHotkey(HOTKEY_ID_CHANGE_VOICE, changeVoiceHotkeyModifiers, changeVoiceHotkeyVk, "Change Voice");
            RegisterHotkey(HOTKEY_ID_SPEED_INCREASE, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, "Speed Increase");
            RegisterHotkey(HOTKEY_ID_SPEED_DECREASE, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, "Speed Decrease");

            Log("Hotkey settings loaded and applied");
        }

        protected override void WndProc(ref Message m)
        {
            Log($"[WndProc] Message received: 0x{m.Msg:X4}, WParam: 0x{m.WParam:X8}, LParam: 0x{m.LParam:X8}");

            switch (m.Msg)
            {
                case 0x0312: // WM_HOTKEY
                    int id = m.WParam.ToInt32();
                    uint modifiers = (uint)((int)m.LParam & 0xFFFF);
                    uint vk = (uint)((int)m.LParam >> 16);
                    Log($"[WndProc] WM_HOTKEY received. ID: {id}, Modifiers: 0x{modifiers:X2}, VK: 0x{vk:X2}");

                    if (hotkeyActions.TryGetValue(id, out Action action))
                    {
                        Log($"[WndProc] Action found for hotkey ID: {id}. Invoking action.");
                        action.Invoke();
                    }
                    else
                    {
                        Log($"[WndProc] No action found for hotkey ID: {id}");
                    }
                    break;

                case 0x0001: // WM_CREATE
                    Log($"[WndProc] WM_CREATE message received");
                    break;

                case 0x0002: // WM_DESTROY
                    Log($"[WndProc] WM_DESTROY message received");
                    UnregisterAllHotkeys();
                    break;

                case 0x0010: // WM_CLOSE
                    Log($"[WndProc] WM_CLOSE message received");
                    break;

                default:
                    Log($"[WndProc] Unhandled message: 0x{m.Msg:X4}");
                    break;
            }

            base.WndProc(ref m);
            Log($"[WndProc] Message 0x{m.Msg:X4} processed");
        }


        private void AddPresetsToContextMenu()
        {
            presetsMenuItem.DropDownItems.Clear();
            for (int i = 0; i < 4; i++)
            {
                var settings = LoadPreset(i);
                string presetName = settings?.Name ?? $"Preset {i + 1}";
                var presetItem = new CustomVoiceMenuItem
                {
                    Text = presetName,
                    IsSelected = (i == currentPresetIndex),
                    SelectionColor = Color.LightBlue
                };
                int presetIndex = i;
                presetItem.Click += (s, e) =>
                {
                    ApplyPreset(presetIndex);
                    currentPresetIndex = presetIndex;
                    UpdatePresetMenuSelection(presetIndex);
                };
                presetsMenuItem.DropDownItems.Add(presetItem);
            }

            // Load and apply the last used preset index
            var lastUsedPresetIndex = ReadSettingValue("LastUsedPreset");
            if (!string.IsNullOrEmpty(lastUsedPresetIndex) && int.TryParse(lastUsedPresetIndex, out int savedIndex))
            {
                currentPresetIndex = savedIndex;
                UpdatePresetMenuSelection(savedIndex);
            }
        }

        private void SaveLastUsedPreset(int presetIndex)
        {
            string configPath = GetConfigPath();
            var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
            UpdateOrAddSetting(lines, "LastUsedPreset", presetIndex.ToString());
            File.WriteAllLines(configPath, lines);
        }

        private void UpdatePresetMenuSelection(int selectedIndex)
        {
            foreach (ToolStripItem item in presetsMenuItem.DropDownItems)
            {
                if (item is CustomVoiceMenuItem customItem)
                {
                    customItem.IsSelected = (presetsMenuItem.DropDownItems.IndexOf(item) == selectedIndex);
                }
            }
            trayIcon.ContextMenuStrip.Refresh();
        }

        private void UpdatePresetMenuChecks(int selectedIndex)
        {
            foreach (ToolStripItem item in presetsMenuItem.DropDownItems)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    menuItem.Checked = (presetsMenuItem.DropDownItems.IndexOf(item) == selectedIndex);
                }
            }
        }

        public void ApplyPreset(int presetIndex)
        {
            var settings = LoadPreset(presetIndex);
            if (settings != null)
            {
                UpdateVoiceModel(settings.VoiceModel);
                UpdateCurrentSpeaker(settings.Speaker);
                UpdateSpeedFromSettings(settings.Speed);
                SaveSettings(
                    speed: settings.Speed,
                    voiceModel: settings.VoiceModel,
                    speaker: settings.Speaker,
                    sentenceSilence: settings.SentenceSilence
                );
                currentPresetIndex = presetIndex;
                SaveLastUsedPreset(presetIndex);
                UpdatePresetMenuSelection(presetIndex);
            }
        }

        public PresetSettings LoadPreset(int index)
        {
            string configPath = GetConfigPath();
            var settings = ReadCurrentSettings();
            string presetKey = $"Preset{index + 1}";

            if (settings.TryGetValue(presetKey, out string presetJson))
            {
                return JsonSerializer.Deserialize<PresetSettings>(presetJson);
            }
            return null;
        }

        public void SavePreset(int index, PresetSettings settings)
        {
            string configPath = GetConfigPath();
            var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
            string presetKey = $"Preset{index + 1}";
            string presetJson = JsonSerializer.Serialize(settings);
            UpdateOrAddSetting(lines, presetKey, presetJson);
            File.WriteAllLines(configPath, lines);
        }

        private void InitializeClipboardMonitoring()
        {
            clipboardTimer = new System.Windows.Forms.Timer();
            clipboardTimer.Interval = 1000;
            clipboardTimer.Tick += ClipboardTimer_Tick;
            ignoreCurrentClipboard = true;
            Log($"Clipboard monitoring initialized with {clipboardTimer.Interval}ms interval");
        }

        private void SetPiperPath()
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            piperPath = Path.Combine(assemblyDirectory, "piper.exe");

            if (!File.Exists(piperPath))
            {
                Log($"Error: piper.exe not found at {piperPath}");
                MessageBox.Show("piper.exe not found in the application directory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            else
            {
                Log($"Piper executable found at {piperPath}");
            }
        }

        private void SetLogFilePath()
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            LogFilePath = Path.Combine(assemblyDirectory, "system.log");
            Log($"Log file initialized at {LogFilePath}");
        }

        private void Log(string message)
        {
            if (!PiperTrayApp.IsLoggingEnabled)
            {
                return;
            }

            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(PiperTrayApp.LogFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }


        private void LogAudioDevices()
        {
            Log($"Available audio devices:");
            for (int i = 0; i < WaveOut.DeviceCount; i++)
            {
                var capabilities = WaveOut.GetCapabilities(i);
                Log($"Device {i}: {capabilities.ProductName}");
            }
        }

        private void LogDefaultAudioDevice()
        {
            try
            {
                var defaultDevice = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                Log($"Default audio device: {defaultDevice.FriendlyName}");
            }
            catch (Exception ex)
            {
                Log($"Error getting default audio device: {ex.Message}");
            }
        }

        private void StartMonitoring()
        {
            isMonitoring = true;
            ignoreCurrentClipboard = true;
            toggleMonitoringMenuItem.Checked = true;
            toggleMonitoringMenuItem.Text = "Monitoring";
            clipboardTimer.Start();
            Log($"Clipboard monitoring started");
        }

        private void StopMonitoring()
        {
            isMonitoring = false;
            toggleMonitoringMenuItem.Checked = false;
            toggleMonitoringMenuItem.Text = "Monitoring";
            clipboardTimer.Stop();
            Log($"Clipboard monitoring stopped");
        }

        public void ToggleMonitoring(object sender, EventArgs e)
        {
            Log($"[ToggleMonitoring] Starting toggle operation. Current state: {isMonitoring}");
            try
            {
                isMonitoring = !isMonitoring;
                Log($"[ToggleMonitoring] State toggled to: {isMonitoring}");

                toggleMonitoringMenuItem.Checked = isMonitoring;
                Log($"[ToggleMonitoring] Menu item checked state updated: {toggleMonitoringMenuItem.Checked}");

                if (isMonitoring)
                {
                    Log($"[ToggleMonitoring] Calling StartMonitoring()");
                    StartMonitoring();
                }
                else
                {
                    Log($"[ToggleMonitoring] Calling StopMonitoring()");
                    StopMonitoring();
                }

                string configPath = GetConfigPath();
                Log($"[ToggleMonitoring] Config path: {configPath}");
                Log($"[ToggleMonitoring] Config file exists: {File.Exists(configPath)}");

                SaveSettings();
                Log($"[ToggleMonitoring] SaveSettings() called");

                // Verify the save
                var currentSettings = ReadCurrentSettings();
                Log($"[ToggleMonitoring] Verification - MonitoringEnabled in settings: {currentSettings.GetValueOrDefault("MonitoringEnabled")}");

                toggleMonitoringMenuItem.Text = isMonitoring ? "Monitoring (On)" : "Monitoring (Off)";
                Log($"[ToggleMonitoring] Menu text updated to: {toggleMonitoringMenuItem.Text}");
            }
            catch (Exception ex)
            {
                Log($"[ToggleMonitoring] Error occurred: {ex.Message}");
                Log($"[ToggleMonitoring] Stack trace: {ex.StackTrace}");
            }
            Log($"[ToggleMonitoring] Toggle operation completed");
        }

        private void SaveMonitoringState(bool enabled)
        {
            string configPath = GetConfigPath();
            try
            {
                var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
                UpdateOrAddSetting(lines, "MonitoringEnabled", enabled.ToString());
                File.WriteAllLines(configPath, lines);
                Log($"[SaveMonitoringState] MonitoringEnabled set to {enabled}");
            }
            catch (Exception ex)
            {
                Log($"[SaveMonitoringState] Error updating MonitoringEnabled: {ex.Message}");
                MessageBox.Show($"Failed to save Monitoring state: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChangeVoice()
        {
            Log($"[ChangeVoice] Method called. Current voice model index: {currentVoiceModelIndex}");

            if (voiceModels != null && voiceModels.Count > 0 && currentVoiceModelIndex < voiceModels.Count)
            {
                Log($"[ChangeVoice] Current voice model: {voiceModels[currentVoiceModelIndex]}");

                CycleVoiceModel();

                Log($"[ChangeVoice] Voice model cycled");
                Log($"[ChangeVoice] New voice model index: {currentVoiceModelIndex}");
                Log($"[ChangeVoice] New voice model: {voiceModels[currentVoiceModelIndex]}");

                UpdateVoiceModelUI();
                Log($"[ChangeVoice] UI updated with new voice model");
            }
            else
            {
                Log($"[ChangeVoice] Error: Voice models not properly initialized or current index out of range");
            }
        }

        private void CycleVoiceModel()
        {
            Log($"[CycleVoiceModel] Entering method. Current voice model index: {currentVoiceModelIndex}");

            if (voiceModels == null || voiceModels.Count == 0)
            {
                Log("[CycleVoiceModel] Voice models not loaded. Calling LoadVoiceModels()");
                LoadVoiceModels();
            }

            Log($"[CycleVoiceModel] Number of voice models: {voiceModels?.Count ?? 0}");

            if (voiceModels != null && voiceModels.Count > 0)
            {
                int oldIndex = currentVoiceModelIndex;
                currentVoiceModelIndex = (currentVoiceModelIndex + 1) % voiceModels.Count;
                Log($"[CycleVoiceModel] Voice model index updated. Old: {oldIndex}, New: {currentVoiceModelIndex}");

                string newModel = Path.GetFileName(voiceModels[currentVoiceModelIndex]);
                Log($"[CycleVoiceModel] New voice model selected: {newModel}");

                UpdateVoiceModelSetting(newModel);
                Log("[CycleVoiceModel] UpdateVoiceModelSetting called with new model");
            }
            else
            {
                Log("[CycleVoiceModel] No voice models available to cycle through.");
            }

            Log("[CycleVoiceModel] Exiting method");
        }

        public void UpdateCurrentSpeaker(int speakerId)
        {
            currentSpeaker = speakerId;
            Log($"[UpdateCurrentSpeaker] Updated current speaker to: {speakerId}");
        }

        private void IncreaseSpeed()
        {
            Log($"[IncreaseSpeed] Entering method. Current speed: {currentSpeed}");
            if (currentSpeed > 0.1)
            {
                currentSpeed = Math.Max(0.1, currentSpeed - 0.1);
                Log($"[IncreaseSpeed] New speed after increase: {currentSpeed}");
                UpdateSpeedDisplay();
                SaveSettings(currentSpeed);
                Log($"[IncreaseSpeed] Settings saved with new speed: {currentSpeed}");
            }
            else
            {
                Log($"[IncreaseSpeed] Speed not increased, already at maximum");
            }
            Log($"[IncreaseSpeed] Exiting method");
        }

        private void DecreaseSpeed()
        {
            Log($"[DecreaseSpeed] Entering method. Current speed: {currentSpeed}");
            if (currentSpeed < 1.0)
            {
                currentSpeed = Math.Min(1.0, currentSpeed + 0.1);
                Log($"[DecreaseSpeed] New speed after decrease: {currentSpeed}");
                UpdateSpeedDisplay();
                SaveSettings(currentSpeed);
                Log($"[DecreaseSpeed] Settings saved with new speed: {currentSpeed}");
            }
            else
            {
                Log($"[DecreaseSpeed] Speed not decreased, already at minimum");
            }
            Log($"[DecreaseSpeed] Exiting method");
        }

        private void ResetSpeed(object sender, EventArgs e)
        {
            currentSpeedIndex = speedOptions.Length - 1; // Reset to default speed (1.0)
            currentSpeed = speedOptions[currentSpeedIndex];
            UpdateSpeedDisplay();
            SaveSettings(currentSpeed);
        }

        public void UpdateSpeedFromSettings(double newSpeed)
        {
            syncContext.Post(_ =>
            {
                Log($"[UpdateSpeedFromSettings] Entering method. New speed: {newSpeed}");
                Log($"[UpdateSpeedFromSettings] Current Instance Hash: {Instance?.GetHashCode() ?? 0}");

                currentSpeed = newSpeed;
                currentSpeedIndex = Array.IndexOf(speedOptions, newSpeed);
                Log($"[UpdateSpeedFromSettings] Updated currentSpeed: {currentSpeed}, currentSpeedIndex: {currentSpeedIndex}");

                if (currentSpeedIndex == -1)
                {
                    currentSpeedIndex = speedOptions.Length - 1;
                    for (int i = 0; i < speedOptions.Length; i++)
                    {
                        if (newSpeed >= speedOptions[i])
                        {
                            currentSpeedIndex = i;
                            break;
                        }
                    }
                    Log($"[UpdateSpeedFromSettings] Adjusted currentSpeedIndex: {currentSpeedIndex}");
                }

                UpdateSpeedDisplay();
                Log($"[UpdateSpeedFromSettings] Speed display updated");

                Log($"[UpdateSpeedFromSettings] Exiting method");
            }, null);
        }


        private void UpdateSpeedDisplay()
        {
            Log($"[UpdateSpeedDisplay] Entering method. Current speed: {currentSpeed}");
            int displaySpeed = 11 - (int)Math.Round(currentSpeed * 10);
            Log($"[UpdateSpeedDisplay] Calculated display speed: {displaySpeed}");

            speedMenuItem.Text = $"Speed: {displaySpeed}";
            Log($"[UpdateSpeedDisplay] Updated speedMenuItem text: {speedMenuItem.Text}");

            fasterMenuItem.Enabled = !IsApproximatelyEqual(currentSpeed, 0.1);
            slowerMenuItem.Enabled = !IsApproximatelyEqual(currentSpeed, 1.0);
            resetSpeedMenuItem.Enabled = !IsApproximatelyEqual(currentSpeed, 1.0);
            Log($"[UpdateSpeedDisplay] Menu item states - Faster: {fasterMenuItem.Enabled}, Slower: {slowerMenuItem.Enabled}, Reset: {resetSpeedMenuItem.Enabled}");
            Log($"[UpdateSpeedDisplay] Exiting method");
        }

        private bool IsApproximatelyEqual(double a, double b, double epsilon = 1e-6)
        {
            return Math.Abs(a - b) < epsilon;
        }

        private void LoadVoiceModels()
        {
            ScanForVoiceModels();
            Log($"[LoadVoiceModels] Loaded {voiceModelState.Models.Count} voice models");
            foreach (var model in voiceModelState.Models)
            {
                Log($"[LoadVoiceModels] Loaded model: {model}");
            }
        }

        private void UpdateVoiceModelAndRefresh(string newModel)
        {

            LoadVoiceModels();
            int newIndex = voiceModelState.Models.FindIndex(m => Path.GetFileNameWithoutExtension(m).Equals(newModel, StringComparison.OrdinalIgnoreCase));

            if (newIndex != -1)
            {
                voiceModelState.CurrentIndex = newIndex;
                UpdateVoiceModelSetting(newModel);
            }
            else
            {
            }

            RebuildContextMenu();
            ForceMenuRefresh();
            UpdateVoiceMenuCheckedState();
        }

        public void ScanForVoiceModels()
        {
            if (voiceModelState == null)
            {
                voiceModelState = new VoiceModelState
                {
                    CurrentIndex = 0,
                    Models = new List<string>(),
                    IsDirty = false
                };
            }

            if ((DateTime.Now - lastScanTime).TotalSeconds < ScanCooldownSeconds)
            {
                Log($"[ScanForVoiceModels] Scan skipped, cooldown active");
                return;
            }

            Log($"[ScanForVoiceModels] Starting voice model scan");
            string appDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Log($"[ScanForVoiceModels] Scanning directory: {appDirectory}");

            var newModels = Directory.GetFiles(appDirectory, "*.onnx")
                .Select(Path.GetFileNameWithoutExtension)
                .ToList();

            Log($"[ScanForVoiceModels] Found {newModels.Count} voice models");
            foreach (var model in newModels)
            {
                Log($"[ScanForVoiceModels] Found model: {model}");
            }

            voiceModelState.Models = newModels;
            voiceModels = new List<string>(newModels);
            voiceModelState.IsDirty = true;

            lastScanTime = DateTime.Now;
            Log($"[ScanForVoiceModels] Voice model scan completed");
        }

        public List<string> GetVoiceModels()
        {
            Log($"[GetVoiceModels] Returning {voiceModelState.Models.Count} voice models");
            return voiceModelState.Models;
        }

        private void RefreshVoiceModels_Click(object sender, EventArgs e)
        {
            ScanForVoiceModels();
            PopulateVoiceMenu();
            ForceMenuRefresh();
        }

        private void RebuildContextMenu()
        {
            if (trayIcon.ContextMenuStrip.InvokeRequired)
            {
                trayIcon.ContextMenuStrip.Invoke(new Action(RebuildContextMenu));
                return;
            }

            trayIcon.ContextMenuStrip.SuspendLayout();
            trayIcon.ContextMenuStrip.Items.Clear();
            PopulateContextMenu();
            trayIcon.ContextMenuStrip.ResumeLayout(true);
            trayIcon.ContextMenuStrip.Refresh();
        }

        private void UpdateVoiceModelSetting(string newModel)
        {
            string fileName = Path.GetFileName(newModel);
            string configPath = GetConfigPath();
            try
            {
                var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
                UpdateOrAddSetting(lines, "VoiceModel", fileName);
                File.WriteAllLines(configPath, lines);
                Log($"[UpdateVoiceModelSetting] VoiceModel set to {fileName}");
            }
            catch (Exception ex)
            {
                Log($"[UpdateVoiceModelSetting] Error updating VoiceModel: {ex.Message}");
                MessageBox.Show($"Failed to save VoiceModel setting: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string ReadSettingValue(string key)
        {
            string configPath = GetConfigPath();
            if (File.Exists(configPath))
            {
                try
                {
                    var lines = File.ReadAllLines(configPath);
                    var setting = lines.FirstOrDefault(l => l.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase));
                    if (setting != null)
                    {
                        return setting.Substring(key.Length + 1).Trim();
                    }
                }
                catch (Exception ex)
                {
                    Log($"[ReadSettingValue] Exception reading key '{key}': {ex.Message}");
                }
            }
            else
            {
                Log($"[ReadSettingValue] settings.conf not found. Key '{key}' not found.");
            }
            return string.Empty;
        }

        private void UpdateVoiceModelUI()
        {
            PopulateVoiceMenu();
            try
            {
                Log($"[UpdateVoiceModelUI] Attempting to force menu refresh");
                ForceMenuRefresh();
                Log($"[UpdateVoiceModelUI] Menu refresh completed successfully");
            }
            catch (Exception ex)
            {
                Log($"[UpdateVoiceModelUI] Error during menu refresh: {ex.Message}");
                Log($"[UpdateVoiceModelUI] Stack trace: {ex.StackTrace}");
            }
        }

        private void PopulateVoiceMenu()
        {
            voiceMenuItem.DropDownItems.Clear();
            string currentModel = ReadSettingValue("VoiceModel");

            foreach (string model in voiceModelState.Models)
            {
                CustomVoiceMenuItem item = new CustomVoiceMenuItem
                {
                    Text = model,
                    IsSelected = (Path.GetFileNameWithoutExtension(model) == currentModel),
                    SelectionColor = Color.Green,
                    CheckMarkEnabled = false
                };
                item.Click += VoiceMenuItem_Click;
                voiceMenuItem.DropDownItems.Add(item);
            }

            voiceMenuItem.DropDownItems.Add(new ToolStripSeparator());
            var refreshItem = new ToolStripMenuItem("Refresh");
            refreshItem.Click += RefreshVoiceModels_Click;
            voiceMenuItem.DropDownItems.Add(refreshItem);
        }


        private void VoiceMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is CustomVoiceMenuItem clickedItem)
            {
                string selectedModel = clickedItem.Text;

                UpdateVoiceModelSetting(selectedModel);

                foreach (ToolStripItem item in voiceMenuItem.DropDownItems)
                {
                    if (item is CustomVoiceMenuItem customItem)
                    {
                        customItem.IsSelected = (customItem.Text == selectedModel);
                    }
                }

                voiceMenuItem.Invalidate(); // Force redraw
            }
        }

        private void SaveVoiceModelSetting(string modelName)
        {
            string configPath = GetConfigPath();
            try
            {
                var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
                UpdateOrAddSetting(lines, "VoiceModel", modelName);
                File.WriteAllLines(configPath, lines);
                Log($"[SaveVoiceModelSetting] VoiceModel set to {modelName}");
            }
            catch (Exception ex)
            {
                Log($"[SaveVoiceModelSetting] Error updating VoiceModel: {ex.Message}");
                MessageBox.Show($"Failed to save VoiceModel setting: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshVoiceModels()
        {
            LoadVoiceModels();
            UpdateVoiceModelUI();
            ForceMenuRefresh();
        }

        private void RefreshVoiceModelList()
        {
            try
            {
                LoadVoiceModels();
                Log($"[RefreshVoiceModelList] Loaded {voiceModelState?.Models?.Count ?? 0} voice models");

                if (voiceModelState?.Models?.Any() == true)
                {
                    var (model, speed, logging, monitoringHotkeyModifiers, monitoringHotkeyVk, changeVoiceHotkeyModifiers, changeVoiceVk, monitoringEnabled, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence) = ReadSettings();
                    Log($"[RefreshVoiceModelList] Current model from settings: {model}");

                    if (!string.IsNullOrEmpty(model))
                    {
                        voiceModelState.CurrentIndex = voiceModelState.Models.FindIndex(m =>
                            Path.GetFileName(m).Equals(model, StringComparison.OrdinalIgnoreCase));
                        Log($"[RefreshVoiceModelList] Updated currentVoiceModelIndex: {voiceModelState.CurrentIndex}");

                        if (voiceModelState.CurrentIndex == -1)
                        {
                            voiceModelState.CurrentIndex = 0;
                            Log($"[RefreshVoiceModelList] Voice model not found in list, reset to index 0");
                        }
                    }
                    else
                    {
                        Log($"[RefreshVoiceModelList] Invalid model name from settings");
                    }
                }
                else
                {
                    Log($"[RefreshVoiceModelList] No voice models available or voiceModelState is null");
                }
            }
            catch (Exception ex)
            {
                Log($"[RefreshVoiceModelList] Error: {ex.Message}");
                Log($"[RefreshVoiceModelList] Stack trace: {ex.StackTrace}");
            }
        }

        private void UpdateVoiceMenu()
        {
            for (int i = 0; i < voiceMenuItem.DropDownItems.Count; i++)
            {
                if (voiceMenuItem.DropDownItems[i] is ToolStripMenuItem item)
                {
                    bool oldChecked = item.Checked;
                    item.Checked = (i == voiceModelState.CurrentIndex);
                    Log($"[{DateTime.Now:HH:mm:ss.fff}] [UpdateVoiceMenu] Item {i}: '{item.Text}', Old Checked: {oldChecked}, New Checked: {item.Checked}");
                }
            }
        }

        private void UpdateVoiceMenuCheckedState()
        {

            for (int i = 0; i < voiceMenuItem.DropDownItems.Count; i++)
            {
                if (voiceMenuItem.DropDownItems[i] is ToolStripMenuItem menuItem)
                {
                    bool oldCheckedState = menuItem.Checked;
                    bool newCheckedState = (i == voiceModelState.CurrentIndex);
                    menuItem.Checked = newCheckedState;

                }
            }

            voiceMenuItem.Invalidate();
            trayIcon.ContextMenuStrip.Refresh();

        }

        private void ForceMenuRefresh()
        {

            trayIcon.ContextMenuStrip.SuspendLayout();
            PopulateContextMenu();
            trayIcon.ContextMenuStrip.ResumeLayout();

        }

        private void PopulateContextMenu()
        {
            Log("[PopulateContextMenu] Starting menu population");

            toggleMonitoringMenuItem = new ToolStripMenuItem("Monitoring")
            {
                Checked = isMonitoring
            };
            toggleMonitoringMenuItem.Click += (s, e) =>
            {
                Log("[PopulateContextMenu] Monitoring menu item clicked");
                ToggleMonitoring(s, e);
            };

            trayIcon.ContextMenuStrip.Items.Add(toggleMonitoringMenuItem);
            trayIcon.ContextMenuStrip.Items.Add(speedMenuItem);
            trayIcon.ContextMenuStrip.Items.Add(voiceMenuItem);
            PopulateVoiceMenu();
            presetsMenuItem = new ToolStripMenuItem("Presets");
            AddPresetsToContextMenu();

            exportMenuItem = new ToolStripMenuItem("Export to WAV", null, SafeEventHandler(ExportWav));

            // Apply visibility setting before adding to the menu
            if (menuVisibilitySettings.TryGetValue("Export to WAV", out bool isVisible))
            {
                exportMenuItem.Visible = isVisible;
            }

            trayIcon.ContextMenuStrip.Items.Add(exportMenuItem);
            trayIcon.ContextMenuStrip.Items.Add(presetsMenuItem);
            var settingsMenuItem = new ToolStripMenuItem("Settings", null, SafeEventHandler(OpenSettings));
            trayIcon.ContextMenuStrip.Items.Add(settingsMenuItem);
            trayIcon.ContextMenuStrip.Items.Add(exitMenuItem);
        }

        private async void ExportWav(object sender, EventArgs e)
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "WAV files (*.wav)|*.wav";
                saveDialog.DefaultExt = "wav";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string clipboardText = Clipboard.GetText();
                    if (string.IsNullOrEmpty(clipboardText))
                    {
                        MessageBox.Show("No text in clipboard to convert.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    using (var memoryStream = new MemoryStream())
                    {
                        var processedText = ProcessLine(clipboardText);
                        var (model, speed, _, _, _, _, _, _, _, _, _, _, speaker, sentenceSilence) = ReadSettings();

                        ProcessStartInfo psi = new ProcessStartInfo
                        {
                            FileName = piperPath,
                            Arguments = $"--model {model}.onnx --output-raw --length-scale {speed} --speaker {speaker} --sentence-silence {sentenceSilence}",
                            UseShellExecute = false,
                            RedirectStandardInput = true,
                            RedirectStandardOutput = true,
                            CreateNoWindow = true
                        };

                        using (Process process = new Process())
                        {
                            process.StartInfo = psi;
                            process.Start();

                            using (var writer = new StreamWriter(process.StandardInput.BaseStream, Encoding.UTF8))
                            {
                                await writer.WriteLineAsync(processedText);
                                writer.Close();
                            }

                            await process.StandardOutput.BaseStream.CopyToAsync(memoryStream);
                            memoryStream.Position = 0;

                            using (var rawStream = new RawSourceWaveStream(memoryStream, new WaveFormat(22050, 16, 1)))
                            using (var waveStream = new WaveFileWriter(saveDialog.FileName, rawStream.WaveFormat))
                            {
                                await rawStream.CopyToAsync(waveStream);
                            }
                        }
                    }
                    MessageBox.Show($"Audio exported successfully to {saveDialog.FileName}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }


        private void OpenSettings(object sender, EventArgs e)
        {
            try
            {
                Log($"Opening Settings form");
                var settingsForm = SettingsForm.GetInstance();
                settingsForm.ShowSettingsForm();
                Log($"Settings form displayed successfully");
            }
            catch (Exception ex)
            {
                Log($"Error in OpenSettings: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
            }
        }

        private void SettingsForm_VoiceModelChanged(object sender, EventArgs e)
        {
            var (model, speed, logging, monitoringHotkeyModifiers, monitoringHotkeyVk, changeVoiceHotkeyModifiers, changeVoiceVk, monitoringEnabled, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence) = ReadSettings();

            UpdateVoiceModelAndRefresh(model);

            UpdateVoiceMenu();
            ForceMenuRefresh();
            UpdateVoiceMenuCheckedState();

        }

        private void SettingsForm_SpeedChanged(object sender, double newSpeed)
        {
            Log($"[SettingsForm_SpeedChanged] Entering method. New speed: {newSpeed}");
            UpdateSpeedFromSettings(newSpeed);
            Log($"[SettingsForm_SpeedChanged] Exiting method");
        }

        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            RefreshVoiceModels();
        }

        public void UpdateVoiceModel(string newModel)
        {
            LoadVoiceModels(); // Refresh the list of voice models
            if (voiceModelState?.Models == null)
            {
                Log($"[UpdateVoiceModel] voiceModelState.Models is null. Cannot update voice model.");
                return;
            }
            string newModelName = Path.GetFileNameWithoutExtension(newModel);
            voiceModelState.CurrentIndex = voiceModelState.Models.FindIndex(m => Path.GetFileNameWithoutExtension(m) == newModelName);
            Log($"[UpdateVoiceModel] New currentVoiceModelIndex: {voiceModelState.CurrentIndex}");
            if (voiceModelState.CurrentIndex == -1)
            {
                voiceModelState.CurrentIndex = 0;
                Log($"[UpdateVoiceModel] Voice model not found in list, reset to index 0");
            }
            UpdateVoiceModelSetting(newModel);
            UpdateVoiceMenu();
            ForceMenuRefresh();
        }

        private async void ClipboardTimer_Tick(object sender, EventArgs e)
        {
            if (CurrentAudioState != AudioPlaybackState.Idle)
            {
                if (currentWaveOut?.PlaybackState != PlaybackState.Playing)
                {
                    Log($"[ClipboardTimer_Tick] Resetting audio state from {CurrentAudioState} to Idle.");
                    CurrentAudioState = AudioPlaybackState.Idle;
                }
            }

            if (!isMonitoring || isProcessing)
            {
                return;
            }

            if (Clipboard.ContainsText())
            {
                string clipboardContent = Clipboard.GetText();
                if (clipboardContent != lastClipboardContent)
                {
                    lastClipboardContent = clipboardContent;
                    if (ignoreCurrentClipboard)
                    {
                        ignoreCurrentClipboard = false;
                        return;
                    }
                    Log($"New clipboard content detected: {clipboardContent.Substring(0, Math.Min(50, clipboardContent.Length))}...");
                    isProcessing = true;
                    await ConvertAndPlayTextToSpeechAsync(clipboardContent);
                    isProcessing = false;
                }
            }
        }

        private (string model, float speed, bool logging, uint monitoringHotkeyModifiers, uint monitoringHotkeyVk, uint changeVoiceHotkeyModifiers, uint changeVoiceHotkeyVk, bool monitoringEnabled, uint speedIncreaseHotkeyModifiers, uint speedIncreaseHotkeyVk, uint speedDecreaseHotkeyModifiers, uint speedDecreaseHotkeyVk, int speaker, float sentenceSilence) ReadSettings()
        {
            string settingsPath = GetConfigPath();
            string model = "";
            float speed = 0f;
            bool logging = false;
            uint monitoringHotkeyModifiers = 0;
            uint monitoringHotkeyVk = 0;
            uint changeVoiceHotkeyModifiers = 0;
            uint changeVoiceHotkeyVk = 0;
            uint speedIncreaseHotkeyModifiers = 0;
            uint speedIncreaseHotkeyVk = 0;
            uint speedDecreaseHotkeyModifiers = 0;
            uint speedDecreaseHotkeyVk = 0;
            bool monitoringEnabled = true;
            int speaker = 0;
            float sentenceSilence = 0.2f;
            bool defaultsUsed = false;

            if (!File.Exists(settingsPath))
            {
                defaultsUsed = true;
                SaveSettings();
            }

            Log($"[ReadSettings] Reading settings from: {settingsPath}");
            if (File.Exists(settingsPath))
            {
                foreach (string line in File.ReadAllLines(settingsPath))
                {
                    var parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        switch (parts[0].Trim())
                        {
                            case "VoiceModel":
                                model = Path.GetFileNameWithoutExtension(parts[1].Trim());
                                break;
                            case "Speed":
                                if (double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double parsedSpeed))
                                {
                                    speed = (float)parsedSpeed;
                                }
                                break;
                            case "Logging":
                                bool.TryParse(parts[1].Trim(), out logging);
                                break;
                            case "MonitoringModifier":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out monitoringHotkeyModifiers);
                                break;
                            case "MonitoringKey":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out monitoringHotkeyVk);
                                break;
                            case "StopSpeechModifier":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out stopSpeechModifiers);
                                break;
                            case "StopSpeechKey":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out stopSpeechVk);
                                break;
                            case "ChangeVoiceModifier":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out changeVoiceHotkeyModifiers);
                                break;
                            case "ChangeVoiceKey":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out changeVoiceHotkeyVk);
                                break;
                            case "SpeedIncreaseModifier":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out speedIncreaseHotkeyModifiers);
                                break;
                            case "SpeedIncreaseKey":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out speedIncreaseHotkeyVk);
                                break;
                            case "SpeedDecreaseModifier":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out speedDecreaseHotkeyModifiers);
                                break;
                            case "SpeedDecreaseKey":
                                uint.TryParse(parts[1].Trim().Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out speedDecreaseHotkeyVk);
                                break;
                            case "MonitoringEnabled":
                                bool.TryParse(parts[1].Trim(), out monitoringEnabled);
                                break;
                            case "Speaker":
                                int.TryParse(parts[1].Trim(), out speaker);
                                break;
                            case "SentenceSilence":
                                if (float.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float parsedSilence))
                                {
                                    sentenceSilence = parsedSilence;
                                    Log($"[ReadSettings] Parsed SentenceSilence value: {sentenceSilence}");
                                }
                                break;
                        }
                    }
                }
            }

            if (defaultsUsed)
            {
                SaveSettings();
            }

            Log($"Voice model read from settings: {model}");
            return (model, speed, logging, monitoringHotkeyModifiers, monitoringHotkeyVk, changeVoiceHotkeyModifiers, changeVoiceHotkeyVk, monitoringEnabled, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence);
        }

        private void ApplyMonitoringState(bool enabled)
        {
            Log($"[ApplyMonitoringState] Setting monitoring state to: {enabled}");
            if (enabled)
            {
                StartMonitoring();
            }
            else
            {
                StopMonitoring();
            }
            SaveMonitoringState(enabled);
        }

        private (HashSet<string> ignoreWords, HashSet<string> bannedWords, Dictionary<string, string> replaceWords) LoadDictionaries()
        {
            var ignoreWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "#", "*" };
            var bannedWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var replaceWords = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (File.Exists(Path.Combine(baseDir, "ignore.dict")))
                ignoreWords.UnionWith(File.ReadAllLines(Path.Combine(baseDir, "ignore.dict")));

            if (File.Exists(Path.Combine(baseDir, "banned.dict")))
                bannedWords = new HashSet<string>(File.ReadAllLines(Path.Combine(baseDir, "banned.dict")), StringComparer.OrdinalIgnoreCase);

            if (File.Exists(Path.Combine(baseDir, "replace.dict")))
            {
                foreach (var line in File.ReadAllLines(Path.Combine(baseDir, "replace.dict")))
                {
                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                        replaceWords[parts[0].Trim()] = parts[1].Trim();
                }
            }

            return (ignoreWords, bannedWords, replaceWords);
        }

        private async Task ConvertAndPlayTextToSpeechAsync(string text)
        {
            var processedTextBuilder = new StringBuilder();
            Log($"Original text: {text}");

            using (var reader = new StringReader(text))
            {
                string line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    if (!bannedWords.Any(word => line.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        var processedLine = ProcessLine(line);
                        Log($"Processed line before Piper: {processedLine}");
                        if (!string.IsNullOrWhiteSpace(processedLine))
                        {
                            processedTextBuilder.AppendLine(processedLine);
                        }
                    }
                }
            }

            var processedText = processedTextBuilder.ToString();
            Log($"Final text sent to Piper: {processedText}");
            var (model, speed, logging, monitoringHotkeyModifiers, monitoringHotkeyVk, changeVoiceHotkeyModifiers, changeVoiceHotkeyVk, monitoringEnabled, speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk, speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence) = ReadSettings();

            Log($"Initializing Piper process with model: {model}");
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = piperPath,
                Arguments = $"--model {model}.onnx --output-raw --length-scale {speed.ToString(System.Globalization.CultureInfo.InvariantCulture)} --speaker {currentSpeaker} --sentence-silence {sentenceSilence.ToString(System.Globalization.CultureInfo.InvariantCulture)}",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(piperPath),
                StandardInputEncoding = Encoding.UTF8,
                StandardOutputEncoding = Encoding.UTF8
            };

            try
            {
                if (CurrentAudioState == AudioPlaybackState.Playing)
                {
                    CurrentAudioState = AudioPlaybackState.Stopping;
                    if (currentWaveOut != null)
                    {
                        currentWaveOut.Stop();
                        currentWaveOut.Dispose();
                        currentWaveOut = null;
                    }
                }

                using (Process process = new Process())
                {
                    process.StartInfo = psi;
                    process.Start();
                    Log($"Piper process started with model: {model}");

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            Log($"Piper: {e.Data}");
                        }
                    };
                    process.BeginErrorReadLine();

                    playbackCancellationTokenSource = new CancellationTokenSource();
                    CurrentAudioState = AudioPlaybackState.Playing;

                    using (var writer = new StreamWriter(process.StandardInput.BaseStream, Encoding.UTF8))
                    {
                        await writer.WriteLineAsync(processedText);
                        writer.Close();
                    }

                    await StreamAudioPlayback(process);

                    if (!process.HasExited)
                    {
                        process.Kill();
                        Log($"Piper process terminated after audio playback");
                    }

                    await Task.Run(() => process.WaitForExit());
                    Log($"Piper process exited with code: {process.ExitCode}");
                }
            }
            catch (Exception ex)
            {
                Log($"Error in ConvertAndPlayTextToSpeech: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CurrentAudioState = AudioPlaybackState.Idle;
                playbackCancellationTokenSource?.Dispose();
                playbackCancellationTokenSource = null;
            }
        }

        private string ProcessLine(string line)
        {
            Log($"Step 1 - Raw input: {line}");

            // Currency abbreviations
            line = Regex.Replace(line, @"GBP\s*(\d+)", "$1 pounds");
            line = Regex.Replace(line, @"USD\s*(\d+)", "$1 dollars");
            line = Regex.Replace(line, @"EUR\s*(\d+)", "$1 euros");
            line = Regex.Replace(line, @"JPY\s*(\d+)", "$1 yen");
            line = Regex.Replace(line, @"AUD\s*(\d+)", "$1 australian dollars");
            line = Regex.Replace(line, @"CAD\s*(\d+)", "$1 canadian dollars");
            line = Regex.Replace(line, @"CHF\s*(\d+)", "$1 swiss francs");
            line = Regex.Replace(line, @"CNY\s*(\d+)", "$1 yuan");
            line = Regex.Replace(line, @"INR\s*(\d+)", "$1 rupees");

            // Currency symbols with decimals
            line = Regex.Replace(line, @"£(\d+)\.(\d{2})", "$1 pounds $2 pence");
            line = Regex.Replace(line, @"£(\d+)", "$1 pounds");

            line = Regex.Replace(line, @"\$(\d+)\.(\d{2})", "$1 dollars $2 cents");
            line = Regex.Replace(line, @"\$(\d+)", "$1 dollars");

            line = Regex.Replace(line, @"€(\d+)\.(\d{2})", "$1 euros $2 cents");
            line = Regex.Replace(line, @"€(\d+)", "$1 euros");

            line = Regex.Replace(line, @"¥(\d+)", "$1 yen");
            line = Regex.Replace(line, @"₹(\d+)", "$1 rupees");
            line = Regex.Replace(line, @"₣(\d+)", "$1 francs");
            line = Regex.Replace(line, @"元(\d+)", "$1 yuan");

            // Handle currency at the end of amount
            line = Regex.Replace(line, @"(\d+)\s*pounds?", "$1 pounds");
            line = Regex.Replace(line, @"(\d+)\s*dollars?", "$1 dollars");
            line = Regex.Replace(line, @"(\d+)\s*euros?", "$1 euros");
            line = Regex.Replace(line, @"(\d+)\s*yen", "$1 yen");
            line = Regex.Replace(line, @"(\d+)\s*rupees?", "$1 rupees");
            line = Regex.Replace(line, @"(\d+)\s*francs?", "$1 francs");
            line = Regex.Replace(line, @"(\d+)\s*yuan", "$1 yuan");

            line = line.Replace("#", "").Replace("*", "");
            Log($"Step 2 - After character cleanup: {line}");

            line = Regex.Replace(line, @"(\w+)([.!?])(\s*)$", "$1, . . . . . . . $2$3");

            char[] preservePunctuation = { '.', ',', '?', '!', ':', ';' };
            var words = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var processedWords = new List<string>();

            foreach (var word in words)
            {
                if (!ignoreWords.Contains(word))
                {
                    var processedWord = ApplyReplacements(word);
                    foreach (char punct in preservePunctuation)
                    {
                        if (word.EndsWith(punct.ToString()) && !processedWord.EndsWith(punct.ToString()))
                        {
                            processedWord += punct;
                        }
                    }
                    processedWords.Add(processedWord);
                }
            }

            var result = string.Join(" ", processedWords);
            Log($"Final processed output: {result}");

            return result;
        }

        private string ApplyReplacements(string word)
        {

            // Extract quoted content if present
            Match quotedMatch = Regex.Match(word, @"""(\w+)""");
            if (quotedMatch.Success)
            {
                string quotedWord = quotedMatch.Groups[1].Value;
                Log($"Found quoted word: {quotedWord}");
                return $"\"{quotedWord}\"";  // Return quoted word unchanged
            }

            // Apply replacements only for non-quoted text
            foreach (var replace in replaceWords)
            {
                var beforeReplace = word;
                word = Regex.Replace(word, replace.Key, replace.Value, RegexOptions.IgnoreCase);
                if (beforeReplace != word)
                {
                    Log($"Replacement changed word from '{beforeReplace}' to '{word}' using pattern '{replace.Key}'");
                }
            }

            return word;
        }

        private async Task StreamAudioPlayback(Process piperProcess)
        {
            if (piperProcess?.StandardOutput?.BaseStream == null)
            {
                Log($"[StreamAudioPlayback] Process or stream is null");
                return;
            }

            using (var memoryStream = new MemoryStream())
            {
                try
                {
                    await piperProcess.StandardOutput.BaseStream.CopyToAsync(memoryStream);
                    memoryStream.Position = 0;

                    if (CurrentAudioState == AudioPlaybackState.StopRequested)
                    {
                        Log($"[StreamAudioPlayback] Stop requested before playback started");
                        return;
                    }

                    using (var waveStream = new RawSourceWaveStream(memoryStream, new WaveFormat(22050, 16, 1)))
                    using (currentWaveOut = new WaveOutEvent())
                    {
                        currentWaveOut.Init(waveStream);
                        CurrentAudioState = AudioPlaybackState.Playing;
                        currentWaveOut.Play();

                        while (currentWaveOut.PlaybackState == PlaybackState.Playing)
                        {
                            if (playbackCancellationTokenSource?.Token.IsCancellationRequested == true)
                            {
                                break;
                            }
                            await Task.Delay(10);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[StreamAudioPlayback] Error during playback: {ex.Message}");
                }
                finally
                {
                    CurrentAudioState = AudioPlaybackState.Idle;
                    currentWaveOut = null;
                }
            }
        }

        private Task PlayAudioWithWaveOutEvent(MemoryStream audioStream)
        {
            playbackCancellationTokenSource = new CancellationTokenSource();
            return Task.Run(() =>
            {
                try
                {
                    Log($"[PlayAudioWithWaveOutEvent] Starting playback. AudioStream length: {audioStream.Length}");
                    audioStream.Position = 0;
                    using (var rawStream = new RawSourceWaveStream(audioStream, new WaveFormat(22050, 16, 1)))
                    using (currentWaveOut = new WaveOutEvent())
                    {
                        currentWaveOut.Init(rawStream);
                        CurrentAudioState = AudioPlaybackState.Playing;
                        currentWaveOut.Play();
                        Log($"[PlayAudioWithWaveOutEvent] Playback started. State: {currentWaveOut.PlaybackState}");

                        while (currentWaveOut.PlaybackState == PlaybackState.Playing)
                        {
                            if (playbackCancellationTokenSource.Token.IsCancellationRequested)
                            {
                                Log($"[PlayAudioWithWaveOutEvent] Cancellation detected. Stopping playback.");
                                currentWaveOut.Stop();
                                break;
                            }
                            Thread.Sleep(1);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[PlayAudioWithWaveOutEvent] Exception: {ex.Message}");
                }
                finally
                {
                    CurrentAudioState = AudioPlaybackState.Idle;
                    currentWaveOut = null;
                    Log($"[PlayAudioWithWaveOutEvent] Playback ended. Final state: {CurrentAudioState}");
                }
            });
        }

        private void Exit(object sender, EventArgs e)
        {
            // Stop all background operations
            clipboardTimer?.Stop();
            UnregisterAllHotkeys();

            // Clean up tray icon
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
                trayIcon = null;
            }

            // Exit the application directly
            Environment.Exit(0);
        }

        public void SaveSettings(double? speed = null, string voiceModel = null, int? speaker = null, float? sentenceSilence = null)
        {
            Log($"[SaveSettings] Entering method. Saving current settings to file.");
            string configPath = GetConfigPath();
            var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();

            UpdateOrAddSetting(lines, "MonitoringEnabled", isMonitoring.ToString());

            if (speed.HasValue)
            {
                UpdateOrAddSetting(lines, "Speed", speed.Value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture));
            }
            if (!string.IsNullOrEmpty(voiceModel))
            {
                UpdateOrAddSetting(lines, "VoiceModel", Path.GetFileName(voiceModel));
            }
            if (speaker.HasValue)
            {
                UpdateOrAddSetting(lines, "Speaker", speaker.Value.ToString());
            }
            if (sentenceSilence.HasValue)
            {
                UpdateOrAddSetting(lines, "SentenceSilence", sentenceSilence.Value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture));
            }

            UpdateOrAddSetting(lines, "Logging", isLoggingEnabled.ToString());
            UpdateOrAddSetting(lines, "MonitoringModifier", $"0x{monitoringModifiers:X2}");
            UpdateOrAddSetting(lines, "MonitoringKey", $"0x{monitoringVk:X2}");
            UpdateOrAddSetting(lines, "StopSpeechModifier", $"0x{stopSpeechModifiers:X2}");
            UpdateOrAddSetting(lines, "StopSpeechKey", $"0x{stopSpeechVk:X2}");
            UpdateOrAddSetting(lines, "ChangeVoiceModifier", $"0x{changeVoiceModifiers:X2}");
            UpdateOrAddSetting(lines, "ChangeVoiceKey", $"0x{changeVoiceVk:X2}");
            UpdateOrAddSetting(lines, "SpeedIncreaseModifier", $"0x{speedIncreaseModifiers:X2}");
            UpdateOrAddSetting(lines, "SpeedIncreaseKey", $"0x{speedIncreaseVk:X2}");
            UpdateOrAddSetting(lines, "SpeedDecreaseModifier", $"0x{speedDecreaseModifiers:X2}");
            UpdateOrAddSetting(lines, "SpeedDecreaseKey", $"0x{speedDecreaseVk:X2}");

            try
            {
                File.WriteAllLines(configPath, lines);
                Log($"Settings saved successfully.");
            }
            catch (Exception ex)
            {
                Log($"[SaveSettings] Error writing to settings.conf: {ex.Message}");
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SaveHotkeySettings(uint modifiers, uint vk)
        {
            string configPath = GetConfigPath();
            try
            {
                var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
                UpdateOrAddSetting(lines, "PauseResumeModifier", $"0x{modifiers:X2}");
                UpdateOrAddSetting(lines, "PauseResumeKey", $"0x{vk:X2}");
                File.WriteAllLines(configPath, lines);
                Log($"[SaveHotkeySettings] PauseResume Hotkey saved. Modifier: 0x{modifiers:X2}, Key: 0x{vk:X2}");
            }
            catch (Exception ex)
            {
                Log($"[SaveHotkeySettings] Error saving hotkey settings: {ex.Message}");
                MessageBox.Show($"Failed to save hotkey settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateOrAddSetting(List<string> lines, string key, string value)
        {
            int index = lines.FindIndex(l => l.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase));
            if (index != -1)
            {
                lines[index] = $"{key}={value}";
                Log($"[UpdateOrAddSetting] Updated setting '{key}' to '{value}' at line {index + 1}.");
            }
            else
            {
                lines.Add($"{key}={value}");
                Log($"[UpdateOrAddSetting] Added new setting '{key}' with value '{value}'.");
            }
        }

        private Dictionary<string, string> ReadCurrentSettings()
        {
            var settings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string configPath = GetConfigPath();
            if (File.Exists(configPath))
            {
                try
                {
                    foreach (var line in File.ReadAllLines(configPath))
                    {
                        var parts = line.Split(new[] { '=' }, 2);
                        if (parts.Length == 2)
                        {
                            string key = parts[0].Trim();
                            string value = parts[1].Trim();
                            if (!settings.ContainsKey(key))
                            {
                                settings.Add(key, value);
                            }
                            else
                            {
                                Log($"[ReadCurrentSettings] Duplicate key '{key}' found in settings.conf. Ignoring duplicate.");
                            }
                        }
                        else
                        {
                            Log($"[ReadCurrentSettings] Invalid line format: '{line}'. Skipping.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[ReadCurrentSettings] Exception while reading settings: {ex.Message}");
                }
            }
            else
            {
                Log("[ReadCurrentSettings] settings.conf not found. Returning empty settings.");
            }
            return settings;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                var (model, speed, logging, monitoringHotkeyModifiers, monitoringHotkeyVk,
                    changeVoiceHotkeyModifiers, changeVoiceHotkeyVk, monitoringEnabled,
                    speedIncreaseHotkeyModifiers, speedIncreaseHotkeyVk,
                    speedDecreaseHotkeyModifiers, speedDecreaseHotkeyVk, speaker, sentenceSilence) = ReadSettings();
                UnregisterAllHotkeys();
                trayIcon?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
