using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Globalization;

namespace PiperTray
{
    public enum HotkeyId
    {
        SwitchPreset = 9005
    }

    public class SettingsForm : Form
    {
        private PiperTrayApp app;
        private bool _isInitialized = false;
        private bool isInitializing = true;
        private bool presetsInitialized = false;
        private static readonly Lazy<SettingsForm> lazy = new Lazy<SettingsForm>(() => new SettingsForm());
        private static readonly string[] MODIFIER_OPTIONS = new[] { "ALT", "CTRL", "SHIFT" };
        private static SettingsForm instance;
        private static readonly object _lock = new object();
        public TabControl TabControl => tabControl;
        private Label[] presetLabels;
        private TabPage appearanceTab;

        private Dictionary<string, CheckBox> hotkeyEnableCheckboxes = new Dictionary<string, CheckBox>();
        private Dictionary<string, CheckBox> menuVisibilityCheckboxes;
        private readonly double[] speedOptions = {
            2.0, 1.9, 1.8, 1.7, 1.6, 1.5, 1.4, 1.3, 1.2, 1.1,  // -9 to 0
            1.0, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1   // 1 to 10
        };
        private string currentVoiceModel;

        private ComboBox switchPresetModifierComboBox;
        private TextBox switchPresetKeyTextBox;
        private ComboBox speakerComboBox;
        private int currentSpeaker = 0;
        private int currentPresetIndex = -1;
        private Dictionary<string, int> speakerIdMap = new Dictionary<string, int>();
        private Dictionary<string, int> voiceModelSpeakerCache = new Dictionary<string, int>();

        private NumericUpDown sentenceSilenceNumeric;
        private Label sentenceSilenceLabel;

        public event EventHandler VoiceModelChanged;
        public event EventHandler<double> SpeedChanged;
        protected virtual void OnVoiceModelChanged()
        {
            Log($"[OnVoiceModelChanged] Triggering VoiceModelChanged event");
            VoiceModelChanged?.Invoke(this, EventArgs.Empty);
        }
        private const int HOTKEY_ID_MONITORING = 9001;

        private const int topMargin = 10;
        private const int controlSpacing = 5;
        private const int labelHeight = 20;
        private const int controlHeight = 25;
        private const int columnWidth = 100;
        private const int rowSpacing = 35;

        private TabControl tabControl;
        private TabPage hotkeysTab;
        private TabPage presetsTab;

        private TextBox[] presetNameTextBoxes = new TextBox[4];
        private ComboBox[] presetVoiceModelComboBoxes = new ComboBox[4];
        private ComboBox[] presetSpeakerComboBoxes = new ComboBox[4];
        private ComboBox[] presetSpeedComboBoxes = new ComboBox[4];
        private NumericUpDown[] presetSilenceNumericUpDowns = new NumericUpDown[4];
        private CheckBox[] presetEnableCheckBoxes = new CheckBox[4];

        public class PresetSettings
        {
            public string Name { get; set; }
            public string VoiceModel { get; set; }
            public string Speaker { get; set; }
            public string Speed { get; set; }
            public string SentenceSilence { get; set; }
            public string Enabled { get; set; }
        }

        private Button[] presetApplyButtons;
        private ComboBox VoiceModelComboBox;
        private List<string> voiceModels;
        private List<string> cachedVoiceModels;
        private List<Panel> presetPanels = new List<Panel>();
        private bool isMonitoring;
        private ComboBox speedComboBox;
        private ComboBox stopSpeechModifierComboBox;
        private TextBox stopSpeechKeyTextBox;
        private ComboBox monitoringModifierComboBox;
        private TextBox monitoringKeyTextBox;
        private ComboBox changeVoiceModifierComboBox;
        private TextBox changeVoiceKeyTextBox;
        private ComboBox speedIncreaseModifierComboBox;
        private TextBox speedIncreaseKeyTextBox;
        private ComboBox speedDecreaseModifierComboBox;
        private TextBox speedDecreaseKeyTextBox;

        private uint stopSpeechModifiers;
        private uint stopSpeechVk;
        private uint monitoringModifiers;
        private uint monitoringVk;
        private uint pauseModifiers;
        private uint pauseVk;
        private uint changeVoiceModifiers;
        private uint changeVoiceVk;
        private uint speedIncreaseModifiers;
        private uint speedIncreaseVk;
        private uint speedDecreaseModifiers;
        private uint speedDecreaseVk;
        private uint switchPresetVk;
        private uint switchPresetModifiers;

        private Button saveButton;
        private Button cancelButton;
        private Button refreshVoiceModelsButton;

        private bool isScanning = false;
        private DateTime lastScanTime = DateTime.MinValue;


        public static SettingsForm GetInstance()
        {
            return lazy.Value;
        }


        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(value);
        }

        private string logFilePath;

        [DllImport("kernel32.dll")]
        static extern uint GetLastError();


        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private SettingsForm()
        {
            isInitializing = true;

            // Base setup
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            logFilePath = Path.Combine(assemblyDirectory, "system.log");
            Log($"SettingsForm constructor called");

            // Initialize the PiperTrayApp instance
            app = PiperTrayApp.Instance;
            app.ActivePresetChanged += App_ActivePresetChanged;
            currentPresetIndex = app.GetCurrentPresetIndex();

            // Initialize form controls and arrays
            InitializeComponent();
            InitializeFields();
            CreateControls();
            InitializeCurrentPreset();

            // Configure window properties
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Icon = PiperTrayApp.GetApplicationIcon();

            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // Load voice models
            voiceModels = Directory.GetFiles(assemblyDirectory, "*.onnx")
                .Select(Path.GetFileName)
                .ToList();

            // Initialize UI and settings
            InitializeUIControls();
            if (IsHandleCreated)
            {
                LoadSettingsIntoUI();
            }

            _isInitialized = true;
            isInitializing = false;
        }

        private void InitializePresetArrays()
        {
            int presetCount = 4; // Number of presets

            // Initialize arrays with the specified size
            presetNameTextBoxes = new TextBox[presetCount];
            presetVoiceModelComboBoxes = new ComboBox[presetCount];
            presetSpeakerComboBoxes = new ComboBox[presetCount];
            presetSpeedComboBoxes = new ComboBox[presetCount];
            presetSilenceNumericUpDowns = new NumericUpDown[presetCount];
            presetApplyButtons = new Button[presetCount];
            presetEnableCheckBoxes = new CheckBox[presetCount];
            presetLabels = new Label[presetCount];
            presetPanels = new List<Panel>();

            // Initialize each control in the arrays
            for (int i = 0; i < presetCount; i++)
            {
                // Initialize TextBoxes
                presetNameTextBoxes[i] = new TextBox();

                // Initialize ComboBoxes
                presetVoiceModelComboBoxes[i] = new ComboBox();
                presetSpeakerComboBoxes[i] = new ComboBox();
                presetSpeedComboBoxes[i] = new ComboBox();

                // Initialize NumericUpDowns
                presetSilenceNumericUpDowns[i] = new NumericUpDown();

                // Initialize Buttons
                presetApplyButtons[i] = new Button();

                // Initialize Labels and CheckBoxes
                presetLabels[i] = new Label();
                presetEnableCheckBoxes[i] = new CheckBox();
            }
        }

        private void App_ActivePresetChanged(object sender, int newPresetIndex)
        {
            Log($"[App_ActivePresetChanged] Active preset changed to index {newPresetIndex}");

            // Update currentPresetIndex
            currentPresetIndex = newPresetIndex;

            // Update the UI on the UI thread
            if (InvokeRequired)
            {
                Invoke(new Action(UpdatePresetUI));
            }
            else
            {
                UpdatePresetUI();
            }
        }

        private void UpdatePresetUI()
        {
            Log($"[UpdatePresetUI] Updating preset UI for current preset index {currentPresetIndex}");
            RefreshPresetPanels();
        }

        private void CreateControls()
        {
            monitoringModifierComboBox = new ComboBox();
            stopSpeechModifierComboBox = new ComboBox();
            changeVoiceModifierComboBox = new ComboBox();
            speedIncreaseModifierComboBox = new ComboBox();
            speedDecreaseModifierComboBox = new ComboBox();
            switchPresetModifierComboBox = new ComboBox();

            monitoringKeyTextBox = new TextBox();
            stopSpeechKeyTextBox = new TextBox();
            changeVoiceKeyTextBox = new TextBox();
            speedIncreaseKeyTextBox = new TextBox();
            speedDecreaseKeyTextBox = new TextBox();
            switchPresetKeyTextBox = new TextBox();
        }

        private void InitializeCurrentPreset()
        {
            var settings = PiperTrayApp.GetInstance().ReadCurrentSettings();
            if (settings.TryGetValue("LastUsedPreset", out string lastUsedPreset))
            {
                currentPresetIndex = int.Parse(lastUsedPreset);
            }
        }

        private void InitializeFields()
        {
            tabControl = new TabControl();
        }



        public void ShowSettingsForm()
        {
            // Force window state to normal before showing
            this.WindowState = FormWindowState.Normal;

            // Refresh active preset status
            InitializeCurrentPreset();
            RefreshPresetPanels();

            // Ensure UI is ready
            if (!_isInitialized)
            {
                InitializeUIControls();
                LoadSettingsIntoUI();
            }

            // Position on primary screen
            Screen primaryScreen = Screen.PrimaryScreen;
            Rectangle workingArea = primaryScreen.WorkingArea;
            this.Location = new Point(
                workingArea.Left + (workingArea.Width - this.Width) / 2,
                workingArea.Top + (workingArea.Height - this.Height) / 2
            );

            // Force visibility and focus
            this.Show();
            this.BringToFront();
            this.Activate();
            this.Focus();
        }

        private void InitializeUIControls()
        {
            Log("[InitializeUIControls] Initializing UI controls");
        }

        private void InitializeComboBoxes()
        {

        }

        private void RefreshPresetPanels()
        {
            foreach (var panel in presetPanels)
            {
                panel.Paint -= Panel_Paint; // Remove any existing paint handlers
                panel.BorderStyle = BorderStyle.None;
                panel.BackColor = SystemColors.Control;

                // Get the index of this panel
                int panelIndex = presetPanels.IndexOf(panel);

                if (panelIndex == currentPresetIndex)
                {
                    panel.Paint += Panel_Paint; // Add the paint handler for the active preset
                }

                panel.Invalidate(); // Force the panel to repaint
            }
        }

        private void Panel_Paint(object sender, PaintEventArgs e)
        {
            Panel panel = sender as Panel;
            if (panel != null)
            {
                using (Pen borderPen = new Pen(Color.LightGreen, 3.0f))
                {
                    Rectangle borderRect = panel.ClientRectangle;
                    borderRect.Inflate(-1, -1);
                    e.Graphics.DrawRectangle(borderPen, borderRect);
                }
            }
        }

        private void PopulateVoiceModelComboBox(ComboBox comboBox)
        {
            if (cachedVoiceModels == null)
            {
                cachedVoiceModels = Directory.GetFiles(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "*.onnx"
                ).Select(Path.GetFileNameWithoutExtension).ToList();
            }

            comboBox.Items.Clear();
            comboBox.Items.AddRange(cachedVoiceModels.ToArray());
            if (comboBox.Items.Count > 0)
                comboBox.SelectedIndex = 0;
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ApplyPreset(int index)
        {
            var preset = PiperTrayApp.GetInstance().LoadPreset(index);
            if (preset == null)
            {
                Log($"[ApplyPreset] Preset {index + 1} is null.");
                return;
            }

            // Parse and apply the preset settings
            double speed = 0;
            if (double.TryParse(preset.Speed, NumberStyles.Float, CultureInfo.InvariantCulture, out speed))
            {
                Log($"[ApplyPreset] Parsed preset speed: {speed}");
                PiperTrayApp.GetInstance().SaveSettings(speed: speed);
            }

            int speaker = 0; // Default speaker ID
            if (int.TryParse(preset.Speaker, out int spk))
            {
                speaker = spk;
                PiperTrayApp.GetInstance().UpdateCurrentSpeaker(speaker);
            }

            float silence = 0.5f; // Default value
            if (float.TryParse(preset.SentenceSilence, NumberStyles.Float, CultureInfo.InvariantCulture, out float s))
            {
                silence = s;
            }

            // Save all settings at once to prevent overwrites
            PiperTrayApp.GetInstance().SaveSettings(
                voiceModel: preset.VoiceModel,
                speaker: speaker,
                speed: speed,
                sentenceSilence: silence);

            Log($"[ApplyPreset] Applied preset {index + 1} with Speed {speed}, Speaker {speaker}, VoiceModel {preset.VoiceModel}");
        }

        private void speakerComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isInitializing && speakerComboBox.SelectedItem != null)
            {
                int selectedSpeaker = int.Parse(speakerComboBox.SelectedItem.ToString());
                var app = PiperTrayApp.GetInstance();
                app.UpdateCurrentSpeaker(selectedSpeaker);
                app.SaveSettings(speaker: selectedSpeaker);
            }
        }

        private void RegisterHotkeys()
        {
            Log($"[RegisterHotkeys] Starting hotkey registration process");
            var mainForm = PiperTrayApp.GetInstance();
            mainForm.UnregisterAllHotkeys();

            var settings = ReadCurrentSettings();

            // Only register hotkeys if they are enabled
            if (settings.TryGetValue("MonitoringHotkeyEnabled", out string monEnabled) && bool.Parse(monEnabled))
            {
                string monitoringMod = monitoringModifierComboBox?.SelectedItem?.ToString() ?? "NONE";
                string monitoringKey = monitoringKeyTextBox?.Text ?? string.Empty;
                uint monitoringModifiers = GetModifierVirtualKeyCode(monitoringMod);
                uint monitoringVk = GetVirtualKeyCode(monitoringKey);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_MONITORING, monitoringModifiers, monitoringVk, "Monitoring");
            }

            if (settings.TryGetValue("StopSpeechHotkeyEnabled", out string stopEnabled) && bool.Parse(stopEnabled))
            {
                string stopSpeechMod = stopSpeechModifierComboBox?.SelectedItem?.ToString() ?? "NONE";
                string stopSpeechKey = stopSpeechKeyTextBox?.Text ?? string.Empty;
                uint stopSpeechModifiers = GetModifierVirtualKeyCode(stopSpeechMod);
                uint stopSpeechVk = GetVirtualKeyCode(stopSpeechKey);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_STOP_SPEECH, stopSpeechModifiers, stopSpeechVk, "Stop Speech");
            }

            if (settings.TryGetValue("ChangeVoiceHotkeyEnabled", out string voiceEnabled) && bool.Parse(voiceEnabled))
            {
                string changeVoiceMod = changeVoiceModifierComboBox?.SelectedItem?.ToString() ?? "NONE";
                string changeVoiceKey = changeVoiceKeyTextBox?.Text ?? string.Empty;
                uint changeVoiceModifiers = GetModifierVirtualKeyCode(changeVoiceMod);
                uint changeVoiceVk = GetVirtualKeyCode(changeVoiceKey);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_CHANGE_VOICE, changeVoiceModifiers, changeVoiceVk, "Change Voice");
            }

            if (settings.TryGetValue("SpeedIncreaseHotkeyEnabled", out string speedIncEnabled) && bool.Parse(speedIncEnabled))
            {
                string speedIncreaseMod = speedIncreaseModifierComboBox?.SelectedItem?.ToString() ?? "NONE";
                string speedIncreaseKey = speedIncreaseKeyTextBox?.Text ?? string.Empty;
                uint speedIncreaseModifiers = GetModifierVirtualKeyCode(speedIncreaseMod);
                uint speedIncreaseVk = GetVirtualKeyCode(speedIncreaseKey);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_SPEED_INCREASE, speedIncreaseModifiers, speedIncreaseVk, "Speed Increase");
            }

            if (settings.TryGetValue("SpeedDecreaseHotkeyEnabled", out string speedDecEnabled) && bool.Parse(speedDecEnabled))
            {
                string speedDecreaseMod = speedDecreaseModifierComboBox?.SelectedItem?.ToString() ?? "NONE";
                string speedDecreaseKey = speedDecreaseKeyTextBox?.Text ?? string.Empty;
                uint speedDecreaseModifiers = GetModifierVirtualKeyCode(speedDecreaseMod);
                uint speedDecreaseVk = GetVirtualKeyCode(speedDecreaseKey);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_SPEED_DECREASE, speedDecreaseModifiers, speedDecreaseVk, "Speed Decrease");
            }

            if (settings.TryGetValue("SwitchPresetHotkeyEnabled", out string switchEnabled) && bool.Parse(switchEnabled))
            {
                string switchPresetMod = switchPresetModifierComboBox?.SelectedItem?.ToString() ?? "NONE";
                string switchPresetKey = switchPresetKeyTextBox?.Text ?? string.Empty;
                uint switchPresetModifiers = GetModifierVirtualKeyCode(switchPresetMod);
                uint switchPresetVk = GetVirtualKeyCode(switchPresetKey);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_SWITCH_PRESET, switchPresetModifiers, switchPresetVk, "Switch Preset");
            }
        }


        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312) // WM_HOTKEY
            {
                int id = m.WParam.ToInt32();
                uint modifiers = (uint)((int)m.LParam & 0xFFFF);
                uint vk = (uint)((int)m.LParam >> 16);

                Log($"[WndProc] Hotkey received. ID: {id}, Modifiers: 0x{modifiers:X2}, VK: 0x{vk:X2}");

                var app = PiperTrayApp.GetInstance();

                switch (id)
                {
                    case PiperTrayApp.HOTKEY_ID_SWITCH_PRESET:
                        Log($"[WndProc] Switch Preset hotkey pressed");
                        app.SwitchPreset();
                        break;
                }
            }
            base.WndProc(ref m);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                base.OnFormClosing(e);
            }
        }

        private void InitializeComponent()
        {
            isInitializing = true;

            this.Text = "Settings";
            this.Size = new System.Drawing.Size(400, 300);

            tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            appearanceTab = new TabPage("Appearance");
            hotkeysTab = new TabPage("Hotkeys");
            presetsTab = new TabPage("Presets");

            tabControl.TabPages.Add(appearanceTab);
            tabControl.TabPages.Add(hotkeysTab);
            tabControl.TabPages.Add(presetsTab);

            InitializePresetArrays();

            // Create preset panels
            for (int i = 0; i < 4; i++)
            {
                CreatePresetPanel(i);
            }

            CreatePresetCheckboxes();

            LoadSavedPresets();

            // Create the menu visibility controls
            CreateMenuVisibilityControls();

            // **Add this line to load the settings**
            LoadMenuVisibilitySettings();

            // Hotkeys Tab
            AddHotkeyControls("Switch Preset:", 10, MODIFIER_OPTIONS);
            AddHotkeyControls("Monitoring:", 40, MODIFIER_OPTIONS);
            AddHotkeyControls("Stop Speech:", 70, MODIFIER_OPTIONS);
            AddHotkeyControls("Change Voice:", 100, MODIFIER_OPTIONS);
            AddHotkeyControls("Speed Increase:", 130, MODIFIER_OPTIONS);
            AddHotkeyControls("Speed Decrease:", 160, MODIFIER_OPTIONS);

            this.Controls.Add(tabControl);

            saveButton = new Button();
            saveButton.Text = "Save";
            saveButton.Location = new System.Drawing.Point(200, this.ClientSize.Height - 40);
            saveButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.Controls.Add(saveButton);
            saveButton.BringToFront();
            saveButton.Click += SaveButton_Click;

            cancelButton = new Button();
            cancelButton.Text = "Cancel";
            cancelButton.Location = new System.Drawing.Point(300, this.ClientSize.Height - 40);
            cancelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.Controls.Add(cancelButton);
            cancelButton.BringToFront();
            cancelButton.Click += CancelButton_Click;

            InitializeHotkeyControls();
            LoadHotkeySettings(ReadCurrentSettings());

            LoadHotkeyEnabledStates(ReadCurrentSettings());

            isInitializing = false;
        }

        private void InitializeHotkeyControls()
        {
            foreach (ComboBox combo in new[] {
                monitoringModifierComboBox,
                stopSpeechModifierComboBox,
                changeVoiceModifierComboBox,
                speedIncreaseModifierComboBox,
                speedDecreaseModifierComboBox,
                switchPresetModifierComboBox
            })
            {
                combo.SelectedIndex = 0;
            }
        }

        private void AddSwitchPresetControls()
        {
            Label switchPresetLabel = new Label
            {
                Text = "Switch Preset:",
                Location = new System.Drawing.Point(10, 190)
            };
            hotkeysTab.Controls.Add(switchPresetLabel);

            switchPresetModifierComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(130, 190),
                Width = 100,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            switchPresetModifierComboBox.Items.AddRange(new[] { "NONE", "ALT", "CTRL", "SHIFT" });
            switchPresetModifierComboBox.SelectedIndex = 0;
            hotkeysTab.Controls.Add(switchPresetModifierComboBox);

            switchPresetKeyTextBox = new TextBox
            {
                Location = new System.Drawing.Point(240, 190),
                Width = 100,
                ReadOnly = true
            };
            hotkeysTab.Controls.Add(switchPresetKeyTextBox);

            // Add key capture handler
            switchPresetKeyTextBox.KeyDown += (sender, e) =>
            {
                if (sender is TextBox textBox)
                {
                    e.Handled = true;
                    if (e.KeyCode != Keys.None && e.KeyCode != Keys.Shift &&
                        e.KeyCode != Keys.Control && e.KeyCode != Keys.Alt)
                    {
                        textBox.Text = e.KeyCode.ToString().ToUpper();
                        Log($"Captured key for Switch Preset: {textBox.Text}");
                    }
                }
            };
        }

        // Add this to LoadHotkeySettings()
        private void LoadSwitchPresetSettings(Dictionary<string, string> settings)
        {
            if (settings.TryGetValue("SwitchPresetModifier", out string switchPresetMod))
            {
                uint modValue = uint.Parse(switchPresetMod.Replace("0x", ""),
                    System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (switchPresetModifierComboBox.Items.Contains(modString))
                {
                    switchPresetModifierComboBox.SelectedItem = modString;
                    Log($"[LoadHotkeySettings] Set Switch Preset modifier to: {modString}");
                }
            }

            if (settings.TryGetValue("SwitchPresetKey", out string switchPresetKey))
            {
                uint keyValue = uint.Parse(switchPresetKey.Replace("0x", ""),
                    System.Globalization.NumberStyles.HexNumber);
                switchPresetKeyTextBox.Text = ((Keys)keyValue).ToString();
                Log($"[LoadHotkeySettings] Set Switch Preset key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void CreateMenuVisibilityControls()
        {
            menuVisibilityCheckboxes = new Dictionary<string, CheckBox>();
            string[] menuItems = new[] {
                "Monitoring",
                "Stop Speech",
                "Speed",
                "Voice",
                "Export to WAV"
            };

            Label headerLabel = new Label
            {
                Text = "Show/Hide Menu Items:",
                Location = new Point(10, 10),
                AutoSize = true
            };
            appearanceTab.Controls.Add(headerLabel);

            int yOffset = 40;
            foreach (string menuItem in menuItems)
            {
                var checkbox = new CheckBox
                {
                    Text = $"Show '{menuItem}'",
                    Location = new Point(20, yOffset),
                    Checked = true,
                    AutoSize = true
                };
                checkbox.CheckedChanged += MenuVisibilityChanged;
                menuVisibilityCheckboxes[menuItem] = checkbox;
                appearanceTab.Controls.Add(checkbox);
                yOffset += 30;
            }
        }

        private void MenuVisibilityChanged(object sender, EventArgs e)
        {
            if (sender is CheckBox checkbox)
            {
                string menuItem = menuVisibilityCheckboxes.First(x => x.Value == checkbox).Key;
                SaveMenuVisibilitySetting(menuItem, checkbox.Checked);
                PiperTrayApp.GetInstance().UpdateMenuVisibility(menuItem, checkbox.Checked);
            }
        }

        private void SaveMenuVisibilitySetting(string menuItem, bool isVisible)
        {
            string configPath = GetConfigPath();
            var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
            UpdateOrAddSetting(lines, $"MenuVisible_{menuItem.Replace(" ", "_")}", isVisible.ToString());
            File.WriteAllLines(configPath, lines);
        }

        private void LoadMenuVisibilitySettings()
        {
            foreach (var kvp in menuVisibilityCheckboxes)
            {
                string settingKey = $"MenuVisible_{kvp.Key.Replace(" ", "_")}";
                string value = ReadSettingValue(settingKey);
                bool isVisible = string.IsNullOrEmpty(value) ? true : bool.Parse(value);
                kvp.Value.Checked = isVisible;
            }
        }

        private void CreatePresetPanel(int index)
        {
            Log($"[CreatePresetPanel] ===== Creating Preset Panel {index + 1} =====");
            presetsTab.SuspendLayout();

            // Create the panel for the preset
            var presetPanel = CreateBasePanelStructure(index);

            // List to hold header labels
            List<Control> controls = new List<Control>();

            if (index == 0)
            {
                CreateHeaderLabels(controls);
            }

            CreatePresetControls(index, presetPanel);
            FinalizePresetPanel(index, controls, presetPanel);

            presetsTab.ResumeLayout(true);
            Log($"[CreatePresetPanel] ===== Preset Panel {index + 1} Creation Complete =====");
        }

        private bool VerifyPresetControls()
        {
            Log("[VerifyPresetControls] Starting control verification");
            bool allValid = true;

            for (int i = 0; i < 4; i++)
            {
                bool presetValid = presetNameTextBoxes[i]?.IsHandleCreated == true &&
                                  presetVoiceModelComboBoxes[i]?.IsHandleCreated == true &&
                                  presetSpeakerComboBoxes[i]?.IsHandleCreated == true &&
                                  presetSpeedComboBoxes[i]?.IsHandleCreated == true &&
                                  presetSilenceNumericUpDowns[i]?.IsHandleCreated == true &&
                                  presetEnableCheckBoxes[i]?.IsHandleCreated == true;

                Log($"[VerifyPresetControls] Preset {i + 1} validation: {(presetValid ? "Valid" : "Invalid")}");
                allValid &= presetValid;
            }

            return allValid;
        }

        private Panel CreateBasePanelStructure(int index)
        {
            int rowY = topMargin + labelHeight + controlSpacing + (index * rowSpacing);
            var presetPanel = new Panel
            {
                Location = new Point(8, rowY - 2),
                Height = controlHeight + 4,
                BorderStyle = BorderStyle.None
            };
            presetPanels.Add(presetPanel);
            return presetPanel;
        }

        private void CreateHeaderLabels(List<Control> controls)
        {
            // Positions based on your layout; adjust as needed
            int labelY = topMargin; // Position Y for the labels
            int nameLabelX = 8; // Starting X position for the first label

            // Calculate X positions based on control widths and spacings
            int nameWidth = (columnWidth - 24);
            int modelWidth = columnWidth;
            int speakerWidth = columnWidth / 2;
            int speedWidth = columnWidth / 2;
            int silenceWidth = columnWidth / 2;

            int modelLabelX = nameLabelX + nameWidth + controlSpacing;
            int speakerLabelX = modelLabelX + modelWidth + controlSpacing;
            int speedLabelX = speakerLabelX + speakerWidth + controlSpacing;
            int silenceLabelX = speedLabelX + speedWidth + controlSpacing;

            // Create labels and set their positions
            Label nameLabel = new Label
            {
                Text = "Name",
                AutoSize = true,
                Location = new Point(nameLabelX, labelY),
                Padding = new Padding(2)
            };

            Label modelLabel = new Label
            {
                Text = "Model",
                AutoSize = true,
                Location = new Point(modelLabelX, labelY),
                Padding = new Padding(2)
            };

            Label speakerLabel = new Label
            {
                Text = "Speaker",
                AutoSize = true,
                Location = new Point(speakerLabelX, labelY),
                Padding = new Padding(2)
            };

            Label speedLabel = new Label
            {
                Text = "Speed",
                AutoSize = true,
                Location = new Point(speedLabelX, labelY),
                Padding = new Padding(2)
            };

            Label silenceLabel = new Label
            {
                Text = "Silence",
                AutoSize = true,
                Location = new Point(silenceLabelX, labelY),
                Padding = new Padding(2)
            };

            // Add the labels to the controls list
            controls.Add(nameLabel);
            controls.Add(modelLabel);
            controls.Add(speakerLabel);
            controls.Add(speedLabel);
            controls.Add(silenceLabel);
        }

        private void CreatePresetCheckboxes()
        {
            int checkboxStartX = 10;
            int checkboxSpacing = 30;
            int bottomMargin = -80;

            for (int i = 0; i < 4; i++)
            {
                var numberLabel = new Label
                {
                    Text = $"{i + 1}",
                    Location = new Point(checkboxStartX + (i * checkboxSpacing), presetsTab.Height - bottomMargin),
                    AutoSize = true
                };

                var checkBox = new CheckBox
                {
                    Text = "",
                    Location = new Point(checkboxStartX + (i * checkboxSpacing), presetsTab.Height - bottomMargin + 20),
                    AutoSize = true,
                    Tag = i
                };

                checkBox.CheckedChanged += presetEnableCheckBoxes_CheckedChanged;

                // Store the checkbox in the array
                presetEnableCheckBoxes[i] = checkBox;

                // Add controls to the UI
                presetsTab.Controls.Add(numberLabel);
                presetsTab.Controls.Add(checkBox);
                numberLabel.BringToFront();
                checkBox.BringToFront();
            }
        }

        private void CreatePresetCheckbox(int i, int startX, int spacing, int bottomMargin)
        {
            var numberLabel = new Label
            {
                Text = $"{i + 1}",
                Location = new Point(startX + (i * spacing), presetsTab.Height - bottomMargin),
                AutoSize = true
            };

            presetEnableCheckBoxes[i] = new CheckBox
            {
                Text = "",
                Location = new Point(startX + (i * spacing), presetsTab.Height - bottomMargin + 20),
                AutoSize = true
            };

            SetupCheckboxEvents(i);
            AddCheckboxControls(numberLabel, presetEnableCheckBoxes[i]);
        }

        private void CreatePresetControls(int index, Panel presetPanel)
        {
            // Load the preset data first
            var preset = PiperTrayApp.GetInstance().LoadPreset(index);

            CreateNameTextBox(index);
            if (preset?.Name != null)
            {
                presetNameTextBoxes[index].Text = preset.Name;
            }

            CreateVoiceModelComboBox(index);
            if (preset?.VoiceModel != null)
            {
                int modelIndex = presetVoiceModelComboBoxes[index].Items.IndexOf(preset.VoiceModel);
                if (modelIndex >= 0)
                {
                    presetVoiceModelComboBoxes[index].SelectedIndex = modelIndex;
                }
            }

            CreateSpeakerComboBox(index);
            if (preset?.Speaker != null)
            {
                presetSpeakerComboBoxes[index].SelectedItem = preset.Speaker;
            }

            CreateSpeedComboBox(index);
            if (preset?.Speed != null && double.TryParse(preset.Speed, out double speed))
            {
                int speedIndex = GetSpeedIndex(speed);
                presetSpeedComboBoxes[index].SelectedIndex = speedIndex + 9;
            }

            CreateSilenceNumericUpDown(index);
            if (preset?.SentenceSilence != null && decimal.TryParse(preset.SentenceSilence, out decimal silence))
            {
                presetSilenceNumericUpDowns[index].Value = silence;
            }

            // Set enabled state from preset
            if (preset != null)
            {
                bool isEnabled = bool.Parse(preset.Enabled);
                presetEnableCheckBoxes[index].Checked = isEnabled;
                UpdatePresetControlsEnabled(index, isEnabled);
            }

            presetPanel.Width = presetSilenceNumericUpDowns[index].Right + 4;
            AddControlsToPanel(index, presetPanel);
        }


        private void FinalizePresetPanel(int index, List<Control> controls, Panel presetPanel)
        {
            presetsTab.Controls.Add(presetPanel);
            controls.Add(presetPanel);
            presetsTab.Controls.AddRange(controls.ToArray());

            if (index == 0)
            {
                BringLabelsToFront(controls);
            }
        }

        private void CreateNameTextBox(int index)
        {
            presetNameTextBoxes[index] = new TextBox
            {
                Location = new Point(2, 2),
                Width = (columnWidth - 24),
                Text = $"Preset {index + 1}"
            };
            presetNameTextBoxes[index].TextChanged += (s, e) => UpdatePresetName(index);
        }

        private void CreateVoiceModelComboBox(int index)
        {
            presetVoiceModelComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetNameTextBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth
            };
            PopulateVoiceModelComboBox(presetVoiceModelComboBoxes[index]);
            presetVoiceModelComboBoxes[index].SelectedIndexChanged += (s, e) => UpdatePresetSpeakers(index);
        }

        private void CreateSpeakerComboBox(int index)
        {
            presetSpeakerComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetVoiceModelComboBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth / 2
            };

            string selectedModel = presetVoiceModelComboBoxes[index].SelectedItem?.ToString();
            if (selectedModel != null)
            {
                int speakerCount = GetSpeakerCountFromCache(selectedModel);

                for (int i = 0; i < speakerCount; i++)
                {
                    presetSpeakerComboBoxes[index].Items.Add(i.ToString());
                }
            }

            // Get preset value from settings
            var preset = PiperTrayApp.GetInstance().LoadPreset(index);
            if (preset?.Speaker != null)
            {
                presetSpeakerComboBoxes[index].SelectedItem = preset.Speaker;
            }
            else if (presetSpeakerComboBoxes[index].Items.Count > 0)
            {
                presetSpeakerComboBoxes[index].SelectedIndex = 0;
            }
        }

        private int GetSpeakerCountFromCache(string modelName)
        {
            if (voiceModelSpeakerCache.TryGetValue(modelName, out int cachedCount))
            {
                return cachedCount;
            }
            else
            {
                int speakerCount = LoadSpeakerCountFromFile(modelName);
                voiceModelSpeakerCache[modelName] = speakerCount;
                return speakerCount;
            }
        }

        private int LoadSpeakerCountFromFile(string modelName)
        {
            string jsonPath = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                $"{modelName}.onnx.json"
            );

            if (File.Exists(jsonPath))
            {
                string jsonContent = File.ReadAllText(jsonPath);
                using JsonDocument doc = JsonDocument.Parse(jsonContent);
                if (doc.RootElement.TryGetProperty("num_speakers", out JsonElement numSpeakers))
                {
                    return numSpeakers.GetInt32();
                }
            }
            // Default to 1 speaker if not found
            return 1;
        }

        private void CreateSpeedComboBox(int index)
        {
            presetSpeedComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetSpeakerComboBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth / 2
            };

            presetSpeedComboBoxes[index].Items.Clear();

            for (int i = -9; i <= 10; i++)
            {
                presetSpeedComboBoxes[index].Items.Add(i.ToString());
            }

            SetupSpeedComboBoxEvents(index);
            LoadSpeedSettings(index);
        }

        private void CreateSilenceNumericUpDown(int index)
        {
            presetSilenceNumericUpDowns[index] = new NumericUpDown
            {
                Location = new Point(presetSpeedComboBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth / 2,
                DecimalPlaces = 1,
                Increment = 0.1m,
                Minimum = 0.0m,
                Maximum = 2.0m,
                Value = 0.5m
            };
        }

        private void SetupSpeedComboBoxEvents(int index)
        {
            presetSpeedComboBoxes[index].SelectedIndex = 9;
            presetSpeedComboBoxes[index].SelectedIndexChanged += (s, e) =>
            {
                if (!isInitializing && index == currentPresetIndex && s is ComboBox cb && cb.SelectedItem != null)
                {
                    Log($"[SpeedComboBox_SelectedIndexChanged] Preset {index} - Selected item: {cb.SelectedItem}");
                    if (int.TryParse(cb.SelectedItem.ToString(), out int selectedIndex))
                    {
                        double speed = GetSpeedValue(selectedIndex);
                        Log($"[SpeedComboBox_SelectedIndexChanged] Converted speed index {selectedIndex} to value {speed}");
                        PiperTrayApp.GetInstance().SaveSettings(speed: speed);
                    }
                }
            };
        }

        private void LoadSpeedSettings(int index)
        {
            if (index == currentPresetIndex)
            {
                var settings = PiperTrayApp.GetInstance().ReadCurrentSettings();
                if (settings.TryGetValue("Speed", out string speedValue))
                {
                    if (double.TryParse(speedValue, out double speed))
                    {
                        int speedIndex = GetSpeedIndex(speed);
                        presetSpeedComboBoxes[index].SelectedIndex = speedIndex + 9;
                    }
                }
            }
        }

        private void SetupCheckboxEvents(int index)
        {
            presetEnableCheckBoxes[index].CheckedChanged += presetEnableCheckBoxes_CheckedChanged;
        }

        private void presetEnableCheckBoxes_CheckedChanged(object sender, EventArgs e)
        {
            if (!isInitializing && sender is CheckBox checkBox && int.TryParse(checkBox.Tag.ToString(), out int presetIndex))
            {
                bool isEnabled = checkBox.Checked;
                UpdatePresetControlsState(presetIndex, isEnabled);

                var preset = PiperTrayApp.GetInstance().LoadPreset(presetIndex);
                if (preset != null)
                {
                    preset.Enabled = isEnabled.ToString();
                    PiperTrayApp.GetInstance().SavePreset(presetIndex, preset);
                }
            }
        }

        private void PresetEnableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (sender is CheckBox checkBox && int.TryParse(checkBox.Tag.ToString(), out int presetIndex))
            {
                bool isEnabled = checkBox.Checked;
                Log($"[PresetEnableCheckBox_CheckedChanged] Preset {presetIndex} checkbox changed to: {isEnabled}");

                // Suspend layout updates
                presetPanels[presetIndex].SuspendLayout();

                try
                {
                    UpdatePresetControlsState(presetIndex, isEnabled);

                    var preset = PiperTrayApp.GetInstance().LoadPreset(presetIndex);
                    if (preset != null)
                    {
                        preset.Enabled = isEnabled.ToString();
                        PiperTrayApp.GetInstance().SavePreset(presetIndex, preset);
                        Log($"[PresetEnableCheckBox_CheckedChanged] Saved preset {presetIndex} enabled state: {isEnabled}");
                    }
                }
                finally
                {
                    // Resume and force layout update
                    presetPanels[presetIndex].ResumeLayout(true);
                    presetPanels[presetIndex].Refresh();
                }
            }
        }

        private void UpdateActivePreset()
        {
            int enabledCount = 0;
            int enabledIndex = -1;

            for (int j = 0; j < presetEnableCheckBoxes.Length; j++)
            {
                if (presetEnableCheckBoxes[j].Checked)
                {
                    enabledCount++;
                    enabledIndex = j;
                }
            }

            if (enabledCount == 1)
            {
                currentPresetIndex = enabledIndex;
                PiperTrayApp.GetInstance().ApplyPreset(enabledIndex);
                RefreshPresetPanels();
                Log($"[UpdateActivePreset] Set active preset to index: {enabledIndex}");
            }
        }

        private void AddControlsToPanel(int index, Panel presetPanel)
        {
            presetPanel.Controls.AddRange(new Control[] {
        presetNameTextBoxes[index],
        presetVoiceModelComboBoxes[index],
        presetSpeakerComboBoxes[index],
        presetSpeedComboBoxes[index],
        presetSilenceNumericUpDowns[index]
    });
        }

        private void AddCheckboxControls(Label numberLabel, CheckBox checkbox)
        {
            presetsTab.Controls.Add(numberLabel);
            presetsTab.Controls.Add(checkbox);
            numberLabel.BringToFront();
            checkbox.BringToFront();
        }

        private void BringLabelsToFront(List<Control> controls)
        {
            foreach (Control control in controls)
            {
                if (control is Label)
                {
                    control.BringToFront();
                }
            }
        }

        private void UpdatePresetIndicators()
        {
            foreach (Control control in presetsTab.Controls)
            {
                if (control is Label label && label.Text == "►")
                {
                    int presetIndex = (label.Location.Y - (topMargin + labelHeight + controlSpacing)) / rowSpacing;
                    label.Visible = (presetIndex == currentPresetIndex);
                }
            }
        }

        private void UpdatePresetName(int index)
        {
            var preset = new PiperTrayApp.PresetSettings
            {
                Name = presetNameTextBoxes[index].Text,
                VoiceModel = presetVoiceModelComboBoxes[index].SelectedItem?.ToString(),
                Speaker = presetSpeakerComboBoxes[index].SelectedItem?.ToString() ?? "0",
                Speed = GetSpeedValue(presetSpeedComboBoxes[index].SelectedIndex).ToString(CultureInfo.InvariantCulture),
                SentenceSilence = presetSilenceNumericUpDowns[index].Value.ToString(CultureInfo.InvariantCulture),
                Enabled = "true"
            };

            PiperTrayApp.GetInstance().SavePreset(index, preset);
        }

        private void UpdatePresetSpeakers(int index)
        {
            if (presetVoiceModelComboBoxes[index].SelectedItem != null)
            {
                string selectedModel = presetVoiceModelComboBoxes[index].SelectedItem.ToString();
                string jsonPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    $"{selectedModel}.onnx.json"  // Updated extension
                );

                Log($"[UpdatePresetSpeakers] Loading speakers for model: {selectedModel}");
                Log($"[UpdatePresetSpeakers] JSON path: {jsonPath}");

                presetSpeakerComboBoxes[index].Items.Clear();

                if (File.Exists(jsonPath))
                {
                    string jsonContent = File.ReadAllText(jsonPath);
                    using JsonDocument doc = JsonDocument.Parse(jsonContent);
                    if (doc.RootElement.TryGetProperty("num_speakers", out JsonElement numSpeakers))
                    {
                        int speakerCount = numSpeakers.GetInt32();
                        Log($"[UpdatePresetSpeakers] Found {speakerCount} speakers for model {selectedModel}");

                        for (int i = 0; i < speakerCount; i++)
                        {
                            presetSpeakerComboBoxes[index].Items.Add(i.ToString());
                        }
                    }
                    else
                    {
                        Log($"[UpdatePresetSpeakers] No num_speakers found for {selectedModel}, defaulting to single speaker");
                        presetSpeakerComboBoxes[index].Items.Add("0");
                    }
                }
                else
                {
                    Log($"[UpdatePresetSpeakers] JSON file not found for {selectedModel}, defaulting to single speaker");
                    presetSpeakerComboBoxes[index].Items.Add("0");
                }

                if (presetSpeakerComboBoxes[index].Items.Count > 0)
                {
                    presetSpeakerComboBoxes[index].SelectedIndex = 0;
                }
            }
        }

        private void LoadSavedPresets()
        {
            for (int i = 0; i < 4; i++)
            {
                var preset = app.LoadPreset(i);
                if (preset != null)
                {
                    bool isEnabled = preset.EnabledBool;
                    presetEnableCheckBoxes[i].Checked = isEnabled;
                    UpdatePresetControlsState(i, isEnabled);
                }
            }
        }

        private void UpdatePresetControlsState(int presetIndex, bool enabled)
        {
            Log($"[UpdatePresetControlsState] Updating controls for preset {presetIndex + 1}");
            Log($"  Enabled state: {enabled}");

            // Get the panel containing all controls for this preset
            Panel presetPanel = presetPanels[presetIndex];

            // Update all controls within the preset panel
            foreach (Control control in presetPanel.Controls)
            {
                control.Enabled = enabled;

                // Apply BackColor and ForeColor for relevant control types
                if (control is TextBox || control is ComboBox || control is NumericUpDown)
                {
                    control.BackColor = enabled ? SystemColors.Window : SystemColors.Control;
                    control.ForeColor = enabled ? SystemColors.WindowText : SystemColors.GrayText;
                }

                // Force immediate visual refresh
                control.Refresh();
            }

            // Force panel refresh
            presetPanel.Refresh();

            Log($"[UpdatePresetControlsState] Controls updated and refreshed for preset {presetIndex + 1}");
        }

        private void UpdatePresetControlsEnabled(int index, bool enabled)
        {
            var controls = new Control[]
            {
                presetNameTextBoxes[index],
                presetVoiceModelComboBoxes[index],
                presetSpeakerComboBoxes[index],
                presetSpeedComboBoxes[index],
                presetSilenceNumericUpDowns[index]
            };

            foreach (var control in controls)
            {
                control.Enabled = enabled;
                control.BackColor = enabled ? SystemColors.Window : SystemColors.Control;
                control.ForeColor = enabled ? SystemColors.WindowText : SystemColors.GrayText;
            }
        }

        private void presetNameTextBoxes_TextChanged(object sender, EventArgs e)
        {
            if (!isInitializing && sender is TextBox textBox)
            {
                int index = Array.IndexOf(presetNameTextBoxes, textBox);
                if (index >= 0)
                {
                    UpdatePresetName(index);
                }
            }
        }

        private void presetVoiceModelComboBoxes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isInitializing && sender is ComboBox comboBox)
            {
                int index = Array.IndexOf(presetVoiceModelComboBoxes, comboBox);
                if (index >= 0)
                {
                    UpdatePresetSpeakers(index);
                }
            }
        }

        private void PopulateSpeakerComboBox(string modelName)
        {
            Log($"[PopulateSpeakerComboBox] Starting population with current speaker value: {speakerComboBox.SelectedItem}");
            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string jsonPath = Path.Combine(baseDir, $"{modelName}.onnx.json");

            Log($"[PopulateSpeakerComboBox] Base directory: {baseDir}");
            Log($"[PopulateSpeakerComboBox] Model name: {modelName}");
            Log($"[PopulateSpeakerComboBox] Looking for JSON at: {jsonPath}");
            Log($"[PopulateSpeakerComboBox] File exists: {File.Exists(jsonPath)}");

            if (File.Exists(jsonPath))
            {
                try
                {
                    string jsonContent = File.ReadAllText(jsonPath);
                    Log($"[PopulateSpeakerComboBox] JSON content loaded successfully");

                    using JsonDocument doc = JsonDocument.Parse(jsonContent);
                    var root = doc.RootElement;

                    Log($"[PopulateSpeakerComboBox] JSON document parsed successfully");
                    foreach (JsonProperty prop in root.EnumerateObject())
                    {
                        Log($"[PopulateSpeakerComboBox] Found root property: {prop.Name}");
                    }

                    if (root.TryGetProperty("num_speakers", out JsonElement numSpeakers))
                    {
                        int speakerCount = numSpeakers.GetInt32();
                        Log($"[PopulateSpeakerComboBox] Found {speakerCount} speakers");

                        speakerComboBox.Items.Clear();
                        speakerIdMap.Clear();

                        for (int i = 0; i < speakerCount; i++)
                        {
                            speakerComboBox.Items.Add(i.ToString());
                            speakerIdMap[i.ToString()] = i;
                        }

                        if (speakerComboBox.Items.Count > 0)
                        {
                            speakerComboBox.SelectedIndex = 0;
                            Log($"[PopulateSpeakerComboBox] Set default selection to 0");
                            Log($"[PopulateSpeakerComboBox] Populated {speakerComboBox.Items.Count} speakers");
                            Log($"[PopulateSpeakerComboBox] Current selection: {speakerComboBox.SelectedItem}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[PopulateSpeakerComboBox] Error processing JSON: {ex.Message}");
                    Log($"[PopulateSpeakerComboBox] Stack trace: {ex.StackTrace}");
                }
            }
        }

        private int GetSelectedSpeakerId()
        {
            if (speakerComboBox.SelectedItem != null)
            {
                string selectedName = speakerComboBox.SelectedItem.ToString();
                if (speakerIdMap.TryGetValue(selectedName, out int id))
                {
                    return id;
                }
            }
            return 0;
        }

        private string GetConfigPath()
        {
            return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "settings.conf");
        }

        private string ReadSettingValue(string key)
        {
            var settings = ReadCurrentSettings();
            if (settings.TryGetValue(key, out string value))
            {
                return value;
            }
            return null;
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

        private void InitializeSentenceSilence()
        {
            var settings = ReadCurrentSettings();
            if (settings.TryGetValue("SentenceSilence", out string silenceValue))
            {
                if (float.TryParse(silenceValue, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out float value))
                {
                    sentenceSilenceNumeric.Value = (decimal)value;
                    Log($"[InitializeSentenceSilence] Set value to: {value}");
                }
            }
        }

        private void SentenceSilenceNumeric_ValueChanged(object sender, EventArgs e)
        {
            float newSilence = (float)sentenceSilenceNumeric.Value;
            PiperTrayApp.GetInstance().SaveSettings(sentenceSilence: newSilence);
        }

        private void AddHotkeyControls(string labelText, int yPosition, string[] modifiers)
        {
            Log($"[AddHotkeyControls] Adding controls for: {labelText}");

            Label label = new Label();
            label.Text = labelText;
            label.Location = new System.Drawing.Point(10, yPosition);
            hotkeysTab.Controls.Add(label);

            ComboBox modifierComboBox = new ComboBox();
            modifierComboBox.Location = new System.Drawing.Point(130, yPosition);
            modifierComboBox.Width = 100;
            modifierComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            modifierComboBox.Items.AddRange(MODIFIER_OPTIONS);
            modifierComboBox.SelectedIndex = 0;
            hotkeysTab.Controls.Add(modifierComboBox);

            TextBox keyTextBox = new TextBox();
            keyTextBox.Location = new System.Drawing.Point(240, yPosition);
            keyTextBox.Width = 100;
            keyTextBox.ReadOnly = true;
            hotkeysTab.Controls.Add(keyTextBox);

            CheckBox enableCheckBox = new CheckBox();
            enableCheckBox.Location = new System.Drawing.Point(350, yPosition);
            enableCheckBox.Width = 20;
            enableCheckBox.Checked = true;
            enableCheckBox.CheckedChanged += (sender, e) =>
            {
                modifierComboBox.Enabled = enableCheckBox.Checked;
                keyTextBox.Enabled = enableCheckBox.Checked;
            };
            hotkeysTab.Controls.Add(enableCheckBox);
            hotkeyEnableCheckboxes[labelText] = enableCheckBox;

            keyTextBox.KeyDown += (sender, e) =>
            {
                if (sender is TextBox textBox)
                {
                    e.Handled = true;
                    if (e.KeyCode != Keys.None && e.KeyCode != Keys.Shift &&
                        e.KeyCode != Keys.Control && e.KeyCode != Keys.Alt)
                    {
                        textBox.Text = e.KeyCode.ToString().ToUpper();
                        Log($"Captured key for {labelText}: {textBox.Text}");

                        // Store the values immediately based on the control type
                        switch (labelText)
                        {
                            case "Switch Preset:":
                                switchPresetModifierComboBox = modifierComboBox;
                                switchPresetKeyTextBox = keyTextBox;
                                break;
                            case "Monitoring:":
                                monitoringVk = (uint)e.KeyCode;
                                monitoringModifiers = GetModifierVirtualKeyCode(monitoringModifierComboBox.SelectedItem?.ToString() ?? "NONE");
                                Log($"Stored Monitoring values - Modifier: 0x{monitoringModifiers:X2}, Key: 0x{monitoringVk:X2}");
                                break;
                            case "Stop Speech:":
                                stopSpeechVk = (uint)e.KeyCode;
                                stopSpeechModifiers = GetModifierVirtualKeyCode(stopSpeechModifierComboBox.SelectedItem?.ToString() ?? "NONE");
                                Log($"Stored Stop Speech values - Modifier: 0x{stopSpeechModifiers:X2}, Key: 0x{stopSpeechVk:X2}");
                                break;
                            case "Change Voice:":
                                changeVoiceVk = (uint)e.KeyCode;
                                changeVoiceModifiers = GetModifierVirtualKeyCode(changeVoiceModifierComboBox.SelectedItem?.ToString() ?? "NONE");
                                Log($"Stored Change Voice values - Modifier: 0x{changeVoiceModifiers:X2}, Key: 0x{changeVoiceVk:X2}");
                                break;
                            case "Speed Increase:":
                                speedIncreaseVk = (uint)e.KeyCode;
                                speedIncreaseModifiers = GetModifierVirtualKeyCode(speedIncreaseModifierComboBox.SelectedItem?.ToString() ?? "NONE");
                                Log($"Stored Speed Increase values - Modifier: 0x{speedIncreaseModifiers:X2}, Key: 0x{speedIncreaseVk:X2}");
                                break;
                            case "Speed Decrease:":
                                speedDecreaseVk = (uint)e.KeyCode;
                                speedDecreaseModifiers = GetModifierVirtualKeyCode(speedDecreaseModifierComboBox.SelectedItem?.ToString() ?? "NONE");
                                Log($"Stored Speed Decrease values - Modifier: 0x{speedDecreaseModifiers:X2}, Key: 0x{speedDecreaseVk:X2}");
                                break;
                        }
                    }
                }
            };

            modifierComboBox.SelectedIndexChanged += (sender, e) =>
            {
                if (sender is ComboBox combo)
                {
                    switch (labelText)
                    {
                        case "Monitoring:":
                            monitoringModifiers = GetModifierVirtualKeyCode(combo.SelectedItem?.ToString() ?? "NONE");
                            Log($"Updated Monitoring modifier: 0x{monitoringModifiers:X2}");
                            break;
                        case "Stop Speech:":
                            stopSpeechModifiers = GetModifierVirtualKeyCode(combo.SelectedItem?.ToString() ?? "NONE");
                            Log($"Updated Stop Speech modifier: 0x{stopSpeechModifiers:X2}");
                            break;
                        case "Change Voice:":
                            changeVoiceModifiers = GetModifierVirtualKeyCode(combo.SelectedItem?.ToString() ?? "NONE");
                            Log($"Updated Change Voice modifier: 0x{changeVoiceModifiers:X2}");
                            break;
                        case "Speed Increase:":
                            speedIncreaseModifiers = GetModifierVirtualKeyCode(combo.SelectedItem?.ToString() ?? "NONE");
                            Log($"Updated Speed Increase modifier: 0x{speedIncreaseModifiers:X2}");
                            break;
                        case "Speed Decrease:":
                            speedDecreaseModifiers = GetModifierVirtualKeyCode(combo.SelectedItem?.ToString() ?? "NONE");
                            Log($"Updated Speed Decrease modifier: 0x{speedDecreaseModifiers:X2}");
                            break;
                    }
                }
            };

            // Store references to controls based on label text
            switch (labelText)
            {
                case "Switch Preset:":
                    switchPresetModifierComboBox = modifierComboBox;
                    switchPresetKeyTextBox = keyTextBox;
                    break;
                case "Monitoring:":
                    monitoringModifierComboBox = modifierComboBox;
                    monitoringKeyTextBox = keyTextBox;
                    break;
                case "Stop Speech:":
                    stopSpeechModifierComboBox = modifierComboBox;
                    stopSpeechKeyTextBox = keyTextBox;
                    break;
                case "Change Voice:":
                    changeVoiceModifierComboBox = modifierComboBox;
                    changeVoiceKeyTextBox = keyTextBox;
                    break;
                case "Speed Increase:":
                    speedIncreaseModifierComboBox = modifierComboBox;
                    speedIncreaseKeyTextBox = keyTextBox;
                    break;
                case "Speed Decrease:":
                    speedDecreaseModifierComboBox = modifierComboBox;
                    speedDecreaseKeyTextBox = keyTextBox;
                    break;
            }

            Log($"[AddHotkeyControls] Controls added for {labelText}. ModifierComboBox items: {modifierComboBox.Items.Count}");
        }

        private void LoadHotkeyEnabledStates(Dictionary<string, string> settings)
        {
            foreach (var pair in hotkeyEnableCheckboxes)
            {
                string settingName = pair.Key.Replace(":", "").Replace(" ", "") + "HotkeyEnabled";
                if (settings.TryGetValue(settingName, out string enabledStr))
                {
                    if (bool.TryParse(enabledStr, out bool enabled))
                    {
                        pair.Value.Checked = enabled;
                    }
                }
            }
        }

        private void SaveHotkeyEnabledStates(List<string> lines)
        {
            foreach (var pair in hotkeyEnableCheckboxes)
            {
                string settingName = pair.Key.Replace(":", "").Replace(" ", "") + "HotkeyEnabled";
                UpdateOrAddSetting(lines, settingName, pair.Value.Checked.ToString());
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void UpdateHotkey(string hotkeyName, ComboBox modifierComboBox, TextBox keyTextBox)
        {
            if (modifierComboBox == null || keyTextBox == null)
            {
                Log($"UpdateHotkey called with null controls for {hotkeyName}");
                return;
            }

            Log($"UpdateHotkey called for: {hotkeyName}");
            string selectedModifier = modifierComboBox.SelectedItem as string;
            string enteredKey = keyTextBox.Text;

            // Skip registration if modifier is "NONE" or key is empty
            if (selectedModifier == "NONE" || string.IsNullOrEmpty(enteredKey))
            {
                Log($"Skipping registration for {hotkeyName}: No key combination assigned");
                return;
            }

            uint modifiers = GetModifierVirtualKeyCode(selectedModifier);
            uint vk = GetVirtualKeyCode(enteredKey);

            int hotkeyId = GetHotkeyId(hotkeyName);  // Add this line to get the hotkeyId

            Log($"Attempting to register hotkey: {hotkeyName} Modifiers=0x{modifiers:X2}, VK=0x{vk:X2}");

            var mainForm = Application.OpenForms.OfType<PiperTrayApp>().FirstOrDefault();
            if (mainForm != null)
            {
                if (!mainForm.IsHotkeyRegistered(modifiers, vk))
                {
                    var (success, errorCode) = mainForm.RegisterHotkey(hotkeyId, modifiers, vk, hotkeyName);
                    if (success)
                    {
                        Log($"{hotkeyName} hotkey registered successfully");
                    }
                    else
                    {
                        Log($"Failed to register {hotkeyName} hotkey. Error code: {errorCode}");
                        MessageBox.Show($"Failed to register hotkey for {hotkeyName}. Error code: {errorCode}", "Hotkey Registration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    Log($"{hotkeyName} hotkey already registered");
                }
            }
            else
            {
                Log($"Main form not found. Unable to register {hotkeyName} hotkey.");
            }
        }

        private void LoadSettingsIntoUI()
        {
            Log($"[LoadSettingsIntoUI] ===== Starting Settings Load Process =====");

            if (!IsHandleCreated)
            {
                CreateHandle();
                Log($"[LoadSettingsIntoUI] Handle created for the form");
            }

            isInitializing = true;
            string configPath = PiperTrayApp.GetInstance().GetConfigPath();
            Log($"[LoadSettingsIntoUI] Loading settings from: {configPath}");

            try
            {
                if (File.Exists(configPath))
                {
                    var settings = ReadSettingsFromFile(configPath);
                    InitializeUIWithSettings(settings);
                }
                else
                {
                    Log($"[LoadSettingsIntoUI] Settings file not found at: {configPath}. Using defaults.");
                    SetDefaultValues();
                }
            }
            catch (Exception ex)
            {
                Log($"[LoadSettingsIntoUI] Error loading settings: {ex.Message}");
                SetDefaultValues();
            }
            finally
            {
                _isInitialized = true;
                isInitializing = false;
                Log($"[LoadSettingsIntoUI] ===== Settings Load Complete =====");
            }
        }

        private Dictionary<string, string> ReadSettingsFromFile(string configPath)
        {
            var settings = File.ReadAllLines(configPath)
                .Where(line => !string.IsNullOrWhiteSpace(line) && line.Contains("="))
                .ToDictionary(
                    line => line.Split('=')[0].Trim(),
                    line => line.Split('=')[1].Trim(),
                    StringComparer.OrdinalIgnoreCase
                );
            Log($"[ReadSettingsFromFile] Loaded {settings.Count} settings");
            return settings;
        }

        private void InitializeUIWithSettings(Dictionary<string, string> settings)
        {
            foreach (var preset in settings.Where(s => s.Key.StartsWith("Preset", StringComparison.OrdinalIgnoreCase)))
            {
                int presetIndex = int.Parse(preset.Key.Replace("Preset", "")) - 1;
                var presetSettings = JsonSerializer.Deserialize<PiperTrayApp.PresetSettings>(preset.Value);

                if (presetSettings != null)
                {
                    ApplyPresetToUI(presetIndex, presetSettings);
                }
            }
        }

        private void ApplyPresetToUI(int index, PiperTrayApp.PresetSettings preset)
        {
            if (presetNameTextBoxes[index] != null)
            {
                presetNameTextBoxes[index].Text = preset.Name;
            }

            if (presetVoiceModelComboBoxes[index] != null)
            {
                int modelIndex = presetVoiceModelComboBoxes[index].Items.IndexOf(preset.VoiceModel);
                if (modelIndex >= 0)
                {
                    presetVoiceModelComboBoxes[index].SelectedIndex = modelIndex;
                }
            }

            if (presetSpeakerComboBoxes[index] != null && preset.Speaker != null)
            {
                presetSpeakerComboBoxes[index].SelectedItem = preset.Speaker;
            }

            if (presetSpeedComboBoxes[index] != null && double.TryParse(preset.Speed, out double speed))
            {
                int speedIndex = GetSpeedIndex(speed);
                presetSpeedComboBoxes[index].SelectedIndex = speedIndex + 9;
            }

            if (presetSilenceNumericUpDowns[index] != null && decimal.TryParse(preset.SentenceSilence, out decimal silence))
            {
                presetSilenceNumericUpDowns[index].Value = silence;
            }

            if (presetEnableCheckBoxes[index] != null)
            {
                presetEnableCheckBoxes[index].Checked = bool.Parse(preset.Enabled);
            }
        }

        private void LogSettingsContent(Dictionary<string, string> settings)
        {
            Log($"[LogSettingsContent] ===== Settings Content =====");
            foreach (var preset in settings.Where(s => s.Key.StartsWith("Preset")))
            {
                Log($"[LogSettingsContent] {preset.Key}: {preset.Value}");
            }
            Log($"[LogSettingsContent] Current VoiceModel: {settings.GetValueOrDefault("VoiceModel")}");
            Log($"[LogSettingsContent] Current Speed: {settings.GetValueOrDefault("Speed")}");
            Log($"[LogSettingsContent] Current Speaker: {settings.GetValueOrDefault("Speaker")}");
        }

        private void LoadSpeedSettings(Dictionary<string, string> settings)
        {
            if (settings.TryGetValue("Speed", out string speedValue))
            {
                if (double.TryParse(speedValue, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double speed))
                {
                    int speedIndex = GetSpeedIndex(speed);
                    speedIndex = Math.Max(0, Math.Min(presetSpeedComboBoxes[currentPresetIndex].Items.Count - 1, speedIndex + 9));
                    Log($"[LoadSpeedSettings] Setting speed for preset {currentPresetIndex} to: {speed} (index: {speedIndex})");

                    if (presetSpeedComboBoxes[currentPresetIndex].Items.Count > speedIndex)
                    {
                        presetSpeedComboBoxes[currentPresetIndex].SelectedIndex = speedIndex;
                        Log($"[LoadSpeedSettings] Speed index set successfully");
                    }
                }
            }
        }

        private void SetDefaultValues()
        {
            foreach (var checkbox in hotkeyEnableCheckboxes.Values)
            {
                if (checkbox != null && checkbox.IsHandleCreated)
                {
                    checkbox.Checked = true;
                }
            }
        }


        private void LoadHotkeySettings(Dictionary<string, string> settings)
        {
            try
            {
                LogInitialState(settings);

                LoadSwitchPresetHotkey(settings);
                LoadMonitoringHotkey(settings);
                LoadStopSpeechHotkey(settings);
                LoadChangeVoiceHotkey(settings);
                LoadSpeedHotkeys(settings);
            }
            catch (Exception ex)
            {
                Log($"Error in LoadHotkeySettings: {ex.Message}");
            }
        }

        private void LogInitialState(Dictionary<string, string> settings)
        {
            Log($"[LoadHotkeySettings] Starting to load hotkey settings");
            foreach (var setting in settings)
            {
                Log($"[LoadHotkeySettings] Found setting: {setting.Key}={setting.Value}");
            }

            Log($"[LoadHotkeySettings] Control states:");
            Log($"MonitoringModifierComboBox initialized: {monitoringModifierComboBox?.IsHandleCreated}, Items: {monitoringModifierComboBox?.Items.Count}");
            Log($"StopSpeechModifierComboBox initialized: {stopSpeechModifierComboBox?.IsHandleCreated}, Items: {stopSpeechModifierComboBox?.Items.Count}");
            Log($"ChangeVoiceModifierComboBox initialized: {changeVoiceModifierComboBox?.IsHandleCreated}, Items: {changeVoiceModifierComboBox?.Items.Count}");
            Log($"SpeedIncreaseModifierComboBox initialized: {speedIncreaseModifierComboBox?.IsHandleCreated}, Items: {speedIncreaseModifierComboBox?.Items.Count}");
            Log($"SpeedDecreaseModifierComboBox initialized: {speedDecreaseModifierComboBox?.IsHandleCreated}, Items: {speedDecreaseModifierComboBox?.Items.Count}");
        }

        private void LoadSwitchPresetHotkey(Dictionary<string, string> settings)
        {
            // Switch Preset Modifier
            if (settings.TryGetValue("SwitchPresetModifier", out string spMod) && !string.IsNullOrEmpty(spMod))
            {
                uint modValue = uint.Parse(spMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                Log($"[LoadSwitchPresetHotkey] Modifier value: 0x{modValue:X2}, ModString: {modString}");

                if (switchPresetModifierComboBox.Items.Contains(modString))
                {
                    switchPresetModifierComboBox.SelectedItem = modString;
                    Log($"[LoadSwitchPresetHotkey] Set ComboBox SelectedItem to: {modString}");
                }
                else
                {
                    Log($"[LoadSwitchPresetHotkey] ComboBox does not contain item: {modString}");
                }

                switchPresetModifiers = modValue; // Assign to variable
                Log($"[LoadSwitchPresetHotkey] Set switchPresetModifiers to: 0x{modValue:X2}");
            }
            else
            {
                Log($"[LoadSwitchPresetHotkey] SwitchPresetModifier not found in settings or is empty");
            }

            // Switch Preset Key
            if (settings.TryGetValue("SwitchPresetKey", out string spKey) && !string.IsNullOrEmpty(spKey))
            {
                uint keyValue = uint.Parse(spKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string keyString = ((Keys)keyValue).ToString();
                Log($"[LoadSwitchPresetHotkey] Key value: 0x{keyValue:X2}, KeyString: {keyString}");

                switchPresetKeyTextBox.Text = keyString;
                Log($"[LoadSwitchPresetHotkey] Set TextBox Text to: {keyString}");

                switchPresetVk = keyValue; // Assign to variable
                Log($"[LoadSwitchPresetHotkey] Set switchPresetVk to: 0x{keyValue:X2}");
            }
            else
            {
                Log($"[LoadSwitchPresetHotkey] SwitchPresetKey not found in settings or is empty");
            }
        }

        private void LoadMonitoringHotkey(Dictionary<string, string> settings)
        {
            // Monitoring Modifier
            if (settings.TryGetValue("MonitoringModifier", out string monMod) && !string.IsNullOrEmpty(monMod))
            {
                uint modValue = uint.Parse(monMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (monitoringModifierComboBox.Items.Contains(modString))
                {
                    monitoringModifierComboBox.SelectedItem = modString;
                }
                monitoringModifiers = modValue; // Assign to variable
                Log($"Set Monitoring modifier to: {modString} (0x{modValue:X2})");
            }

            // Monitoring Key
            if (settings.TryGetValue("MonitoringKey", out string monKey) && !string.IsNullOrEmpty(monKey))
            {
                uint keyValue = uint.Parse(monKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                monitoringKeyTextBox.Text = ((Keys)keyValue).ToString();
                monitoringVk = keyValue; // Assign to variable
                Log($"Set Monitoring key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void LoadStopSpeechHotkey(Dictionary<string, string> settings)
        {
            // Stop Speech Modifier
            if (settings.TryGetValue("StopSpeechModifier", out string ssMod) && !string.IsNullOrEmpty(ssMod))
            {
                uint modValue = uint.Parse(ssMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (stopSpeechModifierComboBox.Items.Contains(modString))
                {
                    stopSpeechModifierComboBox.SelectedItem = modString;
                }
                stopSpeechModifiers = modValue; // Assign to variable
                Log($"Set Stop Speech modifier to: {modString} (0x{modValue:X2})");
            }

            // Stop Speech Key
            if (settings.TryGetValue("StopSpeechKey", out string ssKey) && !string.IsNullOrEmpty(ssKey))
            {
                uint keyValue = uint.Parse(ssKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                stopSpeechKeyTextBox.Text = ((Keys)keyValue).ToString();
                stopSpeechVk = keyValue; // Assign to variable
                Log($"Set Stop Speech key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void LoadChangeVoiceHotkey(Dictionary<string, string> settings)
        {
            // Change Voice Modifier
            if (settings.TryGetValue("ChangeVoiceModifier", out string cvMod) && !string.IsNullOrEmpty(cvMod))
            {
                uint modValue = uint.Parse(cvMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (changeVoiceModifierComboBox.Items.Contains(modString))
                {
                    changeVoiceModifierComboBox.SelectedItem = modString;
                }
                changeVoiceModifiers = modValue; // Assign to variable
                Log($"Set Change Voice modifier to: {modString} (0x{modValue:X2})");
            }

            // Change Voice Key
            if (settings.TryGetValue("ChangeVoiceKey", out string cvKey) && !string.IsNullOrEmpty(cvKey))
            {
                uint keyValue = uint.Parse(cvKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                changeVoiceKeyTextBox.Text = ((Keys)keyValue).ToString();
                changeVoiceVk = keyValue; // Assign to variable
                Log($"Set Change Voice key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void LoadSpeedHotkeys(Dictionary<string, string> settings)
        {
            // Speed Increase Modifier
            if (settings.TryGetValue("SpeedIncreaseModifier", out string siMod) && !string.IsNullOrEmpty(siMod))
            {
                uint modValue = uint.Parse(siMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (speedIncreaseModifierComboBox.Items.Contains(modString))
                {
                    speedIncreaseModifierComboBox.SelectedItem = modString;
                }
                speedIncreaseModifiers = modValue; // Assign to variable
                Log($"Set Speed Increase modifier to: {modString} (0x{modValue:X2})");
            }

            // Speed Increase Key
            if (settings.TryGetValue("SpeedIncreaseKey", out string siKey) && !string.IsNullOrEmpty(siKey))
            {
                uint keyValue = uint.Parse(siKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                speedIncreaseKeyTextBox.Text = ((Keys)keyValue).ToString();
                speedIncreaseVk = keyValue; // Assign to variable
                Log($"Set Speed Increase key to: {((Keys)keyValue).ToString()}");
            }

            // Speed Decrease Modifier
            if (settings.TryGetValue("SpeedDecreaseModifier", out string sdMod) && !string.IsNullOrEmpty(sdMod))
            {
                uint modValue = uint.Parse(sdMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (speedDecreaseModifierComboBox.Items.Contains(modString))
                {
                    speedDecreaseModifierComboBox.SelectedItem = modString;
                }
                speedDecreaseModifiers = modValue; // Assign to variable
                Log($"Set Speed Decrease modifier to: {modString} (0x{modValue:X2})");
            }

            // Speed Decrease Key
            if (settings.TryGetValue("SpeedDecreaseKey", out string sdKey) && !string.IsNullOrEmpty(sdKey))
            {
                uint keyValue = uint.Parse(sdKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                speedDecreaseKeyTextBox.Text = ((Keys)keyValue).ToString();
                speedDecreaseVk = keyValue; // Assign to variable
                Log($"Set Speed Decrease key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void LoadSpeedIncreaseHotkey(Dictionary<string, string> settings)
        {
            if (settings.TryGetValue("SpeedIncreaseModifier", out string speedIncMod) && !string.IsNullOrEmpty(speedIncMod))
            {
                uint modValue = uint.Parse(speedIncMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (speedIncreaseModifierComboBox.Items.Contains(modString))
                {
                    speedIncreaseModifierComboBox.SelectedItem = modString;
                    Log($"Set Speed Increase modifier to: {modString}");
                }
            }

            if (settings.TryGetValue("SpeedIncreaseKey", out string speedIncKey) && !string.IsNullOrEmpty(speedIncKey))
            {
                uint keyValue = uint.Parse(speedIncKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                speedIncreaseKeyTextBox.Text = ((Keys)keyValue).ToString();
                Log($"Set Speed Increase key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void LoadSpeedDecreaseHotkey(Dictionary<string, string> settings)
        {
            if (settings.TryGetValue("SpeedDecreaseModifier", out string speedDecMod) && !string.IsNullOrEmpty(speedDecMod))
            {
                uint modValue = uint.Parse(speedDecMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                string modString = GetModifierString(modValue);
                if (speedDecreaseModifierComboBox.Items.Contains(modString))
                {
                    speedDecreaseModifierComboBox.SelectedItem = modString;
                    Log($"Set Speed Decrease modifier to: {modString}");
                }
            }

            if (settings.TryGetValue("SpeedDecreaseKey", out string speedDecKey) && !string.IsNullOrEmpty(speedDecKey))
            {
                uint keyValue = uint.Parse(speedDecKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                speedDecreaseKeyTextBox.Text = ((Keys)keyValue).ToString();
                Log($"Set Speed Decrease key to: {((Keys)keyValue).ToString()}");
            }
        }

        private void PopulateSpeedComboBox(double savedSpeed)
        {
            speedComboBox.Items.Clear();
            for (int i = -9; i <= 10; i++)
            {
                speedComboBox.Items.Add(i.ToString());
            }

            // Find the closest matching speed value
            int index = Array.IndexOf(speedOptions, savedSpeed);
            if (index != -1)
            {
                // Convert speedOptions index to UI index
                index = 10 - index;
                index = Math.Max(-9, Math.Min(10, index));
                speedComboBox.SelectedIndex = index + 9; // Adjust for -9 to 10 range
            }
            else
            {
                speedComboBox.SelectedIndex = 10; // Default to 1.0x speed
            }
        }

        private void SetModifierComboBox(ComboBox comboBox, string value)
        {
            if (comboBox == null || !comboBox.IsHandleCreated)
            {
                Log($"ComboBox not initialized when setting modifier value: {value}");
                return;
            }

            uint modifierCode;
            if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                if (!uint.TryParse(value.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out modifierCode))
                {
                    modifierCode = 0;
                    Log($"Invalid modifier value: {value}. Defaulting to NONE.");
                }
            }
            else if (!uint.TryParse(value, out modifierCode))
            {
                modifierCode = 0;
                Log($"Invalid modifier value: {value}. Defaulting to NONE.");
            }

            string modifierString = GetModifierString(modifierCode);

            if (comboBox.InvokeRequired)
            {
                comboBox.Invoke(new Action(() => comboBox.SelectedItem = modifierString));
            }
            else
            {
                comboBox.SelectedItem = modifierString;
            }

            Log($"Set hotkey modifier: {modifierString}");
        }

        private string GetModifierString(uint modifierValue)
        {
            switch (modifierValue)
            {
                case 0x0001: return "ALT";
                case 0x0002: return "CTRL";
                case 0x0004: return "SHIFT";
                default: return "ALT";  // Default to ALT as a safe fallback
            }
        }

        private void CheckAndRegisterHotkey(string hotkeyName, ComboBox modifierComboBox, TextBox keyTextBox)
        {
            string selectedModifier = modifierComboBox.SelectedItem as string;
            string enteredKey = keyTextBox.Text;

            if (selectedModifier == "NONE" || string.IsNullOrEmpty(enteredKey))
            {
                Log($"Skipping registration for {hotkeyName}: No key combination assigned");
                return;
            }

            uint modifiers = GetModifierVirtualKeyCode(selectedModifier);
            uint vk = GetVirtualKeyCode(enteredKey);

            Log($"Checking hotkey: {hotkeyName} Modifiers=0x{modifiers:X2}, VK=0x{vk:X2}");

            var mainForm = Application.OpenForms.OfType<PiperTrayApp>().FirstOrDefault();
            if (mainForm != null)
            {
                if (!mainForm.IsHotkeyRegistered(modifiers, vk))
                {
                    if (mainForm.TryRegisterHotkey(modifiers, vk))
                    {
                        Log($"Hotkey registered successfully: {hotkeyName}");
                    }
                    else
                    {
                        Log($"Failed to register hotkey: {hotkeyName}");
                        MessageBox.Show($"Failed to register hotkey for {hotkeyName}. The combination may be in use by another application.", "Hotkey Registration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    Log($"Hotkey already registered: {hotkeyName}");
                }
            }
            else
            {
                Log($"Main form not found. Unable to register hotkey.");
            }
        }

        private void SetKeyTextBox(TextBox textBox, string value)
        {
            if (textBox == null || !textBox.IsHandleCreated)
            {
                Log($"TextBox not initialized when setting key value: {value}");
                return;
            }

            if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                if (uint.TryParse(value.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out uint keyCode))
                {
                    string keyString = ((Keys)keyCode).ToString();

                    if (textBox.InvokeRequired)
                    {
                        textBox.Invoke(new Action(() => textBox.Text = keyString));
                    }
                    else
                    {
                        textBox.Text = keyString;
                    }

                    Log($"Set key: {keyString}");
                }
            }
        }

        private bool AreControlsInitialized()
        {
            return monitoringModifierComboBox != null &&
                   monitoringKeyTextBox != null &&
                   stopSpeechModifierComboBox != null &&
                   stopSpeechKeyTextBox != null &&
                   changeVoiceModifierComboBox != null &&
                   changeVoiceKeyTextBox != null &&
                   speedIncreaseModifierComboBox != null &&
                   speedIncreaseKeyTextBox != null &&
                   speedDecreaseModifierComboBox != null &&
                   speedDecreaseKeyTextBox != null;
        }

        private double GetSpeedValue(int index)
        {
            int speedIndex = 10 - index;
            if (speedIndex >= 0 && speedIndex < speedOptions.Length)
            {
                return speedOptions[speedIndex];
            }
            return 1.0; // Default value if out of range
        }

        private int GetSpeedIndex(double speed)
        {
            int index = Array.IndexOf(speedOptions, speed);
            if (index != -1)
            {
                // Convert speedOptions index (0 to 19) to UI index (-9 to 10)
                return 10 - index;
            }
            return 10; // Default to 1.0x speed
        }

        private void SaveSettings(
    uint monitoringModifiers,
    uint monitoringVk,
    uint stopSpeechModifiers,
    uint stopSpeechVk,
    uint changeVoiceModifiers,
    uint changeVoiceVk,
    uint speedIncreaseModifiers,
    uint speedIncreaseVk,
    uint speedDecreaseModifiers,
    uint speedDecreaseVk,
    uint switchPresetModifiers,
    uint switchPresetVk,
    float sentenceSilence)
        {
            var app = PiperTrayApp.GetInstance();
            string configPath = app.GetConfigPath();

            // Read all existing settings into a dictionary to preserve them
            var settingsDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(configPath))
            {
                var lines = File.ReadAllLines(configPath);
                settingsDict = lines
                    .Where(line => !string.IsNullOrWhiteSpace(line) && line.Contains('='))
                    .Select(line => line.Split(new[] { '=' }, 2))
                    .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);
            }

            // Update settings with the new values
            settingsDict["MonitoringModifier"] = $"0x{monitoringModifiers:X2}";
            settingsDict["MonitoringKey"] = $"0x{monitoringVk:X2}";
            settingsDict["StopSpeechModifier"] = $"0x{stopSpeechModifiers:X2}";
            settingsDict["StopSpeechKey"] = $"0x{stopSpeechVk:X2}";
            settingsDict["ChangeVoiceModifier"] = $"0x{changeVoiceModifiers:X2}";
            settingsDict["ChangeVoiceKey"] = $"0x{changeVoiceVk:X2}";
            settingsDict["SpeedIncreaseModifier"] = $"0x{speedIncreaseModifiers:X2}";
            settingsDict["SpeedIncreaseKey"] = $"0x{speedIncreaseVk:X2}";
            settingsDict["SpeedDecreaseModifier"] = $"0x{speedDecreaseModifiers:X2}";
            settingsDict["SpeedDecreaseKey"] = $"0x{speedDecreaseVk:X2}";
            settingsDict["SwitchPresetModifier"] = $"0x{switchPresetModifiers:X2}";
            settingsDict["SwitchPresetKey"] = $"0x{switchPresetVk:X2}";
            settingsDict["SentenceSilence"] = sentenceSilence.ToString("F1", System.Globalization.CultureInfo.InvariantCulture);

            // Reconstruct the lines from the updated settings dictionary
            var linesToWrite = settingsDict.Select(kvp => $"{kvp.Key}={kvp.Value}").ToList();

            // Write all settings back to the config file, preserving existing settings
            File.WriteAllLines(configPath, linesToWrite);
            Log($"[SaveSettings] Settings saved successfully");
        }

        private int GetHotkeyId(string hotkeyName)
        {
            switch (hotkeyName)
            {
                case "Monitoring": return PiperTrayApp.HOTKEY_ID_MONITORING;
                default: throw new ArgumentException($"Unknown hotkey name: {hotkeyName}");
            }
        }
        private static void Log(string message)
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

        private uint GetModifierVirtualKeyCode(string modifier)
        {
            Log($"[GetModifierVirtualKeyCode] Processing modifier: {modifier}");

            switch (modifier.ToUpper())
            {
                case "ALT":
                    return 0x01;
                case "CTRL":
                    return 0x02;
                case "SHIFT":
                    return 0x04;
                default:
                    return 0x01; // Default to ALT as a safe fallback
            }
        }

        private uint GetVirtualKeyCode(string key)
        {
            Log($"[GetVirtualKeyCode] Getting virtual key code for: {key}");

            if (string.IsNullOrWhiteSpace(key))
            {
                return 0;
            }

            // Handle direct key input
            if (Enum.TryParse<Keys>(key, true, out Keys result))
            {
                Log($"[GetVirtualKeyCode] Parsed key code: 0x{(uint)result:X2}");
                return (uint)result;
            }

            // Handle single character input
            if (key.Length == 1 && char.IsLetterOrDigit(key[0]))
            {
                uint vk = (uint)char.ToUpper(key[0]);
                Log($"[GetVirtualKeyCode] Single character key code: 0x{vk:X2}");
                return vk;
            }

            Log($"[GetVirtualKeyCode] Unable to parse key: {key}");
            return 0;
        }

        private void VoiceModelComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (VoiceModelComboBox.SelectedItem != null)
            {
                string selectedModel = VoiceModelComboBox.SelectedItem.ToString();
                PopulateSpeakerComboBox(selectedModel);
                SaveVoiceModelSetting(selectedModel);
                OnVoiceModelChanged();
            }
        }

        private void SaveVoiceModelSetting(string modelName)
        {
            string configPath = PiperTrayApp.GetInstance().GetConfigPath();
            try
            {
                var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
                UpdateOrAddSetting(lines, "VoiceModel", modelName);
                File.WriteAllLines(configPath, lines);
                Log($"[SaveVoiceModelSetting] VoiceModel set to {modelName} in {configPath}");
            }
            catch (Exception ex)
            {
                Log($"[SaveVoiceModelSetting] Error updating VoiceModel: {ex.Message}");
            }
        }

        private void UpdateOrAddSetting(List<string> lines, string key, string value)
        {
            Log($"[UpdateOrAddSetting] ===== Starting Update Operation =====");
            Log($"[UpdateOrAddSetting] Key: {key}");
            Log($"[UpdateOrAddSetting] Value: {value}");
            Log($"[UpdateOrAddSetting] Initial line count: {lines.Count}");

            // Dump current state of all relevant settings
            Log($"[UpdateOrAddSetting] Current settings state:");
            foreach (var line in lines)
            {
                if (line.Contains("SwitchPreset") || line.Contains("Speed") ||
                    line.Contains("Voice") || line.Contains("Monitoring"))
                {
                    Log($"[UpdateOrAddSetting] Current line: {line}");
                }
            }

            // Find and update/add setting with detailed tracking
            int index = lines.FindIndex(l => l.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase));
            Log($"[UpdateOrAddSetting] Search result - Found at index: {index}");

            string newLine = $"{key}={value}";
            if (index >= 0)
            {
                Log($"[UpdateOrAddSetting] Updating existing line:");
                Log($"[UpdateOrAddSetting] Old: {lines[index]}");
                lines[index] = newLine;
                Log($"[UpdateOrAddSetting] New: {lines[index]}");
            }
            else
            {
                Log($"[UpdateOrAddSetting] Adding new line: {newLine}");
                lines.Add(newLine);
                index = lines.Count - 1;
                Log($"[UpdateOrAddSetting] Added at index: {index}");
            }

            // Verify final state
            Log($"[UpdateOrAddSetting] ===== Final State =====");
            Log($"[UpdateOrAddSetting] Final line count: {lines.Count}");
            var finalLine = lines.FirstOrDefault(l => l.StartsWith(key + "="));
            Log($"[UpdateOrAddSetting] Final setting: {finalLine}");
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            Log($"[SaveButton_Click] ===== Starting Save Operation =====");
            var app = PiperTrayApp.GetInstance();
            string configPath = app.GetConfigPath();
            Log($"[SaveButton_Click] Config path: {configPath}");

            if (!VerifyPresetControls())
            {
                Log("[SaveButton_Click] Preset controls not properly initialized. Reinitializing...");
                InitializePresetArrays();
                for (int i = 0; i < 4; i++)
                {
                    CreatePresetPanel(i);
                }
                LoadSavedPresets();

                if (!VerifyPresetControls())
                {
                    Log("[SaveButton_Click] Control initialization failed. Aborting save.");
                    MessageBox.Show("Unable to save settings due to initialization error.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();
            Log($"[SaveButton_Click] Loaded {lines.Count} existing config lines");

            try
            {
                // Save presets
                SavePresets(lines);

                // Save hotkeys
                SaveHotkeys(lines);

                // Add this line to update the main settings from the active preset
                UpdateMainSettingsFromActivePreset(lines);

                // Write updated settings to file
                // Save current preset values with validation
                if (app.TryGetCurrentPreset(out var currentPreset))
                {
                    if (double.TryParse(currentPreset.Speed, NumberStyles.Float, CultureInfo.InvariantCulture, out double speed))
                    {
                        app.SaveSettings(speed: speed);
                        Log($"[SaveButton_Click] Saved speed: {speed}");
                    }
                    
                    // Add similar validation/blocks for other settings
                    app.SaveSettings(
                        speaker: int.Parse(currentPreset.Speaker),
                        sentenceSilence: float.Parse(currentPreset.SentenceSilence, CultureInfo.InvariantCulture),
                        voiceModel: currentPreset.VoiceModel);
                }
                
                File.WriteAllLines(configPath, lines);
                Log($"[SaveButton_Click] Settings persisted successfully");
                app.ReloadSettings(); // Ensure in-memory settings match config file

                // Register hotkeys and update application state
                app.RegisterHotkeys();
                Log($"[SaveButton_Click] Hotkeys re-registered");

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                Log($"[SaveButton_Click] Error during save operation:");
                Log($"  Exception Type: {ex.GetType().Name}");
                Log($"  Message: {ex.Message}");
                Log($"  Stack Trace: {ex.StackTrace}");
                MessageBox.Show($"An error occurred while saving settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Log($"[SaveButton_Click] ===== Save Operation Complete =====");
        }

        // File: Settings.cs
        private void UpdateMainSettingsFromActivePreset(List<string> lines)
        {
            Log("[UpdateMainSettingsFromActivePreset] Updating main settings from active preset");

            if (currentPresetIndex < 0 || currentPresetIndex >= presetNameTextBoxes.Length)
            {
                Log("[UpdateMainSettingsFromActivePreset] Invalid currentPresetIndex");
                return;
            }

            // Get the values from the active preset
            string voiceModel = presetVoiceModelComboBoxes[currentPresetIndex]?.SelectedItem?.ToString();
            string speaker = presetSpeakerComboBoxes[currentPresetIndex]?.SelectedItem?.ToString();

            // Handle nullable SelectedIndex for speed
            int selectedIndex = presetSpeedComboBoxes[currentPresetIndex]?.SelectedIndex ?? -1; // Default to -1 if null
            int adjustedIndex = selectedIndex - 9;

            // Get the speed value
            double speedValue = GetSpeedValue(adjustedIndex);
            string speed = speedValue.ToString(CultureInfo.InvariantCulture);

            // Get the sentence silence value
            string sentenceSilence = presetSilenceNumericUpDowns[currentPresetIndex]?.Value.ToString(CultureInfo.InvariantCulture);

            // Update or add the main settings
            UpdateOrAddSetting(lines, "VoiceModel", voiceModel);
            UpdateOrAddSetting(lines, "Speaker", speaker);
            UpdateOrAddSetting(lines, "Speed", speed);
            UpdateOrAddSetting(lines, "SentenceSilence", sentenceSilence);

            Log("[UpdateMainSettingsFromActivePreset] Main settings updated");
        }

        private void SavePresets(List<string> lines)
        {
            Log($"[SavePresets] ===== Starting Preset Save Operation =====");

            for (int i = 0; i < 4; i++)
            {
                Log($"[SavePresets] Processing preset {i + 1}");
                try
                {
                    // Validate controls
                    if (!ValidatePresetControls(i))
                    {
                        Log($"[SavePresets] Preset {i + 1} controls validation failed");
                        continue;
                    }

                    // Log control values
                    Log($"[SavePresets] Preset {i + 1} values:");
                    Log($"  Name: {presetNameTextBoxes[i].Text}");
                    Log($"  Voice Model: {presetVoiceModelComboBoxes[i].SelectedItem}");
                    Log($"  Speaker: {presetSpeakerComboBoxes[i].SelectedItem}");
                    Log($"  Speed Index: {presetSpeedComboBoxes[i].SelectedIndex}");
                    Log($"  Silence Value: {presetSilenceNumericUpDowns[i].Value}");
                    Log($"  Enabled: {presetEnableCheckBoxes[i].Checked}");

                    var preset = new PiperTrayApp.PresetSettings
                    {
                        Name = presetNameTextBoxes[i].Text ?? $"Preset {i + 1}",
                        VoiceModel = presetVoiceModelComboBoxes[i].SelectedItem.ToString(),
                        Speaker = presetSpeakerComboBoxes[i].SelectedItem.ToString(),
                        Speed = GetSpeedValue(presetSpeedComboBoxes[i].SelectedIndex - 9).ToString(CultureInfo.InvariantCulture),
                        SentenceSilence = presetSilenceNumericUpDowns[i].Value.ToString(CultureInfo.InvariantCulture),
                        Enabled = presetEnableCheckBoxes[i].Checked.ToString()
                    };

                    string presetJson = JsonSerializer.Serialize(preset);
                    Log($"[SavePresets] Serialized preset {i + 1}: {presetJson}");

                    UpdateOrAddSetting(lines, $"Preset{i + 1}", presetJson);
                    Log($"[SavePresets] Successfully saved preset {i + 1}");
                }
                catch (Exception ex)
                {
                    Log($"[SavePresets] Error saving preset {i + 1}:");
                    Log($"  Exception Type: {ex.GetType().Name}");
                    Log($"  Message: {ex.Message}");
                    Log($"  Stack Trace: {ex.StackTrace}");
                }
            }
            Log($"[SavePresets] ===== Preset Save Operation Complete =====");
        }

        private bool ArePresetControlsValid(int index)
        {
            var controls = new Control[]
            {
                presetNameTextBoxes[index],
                presetVoiceModelComboBoxes[index],
                presetSpeakerComboBoxes[index],
                presetSpeedComboBoxes[index],
                presetSilenceNumericUpDowns[index],
                presetEnableCheckBoxes[index]
            };

            bool valid = controls.All(c => c != null && c.Created);
            Log($"[ArePresetControlsValid] Preset {index + 1} controls valid: {valid}");

            if (!valid)
            {
                var nullControls = controls
                    .Select((c, i) => new { Control = c, Index = i })
                    .Where(x => x.Control == null || !x.Control.Created)
                    .Select(x => x.Index)
                    .ToList();

                Log($"[ArePresetControlsValid] Invalid controls at indices: {string.Join(", ", nullControls)}");
            }

            return valid;
        }

        private bool ValidatePresetControls(int index)
        {
            if (presetNameTextBoxes[index]?.IsHandleCreated != true ||
                presetVoiceModelComboBoxes[index]?.IsHandleCreated != true ||
                presetSpeakerComboBoxes[index]?.IsHandleCreated != true ||
                presetSpeedComboBoxes[index]?.IsHandleCreated != true ||
                presetSilenceNumericUpDowns[index]?.IsHandleCreated != true ||
                presetEnableCheckBoxes[index]?.IsHandleCreated != true)
            {
                Log($"[ValidatePresetControls] One or more controls not properly initialized for preset {index + 1}");
                return false;
            }

            if (presetVoiceModelComboBoxes[index].SelectedItem == null ||
                presetSpeakerComboBoxes[index].SelectedItem == null)
            {
                Log($"[ValidatePresetControls] Required selections missing for preset {index + 1}");
                return false;
            }

            return true;
        }

        private bool ValidateSpeakerValue(string modelName, string speakerValue)
        {
            if (string.IsNullOrEmpty(modelName) || string.IsNullOrEmpty(speakerValue))
            {
                return false;
            }

            string jsonPath = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                $"{modelName}.onnx.json"
            );

            try
            {
                if (File.Exists(jsonPath))
                {
                    string jsonContent = File.ReadAllText(jsonPath);
                    using JsonDocument doc = JsonDocument.Parse(jsonContent);
                    if (doc.RootElement.TryGetProperty("num_speakers", out JsonElement numSpeakers))
                    {
                        int maxSpeakers = numSpeakers.GetInt32();
                        if (int.TryParse(speakerValue, out int speakerIndex))
                        {
                            return speakerIndex >= 0 && speakerIndex < maxSpeakers;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[ValidateSpeakerValue] Error validating speaker value: {ex.Message}");
            }
            return false;
        }

        private void SaveHotkeys(List<string> lines)
        {
            Log($"[SaveHotkeys] ===== Starting Hotkey Processing =====");

            // Parse existing settings into a dictionary
            var existingSettings = lines
                .Where(line => !string.IsNullOrWhiteSpace(line) && line.Contains('='))
                .Select(line => line.Split(new[] { '=' }, 2))
                .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);

            // Get Switch Preset values with verification
            string hotkeyName = "Switch Preset";
            bool isEnabled = hotkeyEnableCheckboxes.TryGetValue($"{hotkeyName}:", out CheckBox enableCheckBox) && enableCheckBox.Checked;

            if (isEnabled)
            {
                string selectedModifier = switchPresetModifierComboBox?.SelectedItem?.ToString();
                string selectedKey = switchPresetKeyTextBox?.Text;
                Log($"[SaveHotkeys] {hotkeyName} UI values - Modifier: '{selectedModifier}', Key: '{selectedKey}'");

                // Check if key or modifier is empty
                if (string.IsNullOrEmpty(selectedKey) || string.IsNullOrEmpty(selectedModifier))
                {
                    // Try to get existing values from settings
                    if (existingSettings.TryGetValue("SwitchPresetModifier", out string existingModifier) &&
                        existingSettings.TryGetValue("SwitchPresetKey", out string existingKey))
                    {
                        // Use existing values
                        selectedModifier = GetModifierString(uint.Parse(existingModifier.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber));
                        selectedKey = ((Keys)uint.Parse(existingKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber)).ToString();
                        Log($"[SaveHotkeys] Using existing values - Modifier: '{selectedModifier}', Key: '{selectedKey}'");
                    }
                    else
                    {
                        // No existing key and modifier; prompt the user or set default
                        MessageBox.Show("No key and modifier assigned for the 'Switch Preset' hotkey. Please specify them.", "Hotkey Configuration", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // You can choose to return or set default values here
                        return; // Exit the method to prevent saving invalid settings
                    }
                }

                uint switchPresetModifiers = GetModifierVirtualKeyCode(selectedModifier ?? string.Empty);
                uint switchPresetVk = GetVirtualKeyCode(selectedKey ?? string.Empty);
                Log($"[SaveHotkeys] {hotkeyName} converted values - Modifier: 0x{switchPresetModifiers:X2}, VK: 0x{switchPresetVk:X2}");

                string modifierValue = $"0x{switchPresetModifiers:X2}";
                string keyValue = $"0x{switchPresetVk:X2}";
                Log($"[SaveHotkeys] Writing {hotkeyName} values - Modifier: {modifierValue}, Key: {keyValue}");

                UpdateOrAddSetting(lines, "SwitchPresetModifier", modifierValue);
                UpdateOrAddSetting(lines, "SwitchPresetKey", keyValue);
            }
            else
            {
                Log($"[SaveHotkeys] {hotkeyName} is disabled; skipping key and modifier update");
            }

            SaveHotkeyEnabledStates(lines);

            // Process other hotkeys
            ProcessOtherHotkeys(lines);

            Log($"[SaveHotkeys] ===== Final State =====");
            foreach (var line in lines.Where(l => l.StartsWith("SwitchPreset", StringComparison.OrdinalIgnoreCase)))
            {
                Log($"[SaveHotkeys] Final setting: {line}");
            }
        }

        private void ProcessOtherHotkeys(List<string> lines)
        {
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
        }

        private void ProcessHotkeyUpdate(List<string> lines, string prefix, uint modifiers, uint vk)
        {
            string modifierValue = $"0x{modifiers:X2}";
            string keyValue = $"0x{vk:X2}";

            Log($"[SaveHotkeys] Processing {prefix} - Modifier: {modifierValue}, Key: {keyValue}");

            UpdateOrAddSetting(lines, $"{prefix}Modifier", modifierValue);
            UpdateOrAddSetting(lines, $"{prefix}Key", keyValue);
        }

        private void SaveGlobalSettings(List<string> lines, PiperTrayApp app)
        {
            Log($"[SaveGlobalSettings] Updating global settings");
            var settings = app.ReadCurrentSettings();
            if (settings.TryGetValue("Speed", out string speedValue) &&
                double.TryParse(speedValue, out double currentSpeed))
            {
                UpdateOrAddSetting(lines, "Speed", currentSpeed.ToString("F1", System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        private void RegisterAndUpdateState(PiperTrayApp app)
        {
            Log($"[RegisterAndUpdateState] Registering hotkeys and updating application state");
            RegisterHotkeys();

            var settings = app.ReadCurrentSettings();
            if (settings.TryGetValue("Speed", out string speedValue) &&
                double.TryParse(speedValue, out double currentSpeed))
            {
                app.UpdateSpeedFromSettings(currentSpeed);
            }
        }

        private void SaveSwitchPresetHotkey(List<string> lines)
        {
            string selectedModifier = switchPresetModifierComboBox?.SelectedItem?.ToString();
            string selectedKey = switchPresetKeyTextBox?.Text;
            Log($"[SaveSwitchPresetHotkey] Raw UI values - Modifier: '{selectedModifier}', Key: '{selectedKey}'");

            uint switchPresetModifiers = GetModifierVirtualKeyCode(selectedModifier ?? string.Empty);
            uint switchPresetVk = GetVirtualKeyCode(selectedKey ?? string.Empty);
            Log($"[SaveSwitchPresetHotkey] Converted values - Modifier: 0x{switchPresetModifiers:X2}, VK: 0x{switchPresetVk:X2}");

            // Force update of settings.conf
            string configPath = PiperTrayApp.GetInstance().GetConfigPath();
            var currentLines = File.ReadAllLines(configPath).ToList();

            // Update or add Switch Preset settings
            int modifierIndex = currentLines.FindIndex(l => l.StartsWith("SwitchPresetModifier="));
            int keyIndex = currentLines.FindIndex(l => l.StartsWith("SwitchPresetKey="));

            string modifierLine = $"SwitchPresetModifier=0x{switchPresetModifiers:X2}";
            string keyLine = $"SwitchPresetKey=0x{switchPresetVk:X2}";

            if (modifierIndex >= 0)
                currentLines[modifierIndex] = modifierLine;
            else
                currentLines.Add(modifierLine);

            if (keyIndex >= 0)
                currentLines[keyIndex] = keyLine;
            else
                currentLines.Add(keyLine);

            // Write updated settings
            File.WriteAllLines(configPath, currentLines);

            // Update application state
            var app = PiperTrayApp.GetInstance();
            app.UpdateHotkey(PiperTrayApp.HOTKEY_ID_SWITCH_PRESET, switchPresetModifiers, switchPresetVk);
        }

        private void SaveStandardHotkeys(List<string> lines)
        {
            // Process standard hotkeys (Monitoring, Stop Speech, etc.)
            ProcessHotkeyPair(lines, "Monitoring", monitoringModifierComboBox, monitoringKeyTextBox);
            ProcessHotkeyPair(lines, "StopSpeech", stopSpeechModifierComboBox, stopSpeechKeyTextBox);
            ProcessHotkeyPair(lines, "ChangeVoice", changeVoiceModifierComboBox, changeVoiceKeyTextBox);
            ProcessHotkeyPair(lines, "SpeedIncrease", speedIncreaseModifierComboBox, speedIncreaseKeyTextBox);
            ProcessHotkeyPair(lines, "SpeedDecrease", speedDecreaseModifierComboBox, speedDecreaseKeyTextBox);
        }

        private void ProcessHotkeyPair(List<string> lines, string prefix, ComboBox modifierBox, TextBox keyBox)
        {
            uint modifiers = GetModifierVirtualKeyCode(modifierBox?.SelectedItem?.ToString() ?? string.Empty);
            uint vk = GetVirtualKeyCode(keyBox?.Text ?? string.Empty);

            UpdateOrAddSetting(lines, $"{prefix}Modifier", $"0x{modifiers:X2}");
            UpdateOrAddSetting(lines, $"{prefix}Key", $"0x{vk:X2}");
        }

        private void UpdateGlobalPresetSettings(List<string> lines, PiperTrayApp.PresetSettings preset, double speed)
        {
            // Convert numeric values to strings using invariant culture
            UpdateOrAddSetting(lines, "Speed", speed.ToString(CultureInfo.InvariantCulture));
            UpdateOrAddSetting(lines, "Speaker", preset.Speaker.ToString());
            UpdateOrAddSetting(lines, "VoiceModel", preset.VoiceModel);
            UpdateOrAddSetting(lines, "SentenceSilence", preset.SentenceSilence.ToString(CultureInfo.InvariantCulture));
        }

        private void RefreshVoiceModels_Click(object sender, EventArgs e)
        {
            if (isScanning || (DateTime.Now - lastScanTime).TotalSeconds < 5)
            {
                return;
            }

            isScanning = true;
            try
            {
                PiperTrayApp.GetInstance().ScanForVoiceModels();
                var newModels = PiperTrayApp.GetInstance().GetVoiceModels();
                UpdateVoiceModels(newModels);
            }
            catch (Exception ex)
            {
                Log($"Error refreshing voice models: {ex.Message}");
            }
            finally
            {
                isScanning = false;
                lastScanTime = DateTime.Now;
            }
        }

        public void UpdateVoiceModels(List<string> newModels)
        {
            Log($"[UpdateVoiceModels] Updating voice models with {newModels.Count} models");
            voiceModels = newModels;
            UpdateVoiceModelComboBox();
        }

        public void UpdateSentenceSilence(float value)
        {
            if (sentenceSilenceNumeric != null)
            {
                sentenceSilenceNumeric.Value = (decimal)value;
                Log($"[UpdateSentenceSilence] Updated sentence silence value to: {value}");
            }
        }

        private void UpdateVoiceModelComboBox()
        {
            Log($"[UpdateVoiceModelComboBox] Updating combobox with {voiceModels.Count} models");
            VoiceModelComboBox.Items.Clear();
            foreach (var model in voiceModels)
            {
                VoiceModelComboBox.Items.Add(model);
            }
            VoiceModelComboBox.SelectedItem = currentVoiceModel;
            Log($"[UpdateVoiceModelComboBox] Combobox updated, selected model: {VoiceModelComboBox.SelectedItem}");
        }
    }
}
