using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace PiperTray
{
    public class SettingsForm : Form
    {
        private bool _isInitialized = false;
        private bool isInitializing = true;
        private static SettingsForm instance = null;
        private static readonly object _lock = new object();
        public TabControl TabControl => tabControl;
        public TabPage GeneralTab => generalTab;
        private TabPage appearanceTab;
        private Dictionary<string, CheckBox> menuVisibilityCheckboxes;
        private readonly double[] speedOptions = { 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0 };
        private string currentVoiceModel;

        private ComboBox speakerComboBox;
        private int currentSpeaker = 0;
        private Dictionary<string, int> speakerIdMap = new Dictionary<string, int>();

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
        private TabControl tabControl;
        private TabPage generalTab;
        private TabPage hotkeysTab;
        private TabPage presetsTab;
        private ComboBox[] presetVoiceModelComboBoxes;
        private ComboBox[] presetSpeakerComboBoxes;
        private ComboBox[] presetSpeedComboBoxes;
        private TextBox[] presetNameTextBoxes;

        public class PresetSettings
        {
            public string Name { get; set; }
            public string VoiceModel { get; set; }
            public int Speaker { get; set; }
            public double Speed { get; set; }
            public float SentenceSilence { get; set; }
        }

        private NumericUpDown[] presetSilenceNumericUpDowns;
        private Button[] presetApplyButtons;
        private CheckBox loggingCheckBox;
        private ComboBox VoiceModelComboBox;
        private List<string> voiceModels;
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

        private Button saveButton;
        private Button cancelButton;
        private Button refreshVoiceModelsButton;

        private bool isScanning = false;
        private DateTime lastScanTime = DateTime.MinValue;


        public static SettingsForm GetInstance()
        {
            lock (_lock)
            {
                if (instance == null || instance.IsDisposed)
                {
                    instance = new SettingsForm();
                    instance.CreateHandle();
                    instance.HandleCreated += (s, e) =>
                    {
                        instance.LoadSettingsIntoUI();
                    };
                }
                return instance;
            }
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(value);
        }

        private string logFilePath;
        public bool IsLoggingEnabled { get { return loggingCheckBox?.Checked ?? false; } }

        [DllImport("kernel32.dll")]
        static extern uint GetLastError();


        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private SettingsForm()
        {
            // First create all form controls
            InitializeComponent();

            // Initialize UI controls after form components are ready
            InitializeUIControls();

            // Load settings last
            LoadSettingsIntoUI();

            // Set up base configuration
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
            logFilePath = Path.Combine(assemblyDirectory, "system.log");
            Log($"SettingsForm constructor called");

            // Initialize window properties
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Icon = PiperTrayApp.GetApplicationIcon();

            // Load voice models
            voiceModels = Directory.GetFiles(assemblyDirectory, "*.onnx")
                .Select(Path.GetFileName)
                .ToList();

        }

        public void ShowSettingsForm()
        {
            // Force window state to normal before showing
            this.WindowState = FormWindowState.Normal;

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

        }

        private void InitializeComboBoxes()
        {

        }

        private void PopulateVoiceModelComboBox(ComboBox comboBox)
        {
            comboBox.Items.Clear();
            string[] voiceModels = Directory.GetFiles(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "*.onnx"
            ).Select(Path.GetFileNameWithoutExtension).ToArray();

            comboBox.Items.AddRange(voiceModels);
            if (comboBox.Items.Count > 0)
                comboBox.SelectedIndex = 0;
        }

        private void ApplyPreset(int index)
        {
            var preset = new PiperTrayApp.PresetSettings
            {
                Name = presetNameTextBoxes[index].Text,
                VoiceModel = presetVoiceModelComboBoxes[index].SelectedItem?.ToString(),
                Speaker = int.Parse(presetSpeakerComboBoxes[index].SelectedItem?.ToString() ?? "0"),
                Speed = GetSpeedValue(presetSpeedComboBoxes[index].SelectedIndex),
                SentenceSilence = (float)presetSilenceNumericUpDowns[index].Value
            };

            PiperTrayApp.GetInstance().SavePreset(index, preset);
            PiperTrayApp.GetInstance().ApplyPreset(index);
        }

        private void speakerComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isInitializing && speakerComboBox.SelectedItem != null)
            {
                int selectedSpeaker = int.Parse(speakerComboBox.SelectedItem.ToString());
                var app = PiperTrayApp.GetInstance();
                app.UpdateCurrentSpeaker(selectedSpeaker);
                app.SaveSettings(speaker: selectedSpeaker);
                Log($"[speakerComboBox_SelectedIndexChanged] Updated speaker to: {selectedSpeaker}");
            }
        }

        private void RegisterHotkeys()
        {
            Log($"[RegisterHotkeys] Starting hotkey registration process");
            var mainForm = PiperTrayApp.GetInstance();
            mainForm.UnregisterAllHotkeys();

            // Only register hotkeys if controls are properly initialized
            if (AreControlsInitialized())
            {
                // Register Monitoring hotkey
                uint monitoringModifiers = GetModifierVirtualKeyCode(monitoringModifierComboBox.SelectedItem as string);
                uint monitoringVk = GetVirtualKeyCode(monitoringKeyTextBox.Text);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_MONITORING, monitoringModifiers, monitoringVk, "Monitoring");

                // Register Stop Speech hotkey
                uint stopSpeechModifiers = GetModifierVirtualKeyCode(stopSpeechModifierComboBox.SelectedItem as string);
                uint stopSpeechVk = GetVirtualKeyCode(stopSpeechKeyTextBox.Text);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_STOP_SPEECH, stopSpeechModifiers, stopSpeechVk, "Stop Speech");

                // Register Change Voice hotkey
                uint changeVoiceModifiers = GetModifierVirtualKeyCode(changeVoiceModifierComboBox.SelectedItem as string);
                uint changeVoiceVk = GetVirtualKeyCode(changeVoiceKeyTextBox.Text);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_CHANGE_VOICE, changeVoiceModifiers, changeVoiceVk, "Change Voice");

                // Register Speed hotkeys
                uint speedIncreaseModifiers = GetModifierVirtualKeyCode(speedIncreaseModifierComboBox.SelectedItem as string);
                uint speedIncreaseVk = GetVirtualKeyCode(speedIncreaseKeyTextBox.Text);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_SPEED_INCREASE, speedIncreaseModifiers, speedIncreaseVk, "Speed Increase");

                uint speedDecreaseModifiers = GetModifierVirtualKeyCode(speedDecreaseModifierComboBox.SelectedItem as string);
                uint speedDecreaseVk = GetVirtualKeyCode(speedDecreaseKeyTextBox.Text);
                mainForm.RegisterHotkey(PiperTrayApp.HOTKEY_ID_SPEED_DECREASE, speedDecreaseModifiers, speedDecreaseVk, "Speed Decrease");

            }
            else
            {
                Log($"[RegisterHotkeys] Controls not fully initialized, skipping hotkey registration");
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

                switch (id)
                {
                    case PiperTrayApp.HOTKEY_ID_MONITORING:
                        Log($"[WndProc] Monitoring hotkey pressed");
                        PiperTrayApp.GetInstance().ToggleMonitoring(null, EventArgs.Empty);
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
            this.Text = "Settings";
            this.Size = new System.Drawing.Size(400, 300);

            tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            generalTab = new TabPage("General");
            appearanceTab = new TabPage("Appearance");
            hotkeysTab = new TabPage("Hotkeys");
            presetsTab = new TabPage("Presets");

            tabControl.TabPages.Add(generalTab);
            tabControl.TabPages.Add(appearanceTab);
            tabControl.TabPages.Add(hotkeysTab);
            tabControl.TabPages.Add(presetsTab);

            CreateMenuVisibilityControls();

            presetVoiceModelComboBoxes = new ComboBox[4];
            presetSpeakerComboBoxes = new ComboBox[4];
            presetSpeedComboBoxes = new ComboBox[4];
            presetSilenceNumericUpDowns = new NumericUpDown[4];
            presetApplyButtons = new Button[4];
            presetNameTextBoxes = new TextBox[4];

            // Create preset panels
            for (int i = 0; i < 4; i++)
            {
                CreatePresetPanel(i);
            }

            // Hotkeys Tab
            string[] modifiers = new[] { "NONE", "ALT", "CTRL", "SHIFT" };
            AddHotkeyControls("Monitoring:", 40, modifiers);
            AddHotkeyControls("Stop Speech:", 70, modifiers);
            AddHotkeyControls("Change Voice:", 100, modifiers);
            AddHotkeyControls("Speed Increase:", 130, modifiers);
            AddHotkeyControls("Speed Decrease:", 160, modifiers);

            // General Tab
            loggingCheckBox = new CheckBox();
            loggingCheckBox.Text = "Logging";
            loggingCheckBox.Location = new System.Drawing.Point(10, 10);
            generalTab.Controls.Add(loggingCheckBox);

            Label voiceModelLabel = new Label();
            voiceModelLabel.Text = "Voice Model:";
            voiceModelLabel.Location = new System.Drawing.Point(10, 40);
            generalTab.Controls.Add(voiceModelLabel);

            this.VoiceModelComboBox = new System.Windows.Forms.ComboBox();
            this.VoiceModelComboBox.Name = "VoiceModelComboBox";
            this.VoiceModelComboBox.Location = new System.Drawing.Point(10, 63);
            this.VoiceModelComboBox.Width = 180;
            this.VoiceModelComboBox.SelectedIndexChanged += VoiceModelComboBox_SelectedIndexChanged;
            generalTab.Controls.Add(this.VoiceModelComboBox);

            speakerComboBox = new ComboBox();
            speakerComboBox.Location = new Point(VoiceModelComboBox.Right + 5, VoiceModelComboBox.Top);
            speakerComboBox.Width = 50;
            generalTab.Controls.Add(speakerComboBox);
            speakerComboBox.SelectedIndexChanged += speakerComboBox_SelectedIndexChanged;

            this.refreshVoiceModelsButton = new Button();
            this.refreshVoiceModelsButton.Image = GetRefreshIcon();
            this.refreshVoiceModelsButton.Size = new Size(24, 24);
            this.refreshVoiceModelsButton.Location = new Point(speakerComboBox.Right + 5, VoiceModelComboBox.Top);
            this.refreshVoiceModelsButton.Cursor = Cursors.Hand;
            this.refreshVoiceModelsButton.FlatStyle = FlatStyle.Flat;
            this.refreshVoiceModelsButton.FlatAppearance.BorderSize = 0;
            this.refreshVoiceModelsButton.Click += new EventHandler(this.RefreshVoiceModels_Click);
            generalTab.Controls.Add(this.refreshVoiceModelsButton);

            // Speed controls
            Label speedLabel = new Label();
            speedLabel.Text = "Speed:";
            speedLabel.Location = new System.Drawing.Point(10, 95);
            generalTab.Controls.Add(speedLabel);

            speedComboBox = new ComboBox();
            speedComboBox.Location = new System.Drawing.Point(10, 118);
            speedComboBox.Width = 50;
            generalTab.Controls.Add(speedComboBox);

            for (int i = 1; i <= 10; i++)
            {
                speedComboBox.Items.Add(i.ToString());
            }
            speedComboBox.SelectedIndex = 9; // Default to 1.0x speed

            // Sentence Silence controls
            Label sentenceSilenceLabel = new Label();
            sentenceSilenceLabel.Text = "Sentence Silence:";
            sentenceSilenceLabel.Location = new System.Drawing.Point(10, 150);
            generalTab.Controls.Add(sentenceSilenceLabel);

            sentenceSilenceNumeric = new NumericUpDown();
            sentenceSilenceNumeric.Location = new System.Drawing.Point(10, 173);
            sentenceSilenceNumeric.Width = 60;
            sentenceSilenceNumeric.DecimalPlaces = 1;
            sentenceSilenceNumeric.Increment = 0.1m;
            sentenceSilenceNumeric.Minimum = 0.0m;
            sentenceSilenceNumeric.Maximum = 2.0m;
            sentenceSilenceNumeric.Value = 0.5m;
            sentenceSilenceNumeric.ValueChanged += SentenceSilenceNumeric_ValueChanged;
            generalTab.Controls.Add(sentenceSilenceNumeric);

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

            LoadVoiceModels();
        }

        private void CreateMenuVisibilityControls()
        {
            menuVisibilityCheckboxes = new Dictionary<string, CheckBox>();
            string[] menuItems = new[] {
        "Monitoring",
        "Speed",
        "Voice",
        "Presets",
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
            int topMargin = 10;
            int controlSpacing = 5;
            int labelHeight = 20;
            int controlHeight = 25;
            int columnWidth = 100;
            int rowSpacing = 35;

            // Labels row (only show for first preset)
            if (index == 0)
            {
                Label nameLabel = new Label
                {
                    Text = "Name",
                    Location = new Point(10, topMargin),
                    Width = columnWidth
                };

                Label voiceLabel = new Label
                {
                    Text = "Model",
                    Location = new Point(nameLabel.Right + controlSpacing, topMargin),
                    Width = columnWidth
                };

                Label speedLabel = new Label
                {
                    Text = "Speed",
                    Location = new Point(voiceLabel.Right - (columnWidth / 2) + 105, topMargin),
                    Width = columnWidth / 2
                };

                Label silenceLabel = new Label
                {
                    Text = "Silence",
                    Location = new Point(speedLabel.Right + (controlSpacing / 4) + (columnWidth / 10) - 8, topMargin),
                    Width = columnWidth
                };

                presetsTab.Controls.AddRange(new Control[] { nameLabel, voiceLabel, speedLabel, silenceLabel });
            }

            int rowY = topMargin + labelHeight + controlSpacing + (index * rowSpacing);

            presetNameTextBoxes[index] = new TextBox
            {
                Location = new Point(10, rowY),
                Width = (columnWidth - 4),
                Text = $"Preset {index + 1}"
            };
            presetNameTextBoxes[index].TextChanged += (s, e) => UpdatePresetName(index);

            presetVoiceModelComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetNameTextBoxes[index].Right + controlSpacing, rowY),
                Width = columnWidth
            };
            PopulateVoiceModelComboBox(presetVoiceModelComboBoxes[index]);
            presetVoiceModelComboBoxes[index].SelectedIndexChanged += (s, e) => UpdatePresetSpeakers(index);

            presetSpeakerComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetVoiceModelComboBoxes[index].Right + controlSpacing, rowY),
                Width = columnWidth / 2
            };

            presetSpeedComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetSpeakerComboBoxes[index].Right + controlSpacing, rowY),
                Width = columnWidth / 2
            };
            for (int i = 1; i <= 10; i++)
            {
                presetSpeedComboBoxes[index].Items.Add(i.ToString());
            }

            presetSilenceNumericUpDowns[index] = new NumericUpDown
            {
                Location = new Point(presetSpeedComboBoxes[index].Right + controlSpacing, rowY),
                Width = columnWidth / 2,
                DecimalPlaces = 1,
                Increment = 0.1m,
                Minimum = 0.0m,
                Maximum = 2.0m,
                Value = 0.5m
            };

            presetsTab.Controls.AddRange(new Control[]
            {
        presetNameTextBoxes[index], presetVoiceModelComboBoxes[index],
        presetSpeakerComboBoxes[index], presetSpeedComboBoxes[index],
        presetSilenceNumericUpDowns[index]
            });
        }

        private void UpdatePresetName(int index)
        {
            var preset = new PiperTrayApp.PresetSettings
            {
                Name = presetNameTextBoxes[index].Text,
                VoiceModel = presetVoiceModelComboBoxes[index].SelectedItem?.ToString(),
                Speaker = int.Parse(presetSpeakerComboBoxes[index].SelectedItem?.ToString() ?? "0"),
                Speed = GetSpeedValue(presetSpeedComboBoxes[index].SelectedIndex),
                SentenceSilence = (float)presetSilenceNumericUpDowns[index].Value
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
                    $"{selectedModel}.onnx.json"
                );

                if (File.Exists(jsonPath))
                {
                    string jsonContent = File.ReadAllText(jsonPath);
                    using JsonDocument doc = JsonDocument.Parse(jsonContent);
                    if (doc.RootElement.TryGetProperty("num_speakers", out JsonElement numSpeakers))
                    {
                        presetSpeakerComboBoxes[index].Items.Clear();
                        for (int i = 0; i < numSpeakers.GetInt32(); i++)
                        {
                            presetSpeakerComboBoxes[index].Items.Add(i.ToString());
                        }
                        if (presetSpeakerComboBoxes[index].Items.Count > 0)
                        {
                            presetSpeakerComboBoxes[index].SelectedIndex = 0;
                        }
                    }
                }
            }
        }

        private void LoadSavedPresets()
        {
            for (int i = 0; i < 4; i++)
            {
                var preset = PiperTrayApp.GetInstance().LoadPreset(i);
                if (preset != null)
                {
                    presetNameTextBoxes[i].Text = preset.Name;

                    if (preset.VoiceModel != null && presetVoiceModelComboBoxes[i].Items.Contains(preset.VoiceModel))
                    {
                        presetVoiceModelComboBoxes[i].SelectedItem = preset.VoiceModel;
                        UpdatePresetSpeakers(i);
                        presetSpeakerComboBoxes[i].SelectedItem = preset.Speaker.ToString();
                    }

                    int speedIndex = 10 - (int)(preset.Speed * 10);
                    presetSpeedComboBoxes[i].SelectedIndex = speedIndex;
                    presetSilenceNumericUpDowns[i].Value = (decimal)preset.SentenceSilence;
                }
            }
        }

        private Image GetRefreshIcon()
        {
            Log($"[GetRefreshIcon] Entering method");
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith("refresh-32.png"));

            if (resourceName != null)
            {
                try
                {
                    using (var stream = assembly.GetManifestResourceStream(resourceName))
                    {
                        if (stream != null)
                        {
                            var image = Image.FromStream(stream);
                            Log($"[GetRefreshIcon] Image successfully loaded, size: {image.Width}x{image.Height}");
                            return image;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[GetRefreshIcon] Error loading image from embedded resource: {ex.Message}");
                }
            }

            Log($"[GetRefreshIcon] Embedded resource not found or failed to load");
            return null;
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
            modifierComboBox.Items.AddRange(modifiers);
            modifierComboBox.SelectedIndex = 0;
            hotkeysTab.Controls.Add(modifierComboBox);

            TextBox keyTextBox = new TextBox();
            keyTextBox.Location = new System.Drawing.Point(240, yPosition);
            keyTextBox.Width = 100;
            keyTextBox.ReadOnly = true;
            hotkeysTab.Controls.Add(keyTextBox);

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

        private void LoadVoiceModels()
        {
            string appDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            voiceModels = Directory.GetFiles(appDirectory, "*.onnx")
                .Select(Path.GetFileName)
                .ToList();
            VoiceModelComboBox.Items.Clear();
            VoiceModelComboBox.Items.AddRange(voiceModels.Select(Path.GetFileNameWithoutExtension).ToArray());

            // Select the current model if it exists
            if (!string.IsNullOrEmpty(currentVoiceModel) && VoiceModelComboBox.Items.Contains(currentVoiceModel))
            {
                VoiceModelComboBox.SelectedItem = currentVoiceModel;
            }
        }


        private void LoadSettingsIntoUI()
        {
            isInitializing = true;
            string configPath = PiperTrayApp.GetInstance().GetConfigPath();
            Log($"Loading settings from: {configPath}");

            if (File.Exists(configPath))
            {
                var settings = File.ReadAllLines(configPath)
                    .Where(line => line.Contains("="))
                    .ToDictionary(
                        line => line.Split('=')[0].Trim(),
                        line => line.Split('=')[1].Trim(),
                        StringComparer.OrdinalIgnoreCase
                    );

                InitializeSentenceSilence();

                if (settings.TryGetValue("Logging", out string logging))
                {
                    loggingCheckBox.Checked = bool.Parse(logging);
                    Log($"Set logging checkbox to: {logging}");
                }

                if (settings.TryGetValue("VoiceModel", out string voiceModel))
                {
                    string modelName = Path.GetFileNameWithoutExtension(voiceModel);
                    if (VoiceModelComboBox.Items.Contains(modelName))
                    {
                        VoiceModelComboBox.SelectedItem = modelName;
                        currentVoiceModel = modelName;
                        Log($"Set voice model to: {modelName}");
                    }
                }

                if (settings.TryGetValue("Speaker", out string speakerValue) &&
                    int.TryParse(speakerValue, out int speaker))
                {
                    if (speakerComboBox.Items.Contains(speaker.ToString()))
                    {
                        speakerComboBox.SelectedItem = speaker.ToString();
                        Log($"[LoadSettingsIntoUI] Set speaker to: {speaker}");
                    }
                }

                if (settings.TryGetValue("Speed", out string speed))
                {
                    if (double.TryParse(speed, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double speedValue))
                    {
                        PopulateSpeedComboBox(speedValue);
                        Log($"Set speed to: {speedValue}");
                    }
                }
                LoadSavedPresets();
                LoadMenuVisibilitySettings();
                LoadHotkeySettings(settings);
            }
            else
            {
                Log($"Settings file not found at: {configPath}");
            }

            _isInitialized = true;
            isInitializing = false;
        }


        private void LoadHotkeySettings(Dictionary<string, string> settings)
        {
            try
            {
                Log($"[LoadHotkeySettings] Starting to load hotkey settings");
                foreach (var setting in settings)
                {
                    Log($"[LoadHotkeySettings] Found setting: {setting.Key}={setting.Value}");
                }

                // Log UI control states
                Log($"[LoadHotkeySettings] Control states:");
                Log($"MonitoringModifierComboBox initialized: {monitoringModifierComboBox?.IsHandleCreated}, Items: {monitoringModifierComboBox?.Items.Count}");
                Log($"StopSpeechModifierComboBox initialized: {stopSpeechModifierComboBox?.IsHandleCreated}, Items: {stopSpeechModifierComboBox?.Items.Count}");
                Log($"ChangeVoiceModifierComboBox initialized: {changeVoiceModifierComboBox?.IsHandleCreated}, Items: {changeVoiceModifierComboBox?.Items.Count}");
                Log($"SpeedIncreaseModifierComboBox initialized: {speedIncreaseModifierComboBox?.IsHandleCreated}, Items: {speedIncreaseModifierComboBox?.Items.Count}");
                Log($"SpeedDecreaseModifierComboBox initialized: {speedDecreaseModifierComboBox?.IsHandleCreated}, Items: {speedDecreaseModifierComboBox?.Items.Count}");

                // Monitoring hotkey (working correctly)
                if (settings.TryGetValue("MonitoringModifier", out string monMod) && !string.IsNullOrEmpty(monMod))
                {
                    uint modValue = uint.Parse(monMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                    string modString = GetModifierString(modValue);
                    if (monitoringModifierComboBox.Items.Contains(modString))
                    {
                        monitoringModifierComboBox.SelectedItem = modString;
                        Log($"Set Monitoring modifier to: {modString}");
                    }
                }
                if (settings.TryGetValue("MonitoringKey", out string monKey) && !string.IsNullOrEmpty(monKey))
                {
                    uint keyValue = uint.Parse(monKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                    monitoringKeyTextBox.Text = ((Keys)keyValue).ToString();
                    Log($"Set Monitoring key to: {((Keys)keyValue).ToString()}");
                }

                // Stop Speech hotkey with detailed logging
                if (settings.TryGetValue("StopSpeechModifier", out string stopMod))
                {
                    Log($"[LoadHotkeySettings] Processing StopSpeechModifier: {stopMod}");
                    uint modValue = uint.Parse(stopMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                    string modString = GetModifierString(modValue);
                    Log($"[LoadHotkeySettings] Converted to modString: {modString}");

                    if (stopSpeechModifierComboBox.Items.Contains(modString))
                    {
                        stopSpeechModifierComboBox.SelectedItem = modString;
                        stopSpeechModifiers = modValue;
                        Log($"[LoadHotkeySettings] Set Stop Speech modifier to: {modString} (0x{modValue:X2})");
                    }
                    else
                    {
                        Log($"[LoadHotkeySettings] ModString not found in ComboBox items: {modString}");
                    }
                }
                else
                {
                    Log($"[LoadHotkeySettings] StopSpeechModifier not found in settings");
                }

                if (settings.TryGetValue("StopSpeechKey", out string stopKey))
                {
                    Log($"[LoadHotkeySettings] Processing StopSpeechKey: {stopKey}");
                    uint keyValue = uint.Parse(stopKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                    stopSpeechKeyTextBox.Text = ((Keys)keyValue).ToString();
                    stopSpeechVk = keyValue;
                    Log($"[LoadHotkeySettings] Set Stop Speech key to: {((Keys)keyValue).ToString()} (0x{keyValue:X2})");
                }
                else
                {
                    Log($"[LoadHotkeySettings] StopSpeechKey not found in settings");
                }

                // Change Voice hotkey
                if (settings.TryGetValue("ChangeVoiceModifier", out string changeMod) && !string.IsNullOrEmpty(changeMod))
                {
                    uint modValue = uint.Parse(changeMod.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                    string modString = GetModifierString(modValue);
                    if (changeVoiceModifierComboBox.Items.Contains(modString))
                    {
                        changeVoiceModifierComboBox.SelectedItem = modString;
                        Log($"Set Change Voice modifier to: {modString}");
                    }
                }
                if (settings.TryGetValue("ChangeVoiceKey", out string changeKey) && !string.IsNullOrEmpty(changeKey))
                {
                    uint keyValue = uint.Parse(changeKey.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber);
                    changeVoiceKeyTextBox.Text = ((Keys)keyValue).ToString();
                    Log($"Set Change Voice key to: {((Keys)keyValue).ToString()}");
                }

                // Speed Increase hotkey
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

                // Speed Decrease hotkey
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
            catch (Exception ex)
            {
                Log($"Error in LoadHotkeySettings: {ex.Message}");
            }
        }

        private void PopulateSpeedComboBox(double savedSpeed)
        {
            speedComboBox.Items.Clear();
            for (int i = 1; i <= 10; i++)
            {
                speedComboBox.Items.Add(i.ToString());
            }

            int index = 10 - (int)Math.Round(savedSpeed * 10);
            index = Math.Max(0, Math.Min(9, index)); // Ensure index is within 0-9 range
            speedComboBox.SelectedIndex = index;
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

        private string GetModifierString(uint modifier)
        {
            if (modifier == 0) return "NONE";
            if (modifier == 1) return "ALT";
            if (modifier == 2) return "CTRL";
            if (modifier == 4) return "SHIFT";
            return "NONE";
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
            return (10 - index) / 10.0; // Convert 0-9 index to 1.0-0.1 speed
        }

        private int GetSpeedIndex(string value)
        {
            double speed;
            if (double.TryParse(value, out speed))
            {
                return Array.IndexOf(speedOptions, speed);
            }
            return 0; // Default to first index if parsing fails
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
            float sentenceSilence)
        {
            string configPath = PiperTrayApp.GetInstance().GetConfigPath();
            try
            {
                Log($"[SaveSettings] Entering method. Saving current settings to file.");
                var lines = File.Exists(configPath) ? File.ReadAllLines(configPath).ToList() : new List<string>();

                // Save logging state
                UpdateOrAddSetting(lines, "Logging", loggingCheckBox.Checked.ToString());

                // Save hotkey settings with actual values from UI
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
                UpdateOrAddSetting(lines, "SentenceSilence", sentenceSilence.ToString("F1", System.Globalization.CultureInfo.InvariantCulture));

                File.WriteAllLines(configPath, lines);
                Log($"Settings saved successfully.");
            }
            catch (Exception ex)
            {
                Log($"[SaveSettings] Error saving settings: {ex.Message}");
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            switch (modifier)
            {
                case "ALT": return 0x0001; // MOD_ALT
                case "CTRL": return 0x0002; // MOD_CONTROL
                case "SHIFT": return 0x0004; // MOD_SHIFT
                case "WIN": return 0x0008; // MOD_WIN
                default: return 0x0000;
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
            // Log the incoming values for debugging
            Log($"[UpdateOrAddSetting] Updating {key} with value {value}");

            int index = lines.FindIndex(l => l.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase));

            // Ensure hotkey values are properly formatted
            if (key.EndsWith("Modifier") || key.EndsWith("Key"))
            {
                value = value.StartsWith("0x") ? value : $"0x{value}";
            }

            if (index != -1)
            {
                lines[index] = $"{key}={value}";
                Log($"[UpdateOrAddSetting] Updated setting '{key}' to '{value}' at line {index + 1}");
            }
            else
            {
                lines.Add($"{key}={value}");
                Log($"[UpdateOrAddSetting] Added new setting '{key}' with value '{value}'");
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            Log($"[SaveButton_Click] Entering method");

            try
            {
                // Get current speed value
                double speedValue = GetSpeedValue(speedComboBox.SelectedIndex);
                Log($"[SaveButton_Click] New speed value: {speedValue}");

                // Get sentence silence value
                float sentenceSilence = (float)sentenceSilenceNumeric.Value;
                Log($"[SaveButton_Click] New sentence silence value: {sentenceSilence}");

                // Save presets
                for (int i = 0; i < 4; i++)
                {
                    var preset = new PiperTrayApp.PresetSettings
                    {
                        Name = presetNameTextBoxes[i].Text,
                        VoiceModel = presetVoiceModelComboBoxes[i].SelectedItem?.ToString(),
                        Speaker = int.Parse(presetSpeakerComboBoxes[i].SelectedItem?.ToString() ?? "0"),
                        Speed = GetSpeedValue(presetSpeedComboBoxes[i].SelectedIndex),
                        SentenceSilence = (float)presetSilenceNumericUpDowns[i].Value
                    };

                    PiperTrayApp.GetInstance().SavePreset(i, preset);
                }

                // Get all current hotkey values
                uint monitoringModifiers = monitoringModifierComboBox?.SelectedItem != null ?
                    GetModifierVirtualKeyCode(monitoringModifierComboBox.SelectedItem.ToString()) : 0;
                uint monitoringVk = !string.IsNullOrEmpty(monitoringKeyTextBox?.Text) ?
                    GetVirtualKeyCode(monitoringKeyTextBox.Text) : 0;

                // Get Stop Speech hotkey values with logging
                uint stopSpeechModifiers = stopSpeechModifierComboBox?.SelectedItem != null ?
                    GetModifierVirtualKeyCode(stopSpeechModifierComboBox.SelectedItem.ToString()) : 0;
                uint stopSpeechVk = !string.IsNullOrEmpty(stopSpeechKeyTextBox?.Text) ?
                    GetVirtualKeyCode(stopSpeechKeyTextBox.Text) : 0;

                Log($"[SaveButton_Click] Stop Speech values - Modifier: 0x{stopSpeechModifiers:X2}, Key: 0x{stopSpeechVk:X2}");

                uint changeVoiceModifiers = changeVoiceModifierComboBox?.SelectedItem != null ?
                    GetModifierVirtualKeyCode(changeVoiceModifierComboBox.SelectedItem.ToString()) : 0;
                uint changeVoiceVk = !string.IsNullOrEmpty(changeVoiceKeyTextBox?.Text) ?
                    GetVirtualKeyCode(changeVoiceKeyTextBox.Text) : 0;

                uint speedIncreaseModifiers = speedIncreaseModifierComboBox?.SelectedItem != null ?
                    GetModifierVirtualKeyCode(speedIncreaseModifierComboBox.SelectedItem.ToString()) : 0;
                uint speedIncreaseVk = !string.IsNullOrEmpty(speedIncreaseKeyTextBox?.Text) ?
                    GetVirtualKeyCode(speedIncreaseKeyTextBox.Text) : 0;

                uint speedDecreaseModifiers = speedDecreaseModifierComboBox?.SelectedItem != null ?
                    GetModifierVirtualKeyCode(speedDecreaseModifierComboBox.SelectedItem.ToString()) : 0;
                uint speedDecreaseVk = !string.IsNullOrEmpty(speedDecreaseKeyTextBox?.Text) ?
                    GetVirtualKeyCode(speedDecreaseKeyTextBox.Text) : 0;

                // Save all settings including hotkeys
                SaveSettings(
                    monitoringModifiers,
                    monitoringVk,
                    this.stopSpeechModifiers,
                    this.stopSpeechVk,
                    changeVoiceModifiers,
                    changeVoiceVk,
                    speedIncreaseModifiers,
                    speedIncreaseVk,
                    speedDecreaseModifiers,
                    speedDecreaseVk,
                    (float)sentenceSilenceNumeric.Value
                );

                // Update speed in main app
                SpeedChanged?.Invoke(this, speedValue);
                if (PiperTrayApp.GetInstance() != null)
                {
                    PiperTrayApp.GetInstance().UpdateSpeedFromSettings(speedValue);
                }

                // Register the new hotkeys
                RegisterHotkeys();

                this.DialogResult = DialogResult.OK;
                this.Close();

                Log($"[SaveButton_Click] Settings saved successfully");
            }
            catch (Exception ex)
            {
                Log($"[SaveButton_Click] Error saving settings: {ex.Message}");
                MessageBox.Show($"An error occurred while saving settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
