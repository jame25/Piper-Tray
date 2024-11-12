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
        private bool presetsInitialized = false;
        private static SettingsForm instance;
        private static readonly object _lock = new object();
        public TabControl TabControl => tabControl;
        private CheckBox[] presetEnableCheckBoxes;
        private Label[] presetLabels;
        private TabPage appearanceTab;

        private Dictionary<string, CheckBox> menuVisibilityCheckboxes;
        private readonly double[] speedOptions = {
            2.0, 1.9, 1.8, 1.7, 1.6, 1.5, 1.4, 1.3, 1.2, 1.1,  // -9 to 0
            1.0, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1   // 1 to 10
        };
        private string currentVoiceModel;

        private ComboBox speakerComboBox;
        private int currentSpeaker = 0;
        private int currentPresetIndex = -1;
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

        private const int topMargin = 10;
        private const int controlSpacing = 5;
        private const int labelHeight = 20;
        private const int controlHeight = 25;
        private const int columnWidth = 100;
        private const int rowSpacing = 35;

        private TabControl tabControl;
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
            public bool Enabled { get; set; }
        }

        private NumericUpDown[] presetSilenceNumericUpDowns;
        private Button[] presetApplyButtons;
        private ComboBox VoiceModelComboBox;
        private List<string> voiceModels;
        private List<string> cachedVoiceModels;
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
            if (instance == null || instance.IsDisposed)
            {
                instance = new SettingsForm();
            }
            return instance;
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

            // First create all form controls
            InitializeFields();

            InitializeCurrentPreset();

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

            isInitializing = false;
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
            presetVoiceModelComboBoxes = new ComboBox[4];
            presetSpeakerComboBoxes = new ComboBox[4];
            presetSpeedComboBoxes = new ComboBox[4];
            presetSilenceNumericUpDowns = new NumericUpDown[4];
            presetApplyButtons = new Button[4];
            presetNameTextBoxes = new TextBox[4];
            tabControl = new TabControl();
            presetEnableCheckBoxes = new CheckBox[4];
            presetLabels = new Label[4];
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

        }

        private void InitializeComboBoxes()
        {

        }

        private void RefreshPresetPanels()
        {
            foreach (Control control in presetsTab.Controls)
            {
                if (control is Panel panel)
                {
                    int panelIndex = (panel.Location.Y - (topMargin + labelHeight + controlSpacing) + 2) / rowSpacing;

                    // Clear existing paint handlers
                    panel.Paint -= new PaintEventHandler(Panel_Paint);
                    panel.BorderStyle = BorderStyle.None;
                    panel.BackColor = SystemColors.Control;

                    if (panelIndex == currentPresetIndex)
                    {
                        panel.Paint += Panel_Paint;
                        panel.BackColor = Color.FromArgb(245, 255, 245);
                    }
                    panel.Refresh();
                }
            }
        }

        private void Panel_Paint(object sender, PaintEventArgs e)
        {
            if (sender is Panel panel)
            {
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(144, 238, 144)))  // Light green
                {
                    e.Graphics.FillRectangle(brush, 0, 0, panel.Width, panel.Height);
                }
                using (Pen pen = new Pen(Color.FromArgb(152, 251, 152), 2))  // Pale green border
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, panel.Width - 1, panel.Height - 1);
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
            if (tabControl.SelectedTab == presetsTab && !presetsInitialized)
            {
                for (int i = 0; i < 4; i++)
                {
                    CreatePresetPanel(i);
                }
                LoadSavedPresets();
                presetsInitialized = true;
            }
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

            appearanceTab = new TabPage("Appearance");
            hotkeysTab = new TabPage("Hotkeys");
            presetsTab = new TabPage("Presets");

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
            // Create preset panel and controls first
            presetsTab.SuspendLayout();
            var controls = new List<Control>();
            int rowY = topMargin + labelHeight + controlSpacing + (index * rowSpacing);

            // Create panel for preset controls
            var presetPanel = new Panel
            {
                Location = new Point(8, rowY - 2),
                Height = controlHeight + 4,
                BorderStyle = BorderStyle.None
            };

            // Add labels if this is the first preset
            if (index == 0)
            {
                controls.AddRange(new Control[] {
                    new Label { Text = "Name", AutoSize = true, Location = new Point(8, 2), Padding = new Padding(2) },
                    new Label { Text = "Model", AutoSize = true, Location = new Point(columnWidth + controlSpacing - 15, 2), Padding = new Padding(2) },
                    new Label { Text = "Speaker", AutoSize = true, Location = new Point((2 * columnWidth) + controlSpacing - 11, 2), Padding = new Padding(2) },
                    new Label { Text = "Speed", AutoSize = true, Location = new Point((2 * columnWidth) + controlSpacing + 44, 2), Padding = new Padding(2) },
                    new Label { Text = "Silence", AutoSize = true, Location = new Point((3 * columnWidth) + controlSpacing - 4, 2), Padding = new Padding(2) }
                });

                // Create checkbox controls at the bottom
                int checkboxStartX = 10;
                int checkboxSpacing = 30;
                int bottomMargin = -80;

                for (int i = 0; i < 4; i++)
                {
                    // Create label for preset number
                    var numberLabel = new Label
                    {
                        Text = $"{i + 1}",
                        Location = new Point(checkboxStartX + (i * checkboxSpacing), presetsTab.Height - bottomMargin),
                        AutoSize = true
                    };

                    // Create checkbox below number
                    presetEnableCheckBoxes[i] = new CheckBox
                    {
                        Text = "",
                        Location = new Point(checkboxStartX + (i * checkboxSpacing), presetsTab.Height - bottomMargin + 20),
                        AutoSize = true
                    };

                    int currentIndex = i;
                    presetEnableCheckBoxes[i].CheckedChanged += (s, e) =>
                    {
                        bool isEnabled = presetEnableCheckBoxes[currentIndex].Checked;
                        presetNameTextBoxes[currentIndex].Enabled = isEnabled;
                        presetVoiceModelComboBoxes[currentIndex].Enabled = isEnabled;
                        presetSpeakerComboBoxes[currentIndex].Enabled = isEnabled;
                        presetSpeedComboBoxes[currentIndex].Enabled = isEnabled;
                        presetSilenceNumericUpDowns[currentIndex].Enabled = isEnabled;

                        // Count enabled presets
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

                        // If only one preset is enabled, make it active
                        if (enabledCount == 1)
                        {
                            currentPresetIndex = enabledIndex;
                            RefreshPresetPanels();
                        }
                    };

                    presetsTab.Controls.Add(numberLabel);
                    presetsTab.Controls.Add(presetEnableCheckBoxes[i]);

                    numberLabel.BringToFront();
                    presetEnableCheckBoxes[i].BringToFront();
                }
            }

            presetsTab.SuspendLayout();

            // Create panel for active preset

            if (index == 0)
            {
                var nameLabel = new Label
                {
                    Text = "Name",
                    AutoSize = true,
                    Location = new Point(8, 2),
                    Padding = new Padding(2)
                };

                var modelLabel = new Label
                {
                    Text = "Model",
                    AutoSize = true,
                    Location = new Point(columnWidth + controlSpacing - 15, 2),
                    Padding = new Padding(2)
                };

                var speakerLabel = new Label
                {
                    Text = "Speaker",
                    AutoSize = true,
                    Location = new Point((2 * columnWidth) + controlSpacing - 11, 2),
                    Padding = new Padding(2)
                };

                var speedLabel = new Label
                {
                    Text = "Speed",
                    AutoSize = true,
                    Location = new Point((2 * columnWidth) + controlSpacing + 44, 2),
                    Padding = new Padding(2)
                };

                var silenceLabel = new Label
                {
                    Text = "Silence",
                    AutoSize = true,
                    Location = new Point((3 * columnWidth) + controlSpacing - 4, 2),
                    Padding = new Padding(2)
                };

                controls.AddRange(new Control[] { nameLabel, modelLabel, speakerLabel, speedLabel, silenceLabel });
            }

            presetNameTextBoxes[index] = new TextBox
            {
                Location = new Point(2, 2),
                Width = (columnWidth - 24),
                Text = $"Preset {index + 1}"
            };
            presetNameTextBoxes[index].TextChanged += (s, e) => UpdatePresetName(index);

            presetVoiceModelComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetNameTextBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth
            };
            PopulateVoiceModelComboBox(presetVoiceModelComboBoxes[index]);
            presetVoiceModelComboBoxes[index].SelectedIndexChanged += (s, e) => UpdatePresetSpeakers(index);

            presetSpeakerComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetVoiceModelComboBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth / 2
            };

            presetSpeedComboBoxes[index] = new ComboBox
            {
                Location = new Point(presetSpeakerComboBoxes[index].Right + controlSpacing, 2),
                Width = columnWidth / 2
            };

            Log($"[CreatePresetPanel] Initializing speed combobox for preset {index}");

            // Clear existing items
            presetSpeedComboBoxes[index].Items.Clear();

            // Add items from -9 to 10
            for (int i = -9; i <= 10; i++)
            {
                presetSpeedComboBoxes[index].Items.Add(i.ToString());
                Log($"[CreatePresetPanel] Added speed option: {i}");
            }

            presetSpeedComboBoxes[index].SelectedIndex = 9; // Index 9 corresponds to value 0

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

            if (index == currentPresetIndex)
            {
                var settings = PiperTrayApp.GetInstance().ReadCurrentSettings();
                if (settings.TryGetValue("Speed", out string speedValue))
                {
                    if (double.TryParse(speedValue, out double speed))
                    {
                        int speedIndex = GetSpeedIndex(speed);
                        presetSpeedComboBoxes[index].SelectedIndex = speedIndex + 9; // Adjust for -9 offset
                    }
                }
            }

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

            // Set panel width based on last control
            presetPanel.Width = presetSilenceNumericUpDowns[index].Right + 4;

            // Add controls to panel
            presetPanel.Controls.AddRange(new Control[] {
                presetNameTextBoxes[index],
                presetVoiceModelComboBoxes[index],
                presetSpeakerComboBoxes[index],
                presetSpeedComboBoxes[index],
                presetSilenceNumericUpDowns[index]
            });

            controls.Add(presetPanel);
            presetsTab.Controls.AddRange(controls.ToArray());

            if (index == 0)
            {
                foreach (Control control in controls)
                {
                    if (control is Label)
                    {
                        control.BringToFront();
                    }
                }
            }

            presetsTab.ResumeLayout(true);
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
            Log("[LoadSavedPresets] Starting to load saved presets");
            isInitializing = true;  // Add this flag

            for (int i = 0; i < 4; i++)
            {
                var preset = PiperTrayApp.GetInstance().LoadPreset(i);
                if (preset != null)
                {
                    Log($"[LoadSavedPresets] Loading preset {i + 1}: Name={preset.Name}, Model={preset.VoiceModel}, Speed={preset.Speed}, Enabled={preset.Enabled}");

                    // Set values without triggering change events
                    presetNameTextBoxes[i].TextChanged -= new EventHandler((s, e) => UpdatePresetName(i));
                    presetVoiceModelComboBoxes[i].SelectedIndexChanged -= (s, e) => UpdatePresetSpeakers(i);

                    presetNameTextBoxes[i].Text = preset.Name;

                    int modelIndex = presetVoiceModelComboBoxes[i].Items.IndexOf(preset.VoiceModel);
                    if (modelIndex >= 0)
                    {
                        presetVoiceModelComboBoxes[i].SelectedIndex = modelIndex;
                        Log($"[LoadSavedPresets] Set model index to {modelIndex} for preset {i + 1}");
                    }

                    presetSpeakerComboBoxes[i].SelectedItem = preset.Speaker.ToString();
                    presetSpeedComboBoxes[i].SelectedItem = GetSpeedIndex(preset.Speed).ToString();
                    presetSilenceNumericUpDowns[i].Value = (decimal)preset.SentenceSilence;

                    // Set enabled state and update controls
                    presetEnableCheckBoxes[i].Checked = preset.Enabled;
                    UpdatePresetControlsEnabled(i, preset.Enabled);

                    // Reattach event handlers
                    presetNameTextBoxes[i].TextChanged += new EventHandler((s, e) => UpdatePresetName(i));
                    presetVoiceModelComboBoxes[i].SelectedIndexChanged += (s, e) => UpdatePresetSpeakers(i);
                }
            }

            isInitializing = false;  // Reset the flag
            Log("[LoadSavedPresets] Finished loading saved presets");
        }

        private void UpdatePresetControlsEnabled(int index, bool enabled)
        {
            presetNameTextBoxes[index].Enabled = enabled;
            presetVoiceModelComboBoxes[index].Enabled = enabled;
            presetSpeakerComboBoxes[index].Enabled = enabled;
            presetSpeedComboBoxes[index].Enabled = enabled;
            presetSilenceNumericUpDowns[index].Enabled = enabled;
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

                LoadSavedPresets();
                LoadMenuVisibilitySettings();
                LoadHotkeySettings(settings);

                if (settings.TryGetValue("Speed", out string speedValue))
                {
                    if (double.TryParse(speedValue, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double speed))
                    {
                        int speedIndex = GetSpeedIndex(speed);
                        // Ensure index is within valid range (0 to ComboBox items count - 1)
                        speedIndex = Math.Max(0, Math.Min(presetSpeedComboBoxes[currentPresetIndex].Items.Count - 1, speedIndex + 9));
                        Log($"Setting active preset (index: {currentPresetIndex}) speed to: {speed} (adjusted combobox index: {speedIndex})");
                        presetSpeedComboBoxes[currentPresetIndex].SelectedIndex = speedIndex;
                    }
                }
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
            switch (index)
            {
                case -9: return 2.0;
                case -8: return 1.9;
                case -7: return 1.8;
                case -6: return 1.7;
                case -5: return 1.6;
                case -4: return 1.5;
                case -3: return 1.4;
                case -2: return 1.3;
                case -1: return 1.2;
                case 0: return 1.1;
                case 1: return 1.0;
                case 2: return 0.9;
                case 3: return 0.8;
                case 4: return 0.7;
                case 5: return 0.6;
                case 6: return 0.5;
                case 7: return 0.4;
                case 8: return 0.3;
                case 9: return 0.2;
                case 10: return 0.1;
                default: return 1.0;
            }
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
            uint speedDecreaseVk)
        {
            var app = PiperTrayApp.GetInstance();
            var lines = File.Exists(app.GetConfigPath()) ?
                File.ReadAllLines(app.GetConfigPath()).ToList() :
                new List<string>();

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

            File.WriteAllLines(app.GetConfigPath(), lines);
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
                // Save presets
                for (int i = 0; i < 4; i++)
                {
                    string selectedValue = presetSpeedComboBoxes[i].SelectedItem?.ToString();
                    int speedIndex = int.Parse(selectedValue); // Gets the actual -9 to 10 value
                    double speed = GetSpeedValue(speedIndex);
                    int speaker = int.Parse(presetSpeakerComboBoxes[i].SelectedItem?.ToString() ?? "0");

                    var preset = new PiperTrayApp.PresetSettings
                    {
                        Name = presetNameTextBoxes[i].Text,
                        VoiceModel = presetVoiceModelComboBoxes[i].SelectedItem?.ToString(),
                        Speaker = speaker,
                        Speed = speed,
                        SentenceSilence = (float)presetSilenceNumericUpDowns[i].Value,
                        Enabled = presetEnableCheckBoxes[i].Checked
                    };

                    PiperTrayApp.GetInstance().SavePreset(i, preset);

                    if (i == currentPresetIndex)
                    {
                        PiperTrayApp.GetInstance().SaveSettings(
                            sentenceSilence: (float)presetSilenceNumericUpDowns[i].Value,
                            speaker: speaker
                        );
                    }
                }

                // Get all current hotkey values
                uint monitoringModifiers = monitoringModifierComboBox?.SelectedItem != null ?
                    GetModifierVirtualKeyCode(monitoringModifierComboBox.SelectedItem.ToString()) : 0;
                uint monitoringVk = !string.IsNullOrEmpty(monitoringKeyTextBox?.Text) ?
                    GetVirtualKeyCode(monitoringKeyTextBox.Text) : 0;

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
                    speedDecreaseVk
                );

                // Register the new hotkeys
                RegisterHotkeys();

                var app = PiperTrayApp.GetInstance();
                var settings = app.ReadCurrentSettings();
                if (settings.TryGetValue("Speed", out string speedValue) &&
                    double.TryParse(speedValue, out double currentSpeed))
                {
                    app.UpdateSpeedFromSettings(currentSpeed);
                }

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
