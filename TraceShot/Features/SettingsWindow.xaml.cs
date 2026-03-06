using Microsoft.Win32; // OpenFolderDialog のために必要
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TraceShot.Services;


namespace TraceShot.Features
{
    public partial class SettingsWindow : Window
    {
        private Key _tempKey;
        private ModifierKeys _tempMod;
        private bool _isWaitingForKey = false;
        public string SelectedPath { get; private set; } = "";

        public SettingsWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            // 1. 保存されたパスを反映
            SaveFolderTextBox.Text = Properties.Settings.Default.SavePath;

            // 2. 保存されたフレームレートを反映
            int savedFps = Properties.Settings.Default.FrameRate;
            foreach (ComboBoxItem item in FpsComboBox.Items)
            {
                if (int.TryParse(item.Tag.ToString(), out var f) && f == savedFps)
                {
                    item.IsSelected = true;
                    break;
                }
            }

            // 3.ハードウェアアクセル
            HardwareAccelCheckBox.IsChecked = Properties.Settings.Default.UseHardwareAccel;

            // ホットキーの表示
            _tempKey = (Key)Properties.Settings.Default.HotkeyKey;
            _tempMod = (ModifierKeys)Properties.Settings.Default.HotkeyMod;
            HotkeySettingButton.Content = HotkeyRegister.Format(_tempKey, _tempMod);

            // 色リストの準備 (主要な色をピックアップ)
            //var colorList = new[] { "White", "Black", "Red", "Blue", "Green", "Orange", "Purple", "DeepPink", "Aqua", "Gold" };
            var colorList = typeof(Brushes).GetProperties().Select(p => p.Name).ToList();
            MainColorComboBox.ItemsSource = colorList;
            HighlightColorComboBox.ItemsSource = colorList;
            MainTextColorComboBox.ItemsSource = colorList;
            HighlightTextColorComboBox.ItemsSource = colorList;
            CropColorComboBox.ItemsSource = colorList;
            CropFillColorComboBox.ItemsSource = colorList;

            // 保存されている色を反映
            MainColorComboBox.SelectedItem = Properties.Settings.Default.MainColorName;
            HighlightColorComboBox.SelectedItem = Properties.Settings.Default.HighlightColorName;
            MainTextColorComboBox.SelectedItem = Properties.Settings.Default.MainTextColorName;
            HighlightTextColorComboBox.SelectedItem = Properties.Settings.Default.HighlightTextColorName;
            CropColorComboBox.SelectedItem = Properties.Settings.Default.CropColorName;
            CropFillColorComboBox.SelectedItem = Properties.Settings.Default.CropFillColorName;
        }

        private void HotkeySettingButton_Click(object sender, RoutedEventArgs e)
        {
            _isWaitingForKey = true;
            HotkeySettingButton.Content = "キーを入力してください...";
            // PreviewKeyDown ではなく PreviewKeyUp を使う
            this.PreviewKeyUp += SettingsWindow_PreviewKeyUp;
        }

        private void SettingsWindow_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (!_isWaitingForKey) return;

            // Snapshot (PrintScreen) かどうかを判定
            Key key = (e.Key == Key.System) ? e.SystemKey : e.Key;

            // 修飾キー単体は無視（離した時も同様）
            if (key == Key.LeftCtrl || key == Key.RightCtrl || key == Key.LeftAlt ||
                key == Key.RightAlt || key == Key.LeftShift || key == Key.RightShift || key == Key.LWin || key == Key.RWin)
                return;

            e.Handled = true;
            _isWaitingForKey = false;
            this.PreviewKeyUp -= SettingsWindow_PreviewKeyUp;

            _tempKey = key;
            _tempMod = Keyboard.Modifiers;

            HotkeySettingButton.Content = HotkeyRegister.Format(_tempKey, _tempMod);
            HotkeySettingButton.ClearValue(BackgroundProperty);
        }

        // 「参照...」ボタンの処理
        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            // フォルダ選択ダイアログのインスタンスを作成
            var dialog = new OpenFolderDialog
            {
                Title = "保存先フォルダを選択してください",
                InitialDirectory = SaveFolderTextBox.Text // 現在のパスを初期位置にする
            };

            // ダイアログを表示し、ユーザーがフォルダを選択した場合
            if (dialog.ShowDialog() == true)
            {
                // テキストボックスに反映
                SaveFolderTextBox.Text = dialog.FolderName;
            }
        }

        // 「保存」ボタンの処理
        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            // デフォルト保存先
            Properties.Settings.Default.SavePath = SaveFolderTextBox.Text;

            // フレームレート
            if (FpsComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                if (int.TryParse(selectedItem.Tag.ToString(), out var framerate))
                {
                    Properties.Settings.Default.FrameRate = framerate;
                }
            }

            // ハードウェアアクセル
            Properties.Settings.Default.UseHardwareAccel = HardwareAccelCheckBox.IsChecked ?? true;

            // 証跡追加ホットキー
            Properties.Settings.Default.HotkeyKey = (int)_tempKey;
            Properties.Settings.Default.HotkeyMod = (int)_tempMod;

            // 表示色の保存
            Properties.Settings.Default.MainColorName = MainColorComboBox.SelectedItem.ToString();
            Properties.Settings.Default.HighlightColorName = HighlightColorComboBox.SelectedItem.ToString();
            Properties.Settings.Default.MainTextColorName = MainTextColorComboBox.SelectedItem.ToString();
            Properties.Settings.Default.HighlightTextColorName = MainTextColorComboBox.SelectedItem.ToString();

            Properties.Settings.Default.Save();
            this.DialogResult = true;
            this.Close();
        }
    }
}