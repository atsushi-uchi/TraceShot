using Microsoft.Win32; // OpenFolderDialog のために必要
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TraceShot.Services;


namespace TraceShot
{
    public partial class SettingsWindow : Window
    {
        private Key _tempKey;
        private ModifierKeys _tempMod;
        private bool _isWaitingForKey = false;
        public string SelectedPath { get; private set; }

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
                if (int.Parse(item.Tag.ToString()) == savedFps)
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

        //private void SettingsWindow_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        //{
        //    if (!_isWaitingForKey) return;

        //    // 修飾キー単体は無視
        //    if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl || e.Key == Key.LeftAlt ||
        //        e.Key == Key.RightAlt || e.Key == Key.LeftShift || e.Key == Key.RightShift || e.Key == Key.LWin || e.Key == Key.RWin)
        //        return;

        //    e.Handled = true;
        //    _isWaitingForKey = false;
        //    this.PreviewKeyDown -= SettingsWindow_PreviewKeyDown;

        //    _tempKey = (e.Key == Key.System) ? e.SystemKey : e.Key;
        //    _tempMod = Keyboard.Modifiers;

        //    HotkeySettingButton.Content = HotkeyRegister.Format(_tempKey, _tempMod);
        //    HotkeySettingButton.ClearValue(BackgroundProperty);
        //}

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
                Properties.Settings.Default.FrameRate = int.Parse(selectedItem.Tag.ToString());
            }

            // ハードウェアアクセル
            Properties.Settings.Default.UseHardwareAccel = HardwareAccelCheckBox.IsChecked ?? true;

            // 証跡追加ホットキー
            Properties.Settings.Default.HotkeyKey = (int)_tempKey;
            Properties.Settings.Default.HotkeyMod = (int)_tempMod;

            Properties.Settings.Default.Save();
            this.DialogResult = true;
            this.Close();
        }
    }
}