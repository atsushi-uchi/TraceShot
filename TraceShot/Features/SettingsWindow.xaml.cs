using Microsoft.Win32; // OpenFolderDialog のために必要
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TraceShot.Services;
using static TraceShot.Properties.Settings;
using Button = System.Windows.Controls.Button;

namespace TraceShot.Features
{
    public partial class SettingsWindow : Window
    {
        private SettingsService _setting = SettingsService.Instance;
        private Key _tempBookmarkKey;
        private ModifierKeys _tempBookmarkMod;
        private Key _tempVoiceKey;
        private ModifierKeys _tempVoiceMod;
        private bool _isWaitingForKey = false;
        private Button? _activeHotkeyButton = null;
        public string SelectedPath { get; private set; } = "";

        public SettingsWindow()
        {
            InitializeComponent();
            DataContext = _setting;
            LoadSettings();
        }

        private void LoadSettings()
        {
            // 1. 保存されたパスを反映
            SaveFolderTextBox.Text = Default.SavePath;

            // 2. 保存されたフレームレートを反映
            int savedFps = Default.FrameRate;
            foreach (ComboBoxItem item in FpsComboBox.Items)
            {
                if (int.TryParse(item.Tag.ToString(), out var f) && f == savedFps)
                {
                    item.IsSelected = true;
                    break;
                }
            }

            // 3.ハードウェアアクセル
            HardwareAccelCheckBox.IsChecked = Default.UseHardwareAccel;

            // ホットキーの表示
            _tempBookmarkKey = (Key)Default.BookmarkHotkeyKey;
            _tempBookmarkMod = (ModifierKeys)Default.BookmarkHotkeyMod;
            HotkeySettingButton.Content = HotkeyRegister.Format(_tempBookmarkKey, _tempBookmarkMod);

            _tempVoiceKey = (Key)Default.VoiceHotkeyKey;
            _tempVoiceMod = (ModifierKeys)Default.VoiceHotkeyMod;
            VoiceHotkeySettingButton.Content = HotkeyRegister.Format(_tempVoiceKey, _tempVoiceMod);

        }

        private void HotkeySettingButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isWaitingForKey) { return; }

            if (sender is Button button)
            {
                _isWaitingForKey = true;
                _activeHotkeyButton = button;
                button.Content = "キーを入力してください...";
                PreviewKeyUp += SettingsWindow_PreviewKeyUp;
            }
        }
        private void SettingsWindow_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (!_isWaitingForKey || _activeHotkeyButton == null) return;

            Key key = (e.Key == Key.System) ? e.SystemKey : e.Key;

            if (IsModifierKey(key)) return; // 修飾キー判定はメソッドに切り出すとスッキリします

            e.Handled = true;
            _isWaitingForKey = false;
            this.PreviewKeyUp -= SettingsWindow_PreviewKeyUp;

            // 設定一時保持用の変数をボタンごとに使い分ける（あるいは直接ボタンのTagなどに入れる）
            // ここではボタンによって保存先プロパティを分岐させる例
            if (_activeHotkeyButton == HotkeySettingButton)
            {
                _tempBookmarkKey = key;
                _tempBookmarkMod = Keyboard.Modifiers;
            }
            else
            {
                _tempVoiceKey = key;
                _tempVoiceMod = Keyboard.Modifiers;
            }

            _activeHotkeyButton.Content = HotkeyRegister.Format(key, Keyboard.Modifiers);
            _activeHotkeyButton.ClearValue(BackgroundProperty);
            _activeHotkeyButton = null;
        }
        // ヘルパー：修飾キーかどうか
        private bool IsModifierKey(Key key) =>
            key == Key.LeftCtrl || key == Key.RightCtrl || key == Key.LeftAlt ||
            key == Key.RightAlt || key == Key.LeftShift || key == Key.RightShift ||
            key == Key.LWin || key == Key.RWin;

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
            Default.SavePath = SaveFolderTextBox.Text;

            // フレームレート
            if (FpsComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                if (int.TryParse(selectedItem.Tag.ToString(), out var framerate))
                {
                    Default.FrameRate = framerate;
                }
            }

            // ハードウェアアクセル
            Default.UseHardwareAccel = HardwareAccelCheckBox.IsChecked ?? true;

            // 証跡追加ホットキー
            Default.BookmarkHotkeyKey = (int)_tempBookmarkKey;
            Default.BookmarkHotkeyMod = (int)_tempBookmarkMod;

            Default.VoiceHotkeyKey = (int)_tempVoiceKey;
            Default.VoiceHotkeyMod = (int)_tempVoiceMod;

            SettingsService.Instance.Save();
            this.DialogResult = true;
            this.Close();
        }
    }
}