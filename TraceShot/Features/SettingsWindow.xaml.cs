using Microsoft.Win32; // OpenFolderDialog のために必要
using System.Windows;
using System.Windows.Controls;


namespace TraceShot
{
    public partial class SettingsWindow : Window
    {
        public string SelectedPath { get; private set; }

        public SettingsWindow()
        {
            InitializeComponent();

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
                Properties.Settings.Default.FrameRate = int.Parse(selectedItem.Tag.ToString());
            }

            // ハードウェアアクセル
            Properties.Settings.Default.UseHardwareAccel = HardwareAccelCheckBox.IsChecked ?? true;

            Properties.Settings.Default.Save();
            this.DialogResult = true;
            this.Close();
        }
    }
}