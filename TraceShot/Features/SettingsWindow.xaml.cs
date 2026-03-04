using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows;
using Microsoft.Win32; // OpenFolderDialog のために必要


namespace TraceShot
{
    public partial class SettingsWindow : Window
    {
        public string SelectedPath { get; private set; }

        public SettingsWindow()
        {
            InitializeComponent();

            // 初期値として現在の設定を表示（Properties.Settingsなどから）
            SaveFolderTextBox.Text = Properties.Settings.Default.SavePath;
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
            // 設定を永続化保存
            Properties.Settings.Default.SavePath = SaveFolderTextBox.Text;
            Properties.Settings.Default.Save();

            // プロパティにセットしてウィンドウを閉じる
            SelectedPath = SaveFolderTextBox.Text;
            this.DialogResult = true;
            this.Close();
        }
    }
}