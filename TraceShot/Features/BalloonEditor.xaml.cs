using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace TraceShot.Features
{
    public partial class BalloonEditor : System.Windows.Controls.UserControl
    {
        // 外部に結果を伝えるためのイベント
        public event Action<string> Committed;
        public event Action Cancelled;

        public BalloonEditor(string initialText = "")
        {
            InitializeComponent();
            BalloonTextInput.Text = initialText;

            // ロードされたら自動的にフォーカスを当てる
            this.Loaded += (s, e) => {
                BalloonTextInput.Focus();
                if (!string.IsNullOrEmpty(BalloonTextInput.Text))
                    BalloonTextInput.SelectAll();
            };
        }

        private void BalloonTextInput_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                Committed?.Invoke(BalloonTextInput.Text);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                Cancelled?.Invoke();
                e.Handled = true;
            }
        }

        private void BalloonTextInput_LostFocus(object sender, RoutedEventArgs e)
        {
            // フォーカスが外れたらキャンセルとみなす（または確定にするか選択可能）
            Cancelled?.Invoke();
        }

        private void BalloonMicButton_Click(object sender, RoutedEventArgs e)
        {
            // 音声入力ロジックをここに記載、またはイベントで外部に投げる
        }
    }
}