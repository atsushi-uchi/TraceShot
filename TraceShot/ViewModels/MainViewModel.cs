using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TraceShot.Services;
using System.Text.RegularExpressions;
using TraceShot.Models;

namespace TraceShot.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        // 設定サービスをプロパティとして持つ（UIから色をバインドするため）
        public SettingsService Config => SettingsService.Instance;

        // 録画・再生の状態管理（前述のロジックをここに集約）
        [ObservableProperty] private string _currentTestCaseNo = "No.1";

        // 録画サービスやエビデンスデータへの参照
        public RecService Recorder => RecService.Instance;

        [ObservableProperty] private bool _isPlayerMode = false;

        [RelayCommand]
        private void RecordTestResult(string result)
        {
            // 1. 現在の時間を取得 (録画中なら録画時間、再生中ならシークバー位置)
            // RecServiceに現在のタイムスタンプを返すプロパティがあると仮定
            var currentTime = Recorder.CurrentDuration;

            // 2. ブックマーク（テスト結果）を作成
            var newBookmark = new Bookmark
            {
                Id = Guid.NewGuid(),
                Time = currentTime,
                TestCaseNo = CurrentTestCaseNo,
                Result = result == "OK" ? TestResult.OK : TestResult.NG,
                Note = $"{CurrentTestCaseNo}: {result}"
            };

            // 3. エビデンスデータに追加
            Recorder.Evidence.Bookmarks.Add(newBookmark);

            // 4. JSONに即時保存（データの安全性を確保）
            Recorder.SaveEvidenceJson();

            // 5. 次のテストのためにケース番号をカウントアップ
            IncrementTestCaseNo();
        }

        private void IncrementTestCaseNo()
        {
            if (string.IsNullOrWhiteSpace(CurrentTestCaseNo)) return;

            // 文字列の末尾にある数字を抽出して +1 するロジック
            var match = Regex.Match(CurrentTestCaseNo, @"(\d+)$");
            if (match.Success)
            {
                var numberText = match.Groups[1].Value;
                var number = int.Parse(numberText);
                var prefix = CurrentTestCaseNo.Substring(0, match.Index);

                // 桁数を維持（例: "No.001" -> "No.002"）
                CurrentTestCaseNo = $"{prefix}{(number + 1).ToString(new string('0', numberText.Length))}";
            }
        }
    }
}
