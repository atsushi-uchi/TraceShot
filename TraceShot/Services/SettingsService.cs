using CommunityToolkit.Mvvm.ComponentModel;
using System.Windows.Media;
using static TraceShot.Properties.Settings;
using Color = System.Windows.Media.Color;

namespace TraceShot.Services
{
    public partial class SettingsService : ObservableObject
    {
        private static readonly SettingsService _instance = new();
        public static SettingsService Instance => _instance;

        [ObservableProperty] private bool _isPlayerMode = false;

        [ObservableProperty] private bool _isVoiceEnabled = false;

        // --- Colors (Source of truth) ---
        [ObservableProperty] private Color _mainColor;
        [ObservableProperty] private Color _overColor;
        [ObservableProperty] private Color _mainTextColor;
        [ObservableProperty] private Color _overTextColor;
        [ObservableProperty] private Color _cropColor;
        [ObservableProperty] private Color _cropFillColor;

        // --- Brushes (Auto-generated from colors) ---
        [ObservableProperty] private SolidColorBrush _mainTextBrush = default!;
        [ObservableProperty] private SolidColorBrush _overTextBrush = default!;
        [ObservableProperty] private SolidColorBrush _mainBrush = default!;
        [ObservableProperty] private SolidColorBrush _overBrush = default!;
        [ObservableProperty] private SolidColorBrush _cropBrush = default!;
        [ObservableProperty] private SolidColorBrush _mainFillBrush = default!;
        [ObservableProperty] private SolidColorBrush _overFillBrush = default!;
        [ObservableProperty] private SolidColorBrush _cropFillBrush = default!;

        private SettingsService()
        {
            // 初期値の読み込み
            UpdateAllColors();

            IsVoiceEnabled = Default.IsVoiceEnabled;
        }

        private void UpdateAllColors()
        {
            MainColor = GetColorSafe(Default.MainColorName, Colors.Red);
            OverColor = GetColorSafe(Default.OverColorName, Colors.Aqua);
            MainTextColor = GetColorSafe(Default.MainTextColorName, Colors.White);
            OverTextColor = GetColorSafe(Default.OverTextColorName, Colors.Black);
            CropColor = GetColorSafe(Default.CropColorName, Colors.Blue);
            CropFillColor = GetColorSafe(Default.CropFillColorName, Colors.Gray);
        }

        // --- 自動生成プロパティの変更をフックしてブラシを更新 ---
        partial void OnMainTextColorChanged(Color value) => MainTextBrush = CreateFrozenBrush(value);
        partial void OnOverTextColorChanged(Color value) => OverTextBrush = CreateFrozenBrush(value);
        partial void OnCropColorChanged(Color value) => CropBrush = CreateFrozenBrush(value);


        partial void OnMainColorChanged(Color value)
        {
            MainBrush = CreateFrozenBrush(value);
            MainFillBrush = CreateFrozenBrush(Color.FromArgb(180, value.R, value.G, value.B));
        }

        partial void OnOverColorChanged(Color value)
        {
            OverBrush = CreateFrozenBrush(value);
            OverFillBrush = CreateFrozenBrush(Color.FromArgb(180, value.R, value.G, value.B));
        }


        partial void OnCropFillColorChanged(Color value)
        {
            // A: 128 (約50%の透明度) を指定して新しい Color を作成する
            // アルファ値は 0 (完全に透明) ～ 255 (不透明) で指定します
            Color transparentColor = Color.FromArgb(128, value.R, value.G, value.B);

            // その色でフリーズ済みのブラシを作成
            CropFillBrush = CreateFrozenBrush(transparentColor);
        }
        // --- Helpers ---

        private SolidColorBrush CreateFrozenBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze(); // ★重要: これで描画パフォーマンスが上がり、スレッド間共有も安全になります
            return brush;
        }

        private Color GetColorSafe(string colorName, Color fallback)
        {
            if (string.IsNullOrEmpty(colorName)) return fallback;
            try { return (Color)System.Windows.Media.ColorConverter.ConvertFromString(colorName); }
            catch { return fallback; }
        }

        public void Save()
        {
            // 現在の Color プロパティを文字列に変換して Settings.Default へ
            Default.MainColorName = MainColor.ToString();
            Default.OverColorName = OverColor.ToString();
            Default.MainTextColorName = MainTextColor.ToString();
            Default.OverTextColorName = OverTextColor.ToString();
            Default.CropColorName = CropColor.ToString();
            Default.CropFillColorName = CropFillColor.ToString();
            Default.IsVoiceEnabled = IsVoiceEnabled;
            Default.Save(); // 物理ファイルへ書き込み
        }
        partial void OnIsPlayerModeChanged(bool value)
        {
            // 必要であればここでログ出力やモード切替時の共通処理を行う
            System.Diagnostics.Debug.WriteLine($"Mode changed to: {(value ? "Player" : "Record")}");
        }
    }
}