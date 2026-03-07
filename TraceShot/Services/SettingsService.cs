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

        [ObservableProperty]
        private Color _mainColor = GetColorSafe(Default.MainColorName, Colors.Red);

        [ObservableProperty]
        private Color _highlightColor = GetColorSafe(Default.HighlightColorName, Colors.Aqua);

        [ObservableProperty]
        private Color _mainTextColor = GetColorSafe(Default.MainTextColorName, Colors.White);

        [ObservableProperty]
        private Color _highlightTextColor = GetColorSafe(Default.HighlightTextColorName, Colors.Black);

        [ObservableProperty]
        private Color _cropColor = GetColorSafe(Default.CropColorName, Colors.Blue);

        [ObservableProperty]
        private Color _cropFillColor = GetColorSafe(Default.CropFillColorName, Colors.Gray);

        private static Color GetColorSafe(string colorName, Color fallback)
        {
            if (string.IsNullOrEmpty(colorName)) return fallback;
            try
            {
                return (Color)System.Windows.Media.ColorConverter.ConvertFromString(colorName);
            }
            catch
            {
                return fallback; // 変換失敗時（不正な文字列など）も予備の色を返す
            }
        }
    }
}