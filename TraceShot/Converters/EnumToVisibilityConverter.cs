using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace TraceShot.Converters
{
    public class EnumToVisibilityConverter : IValueConverter
    {
        // Source(enum) -> Target(Visibility)
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return Visibility.Collapsed;

            // バインドされている現在の値(enum)を文字列に変換
            string checkValue = value?.ToString() ?? "";
            // XAMLのConverterParameterで指定された値を文字列に変換
            string targetValue = parameter?.ToString() ?? "";

            // 一致すれば Visible、そうでなければ Collapsed
            return checkValue.Equals(targetValue, StringComparison.OrdinalIgnoreCase)
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        // Target -> Source (今回は使わないので未実装でOK)
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
