using System;
using System.Globalization;
using System.Windows.Data;

namespace TraceShot.Converters
{
    public class TypeNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // オブジェクトの型名（クラス名）を文字列で返す
            return value?.GetType().Name ?? string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}