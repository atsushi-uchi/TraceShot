using System.Globalization;
using System.Windows.Data;

namespace TraceShot.Converters
{
    public class RelativeToAbsoluteConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length < 2 || !(values[0] is double rel) || !(values[1] is double total))
                return 0.0;

            return rel * total; // 相対値 * 実際のサイズ
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}