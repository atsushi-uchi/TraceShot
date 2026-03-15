using System.Globalization;
using System.Windows.Data;

namespace TraceShot.Converters
{
    public class RelativeToAbsoluteConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            // 最小構成チェック
            if (values.Length < 2 || !(values[0] is double rel) || !(values[1] is double total))
                return 0.0;

            double absolutePos = rel * total;

            // もし3番目の値（コントロールの実際のサイズ）が渡されていたら、その半分を引く
            if (values.Length >= 3 && values[2] is double controlSize && !double.IsNaN(controlSize))
            {
                return absolutePos - (controlSize / 2.0);
            }

            return absolutePos;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}