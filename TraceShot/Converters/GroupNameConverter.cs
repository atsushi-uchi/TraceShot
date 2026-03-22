using System.Globalization;
using System.Windows.Data;

namespace TraceShot.Converters
{
    public class GroupNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string groupName)
            {
                // "No.1_1" などの文字列から "_" 以降を削除して "No.1" にする
                int index = groupName.IndexOf('_');
                return index > 0 ? groupName.Substring(0, index) : groupName;
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}
