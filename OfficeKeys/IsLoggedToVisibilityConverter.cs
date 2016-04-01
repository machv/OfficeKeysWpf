using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace OfficeKeys
{

    public class IsLoggedToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool negate = parameter is string && (parameter as string) == "1";

            if (value is string)
            {
                if (string.IsNullOrEmpty(value as string) == false)
                {
                    return negate ? Visibility.Collapsed : Visibility.Visible;
                }
            }

            return negate ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
