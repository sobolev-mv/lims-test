using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows.Media;

namespace Smv.XAML.Convertors
{

  public sealed class DecimalToStringConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      string prm = System.Convert.ToString(parameter);
      string ds = null;
      System.Decimal v = System.Convert.ToDecimal(value);
      ds = v.ToString("n" + prm);
      return ds;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      DateTime d = new DateTime(); ;
      string s = (string)value;
      DateTime.TryParse(s, out d);
      return d;
    }
  }

  public sealed class SelectedToVisibleConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      Boolean v = System.Convert.ToBoolean(value);

      if (v)
        return System.Windows.Visibility.Visible;
      else
        return System.Windows.Visibility.Hidden;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      DateTime d = new DateTime(); ;
      string s = (string)value;
      DateTime.TryParse(s, out d);
      return d;
    }
  }

  public sealed class CharYn2BooleanConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      if (value != null)
        if (System.Convert.ToString(value) == "Y")
          return true;
        else if (System.Convert.ToString(value) == "N")
          return false;
        else
          return null;

      return null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      if (value != null)
        if (System.Convert.ToBoolean(value))
          return "Y";
        else
          return "N";

      return null;
    }
  }

  public sealed class IntToBoolConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      if (value != null)
        return (System.Convert.ToInt32(value) > 0);
      else
        return null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      return null;
    }
  }

  public class BooleanToFilterColor : IValueConverter
  {

    private readonly SolidColorBrush checkBrush = new SolidColorBrush();
    private readonly SolidColorBrush unCheckBrush = new SolidColorBrush();

    public BooleanToFilterColor()
    {
      checkBrush.Color = Color.FromArgb(255, 0x89, 0xA8, 0xF9);
      unCheckBrush.Color = Color.FromArgb(255, 0xCC, 0xCC, 0xCC);
    }

    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      var state = System.Convert.ToBoolean(value);
      return state ? checkBrush : unCheckBrush;
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      throw new NotImplementedException();
    }
  }

  public class MultiBooleanToColorBrush : IMultiValueConverter
  {
    private readonly SolidColorBrush checkBrush = new SolidColorBrush();
    private readonly SolidColorBrush unCheckBrush = new SolidColorBrush();

    public MultiBooleanToColorBrush()
    {
      checkBrush.Color = Color.FromArgb(255, 0x89, 0xA8, 0xF9);
      unCheckBrush.Color = Color.FromArgb(255, 0xCC, 0xCC, 0xCC);
    }

    public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      Boolean res = values.Aggregate(false, (current, val) => current || System.Convert.ToBoolean(val));

      if (res)
        return unCheckBrush;
      return checkBrush;
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
    {
      throw new NotImplementedException();
    }

  }



}
