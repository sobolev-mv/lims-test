using System;
using System.Linq;
using System.Windows.Media;
using System.Windows.Data;

namespace Viz.WrkModule.RptManager
{

  public class BooleanToColorBrush : IValueConverter
  {

    private readonly SolidColorBrush checkBrush = new SolidColorBrush();
    private readonly SolidColorBrush unCheckBrush = new SolidColorBrush();

    public BooleanToColorBrush()
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