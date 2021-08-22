using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Media.Imaging;
using System.Windows.Data;

namespace Viz.WrkModule.Thp
{

  internal static class ThpBitmap
  {
    public static BitmapImage EmptyImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.Thp;Component/Images/empty-16x16.png"));
    public static BitmapImage NotEmptyImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.Thp;Component/Images/NotEmpty-16x16.png"));
  }
  
  
  public class IntToImageConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      if (value != null){
        if (System.Convert.ToInt32(value) == 0)
          return ThpBitmap.EmptyImage;
        else
          return ThpBitmap.NotEmptyImage;
      }
      else
        return null; 
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
       return null;
    }
  }

  public class IntToBoolConverter : IValueConverter
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





}  




