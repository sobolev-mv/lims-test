using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media.Imaging;
using System.Windows.Data;

namespace Viz.WrkModule.MagLab
{

  internal static class MagLabBitmap
  {
    public static BitmapImage InWorkImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.MagLab;Component/Images/InWork-16x16.png"));
    public static BitmapImage ToMesImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.MagLab;Component/Images/Check-16x16.png"));
    public static BitmapImage ErrorImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.MagLab;Component/Images/Error-16x16.png"));
    public static BitmapImage InMesImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.MagLab;Component/Images/InMes-16x16.png"));
    //Глифы для Аво-Лазер
    public static BitmapImage LsrImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.MagLab;Component/Images/Validate-16x16.png"));
    public static BitmapImage SlImage = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.MagLab;Component/Images/Slact-16x16.png"));
  }
  
  
  public class IntToImageConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      if (value != null){
        int Status = System.Convert.ToInt32(value);

        switch (Status){
          case 0:
            return MagLabBitmap.InWorkImage;
          case 10:
            return MagLabBitmap.ToMesImage;
          case 20:
            return MagLabBitmap.InMesImage;
          case 40:
            return MagLabBitmap.ErrorImage;
          default:
            return MagLabBitmap.ErrorImage;
        }
      }
      else
        return MagLabBitmap.ErrorImage; 
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
       return null;
    }
  }

  public class SlImageConverter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      if (value == null) return null;
      int slFlg = System.Convert.ToInt32(value);
      switch (slFlg){
        case 0:
          return null;
        case 1:
          return MagLabBitmap.LsrImage;
        case 2:
          return MagLabBitmap.SlImage;
        default:
          return null;
      }
    }

    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
    {
      return null;
    }
  }  



}
