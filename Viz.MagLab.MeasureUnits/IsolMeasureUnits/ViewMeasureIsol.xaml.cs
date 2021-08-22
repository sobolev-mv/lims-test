using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DevExpress.Xpf.Core;

namespace Viz.MagLab.MeasureUnits
{
  /// <summary>
  /// Interaction logic for Window1.xaml
  /// </summary>
  public partial class ViewMeasureIsol : DXWindow
  {
    public ViewMeasureIsol(List<decimal?> MeasureVal, List<Boolean> lstVisibleMeasurePoint)
    {
      InitializeComponent();
      //this.SetValue(DevExpress.Xpf.Core.ThemeManager.ThemeNameProperty, Smv.Utils.WindowsOption.CurrentTheme);
      Smv.Utils.WindowsOption.AdjustWindowTextOption(this);
      this.DataContext = new ViewModelMeasureIsol(this, MeasureVal, lstVisibleMeasurePoint);
    }
        
  }
}
