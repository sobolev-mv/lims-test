using System;
using System.Collections.Generic;
using System.Linq;
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


namespace Viz.WrkModule.MagLab.View
{
  /// <summary>
  /// Interaction logic for ViewCalcValues.xaml
  /// </summary>
  public partial class ViewSampleProp : DXWindow
  {
    public ViewSampleProp(String MatLocalId, System.Data.DataTable dt)
    {
      InitializeComponent();
      Smv.Utils.WindowsOption.AdjustWindowTextOption(this);
      this.DataContext = new ViewModel.ViewModelSampleProp(MatLocalId, dt, this); 
    }
  }
}
