using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DevExpress.Xpf.Core;
using DevExpress.Mvvm.POCO;


namespace Viz.MagLab.MeasureUnits
{
  /// <summary>
  /// Interaction logic for ViewMpg200D.xaml
  /// </summary>
  public partial class ViewBrockhausMpg200D
  {
    public ViewBrockhausMpg200D(int uType, decimal thicknessNominal, string sampNum, DataTable tblCoileSystem, Dictionary<string, decimal> resData)
    {
      InitializeComponent();
      this.DataContext = ViewModelSource.Create(() => new ViewModelBrockhausMpg200D(this, uType, thicknessNominal, sampNum, tblCoileSystem, resData));
    }
  }
}
