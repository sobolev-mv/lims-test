using System.Data;
using DevExpress.Xpf.Core;


namespace Viz.MagLab.MeasureUnits
{
  /// <summary>
  /// Interaction logic for ViewMeasureListAp.xaml
  /// </summary>
  public partial class ViewMeasureListAp : DXWindow
  {
    public ViewMeasureListAp(int uType, DataTable Data, decimal? Mass, decimal? sLen, decimal? sWid, decimal? sDens, string mD, int mesDevice, DataTable crcftData)
    {
      InitializeComponent();
      Smv.Utils.WindowsOption.AdjustWindowTextOption(this);
      this.DataContext = new ViewModelMeasureListAp(this, uType, Data, Mass, sLen, sWid, sDens, mD, mesDevice, crcftData);
      this.teMassa.Focus();
    }
  }
}
