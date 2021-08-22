using System;

namespace Viz.WrkModule.RptOtk
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ViewRptOtk
    {
      public ViewRptOtk(Object Param)
      {
        InitializeComponent();
        this.DataContext = new ViewModelRptOtk(this, Param);
      }
    }
}
