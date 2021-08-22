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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Viz.WrkModule.MagLab
{
  /// <summary>
  /// Interaction logic for UserControl1.xaml
  /// </summary>
  public partial class ViewMagLab
  {
    public ViewMagLab() : base()
    {
      InitializeComponent();
      this.DataContext = new ViewModelMagLab(this);
    }

    /*
    private void TableView_FocusedRowChanged(object sender, DevExpress.Xpf.Grid.FocusedRowChangedEventArgs e)
    {
      //btnXSamplesRowChanged.CommandParameter = (sender as DevExpress.Xpf.Grid.GridViewBase).Grid.GetRow(e.RowData.RowHandle.Value);
      btnXSamplesRowChanged.CommandParameter = e.NewRow;
      if (btnXSamplesRowChanged.CommandParameter == null) return;
      btnXSamplesRowChanged.PerformClick();
    }
    */
  }
}
