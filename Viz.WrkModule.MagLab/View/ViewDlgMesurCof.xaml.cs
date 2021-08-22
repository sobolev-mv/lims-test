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
using DevExpress.Mvvm.POCO;
using Viz.WrkModule.MagLab.Db.DataSets;


namespace Viz.WrkModule.MagLab
{
  /// <summary>
  /// Interaction logic for DlgDcBlMet.xaml
  /// </summary>
  public partial class ViewDlgMesurCof : DXWindow
  {
    public ViewDlgMesurCof(DsMgLab dsMagLab)
    {
      InitializeComponent();
      this.DataContext = ViewModelSource.Create(() => new ViewModelDlgMesurCof(this, dsMagLab));
    }
  }
}
