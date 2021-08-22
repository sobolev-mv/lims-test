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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Viz.WrkModule.RptMagLab
{
  /// <summary>
  /// Interaction logic for UserControl1.xaml
  /// </summary>
  public partial class ViewRptMagLab : Smv.RibbonUserUI.RibbonUserControl
  {
    public ViewRptMagLab(Object Param) : base()
    {
      InitializeComponent();
      this.DataContext = new ViewModelRptMagLab(this, Param);
      //pgbRpt.StyleSettings = new DevExpress.Xpf.Editors.ProgressBarMarqueeStyleSettings();
    }
 
  }
}
