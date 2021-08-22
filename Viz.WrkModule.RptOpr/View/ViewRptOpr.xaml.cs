using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Viz.WrkModule.RptOpr
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class ViewRptOpr
    {
      public ViewRptOpr(Object Param) : base()
      {
        InitializeComponent();
        this.DataContext = new ViewModelRptOpr(this, Param);
      }
    }
}
