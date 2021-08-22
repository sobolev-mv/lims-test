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

namespace Viz.WrkModule.Spep
{
  /// <summary>
  /// Interaction logic for UserControl1.xaml
  /// </summary>
  public partial class ViewSpep : Smv.RibbonUserUI.RibbonUserControl
  {
    public ViewSpep(Object Param)
    {
      InitializeComponent();
      this.DataContext = new ViewModelSpep(this, Param);
    }
  }
}
