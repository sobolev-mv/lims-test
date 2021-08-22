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

namespace Viz.WrkModule.Diagrm
{
  /// <summary>
  /// Interaction logic for UserControl1.xaml
  /// </summary>
  public partial class ViewDiagrm : Smv.RibbonUserUI.RibbonUserControl
  {
    public ViewDiagrm() : base()
    {
      InitializeComponent();
      this.DataContext = new ViewModelDiagrm(this, this.beiGroup);
    }

  }
}
