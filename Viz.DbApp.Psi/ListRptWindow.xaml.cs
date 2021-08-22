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
using DevExpress.Xpf.Core;


namespace Viz.DbApp.Psi
{
  /// <summary>
  /// Interaction logic for Window1.xaml
  /// </summary>
  public partial class ListRptWindow : DXWindow
  {
    public FlowDocument FlowDocHlp
    {
      get { return this.FlowDoc; }
    }

    public ListRptWindow()
    {
      InitializeComponent();
      this.SetValue(TextOptions.TextFormattingModeProperty, Environment.OSVersion.Version.Major >= 6 ? TextFormattingMode.Ideal : TextFormattingMode.Display);
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {

    }
 
  }
}
