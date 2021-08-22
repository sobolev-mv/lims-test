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


namespace Viz.DbModule.Psi
{
  /// <summary>
  /// Interaction logic for Window1.xaml
  /// </summary>
  public partial class ConnectWindow : DXWindow
  {
    public ConnectWindow()
    {
      InitializeComponent();

      //Vista features are supported.
      this.SetValue(TextOptions.TextFormattingModeProperty, Environment.OSVersion.Version.Major >= 6 ? TextFormattingMode.Ideal : TextFormattingMode.Display);
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      pbPassword.Focus();
    }
    
    private void chkBoxParam_Checked(object sender, RoutedEventArgs e)
    {
      tbLogin.IsEnabled = ((sender as CheckBox).IsChecked == true);
      gbStrCoding.IsEnabled = ((sender as CheckBox).IsChecked == true);
      tbBase.IsEnabled = ((sender as CheckBox).IsChecked == true);
    }

    private void button2_Click(object sender, RoutedEventArgs e)
    {
      this.DialogResult = true;
    }
  }
}
