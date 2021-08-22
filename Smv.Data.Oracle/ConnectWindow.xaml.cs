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

namespace Smv.Data.Oracle
{
  /// <summary>
  /// Interaction logic for Window1.xaml
  /// </summary>
  public partial class ConnectWindow : Window
  {
    public ConnectWindow()
    {
      InitializeComponent();

      //Vista features are supported.
      if (Environment.OSVersion.Version.Major >= 6)
        this.SetValue(TextOptions.TextFormattingModeProperty, TextFormattingMode.Ideal);
      else
        this.SetValue(TextOptions.TextFormattingModeProperty, TextFormattingMode.Display);
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      if ((Environment.OSVersion.Version.Major >= 6) && (Environment.Is64BitOperatingSystem)) {
         this.Background = Brushes.Transparent;
         this.tbLogin.Background = Brushes.Transparent;
         this.pbPassword.Background = Brushes.Transparent;
         this.tbzxz.Background = Brushes.Transparent;
         this.tbBase.Background = Brushes.Transparent;
         VistaGlassHelper.ExtendGlass(this, -1, -1, -1, -1);
      }

      pbPassword.Focus();
    }
    
    private void chkBoxParam_Checked(object sender, RoutedEventArgs e)
    {
      tbLogin.IsEnabled = ((sender as CheckBox).IsChecked == true);
      //tbServer.IsEnabled = ((sender as CheckBox).IsChecked == true);
      tbBase.IsEnabled = ((sender as CheckBox).IsChecked == true);
    }

    private void button2_Click(object sender, RoutedEventArgs e)
    {
      this.DialogResult = true;
    }
  }
}
