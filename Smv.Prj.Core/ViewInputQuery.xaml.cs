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
using DevExpress.Xpf.Editors;

namespace Smv.Dialogs
{
  /// <summary>
  /// Interaction logic for InputQuery.xaml
  /// </summary>
  public sealed partial class ViewInputQuery : DevExpress.Xpf.Core.DXWindow
  {
    private Boolean isPassword;
    public TextEditBase teStringValue;
    

    public ViewInputQuery(Boolean IsPassword)
    {
      InitializeComponent();
      Smv.Utils.WindowsOption.AdjustWindowTextOption(this);
      this.isPassword = IsPassword;

      if (!this.isPassword)
        teStringValue = new TextEdit();
      else
        teStringValue = new PasswordBoxEdit();

      teStringValue.Margin = new Thickness(8,2,12,2);
      teStringValue.ShowBorder = true;
      this.DlgGrd.Children.Add(teStringValue);
      Grid.SetRow(teStringValue, 1);
      Grid.SetColumn(teStringValue, 0);
    }

    private void btnOk_Click(object sender, RoutedEventArgs e)
    {
      this.DialogResult = true;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      teStringValue.Focus();
    }

    private void DXWindow_Closed(object sender, EventArgs e)
    {
      //Text = te.Text; 
    }



  }
}
