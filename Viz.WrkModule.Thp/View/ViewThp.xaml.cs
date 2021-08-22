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
using DevExpress.Xpf.Grid;

namespace Viz.WrkModule.Thp
{
    /// <summary>
    /// Interaction logic for ViewThp.xaml
    /// </summary>
    public partial class ViewThp
    {
      public ViewThp()
      {
        InitializeComponent();
        this.DataContext = new ViewModelThp(this);
      }
      
    }
}
