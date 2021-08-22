using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using System.Windows.Media;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;


namespace Viz.WrkModule.MagLab.ViewModel
{
  internal sealed class ViewModelMatGnl : Smv.MVVM.ViewModels.ViewModelBase
  {

    #region Fields
    private Db.DataSets.DsMgLab dsMatGnl;
    private DevExpress.Xpf.Grid.GridControl gcMatGnl;
    #endregion Fields

    #region Public Property
    #endregion Public Property

    #region Private Method
    #endregion Private Method

    #region Constructor
    internal ViewModelMatGnl(System.Windows.Window wnd, Db.DataSets.DsMgLab Ds)
    { 
      this.dsMatGnl = Ds;
      this.gcMatGnl = LogicalTreeHelper.FindLogicalNode(wnd, "gcMatGnl") as DevExpress.Xpf.Grid.GridControl;
      this.gcMatGnl.ItemsSource =  this.dsMatGnl.MlMatGnl;  
      this.dsMatGnl.MlMatGnl.LoadData("212873");

      System.Data.DataRow row = this.dsMatGnl.MlMatGnl.NewRow();
      row[0] = "212873";
      row[1] = DBNull.Value;
      this.dsMatGnl.MlMatGnl.Rows.Add(row);
      this.dsMatGnl.MlMatGnl.AcceptChanges();  
      
    }
    #endregion Constructor

    #region Commands
    #endregion Commands

  }
}

