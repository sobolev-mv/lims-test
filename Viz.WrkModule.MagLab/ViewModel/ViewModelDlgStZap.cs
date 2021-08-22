using System;
using System.Data;
using System.Windows.Controls;
using System.Windows;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;
using Smv.Utils;
using Viz.WrkModule.MagLab.Db;
using Viz.WrkModule.MagLab.Db.DataSets;


namespace Viz.WrkModule.MagLab
{
  public class ViewModelDlgStZap
  {
    #region Fields
    private readonly DsMgLab dsMagLab;
    private Control view;

    #endregion

    #region Public Property 
    public DataTable MlStZap
    {
      get { return dsMagLab.MlStZap; }
    }
    #endregion

    #region Private Method
 

    #endregion

    #region Constructor

    public ViewModelDlgStZap(Control control, DsMgLab dsMagLab)
    {
      this.view = control;
      this.dsMagLab = dsMagLab;
      dsMagLab.MlStZap.LoadData();
    }


    #endregion

    #region Command
    public void CloseWnd(Window wnd)
    {
      
      if (wnd != null)
         wnd.Close();
    }

    public bool CanCloseWnd(Window wnd)
    {
      return true;
    }

    #endregion

  }
}
