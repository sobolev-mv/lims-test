using System;
using System.Data;
using System.Windows.Controls;
using System.Windows;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;
using Smv.Utils;
using Viz.WrkModule.RptManager.Db;
using Viz.WrkModule.RptManager.Db.DataSets;


namespace Viz.WrkModule.RptManager
{
  public class ViewModelDlgDcBlMet
  {
    #region Fields
    private readonly Db.DataSets.DsDcBlMet dsDcBlMet = new DsDcBlMet();
    private readonly GridControl gcDc;
    private Control view;
    private DataRow dcRow;

    #endregion

    #region Public Property
    public DataTable FindModePerson
    {
      get { return this.dsDcBlMet.DcBlMet; }
    }

    public virtual Object NewDateFrom { get; set; }

    #endregion

    #region Private Method
    private void CurrentDcRowChanged(object sender, CurrentItemChangedEventArgs args)
    {
      var dataRowView = args.NewItem as DataRowView;
      if (dataRowView == null)
        return; 

      dcRow = dataRowView.Row;
      gcDc.View.AllowEditing = (Convert.ToString(dcRow["IsLast"]) == "Y");
    }
   

    #endregion

    #region Constructor

    public ViewModelDlgDcBlMet(Control control)
    {
      this.view = control;
      this.gcDc = LogicalTreeHelper.FindLogicalNode(this.view, "GcDc") as GridControl;
      if (this.gcDc != null)
        this.gcDc.CurrentItemChanged += CurrentDcRowChanged;
      



      dsDcBlMet.DcBlMet.LoadData();
    }


    #endregion

    #region Command
    public void CloseWnd(Window wnd)
    {
      if (dsDcBlMet.HasChanges())
        if (DxInfo.ShowDxBoxQuestionYn(view, "Сохранение", "Есть несохраненные данные.\nСохранить?", MessageBoxImage.Question))
          dsDcBlMet.DcBlMet.SaveData();  

      if (wnd != null)
         wnd.Close();
    }

    public bool CanCloseWnd(Window wnd)
    {
      return true;
    }

    public void UndoDate()
    {
      dsDcBlMet.DcBlMet.RejectChanges();
    }

    public bool CanUndoDate()
    {
      return true;
    }

    public void SaveDate()
    {
      dsDcBlMet.DcBlMet.SaveData();
    }

    public bool CanSaveDate()
    {
      return true;
    }

    public void DeleteLastDate()
    {
      if (!DxInfo.ShowDxBoxQuestionYn(view, "Удаление", "Внимание последний период будет удален!\nПродолжить?", MessageBoxImage.Warning))
          return;

      if (DbUtils.DeleteLastDateRange())
        dsDcBlMet.DcBlMet.LoadData();      
    }

    public bool CanDeleteLastDate()
    {
      return (dsDcBlMet.DcBlMet.Rows.Count > 0);
    }

    public void AddNewDate()
    {
      if (dsDcBlMet.HasChanges())
        if (DxInfo.ShowDxBoxQuestionYn(view, "Сохранение", "Есть несохраненные данные, которые будут потеряны.\nСохранить?", MessageBoxImage.Question))
          dsDcBlMet.DcBlMet.SaveData(); 

      if (DbUtils.AddNewDateRange(Convert.ToDateTime(NewDateFrom)))
        dsDcBlMet.DcBlMet.LoadData();

      NewDateFrom = null;

    }

    public bool CanAddNewDate()
    {
      return (NewDateFrom != null);
    }


    #endregion

  }
}
