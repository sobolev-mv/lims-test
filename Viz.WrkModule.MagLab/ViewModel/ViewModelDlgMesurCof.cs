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
  public class ViewModelDlgMesurCof
  {
    #region Fields
    private DataTable utypeTable;
    private DataTable mesDeviceTable;
    private DataTable typCorTable;
    private readonly DsMgLab dsMagLab;
    private readonly GridControl gcDc;
    private Control view;


    #endregion

    #region Public Property
    public DataTable MlMesurCof => this.dsMagLab.MlMesurCof;
    public DataTable Utype => this.utypeTable;
    public DataTable MesDevice => this.mesDeviceTable;
    public DataTable TypCor => this.typCorTable;
    public virtual Object NewDateFrom { get; set; }

    #endregion

    #region Private Method
    private void CurrentDcRowChanged(object sender, CurrentItemChangedEventArgs args)
    {
      /*
      var dataRowView = args.NewItem as DataRowView;
      if (dataRowView == null)
        return; 
      
      dcRow = dataRowView.Row;
      gcDc.View.AllowEditing = (Convert.ToString(dcRow["IsLast"]) == "Y");
      */
    }

    private void CreateUtypeTable()
    {
      utypeTable = new DataTable();
      utypeTable.Columns.Add("Id", Type.GetType("System.Int32"));
      utypeTable.Columns.Add("Name", Type.GetType("System.String"));

      DataRow row = utypeTable.NewRow();
      row[0] = 1;
      row[1] = "Лист";
      utypeTable.Rows.Add(row);

      row = utypeTable.NewRow();
      row[0] = 2;
      row[1] = "Эпштейн";
      utypeTable.Rows.Add(row);

      row = utypeTable.NewRow();
      row[0] = 3;
      row[1] = "Изоляция";
      utypeTable.Rows.Add(row);

      utypeTable.AcceptChanges();
    }

    private void CreateMesDeviceTable()
    {
      mesDeviceTable = new DataTable();
      mesDeviceTable.Columns.Add("Id", Type.GetType("System.Int32"));
      mesDeviceTable.Columns.Add("Name", Type.GetType("System.String"));

      DataRow row = mesDeviceTable.NewRow();
      row[0] = 1;
      row[1] = "УИ5099";
      mesDeviceTable.Rows.Add(row);

      row = mesDeviceTable.NewRow();
      row[0] = 2;
      row[1] = "MK4Э";
      mesDeviceTable.Rows.Add(row);

      row = mesDeviceTable.NewRow();
      row[0] = 3;
      row[1] = "MPG200D";
      mesDeviceTable.Rows.Add(row);


      mesDeviceTable.AcceptChanges();
    }

    private void CreateTypCorTable()
    {
      typCorTable = new DataTable();
      typCorTable.Columns.Add("Id", Type.GetType("System.String"));
      typCorTable.Columns.Add("Name", Type.GetType("System.String"));

      DataRow row = typCorTable.NewRow();
      row[0] = "D";
      row[1] = "На установке";
      typCorTable.Rows.Add(row);

      row = typCorTable.NewRow();
      row[0] = "G";
      row[1] = "В лаб. данных";
      typCorTable.Rows.Add(row);


      typCorTable.AcceptChanges();
    }
    #endregion

    #region Constructor

    public ViewModelDlgMesurCof(Control control, DsMgLab dsMagLab)
    {
      this.view = control;
      this.gcDc = LogicalTreeHelper.FindLogicalNode(this.view, "GcDc") as GridControl;
      if (this.gcDc != null)
        this.gcDc.CurrentItemChanged += CurrentDcRowChanged;

      this.dsMagLab = dsMagLab;
      CreateUtypeTable();
      CreateMesDeviceTable();
      CreateTypCorTable();
    }

    #endregion

    #region Command
    public void CloseWnd(Window wnd)
    {
      
      if (dsMagLab.HasChanges())
        if (DxInfo.ShowDxBoxQuestionYn(view, "Сохранение", "Есть несохраненные данные.\nСохранить?", MessageBoxImage.Question))
          dsMagLab.MlMesurCof.SaveData();  

      if (wnd != null)
         wnd.Close();
      
    }

    public bool CanCloseWnd(Window wnd)
    {
      return true;
    }

    public void UndoDate()
    {
      dsMagLab.MlMesurCof.RejectChanges();
    }

    public bool CanUndoDate()
    {
      return true;
    }

    public void SaveDate()
    {
      dsMagLab.MlMesurCof.SaveData();
    }

    public bool CanSaveDate()
    {
      return true;
    }

    #endregion

  }
}
