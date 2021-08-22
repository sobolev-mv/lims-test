using System;
using System.Data;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.ComponentModel;
using DevExpress.Xpf.Grid;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using System.IO;
using Smv.Utils;


namespace Viz.WrkModule.Thp
{
  internal sealed class ViewModelThp : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields
    private readonly UserControl usrControl;
    private GridControl gcThp;
    private GridControl gcThpDetail;
    private readonly Db.DataSets.DsThp dsThp = new Db.DataSets.DsThp();
    private int prevMasterRowHandle = -1;
    private DataRow currentThpDataRow = null;
    private DataRow currentThpDetailDataRow = null; 
    

    private DateTime dateBeginThp;
    private DateTime dateEndThp;

    #endregion

    #region Public Property
    public DateTime DateBeginThp
    {
      get { return dateBeginThp; }
      set{
        if (value == dateBeginThp) return;
        dateBeginThp = value;
        base.OnPropertyChanged("DateBeginThp");
      }
    }

    public DateTime DateEndThp
    {
      get { return dateEndThp; }
      set{
        if (value == dateEndThp) return;
        dateEndThp = value;
        base.OnPropertyChanged("DateEndThp");
      }
    }

    public DataTable Thp
    {
      get { return dsThp.Thp; }
    }

    public DataTable ThpDetail
    {
      get { return dsThp.ThpDetail; }
    }

    #endregion

    #region Private Method
    private void CurrentItemChanged(object sender, CurrentItemChangedEventArgs args)
    {
      //btnXSamplesRowChanged.CommandParameter = (sender as DevExpress.Xpf.Grid.GridViewBase).Grid.GetRow(e.RowData.RowHandle.Value);
      if (args.NewItem != null){
        currentThpDataRow = (args.NewItem as DataRowView).Row;
        this.dsThp.ThpDetail.LoadData(Convert.ToInt64(this.currentThpDataRow["Id"]));  
      }
      else
        currentThpDataRow = null;
    }

    private void CurrentItemDetailChanged(object sender, CurrentItemChangedEventArgs args)
    {
      if (args.NewItem != null)
        currentThpDetailDataRow = (args.NewItem as DataRowView).Row;
      else
        currentThpDetailDataRow = null;
    }


    private void MasterRowExpanded(object sender, RowEventArgs e)
    {
      GridControl gcDetail = (sender as GridControl).GetDetail(e.RowHandle) as GridControl;

      if ((prevMasterRowHandle >= 0) && e.RowHandle != prevMasterRowHandle)
        (sender as GridControl).CollapseMasterRow(prevMasterRowHandle);

      //this.dsThp.ThpDetail.LoadData(Convert.ToInt64(this.currentThpDataRow["Id"]));  
      gcDetail.ItemsSource = this.ThpDetail;
      prevMasterRowHandle = e.RowHandle;
    }

    private void MasterRowExpanding(object sender, RowAllowEventArgs e)
    {
      //this.dsThp.ThpDetail.LoadData(Convert.ToInt64(this.currentThpDataRow["Id"]));
      //e.Allow = (this.dsThp.ThpDetail.Rows.Count > 0);
    }

    private void LoadDocThpToDb(int typeDoc)
    {
      string fieldNameBlob = (typeDoc == 0) ? "DOC_THP" : "PROT_THP";
      string fieldNameTypeDoc = (typeDoc == 0) ? "TYPE_DOC" : "TYPE_PROT";
      var ofd = new OpenFileDialog 
      { DefaultExt = "*.*", 
        Filter = "text format (.*)|*.*" 
      };

      bool? result = ofd.ShowDialog();
      if (result != true)
        return;

      if (Db.Lob.UploadBlob(ofd.FileName, fieldNameBlob, fieldNameTypeDoc, Convert.ToInt64(this.currentThpDataRow["Id"])))
        Smv.Utils.DxInfo.ShowDxBoxInfo("Информация", "Документ сохранен успешно.", MessageBoxImage.Information);
      else
        return;

      //меняем картинку в гриде
      this.currentThpDataRow.BeginEdit();
      if (typeDoc == 0){
        this.currentThpDataRow["Ldoc"] = 1;
        this.currentThpDataRow["TypeDoc"] = Path.GetExtension(ofd.FileName).ToUpper().Replace(".", ""); 
      }
      else{
        this.currentThpDataRow["Lprot"] = 1;
        this.currentThpDataRow["TypeProt"] = Path.GetExtension(ofd.FileName).ToUpper().Replace(".", ""); 
      }

      this.currentThpDataRow.EndEdit();
      this.currentThpDataRow.AcceptChanges();
      this.currentThpDataRow.Table.AcceptChanges();
    }

    private void ShowDocThpFromDb(int typeDoc)
    {
      string newExt = null;
      string tmpfileName = Path.GetTempFileName();
      string oldExt = Path.GetExtension(tmpfileName).ToUpper().Replace(".", "").ToLower();

      newExt = typeDoc == 0 ? Convert.ToString(this.currentThpDataRow["TypeDoc"]).ToLower() : Convert.ToString(this.currentThpDataRow["TypeProt"]).ToLower();
      tmpfileName = tmpfileName.Replace(oldExt, newExt);

      string fieldNameBlob = (typeDoc == 0) ? "DOC_THP" : "PROT_THP";
      if (!Db.Lob.DownloadBlob(tmpfileName, fieldNameBlob, Convert.ToInt64(this.currentThpDataRow["Id"])))
        return;

      Etc.ExecFileAssociationApp(tmpfileName);
    }

    private void AddDetailThp()
    {
      DataRow rowThp = null;

      //Фокус на главном гриде
      if ((this.gcThp.View.IsFocusedView) && (this.gcThp.View.FocusedRowHandle >= 0)){
        rowThp = (this.gcThp.GetRow(this.gcThp.View.FocusedRowHandle) as DataRowView).Row;
        int cntRowDetail = Convert.ToInt32(rowThp["Cdtl"]);

        if (cntRowDetail == 0){
          rowThp.BeginEdit();
          rowThp["Cdtl"] = 1;
          rowThp.EndEdit();
        }

        //В подчиненом гриде есть записи
        //Проверяем, если свернут, то разворачиваем
        if (!this.gcThp.IsMasterRowExpanded(this.gcThp.View.FocusedRowHandle))
          this.gcThp.ExpandMasterRow(this.gcThp.View.FocusedRowHandle);

        //Добавляем запись
        DataRow newDetailRow = this.dsThp.ThpDetail.NewRow();
        newDetailRow["CodeThp"] = "XXX";
        newDetailRow["NumThp"] = "999";
        newDetailRow["DateThp"] = DateTime.Today;
        newDetailRow["DateBeg"] = DateTime.Today;
        newDetailRow["Pid"] = Convert.ToInt64(rowThp["Id"]);
        this.dsThp.ThpDetail.Rows.Add(newDetailRow);

        (this.gcThp.GetDetail(this.gcThp.View.FocusedRowHandle) as GridControl).View.MoveLastRow();

      } else if (this.gcThpDetail.View.IsFocusedView){

        //this.gcThpDetail.View.MoveLastRow();
        //object ooo = this.gcThpDetail.GetFocusedRow(); 
        //DataRow crRow = (this.gcThpDetail.GetRow(this.gcThpDetail.View.FocusedRowHandle) as DataRowView).Row;
        Int64 pId = Convert.ToInt64(currentThpDetailDataRow["Pid"]);

        //Добавляем запись
        DataRow newDetailRow = this.dsThp.ThpDetail.NewRow();
        newDetailRow["CodeThp"] = "XXX";
        newDetailRow["NumThp"] = "999";
        newDetailRow["DateThp"] = DateTime.Today;
        newDetailRow["DateBeg"] = DateTime.Today;
        newDetailRow["Pid"] = pId;
        this.dsThp.ThpDetail.Rows.Add(newDetailRow);

        //переходим на последнюю 
        this.gcThpDetail.View.MoveLastRow();
      }  
     

    }
    #endregion

    #region Constructor
    internal ViewModelThp(UserControl control)
    {
      usrControl = control;
      DateBeginThp = DateTime.Today;
      DateEndThp = DateTime.Today;

      this.gcThp = LogicalTreeHelper.FindLogicalNode(this.usrControl, "gcThpData") as GridControl;
      if (this.gcThp != null){
        this.gcThp.CurrentItemChanged += CurrentItemChanged;
        this.gcThp.MasterRowExpanding += MasterRowExpanding;
        this.gcThp.MasterRowExpanded += MasterRowExpanded;
      }

      this.gcThpDetail = LogicalTreeHelper.FindLogicalNode(this.usrControl, "gcThpDetailData") as GridControl;
      if (this.gcThpDetail != null)
        this.gcThpDetail.CurrentItemChanged += CurrentItemDetailChanged;
      
      
    }
    #endregion Constructor

    #region Command
    private DelegateCommand<Object> showDataCommand;
    private DelegateCommand<Object> saveDataCommand;
    private DelegateCommand<Object> saveBlobCommand;
    private DelegateCommand<Object> addDetailCommand;
    private DelegateCommand<Object> openBlobCommand;

    public ICommand ShowDataCommand
    {
      get { return showDataCommand ?? (showDataCommand = new DelegateCommand<Object>(ExecuteShowData, CanExecuteShowData)); }
    }

    private void ExecuteShowData(Object parameter)
    {
      this.dsThp.Thp.LoadData(this.DateBeginThp, this.DateEndThp);
    }

    private bool CanExecuteShowData(Object parameter)
    {
      return true;
    }

    public ICommand SaveDataCommand
    {
      get { return saveDataCommand ?? (saveDataCommand = new DelegateCommand<Object>(ExecuteSaveData, CanExecuteSaveData)); }
    }

    private void ExecuteSaveData(Object parameter)
    {
      this.dsThp.Thp.SaveData();
      this.dsThp.ThpDetail.SaveData();
    }

    private bool CanExecuteSaveData(Object parameter)
    {
      return this.dsThp.HasChanges();
    }

    public ICommand SaveBlobCommand
    {
      get { return saveBlobCommand ?? (saveBlobCommand = new DelegateCommand<Object>(ExecuteSaveBlob, CanExecuteSaveBlob)); }
    }

    private void ExecuteSaveBlob(Object parameter)
    {
      this.LoadDocThpToDb(Convert.ToInt32(parameter));
    }

    private bool CanExecuteSaveBlob(Object parameter)
    {
      return (this.currentThpDataRow != null);
    }

    public ICommand OpenBlobCommand
    {
      get { return openBlobCommand ?? (openBlobCommand = new DelegateCommand<Object>(ExecuteOpenBlob, CanExecuteOpenBlob)); }
    }

    private void ExecuteOpenBlob(Object parameter)
    {
      this.ShowDocThpFromDb(Convert.ToInt32(parameter));
    }

    private bool CanExecuteOpenBlob(Object parameter)
    {
      int typeDoc = Convert.ToInt32(parameter);
      Boolean rs;

      if (this.currentThpDataRow == null)
        return false;

      rs = typeDoc == 0 ? Convert.ToInt32(this.currentThpDataRow["Ldoc"]) > 0 : Convert.ToInt32(this.currentThpDataRow["Lprot"]) > 0;
      return rs;
    }

    public ICommand AddDetailCommand
    {
      get { return addDetailCommand ?? (addDetailCommand = new DelegateCommand<Object>(ExecuteAddDetail, CanExecuteAddDetail)); }
    }

    private void ExecuteAddDetail(Object parameter)
    {
      AddDetailThp();
    }

    private bool CanExecuteAddDetail(Object parameter)
    {
      return (this.dsThp.Thp.Rows.Count > 0); 
    }
 

    #endregion
  }
}
