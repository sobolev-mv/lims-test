using System;
using System.Collections.Generic;
using System.Data;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Core;
using Viz.WrkModule.Isc.Db.DataSets;
using Viz.DbApp.Psi;
using DevExpress.Xpf.Grid;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Ribbon;
using Smv.Utils;
using Smv.Xls;
using Viz.WrkModule.Isc.Db;
using static System.String;


namespace Viz.WrkModule.Isc
{

  public class UiLng
  {
    public int Id { get; set; }
    public string Name { get; set; }
  } 


  public class ViewModelIsc
  {
    #region Fields
    private readonly UserControl usrControl;
    private readonly RibbonControl rcMain;
    private readonly BarManager bmMain;
    private readonly DXTabControl tcMain;
    private readonly DXTabControl tcDownTime;
    private readonly DsIsc dsIsc = new DsIsc();
    private readonly GridControl dbgMaterial;
    private readonly GridControl dbgShift;
    private readonly GridControl dbgProduct;
    private readonly GridControl dbgDownTime;

    private readonly ProgressBarEdit pgbWait;
    private BarItem biUnit;
    private DataRow currentProdPsiDataRow = null;
    private DataRow currentShiftDataRow = null;
    private DataRow currentProductDataRow = null;
    private DataRow currentDownTimeDataRow = null;

    private readonly XlsInstanceBackgroundReport rpt = new XlsInstanceBackgroundReport();
    #endregion

    #region Public Property
    public DataTable ShipProdProp => this.dsIsc.ShipProdProp;
    public DataTable Agregate => this.dsIsc.Agregate;
    public DataTable Shift => this.dsIsc.Shift;
    public DataTable Product => this.dsIsc.Product;
    public DataTable DtResp => this.dsIsc.DtResp;
    public DataTable DownTime => this.dsIsc.DownTime;
    public DataTable Mnf => this.dsIsc.Mnf;
    public ObservableCollection<UiLng> Rptlng { get; }

    public virtual DateTime DateFrom { get; set; }
    public virtual DateTime DateTo { get; set; }
    public virtual Boolean IsDateShippingChoose { get; set; }
    public virtual Boolean IsManufacturerChoose { get; set; }
    public virtual Boolean IsContractNoChoose { get; set; }
    public virtual string ContractNoValue { get; set; }
    public virtual Boolean IsSpecificationChoose { get; set; }
    public virtual string SpecificationValue { get; set; }
    public virtual Boolean IsSertNoChoose { get; set; }
    public virtual string SertNoValue { get; set; }
    public virtual Boolean IsPlacementNoChoose { get; set; }
    public virtual string PlacementNoValue { get; set; }



    public virtual DateTime ProdDateFrom { get; set; }
    public virtual DateTime ProdDateTo { get; set; }
    public virtual string ProdAgregateId { get; set; }
    public virtual int MnfId { get; set; }
    public virtual Boolean IsControlEnabled { get; set; }
    public virtual DateTime MinValDateDownTime { get; set; }
    public virtual DateTime MaxValDateDownTime { get; set; }
    public virtual int SelectedReportLng { get; set; }
    public virtual string ShiftForeman { get; set; }
    public virtual string SeniorWorker { get; set; }

    #endregion

    #region Private Method

    private void SetHendlerToUnitChange()
    {
      if (biUnit != null)
        return;

      foreach (object item in bmMain.Items.Cast<object>().Where(item => (item as BarItem).Name == "cbUnit")){
        biUnit = (item as BarItem);
        (item as BarEditItem).EditValueChanged += UnitValueChanged;
        return;
      }
    }

    private void RunXlsRptCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      GC.Collect();
      this.pgbWait.StyleSettings = new ProgressBarStyleSettings();
      IsControlEnabled = true;
    }

    //Обработчик когда из списка выбирается другой агрегат
    private void UnitValueChanged(object sender, RoutedEventArgs e)
    {
      dsIsc.Product.Clear();
      dsIsc.DownTime.Clear();
      dsIsc.Shift.Clear();
      dbgShift.View.AllowEditing = true;
      Shift.Columns["AgrId"].DefaultValue = (e.OriginalSource as BarEditItem).EditValue;
      SetUpProductGrid(Convert.ToString((e.OriginalSource as BarEditItem).EditValue));
    }

    private void RibbonSelectedPageChanged(object sender, RibbonPropertyChangedEventArgs e)
    {
      if (e.OldValue == null)
        return;

      if (e.NewValue == null)
        return;
      
      if ((e.OldValue as FrameworkContentElement).Tag == null)
        return;

      if ((e.NewValue as FrameworkContentElement).Tag == null)
        return;
      

      int oldPage = Convert.ToInt32((e.OldValue as FrameworkContentElement).Tag);
      int newPage = Convert.ToInt32((e.NewValue as FrameworkContentElement).Tag);

      if ((oldPage == 0) && (newPage == 1)){
        tcMain.SelectedIndex = 1;
        (tcMain.Items[1] as DXTabItem).Visibility = Visibility.Visible;
        (tcMain.Items[0] as DXTabItem).Visibility = Visibility.Hidden;
        SetHendlerToUnitChange();
      }
      else{
        tcMain.SelectedIndex = 0;
        (tcMain.Items[1] as DXTabItem).Visibility = Visibility.Hidden;
        (tcMain.Items[0] as DXTabItem).Visibility = Visibility.Visible;
      }
    }
    private void CurrentItemChanged(object sender, CurrentItemChangedEventArgs args)
    {
      int tag = Convert.ToInt32((sender as GridControl).Tag);

      switch (tag){
        case 0:
          currentProdPsiDataRow = args.NewItem != null ? (args.NewItem as DataRowView).Row : null;
          break;
        case 1:
          currentShiftDataRow = args.NewItem != null ? (args.NewItem as DataRowView).Row : null;

          if ((currentShiftDataRow != null) && (Convert.ToInt64(currentShiftDataRow["Id"]) > 0)){

            //MinValDateDownTime = Convert.ToDateTime(currentShiftDataRow["DateShift"]);
            //MaxValDateDownTime = Convert.ToDateTime(currentShiftDataRow["DateShift"]).AddHours(23).AddMinutes(59).AddSeconds(59);

            MinValDateDownTime = new DateTime(1900, 1, 1);
            MaxValDateDownTime = new DateTime(2099, 1, 1);

            this.dsIsc.Product.Columns["ShiftId"].DefaultValue = currentShiftDataRow["Id"];
            this.dsIsc.DownTime.Columns["ShiftId"].DefaultValue = currentShiftDataRow["Id"];
            this.dsIsc.DownTime.Columns["DateFrom"].DefaultValue =
                          this.dsIsc.DownTime.Columns["DateTo"].DefaultValue = currentShiftDataRow["DateShift"];

            var task = Task.Factory.StartNew(GetProdactAndDowntimeData, null).ContinueWith(AfterTaskEndProdactAndDowntime);
            

          } else
            dbgProduct.View.AllowEditing = dbgDownTime.View.AllowEditing = false;
          break;
        case 2:
          currentProductDataRow = args.NewItem != null ? (args.NewItem as DataRowView).Row : null;
          break;
        case 3:
          currentDownTimeDataRow = args.NewItem != null ? (args.NewItem as DataRowView).Row : null;
          break;


       default:
          break;
      }
      
    }

    private void AfterTaskEnd(Task obj)
    {
      this.usrControl.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)(() =>
      {
        this.pgbWait.StyleSettings = new ProgressBarStyleSettings();
        IsControlEnabled = true;
      }));
    }

    private void AfterTaskEndProdactAndDowntime(Task obj)
    {
      this.usrControl.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)(() =>
      {
        if ((currentShiftDataRow != null) && (Convert.ToInt64(currentShiftDataRow["Id"]) > 0)){
          dbgProduct.View.AllowEditing = dbgDownTime.View.AllowEditing = true;
          this.dsIsc.Product.Columns["ShiftId"].DefaultValue = currentShiftDataRow["Id"];
          this.dsIsc.DownTime.Columns["ShiftId"].DefaultValue = currentShiftDataRow["Id"];
        }

        this.pgbWait.StyleSettings = new ProgressBarStyleSettings();
        IsControlEnabled = true;
      }));
    }

    private void GetShipProdPropData(object state)
    {
      DbVar.SetRangeDate(DateFrom, DateTo, 1);
      dsIsc.ShipProdProp.LoadData(DateFrom, DateTo, IsDateShippingChoose, ContractNoValue, IsContractNoChoose, SpecificationValue, IsSpecificationChoose, SertNoValue, IsSertNoChoose, MnfId, IsManufacturerChoose, PlacementNoValue, IsPlacementNoChoose);
    }

    private void GetShiftData(object state)
    {
      dsIsc.Shift.LoadData(ProdDateFrom, ProdDateTo, ProdAgregateId);
    }

    private void GetProdactAndDowntimeData(object state)
    {
      this.dsIsc.Product.LoadData(Convert.ToInt64(currentShiftDataRow["Id"]));
      this.dsIsc.DownTime.LoadData(Convert.ToInt64(currentShiftDataRow["Id"]));
    }


    private void SaveData(object state)
    {
      dsIsc.Shift.SaveData();
      //dbgProduct.View.AllowEditing = dbgDownTime.View.AllowEditing = ((currentShiftDataRow != null) && (Convert.ToInt64(currentShiftDataRow["Id"]) > 0));
      dsIsc.Product.SaveData();
      dsIsc.DownTime.SaveData();
    }

    private void SetUpProductGrid(string unitId)
    {
      foreach (var gridColumn in dbgProduct.Columns)
        gridColumn.Visible = false;

      var typeUnit = unitId.Substring(0, 2);

      if (typeUnit == "SL"){
        var vizIdx = 0;

        foreach (var column in dbgProduct.Columns.Where(t => Convert.ToInt32(t.Tag) == 0 || Convert.ToInt32(t.Tag) == 1)){
          column.VisibleIndex = vizIdx;
          column.Visible = true;
          vizIdx++;
        }
      }
      else{
        var vizIdx = 0;

        foreach (var column in dbgProduct.Columns.Where(t => Convert.ToInt32(t.Tag) == 0 || Convert.ToInt32(t.Tag) == 2)){
          column.VisibleIndex = vizIdx;
          column.Visible = true;
          vizIdx++;
        }
        
      }
    }

    #endregion

    #region Constructor
    public ViewModelIsc(UserControl control, Object mainWindow)
    {
      usrControl = control;
      dbgMaterial = LogicalTreeHelper.FindLogicalNode(this.usrControl, "GcShipProdProp") as GridControl;
      if (this.dbgMaterial != null)
        this.dbgMaterial.CurrentItemChanged += CurrentItemChanged;

      dbgShift = LogicalTreeHelper.FindLogicalNode(this.usrControl, "GcProdShift") as GridControl;
      if (this.dbgShift != null)
        this.dbgShift.CurrentItemChanged += CurrentItemChanged;

      dbgProduct = LogicalTreeHelper.FindLogicalNode(this.usrControl, "GcProduct") as GridControl;
      if (this.dbgProduct != null)
        this.dbgProduct.CurrentItemChanged += CurrentItemChanged;

      dbgDownTime = LogicalTreeHelper.FindLogicalNode(this.usrControl, "GcDownTime") as GridControl;
      if (this.dbgDownTime != null)
        this.dbgDownTime.CurrentItemChanged += CurrentItemChanged;

      pgbWait = LogicalTreeHelper.FindLogicalNode(this.usrControl, "PgbMeasure") as ProgressBarEdit;
      tcMain = LogicalTreeHelper.FindLogicalNode(this.usrControl, "tcMain") as DXTabControl;
      tcDownTime = LogicalTreeHelper.FindLogicalNode(this.usrControl, "tcDownTime") as DXTabControl;
      (tcMain.Items[1] as DXTabItem).Visibility = Visibility.Hidden; 

      rcMain = LogicalTreeHelper.FindLogicalNode(mainWindow as Window, "rcMain") as RibbonControl;
      rcMain.SelectedPageChanged += RibbonSelectedPageChanged;
      bmMain = LogicalTreeHelper.FindLogicalNode(mainWindow as Window, "bmMain") as BarManager;

      DevExpress.Xpf.Core.DXGridDataController.DisableThreadingProblemsDetection = true;
      IsControlEnabled = true;
      IsDateShippingChoose = false;
      IsContractNoChoose = false;
      IsSpecificationChoose = false;
      DateFrom = DateTo = ProdDateFrom = ProdDateTo = DateTime.Today;
      this.dsIsc.Agregate.LoadData();
      this.dsIsc.DtResp.LoadData();
      this.dsIsc.Mnf.LoadData();

      Rptlng = new ObservableCollection<UiLng>()
      {
        new UiLng(){Id=1, Name="English language"},
        new UiLng(){Id=2, Name="Russian language"},
        new UiLng(){Id=10, Name="Other"}
      };
      SelectedReportLng = 1;
      MnfId = 1;
    }

    #endregion

    #region Command
    public void ShowDefectMap()
    {
      var view = new ViewMapDefects(this.currentProdPsiDataRow);
      view.ShowDialog();
    }

    public bool CanShowDefectMap()
    {
      return (IsControlEnabled && (currentProdPsiDataRow != null) && (dsIsc.ShipProdProp.Rows.Count != 0));
    }

    public void RptData()
    {
      var src = Etc.StartPath + ModuleConst.VizPruductSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.VizPruductDest;


      var rptParam = new VizPruductRptParam(src, dst)
      {
        DateBegin = DateFrom,
        DateEnd = DateTo,
        IsDateShippingChoose = IsDateShippingChoose,
        IsManufacturerChoose = IsManufacturerChoose,
        IsContractNoChoose = IsContractNoChoose,
        MnfIdValue = MnfId,
        ContractNoValue = ContractNoValue,
        IsSpecificationChoose = IsSpecificationChoose,
        SpecificationValue = SpecificationValue
      };

      var sp = new VizPruduct();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      IsControlEnabled = false;
      this.pgbWait.StyleSettings = new ProgressBarMarqueeStyleSettings();
      (this.pgbWait.StyleSettings as ProgressBarMarqueeStyleSettings).AccelerateRatio = 10;

    }

    public bool CanRptData()
    {
      return (IsControlEnabled && (dsIsc.ShipProdProp.Rows.Count != 0));
    }


    public void GetData()
    {
      IsControlEnabled = false;
      this.pgbWait.StyleSettings = new ProgressBarMarqueeStyleSettings();
      (this.pgbWait.StyleSettings as ProgressBarMarqueeStyleSettings).AccelerateRatio = 10;
      var task = Task.Factory.StartNew(GetShipProdPropData, null).ContinueWith(AfterTaskEnd);
    }
    
    public bool CanGetData()
    {
      return IsControlEnabled;
    }

    public void CloseModule()
    {
      rcMain.SelectedPageChanged -= RibbonSelectedPageChanged;
    }

    public bool CanCloseModule()
    {
      return IsControlEnabled;
    }

    public void GetShift()
    {
      IsControlEnabled = false;
      this.pgbWait.StyleSettings = new ProgressBarMarqueeStyleSettings();
      (this.pgbWait.StyleSettings as ProgressBarMarqueeStyleSettings).AccelerateRatio = 10;
      var task = Task.Factory.StartNew(GetShiftData, null).ContinueWith(AfterTaskEnd);
    }

    public bool CanGetShift()
    {
      return (IsControlEnabled && !IsNullOrEmpty(ProdAgregateId));
    }

    public void SaveData()
    {
      
      IsControlEnabled = false;
      this.pgbWait.StyleSettings = new ProgressBarMarqueeStyleSettings();
      (this.pgbWait.StyleSettings as ProgressBarMarqueeStyleSettings).AccelerateRatio = 10;
      var task = Task.Factory.StartNew(SaveData, null).ContinueWith(AfterTaskEndProdactAndDowntime);
      //SaveData(null);
    }

    public bool CanSaveData()
    {
      return (IsControlEnabled && dsIsc.HasChanges());
    }

    public void UndoData()
    {
      dsIsc.Product.RejectChanges();
      dsIsc.Shift.RejectChanges();
    }

    public bool CanUndoData()
    {
      return (IsControlEnabled && dsIsc.HasChanges());
    }

    public void DeleteShift()
    {
      currentShiftDataRow.Delete(); 
    }

    public bool CanDeleteShift()
    {
      return (IsControlEnabled && currentShiftDataRow != null);
    }

    public void DeleteProduct()
    {
      currentProductDataRow.Delete();
    }

    public bool CanDeleteProduct()
    {
      return (IsControlEnabled && currentProductDataRow != null && tcDownTime.SelectedIndex == 0);
    }

    public void DeleteDownTime()
    {
      currentDownTimeDataRow.Delete();
    }

    public bool CanDeleteDownTime()
    {
      return (IsControlEnabled && currentDownTimeDataRow != null && tcDownTime.SelectedIndex == 1);
    }

    public void RptProdShift()
    {
      string src, dst, local;
      Boolean res;
      string typeUnit = ProdAgregateId.Substring(0, 2);

      if ((typeUnit == "SL") && (SelectedReportLng == 1)){
        src = Etc.StartPath + ModuleConst.ShiftRptSlEnSource;
        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.ShiftRptSlEnDest;
        local = "en";
      }
      else if ((typeUnit == "SL") && (SelectedReportLng == 2)){
        src = Etc.StartPath + ModuleConst.ShiftRptSlRuSource;
        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.ShiftRptSlRuDest;
        local = "ru";
      }
      else if ((typeUnit == "CT") && (SelectedReportLng == 1)){
        src = Etc.StartPath + ModuleConst.ShiftRptCtlEnSource;
        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.ShiftRptCtlEnDest;
        local = "en";
      }
      else if ((typeUnit == "CT") && (SelectedReportLng == 2)){
        src = Etc.StartPath + ModuleConst.ShiftRptCtlRuSource;
        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.ShiftRptCtlRuDest;
        local = "ru";
      }
      else
        return;

      var rptParam = new ShiftRptParam(src, dst)
      {
        DateBegin = Convert.ToDateTime(currentShiftDataRow["DateShift"]),
        Unit = ProdAgregateId,
        Team = Convert.ToString(currentShiftDataRow["Team"]),
        ShiftForeman = ShiftForeman,
        SeniorWorker = SeniorWorker,
        Shift = Convert.ToString(currentShiftDataRow["Shift"]),
        LngId = SelectedReportLng
      };

      if (typeUnit == "SL"){
        var sp = new ShiftRptSl();
        res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam, local);
      }
      else{
        var sp = new ShiftRptCtl();
        res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam, local);
      }

      if (!res) return;
      IsControlEnabled = false;
      this.pgbWait.StyleSettings = new ProgressBarMarqueeStyleSettings();
      (this.pgbWait.StyleSettings as ProgressBarMarqueeStyleSettings).AccelerateRatio = 10;

    }

    public bool CanRptProdShift()
    {
      return (IsControlEnabled && currentShiftDataRow != null && !IsNullOrEmpty(ProdAgregateId) && Convert.ToInt32(currentShiftDataRow["Id"]) > 0);
    }

    public void ShowLaserDiagram()
    {
      string partUrl = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.IscParamConfig, "UrlLaserChart");
      string locNum = Convert.ToString(currentProdPsiDataRow["LocNo"]);
      string placeNum = Convert.ToString(currentProdPsiDataRow["PlacementNum"]);
      string isRewind = Convert.ToString(IscAction.IsRewindAfterLasScr(Convert.ToString(currentProdPsiDataRow["MeId"])));
      Process.Start("IExplore.exe", Format(partUrl, locNum, placeNum, isRewind));
    }

    public bool CanShowLaserDiagram()
    {
      return (IsControlEnabled && (dsIsc.ShipProdProp.Rows.Count != 0) && (Convert.ToInt32(currentProdPsiDataRow["TypeMnf"]) == 1) && (Convert.ToString(currentProdPsiDataRow["StoGrade"]).EndsWith("L")));
    }

    #endregion
  }
}
