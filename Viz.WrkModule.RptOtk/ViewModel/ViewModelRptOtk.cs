using System;
using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Windows.Controls;
using DevExpress.Xpf.LayoutControl;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows;
using Microsoft.Win32;
using DevExpress.Xpf.Bars;
using Smv.Utils;
using Viz.DbApp.Psi;


namespace Viz.WrkModule.RptOtk
{
  internal sealed class ViewModelRptOtk : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields

    private readonly Smv.Xls.XlsInstanceBackgroundReport rpt;
    private readonly UserControl usrControl;
    private readonly Object param;
    private DateTime dateBegin;
    private DateTime dateEnd;
    private readonly LayoutGroup lg;
    private readonly Db.DataSets.DsRptOtk dsRptOtk = new Db.DataSets.DsRptOtk(); 
    //--Для отчета ОТК "Надавы ВТО"
    private decimal glubina;
    //--Для отчета ОТК "Распределение дефектов"
    private string defect;
    //Поля для Фильтра качества Начало
    private int selectedTabIndexF1;
    private string listStendF1;
    private Boolean isAroF1;
    private Boolean is1200F1;
    private Boolean isApr1F1;
    private Boolean isAooF1;
    private Boolean isVtoF1;
    private Boolean isAvoF1;
    private Boolean isDateAroF1;
    private Boolean isDate1200F1;
    private Boolean isDateApr1F1;
    private Boolean isDateAooF1;
    private Boolean isMgOF1;
    private Boolean isPppF1;
    private Boolean isWgtCoverF1;
    private Boolean isStVtoF1;
    private Boolean isKlpVtoF1;
    private Boolean isDiskVtoF1;
    private Boolean isDateAvoLstF1;
    private Boolean isTimeAooVtoF1;
    private Boolean isBrgAooF1;
    
    private DateTime dateBeginAroF1;
    private DateTime dateEndAroF1;
    private DateTime dateBegin1200F1;
    private DateTime dateEnd1200F1;
    private DateTime dateBeginApr1F1;
    private DateTime dateEndApr1F1;
    private DateTime dateBeginAooF1;
    private DateTime dateEndAooF1;
    private DateTime dateBeginAvoLstF1;
    private DateTime dateEndAvoLstF1;


    private DataRowView selAroF1Item;
    private DataRowView sel1200F1Item;
    private DataRowView selTolsF1Item;
    private DataRowView selBrgApr1F1Item;
    private DataRowView selBrgVtoF1Item;
    private DataRowView selBrgAvoF1Item;
    private DataRowView selAooF1Item;
    private DataRowView selAvoF1Item;
    private DataRowView selShirApr1F1Item;
    private DataRowView selDiskVtoF1Item;
    private DataRowView selBrgAooF1Item;

    private int apr1F1Width;
    private int aooF1MgOFrom;
    private int aooF1MgOTo;
    private decimal aooF1PppFrom;
    private decimal aooF1PppTo;
    private int aooF1WgtCoverFrom;
    private int aooF1WgtCoverTo;
    private string vtoF1Stend;
    private string vtoF1Cap;
    private int typeListValueF1 = 0; //0-№ Стенда, 1-№ Стенда ВТО 
    private int vtoF1TimeAooVto;
    //Поля для Фильтра качества Конец

    private DataRowView selTolsItem;
    #endregion

    #region Public Property
    public DateTime DateBegin
    {
      get { return dateBegin; }
      set{
        if (value == dateBegin) return;
        dateBegin = value;
        OnPropertyChanged("DateBegin");
      }
    }

    public DateTime DateEnd
    {
      get { return dateEnd; }
      set{
        if (value == dateEnd) return;
        dateEnd = value;
        OnPropertyChanged("DateEnd");
      }
    }

    //--Для отчета ОТК "Надавы ВТО"
    public decimal Glubina
    {
      get { return glubina; }
      set{
        if (value == glubina) return;
        glubina = value;
        OnPropertyChanged("Glubina");
      }
    }

    //--Для отчета ОТК "Распределение дефектов"
    public string Defect 
    {
      get { return defect; }
      set{
        if (value == defect) return;
        defect = value;
        OnPropertyChanged("Defect");
      }
    }


    public DataRowView SelTolsItem
    {
      get { return selTolsItem; }
      set{
        if (Equals(value, selTolsItem)) return;
        selTolsItem = value;
        OnPropertyChanged("SelTolsItem");
      }
    }

    //Поля для Фильтра качества
    public string ListStendF1
    {
      get { return listStendF1; }
      set{
        if (value == listStendF1) return;
        listStendF1 = value;
        OnPropertyChanged("ListStendF1");
      }
    }

    public int SelectedTabIndexF1
    {
      get { return selectedTabIndexF1; }
      set{
        if (value == selectedTabIndexF1) return;
        selectedTabIndexF1 = value;
        OnPropertyChanged("SelectedTabIndexF1");
      }
    }


    public Boolean IsAroF1
    {
      get { return isAroF1; }
      set{
        if (value == isAroF1) return;
        isAroF1 = value;
        if (!value) IsDateAroF1 = false;
        OnPropertyChanged("IsAroF1");
      }
    }

    public Boolean Is1200F1
    {
      get { return is1200F1; }
      set{
        if (value == is1200F1) return;
        is1200F1 = value;
        if (!value) IsDate1200F1 = false;
        OnPropertyChanged("Is1200F1");
      }
    }

    public Boolean IsApr1F1
    {
      get { return isApr1F1; }
      set{
        if (value == isApr1F1) return;
        isApr1F1 = value;
        if (!value) IsDateApr1F1 = false;
        OnPropertyChanged("IsApr1F1");
      }
    }

    public Boolean IsAooF1
    {
      get { return isAooF1; }
      set{
        if (value == isAooF1) return;
        isAooF1 = value;
        if (!value){
          IsDateAooF1 = false;
          IsMgOF1 = false;
          IsPppF1 = false;
          IsWgtCoverF1 = false;
          IsBrgAooF1 = false;
        }
        OnPropertyChanged("IsAooF1");
      }
    }

    public Boolean IsVtoF1
    {
      get { return isVtoF1; }
      set{
        if (value == isVtoF1) return;
        isVtoF1 = value;
        if (!value){
          IsStVtoF1 = false;
          IsKlpVtoF1 = false;
          IsDiskVtoF1 = false;
          IsTimeAooVtoF1 = false;
        }
        OnPropertyChanged("IsVtoF1");
      }
    }

    public Boolean IsAvoF1
    {
      get { return isAvoF1; }
      set{
        if (value == isAvoF1) return;
        isAvoF1 = value;
        OnPropertyChanged("IsAvoF1");
      }
    }

    public Boolean IsDateAroF1
    {
      get { return isDateAroF1; }
      set{
        if (value == isDateAroF1) return;
        isDateAroF1 = value;
        OnPropertyChanged("IsDateAroF1");
      }
    }

    public Boolean IsDate1200F1
    {
      get { return isDate1200F1; }
      set{
        if (value == isDate1200F1) return;
        isDate1200F1 = value;
        OnPropertyChanged("IsDate1200F1");
      }
    }

    //
    public Boolean IsDateAvoLstF1
    {
      get { return isDateAvoLstF1; }
      set{
        if (value == isDateAvoLstF1) return;
        isDateAvoLstF1 = value;
        OnPropertyChanged("IsDateAvoLstF1");
      }
    }


    public Boolean IsDateApr1F1
    {
      get { return isDateApr1F1; }
      set{
        if (value == isDateApr1F1) return;
        isDateApr1F1 = value;
        OnPropertyChanged("IsDateApr1F1");
      }
    }

    public Boolean IsDateAooF1
    {
      get { return isDateAooF1; }
      set{
        if (value == isDateAooF1) return;
        isDateAooF1 = value;
        OnPropertyChanged("IsDateAooF1");
      }
    }

    public Boolean IsMgOF1
    {
      get { return isMgOF1; }
      set
      {
        if (value == isMgOF1) return;
        isMgOF1 = value;
        OnPropertyChanged("IsMgOF1");
      }
    }

    public Boolean IsPppF1
    {
      get { return isPppF1; }
      set{
        if (value == isPppF1) return;
        isPppF1 = value;
        OnPropertyChanged("IsPppF1");
      }
    }

    public Boolean IsWgtCoverF1
    {
      get { return isWgtCoverF1; }
      set{
        if (value == isWgtCoverF1) return;
        isWgtCoverF1 = value;
        OnPropertyChanged("IsWgtCoverF1");
      }
    }

    public Boolean IsStVtoF1
    {
      get { return isStVtoF1; }
      set
      {
        if (value == isStVtoF1) return;
        isStVtoF1 = value;
        OnPropertyChanged("IsStVtoF1");
      }
    }

    public Boolean IsKlpVtoF1
    {
      get { return isKlpVtoF1; }
      set{
        if (value == isKlpVtoF1) return;
        isKlpVtoF1 = value;
        OnPropertyChanged("IsKlpVtoF1");
      }
    }

    public Boolean IsDiskVtoF1
    {
      get { return isDiskVtoF1; }
      set{
        if (value == isDiskVtoF1) return;
        isDiskVtoF1 = value;
        OnPropertyChanged("IsDiskVtoF1");
      }
    }

    public Boolean IsTimeAooVtoF1
    {
      get { return isTimeAooVtoF1; }
      set{
        if (value == isTimeAooVtoF1) return;
        isTimeAooVtoF1 = value;
        OnPropertyChanged("IsTimeAooVtoF1");
      }
    }

    public Boolean IsBrgAooF1
    {
      get { return isBrgAooF1; }
      set{
        if (value == isBrgAooF1) return;
        isBrgAooF1 = value;
        OnPropertyChanged("IsBrgAooF1");
      }
    }

    public DateTime DateBeginAroF1
    {
      get { return dateBeginAroF1; }
      set{
        if (value == dateBeginAroF1) return;
        dateBeginAroF1 = value;
        OnPropertyChanged("DateBeginAroF1");
      }
    }

    public DateTime DateEndAroF1
    {
      get { return dateEndAroF1; }
      set{
        if (value == dateEndAroF1) return;
        dateEndAroF1 = value;
        OnPropertyChanged("DateEndAroF1");
      }
    }

    public DateTime DateBegin1200F1
    {
      get { return dateBegin1200F1; }
      set{
        if (value == dateBegin1200F1) return;
        dateBegin1200F1 = value;
        OnPropertyChanged("DateBegin1200F1");
      }
    }

    public DateTime DateEnd1200F1
    {
      get { return dateEnd1200F1; }
      set{
        if (value == dateEnd1200F1) return;
        dateEnd1200F1 = value;
        OnPropertyChanged("DateEnd1200F1");
      }
    }

    public DateTime DateBeginApr1F1
    {
      get { return dateBeginApr1F1; }
      set{
        if (value == dateBeginApr1F1) return;
        dateBeginApr1F1 = value;
        OnPropertyChanged("DateBeginApr1F1");
      }
    }

    public DateTime DateEndApr1F1
    {
      get { return dateEndApr1F1; }
      set{
        if (value == dateEndApr1F1) return;
        dateEndApr1F1 = value;
        OnPropertyChanged("DateEndApr1F1");
      }
    }

    public DateTime DateBeginAooF1
    {
      get { return dateBeginAooF1; }
      set{
        if (value == dateBeginAooF1) return;
        dateBeginAooF1 = value;
        OnPropertyChanged("DateBeginAooF1");
      }
    }

    public DateTime DateEndAooF1
    {
      get { return dateEndAooF1; }
      set{
        if (value == dateEndAooF1) return;
        dateEndAooF1 = value;
        OnPropertyChanged("DateEndAooF1");
      }
    }

    public DateTime DateBeginAvoLstF1
    {
      get { return dateBeginAvoLstF1; }
      set{
        if (value == dateBeginAvoLstF1) return;
        dateBeginAvoLstF1 = value;
        OnPropertyChanged("DateBeginAvoLstF1");
      }
    }

    public DateTime DateEndAvoLstF1
    {
      get { return dateEndAvoLstF1; }
      set{
        if (value == dateEndAvoLstF1) return;
        dateEndAvoLstF1 = value;
        OnPropertyChanged("DateEndAvoLstF1");
      }
    }


    public DataRowView SelAroF1Item
    {
      get { return selAroF1Item; }
      set{
        if (Equals(value, selAroF1Item)) return;
        selAroF1Item = value;
        OnPropertyChanged("SelAroF1Item");
      }
    }

    public DataTable AroTs
    {
      get { return dsRptOtk.AroTs; }
    }

    public DataRowView Sel1200F1Item
    {
      get { return sel1200F1Item; }
      set{
        if (Equals(value, sel1200F1Item)) return;
        sel1200F1Item = value;
        OnPropertyChanged("Sel1200F1Item");
      }
    }

    public DataTable Rm1200Ts
    {
      get { return dsRptOtk.Rm1200Ts; }
    }

    public DataRowView SelTolsF1Item
    {
      get { return selTolsF1Item; }
      set{
        if (Equals(value, selTolsF1Item)) return;
        selTolsF1Item = value;
        OnPropertyChanged("SelTolsF1Item");
      }
    }

    public DataTable Tols
    {
      get { return dsRptOtk.Tols; }
    }

    public DataRowView SelBrgApr1F1Item
    {
      get { return selBrgApr1F1Item; }
      set{
        if (Equals(value, selBrgApr1F1Item)) return;
        selBrgApr1F1Item = value;
        OnPropertyChanged("SelBrgApr1F1Item");
      }
    }

    public DataTable ShirApr1
    {
      get { return dsRptOtk.ShirApr1; }
    }


    public DataRowView SelShirApr1F1Item
    {
      get { return selShirApr1F1Item; }
      set{
        if (Equals(value, selShirApr1F1Item)) return;
        selShirApr1F1Item = value;
        OnPropertyChanged("SelShirApr1F1Item");
      }
    }
   

    public DataRowView SelBrgVtoF1Item
    {
      get { return selBrgVtoF1Item; }
      set{
        if (Equals(value, selBrgVtoF1Item)) return;
        selBrgVtoF1Item = value;
        OnPropertyChanged("SelBrgVtoF1Item");
      }
    }

    public DataRowView SelBrgAvoF1Item
    {
      get { return selBrgAvoF1Item; }
      set{
        if (Equals(value, selBrgAvoF1Item)) return;
        selBrgAvoF1Item = value;
        OnPropertyChanged("SelBrgAvoF1Item");
      }
    }

    public DataRowView SelDiskVtoF1Item
    {
      get { return selDiskVtoF1Item; }
      set{
        if (Equals(value, selDiskVtoF1Item)) return;
        selDiskVtoF1Item = value;
        OnPropertyChanged("SelDiskVtoF1Item");
      }
    }

    public DataRowView SelBrgAooF1Item
    {
      get { return selBrgAooF1Item; }
      set{
        if (Equals(value, selBrgAooF1Item)) return;
        selBrgAooF1Item = value;
        OnPropertyChanged("SelBrgAooF1Item");
      }
    }

    public DataTable Brg
    {
      get { return dsRptOtk.Brg; }
    }

    public DataRowView SelAooF1Item
    {
      get { return selAooF1Item; }
      set{
        if (Equals(value, selAooF1Item)) return;
        selAooF1Item = value;
        OnPropertyChanged("SelAooF1Item");
      }
    }

    public DataTable AooTs
    {
      get { return dsRptOtk.AooTs; }
    }

    public DataRowView SelAvoF1Item
    {
      get { return selAvoF1Item; }
      set{
        if (Equals(value, selAvoF1Item)) return;
        selAvoF1Item = value;
        OnPropertyChanged("SelAvoF1Item");
      }
    }

    public DataTable AvoTs
    {
      get { return dsRptOtk.AvoTs; }
    }

    public DataTable DiskVtoTs
    {
      get { return dsRptOtk.DiskVto; }
    }

    public int Apr1F1Width
    {
      get { return apr1F1Width; }
      set{
        if (value == apr1F1Width) return;
        apr1F1Width = value;
        OnPropertyChanged("Apr1F1Width");
      }
    }

    public int AooF1MgOFrom
    {
      get { return aooF1MgOFrom; }
      set{
        if (value == aooF1MgOFrom) return;
        aooF1MgOFrom = value;
        OnPropertyChanged("AooF1MgOFrom");
      }
    }

    public int AooF1MgOTo
    {
      get { return aooF1MgOTo; }
      set{
        if (value == aooF1MgOTo) return;
        aooF1MgOTo = value;
        OnPropertyChanged("AooF1MgOTo");
      }
    }

    public decimal AooF1PppFrom
    {
      get { return aooF1PppFrom; }
      set{
        if (value == aooF1PppFrom) return;
        aooF1PppFrom = value;
        OnPropertyChanged("AooF1PppFrom");
      }
    }

    public decimal AooF1PppTo
    {
      get { return aooF1PppTo; }
      set{
        if (value == aooF1PppTo) return;
        aooF1PppTo = value;
        OnPropertyChanged("AooF1PppTo");
      }
    }

    public int AooF1WgtCoverFrom
    {
      get { return aooF1WgtCoverFrom; }
      set{
        if (value == aooF1WgtCoverFrom) return;
        aooF1WgtCoverFrom = value;
        OnPropertyChanged("AooF1WgtCoverFrom");
      }
    }

    public int AooF1WgtCoverTo
    {
      get { return aooF1WgtCoverTo; }
      set{
        if (value == aooF1WgtCoverTo) return;
        aooF1WgtCoverTo = value;
        OnPropertyChanged("AooF1WgtCoverTo");
      }
    }
    
    public string VtoF1Stend
    {
      get { return vtoF1Stend; }
      set{
        if (value == vtoF1Stend) return;
        vtoF1Stend = value;
        OnPropertyChanged("VtoF1Stend");
      }
    }

    public string VtoF1Cap
    {
      get { return vtoF1Cap; }
      set{
        if (value == vtoF1Cap) return;
        vtoF1Cap = value;
        OnPropertyChanged("VtoF1Cap");
      }
    }

    public int VtoF1TimeAooVto
    {
      get { return vtoF1TimeAooVto; }
      set{
        if (value == vtoF1TimeAooVto) return;
        vtoF1TimeAooVto = value;
        OnPropertyChanged("VtoF1TimeAooVto");
      }
    }


    #endregion

    #region Private Method
    private void RunXlsRptCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      var barEditItem = param as BarEditItem;
      if (barEditItem != null)
        barEditItem.IsVisible = false;
      GC.Collect();
    }

    private void LayoutGroupExpanded(object sender, EventArgs e)
    {
      var layoutGroup = sender as LayoutGroup;
      if (layoutGroup != null){
        var i = Convert.ToInt32(layoutGroup.Tag);
        switch (i){
          case 0:
            this.LgExpanded0();
            break;
          case 1:
            this.LgExpandedF1();
            break;
          case 2:
            break;
          case 3:
            break;
          case 4:
            break;
          case 5:
            break;
        }
      }
    }

    private void LayoutGroupCollapsed(object sender, EventArgs e)
    {
      var layoutGroup = sender as LayoutGroup;
      if (layoutGroup != null){
        var i = Convert.ToInt32(layoutGroup.Tag);
        switch (i){
          case 0:
            break;
          case 1:
            this.LgCollapsedF1();
            break;
          case 2:
            break;
          case 3:
            break;
          case 4:
            break;
        }
      }
    }

    private void LgExpandedF1()
    {
      this.Glubina = 30;
      this.Defect = "501";

      if (dsRptOtk.AroTs.Rows.Count == 0) 
         dsRptOtk.AroTs.LoadData(2);

      if (dsRptOtk.Rm1200Ts.Rows.Count == 0) 
        dsRptOtk.Rm1200Ts.LoadData(1);

      if (dsRptOtk.Tols.Rows.Count == 0) 
        dsRptOtk.Tols.LoadData(12);

      if (dsRptOtk.Brg.Rows.Count == 0) 
        dsRptOtk.Brg.LoadData(13);

      if (dsRptOtk.AooTs.Rows.Count == 0) 
        dsRptOtk.AooTs.LoadData(3);

      if (dsRptOtk.AvoTs.Rows.Count == 0) 
        dsRptOtk.AvoTs.LoadData(4);

      if (dsRptOtk.ShirApr1.Rows.Count == 0) 
        dsRptOtk.ShirApr1.LoadData(16);

      if (dsRptOtk.DiskVto.Rows.Count == 0)
        dsRptOtk.DiskVto.LoadData(17);

      for (int i = 0; i < 11; i++){
        var cbe = LogicalTreeHelper.FindLogicalNode(this.usrControl, "CbeTypeF1Ts" + i.ToString(CultureInfo.InvariantCulture)) as DevExpress.Xpf.Editors.ComboBoxEdit;
        if (cbe != null)
          cbe.SelectedIndex = 0;
      }

      var layoutGroup = LogicalTreeHelper.FindLogicalNode(this.usrControl, "lgFilter1") as LayoutGroup;
      if (layoutGroup != null)
        layoutGroup.SelectedTabChildChanged += SelectedTypeFilter;

    }

    private void LgExpanded0()
    {
      if (dsRptOtk.Tols.Rows.Count == 0) 
        dsRptOtk.Tols.LoadData(12);

      var cbe = LogicalTreeHelper.FindLogicalNode(this.usrControl, "CbeTypeTs") as DevExpress.Xpf.Editors.ComboBoxEdit;
      if (cbe != null)
        cbe.SelectedIndex = 0;
    }

    private void LgCollapsedF1()
    {
      var layoutGroup = LogicalTreeHelper.FindLogicalNode(this.usrControl, "lgFilter1") as LayoutGroup;
      if (layoutGroup != null)
        layoutGroup.SelectedTabChildChanged -= SelectedTypeFilter;
    }


    void SelectedTypeFilter(object sender, DevExpress.Xpf.Core.ValueChangedEventArgs<FrameworkElement> e)
    {
      var layoutGroup = sender as LayoutGroup;
      if (layoutGroup != null)
        this.SelectedTabIndexF1 = layoutGroup.SelectedTabIndex;
    }

    private string GetStringFromTxtFile()
    {
      var ofd = new OpenFileDialog { DefaultExt = ".txt", Filter = "text format (.txt)|*.txt" };
      bool? result = ofd.ShowDialog();
      if (result != true)
        return string.Empty;
      return System.IO.File.ReadAllText(ofd.FileName, System.Text.Encoding.GetEncoding(1251)).Replace(" ", "").Replace("\r\n", " ").Trim().Replace(" ", ",");
    }

    /*
    private Boolean IsFilter1Applayed()
    {
      return (isAroF1 || is1200F1 || isApr1F1 || isAooF1 || isVtoF1 || isAvoF1); 
    }
    */ 
    #endregion

    #region Constructor
    internal ViewModelRptOtk(System.Windows.Controls.UserControl control, Object Param)
    {
      param = Param;
      rpt = new Smv.Xls.XlsInstanceBackgroundReport();
      usrControl = control;
      DateBegin = DateEnd = DateBeginAroF1 = DateEndAroF1 = DateBegin1200F1 = DateEnd1200F1 = DateBeginApr1F1 = DateEndApr1F1 = DateBeginAooF1 = DateEndAooF1 = DateBeginAvoLstF1 = DateEndAvoLstF1 = DateTime.Today;
 
      for (int i = ModuleConst.AccGrpLg0; i < ModuleConst.AccGrpLg2 + 1; i++){
        this.lg = LogicalTreeHelper.FindLogicalNode(this.usrControl, "Lg" + ModuleConst.ModuleId + "_" + i.ToString()) as LayoutGroup;

        if (this.lg != null){
          if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            lg.Visibility = Visibility.Hidden;
            continue;
          }
          this.lg.Expanded += LayoutGroupExpanded;
          this.lg.Collapsed += LayoutGroupCollapsed;
        }
      }

      //Группы 2-уровня
      for (int i = ModuleConst.AccGrpNadavDefect; i < ModuleConst.AccGrpTo2Sort + 1; i++){
        var uie = LogicalTreeHelper.FindLogicalNode(this.usrControl, "G2" + ModuleConst.ModuleId + "_" + i) as UIElement;

        if (uie != null){
          if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            uie.Visibility = Visibility.Hidden;
            continue;
          }
        }
      }

      //Делаем кнопки управления невидимыми 
      for (int i = ModuleConst.AccCmdOtkAvoBonus; i < ModuleConst.AccCmdFinCutByCat + 1; i++){
        var uie = LogicalTreeHelper.FindLogicalNode(this.usrControl, "b" + ModuleConst.ModuleId + "_" + i) as UIElement;
        if (uie == null) continue;

        if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
          uie.Visibility = Visibility.Hidden;
          continue;
        }
      }

      //DxInfo.ShowDxBoxInfo("Внимание", "Файлы созданных отчетов будут находиться в папке Документы (Мои документы)!", MessageBoxImage.Information);
    }
    #endregion

    #region Commands
    //--Для отчетов с учетом фильтра1
    private DelegateCommand<Object> selectTypeListValueF1Command;
    //--Для отчета Карениной ХЛ "Диэлектрич. свойство покрытия"
    private DelegateCommand<Object> showListRptCommand;
    private DelegateCommand<Object> loadFromTxtFileCommand;
    private DelegateCommand<Object> otkQntDefMonthCommand;
    private DelegateCommand<Object> otkAvoBonusCommand;
    private DelegateCommand<Object> catBrAooCommand;
    private DelegateCommand<Object> devDayKesiCommand;
    private DelegateCommand<Object> otkQualityAvoCommand;
    private DelegateCommand<Object> otkDefectAvoCommand;
    private DelegateCommand<Object> otkDefectAvoYearCommand;
    private DelegateCommand<Object> otkDefect501Command;
    private DelegateCommand<Object> otkDistrib501OnLengthCommand;
    private DelegateCommand<Object> оtkShirApr1Command;
    private DelegateCommand<Object> otkNadavVtoCommand;
    private DelegateCommand<Object> otkOutMe1Cls1SrtCommand;
    private DelegateCommand<Object> otkOutMeWdt1000Command;
    private DelegateCommand<Object> otkTo2SortCommand;
    private DelegateCommand<Object> otkInfoShovProryvCommand;
    private DelegateCommand<Object> otkSgpDefectsCommand;
    private DelegateCommand<Object> otkSgpDefectsSort1GostCommand;
    private DelegateCommand<Object> otkDistribDefectsOnLengthCommand;
    private DelegateCommand<Object> otkDistribDefectsOnSurfaceCommand;
    private DelegateCommand<Object> otkSeqCoilLineAooCommand;
    private DelegateCommand<Object> otkChratcerListCoilsCommand;
    private DelegateCommand<Object> otkFreqDistrDefectAvoCommand;
    private DelegateCommand<Object> otkFinCutByCatCommand;
    public ICommand ShowListRptCommand
    {
      get { return showListRptCommand ?? (showListRptCommand = new DelegateCommand<Object>(ExecuteShowListRpt, CanExecuteShowListRpt)); }
    }

    private void ExecuteShowListRpt(Object parameter)
    {
      var im = new Image
      {
        Height = 64,
        Width = 64,
        Stretch = Stretch.None,
        HorizontalAlignment = HorizontalAlignment.Left,
        VerticalAlignment = VerticalAlignment.Top,
        Margin = new Thickness(0, 0, 20, 0),
        Source = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.RptOtk;Component/Images/BarImage.png"))
      };

      ListRpt.ShowListRpt(ModuleConst.ModuleId, im);
    }

    private bool CanExecuteShowListRpt(Object parameter)
    {
      return true;
    }

    public ICommand LoadFromTxtFileCommand
    {
      get { return loadFromTxtFileCommand ?? (loadFromTxtFileCommand = new DelegateCommand<Object>(ExecuteLoadFromTxtFile, CanExecuteLoadFromTxtFile)); }
    }

    private void ExecuteLoadFromTxtFile(Object parameter)
    {
      var i = Convert.ToInt32(parameter);

      if (i == 60)
        this.VtoF1Stend = GetStringFromTxtFile();
      else
        this.ListStendF1 = GetStringFromTxtFile();
    }

    private bool CanExecuteLoadFromTxtFile(Object parameter)
    {
      return true;
    }

    public ICommand SelectTypeListValueF1Command
    {
      get { return selectTypeListValueF1Command ?? (selectTypeListValueF1Command = new DelegateCommand<Object>(ExecuteSelectTypeListValueF1, CanExecuteSelectTypeListValueF1)); }
    }

    private void ExecuteSelectTypeListValueF1(Object parameter)
    {
      this.typeListValueF1 = Convert.ToInt32(parameter);
    }

    private bool CanExecuteSelectTypeListValueF1(Object parameter)
    {
      return true;
    }

    public ICommand OtkQntDefMonthCommand
    {
      get { return otkQntDefMonthCommand ?? (otkQntDefMonthCommand = new DelegateCommand<Object>(ExecuteOtkQntDefMonth, CanExecuteOtkQntDefMonth)); }
    }

    private void ExecuteOtkQntDefMonth(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkQntDefMonthSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkQntDefMonthDest;
      var rptParam = new Db.OtkQntDefMonthRptParam(src, dst, this.DateBegin, this.DateEnd)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        //DateBegin = this.DateBegin,
        //DateEnd = this.DateEnd,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.OtkQntDefMonth();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkQntDefMonth(Object parameter)
    {
      return true;
    }

    public ICommand OtkAvoBonusCommand
    {
      get { return otkAvoBonusCommand ?? (otkAvoBonusCommand = new DelegateCommand<Object>(ExecuteOtkAvoBonus, CanExecuteOtkAvoBonus)); }
    }

    private void ExecuteOtkAvoBonus(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.OtkAvoBonusSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkAvoBonusDest;

      var sp = new Db.OtkAvoBonus()
      /*
      {
        IdReport = ModuleConst.AccCmdOtkAvoBonus,
        ConnectToTargetDb = DbSelector.ConnectToTargetDb,
        GetCurrentDbAlias = DbSelector.GetCurrentDbAlias
      }*/;

      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.OtkAvoBonusRptParam(src, dst, this.DateBegin, this.DateEnd));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkAvoBonus(Object parameter)
    {
      return true;
    }


    public ICommand CatBrAooCommand
    {
      get { return catBrAooCommand ?? (catBrAooCommand = new DelegateCommand<Object>(ExecuteCatBrAoo, CanExecuteCatBrAoo)); }
    }

    private void ExecuteCatBrAoo(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.CatBrAooSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CatBrAooDest;

      var sp = new Db.CatBrAoo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.CatBrAooRptParam(src, dst, this.DateBegin, this.DateEnd));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteCatBrAoo(Object parameter)
    {
      return true;
    }

    public ICommand DevDayKesiCommand
    {
      get { return devDayKesiCommand ?? (devDayKesiCommand = new DelegateCommand<Object>(ExecuteDevDayKesi, CanExecuteDevDayKesi)); }
    }

    private void ExecuteDevDayKesi(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.DevDayKesiSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.DevDayKesiDest;

      var sp = new Db.DevDayKesi();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.DevDayKesiRptParam(src, dst, this.DateBegin, this.DateEnd, "AVO%"));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteDevDayKesi(Object parameter)
    {
      return true;
    }

    //--Для отчета Фадеева ОТК "Качество ЭАС на АВО" 
    public ICommand OtkQualityAvoCommand
    {
      get { return otkQualityAvoCommand ?? (otkQualityAvoCommand = new DelegateCommand<Object>(ExecuteOtkQualityAvo, CanExecuteOtkQualityAvo)); }
    }

    private void ExecuteOtkQualityAvo(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkQualityAvoSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkQualityAvoDest;
      var rptParam = new Db.OtkQualityAvoRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1,
        IsRpt2 = Convert.ToBoolean(parameter)
      };

      var sp = new Db.OtkQualityAvo();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkQualityAvo(Object parameter)
    {
      return true;
    }

    //--Для отчета Фадеева ОТК "Дефекты АВО" 
    public ICommand OtkDefectAvoCommand
    {
      get { return otkDefectAvoCommand ?? (otkDefectAvoCommand = new DelegateCommand<Object>(ExecuteOtkDefectAvo, CanExecuteOtkDefectAvo)); }
    }

    private void ExecuteOtkDefectAvo(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkDefectAvoSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkDefectAvoDest;
      var rptParam = new Db.OtkDefectAvoRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]), 
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.OtkDefectAvo();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkDefectAvo(Object parameter)
    {
      return true;
    }

    //--Для отчета Фадеева ОТК "Дефекты АВО годовые" 
    public ICommand OtkDefectAvoYearCommand
    {
      get { return otkDefectAvoYearCommand ?? (otkDefectAvoYearCommand = new DelegateCommand<Object>(ExecuteOtkDefectAvoYear, CanExecuteOtkDefectAvoYear)); }
    }

    private void ExecuteOtkDefectAvoYear(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkDefYearAvoSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkDefYearAvoDest;
      var rptParam = new Db.OtkDefYearAvoRptParam(src, dst, this.DateBegin, this.DateEnd)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        //DateBegin = this.DateBegin,
        //DateEnd = this.DateEnd,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.OtkDefYearAvo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkDefectAvoYear(Object parameter)
    {
      return true;
    }

    //--Для отчета Фадеева ОТК "Дефект 501..." 
    public ICommand OtkDefect501Command
    {
      get { return otkDefect501Command ?? (otkDefect501Command = new DelegateCommand<Object>(ExecuteOtkDefect501, CanExecuteOtkDefect501)); }
    }

    private void ExecuteOtkDefect501(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.OtkDefect501Source;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkDefect501Dest;
      var rptParam = new Db.OtkDefect501RptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        Glubina = this.Glubina,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.OtkDefect501();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkDefect501(Object parameter)
    {
      return true;
    }

    public ICommand OtkDistrib501OnLengthCommand
    {
      get { return otkDistrib501OnLengthCommand ?? (otkDistrib501OnLengthCommand = new DelegateCommand<Object>(ExecuteOtkDistrib501OnLength, CanExecuteOtkDistrib501OnLength)); }
    }

    private void ExecuteOtkDistrib501OnLength(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.OtkDistrib501OnLengthSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkDistrib501OnLengthDest;
      var rptParam = new Db.Distrib501OnLengthRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        Glubina = this.Glubina,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.Distrib501OnLength();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkDistrib501OnLength(Object parameter)
    {
      return true;
    }


    //--Для отчета Кутепова ОТК "Ширина по бригадам на АПР1" 
    public ICommand OtkShirApr1Command
    {
      get { return оtkShirApr1Command ?? (оtkShirApr1Command = new DelegateCommand<Object>(ExecuteOtkShirApr1, CanExecuteOtkShirApr1)); }
    }

    private void ExecuteOtkShirApr1(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkShirApr1Source;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkShirApr1Dest;

      var sp = new Db.OtkShirApr1();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.OtkShirApr1RptParam(src, dst, this.DateBegin, this.DateEnd));

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkShirApr1(Object parameter)
    {
      return true;
    }

    //--Для отчета Фадеева ОТК "Надав ВТО" 
    public ICommand OtkNadavVtoCommand
    {
      get { return otkNadavVtoCommand ?? (otkNadavVtoCommand = new DelegateCommand<Object>(ExecuteOtkNadavVto, CanExecuteOtkNadavVto)); }
    }

    private void ExecuteOtkNadavVto(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkNadavVtoSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkNadavVtoDest;

      var rptParam = new Db.OtkNadavVtoRptParam(src, dst)
                       {
                         PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder, 
                         DateBegin = this.DateBegin,
                         DateEnd = this.DateEnd,
                         TypeFilter = this.SelectedTabIndexF1,
                         ListStendF1 = this.ListStendF1,
                         TypeListValueF1 = this.typeListValueF1,
                         Glubina = this.Glubina,
                         Defect = this.Defect,
                         IsAroF1 = this.IsAroF1,
                         Is1200F1 = this.Is1200F1,
                         IsApr1F1 = this.IsApr1F1,
                         IsAooF1 = this.IsAooF1,
                         IsVtoF1 = this.IsVtoF1,
                         IsAvoF1 = this.IsAvoF1,
                         IsMgOF1 =  this.IsMgOF1,
                         IsPppF1 = this.IsPppF1,
                         IsWgtCoverF1 = this.IsWgtCoverF1,
                         IsStVtoF1 = this.IsStVtoF1,
                         IsKlpVtoF1  = this.IsKlpVtoF1,
                         IsDiskVtoF1 = this.IsDiskVtoF1,
                         IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
                         IsDateAroF1 = this.IsDateAroF1,
                         IsDate1200F1 = this.IsDate1200F1,
                         IsDateApr1F1 = this.IsDateApr1F1,
                         IsDateAooF1 = this.IsDateAooF1,
                         DateBeginAroF1 = this.DateBeginAroF1,
                         DateEndAroF1 = this.DateEndAroF1,
                         DateBegin1200F1 = this.DateBegin1200F1,
                         DateEnd1200F1 = this.DateEnd1200F1,
                         DateBeginApr1F1 = this.DateBeginApr1F1,
                         DateEndApr1F1 = this.DateEndApr1F1,
                         DateBeginAooF1 = this.DateBeginAooF1,
                         DateEndAooF1 = this.DateEndAooF1,
                         AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
                         Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
                         TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
                         BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
                         BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
                         BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
                         AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
                         AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
                         ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
                         DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
                         //Apr1F1Width = this.Apr1F1Width,
                         AooF1MgOFrom = this.AooF1MgOFrom,
                         AooF1MgOTo = this.AooF1MgOTo,
                         AooF1PppFrom = this.AooF1PppFrom,
                         AooF1PppTo = this.AooF1PppTo,
                         AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
                         AooF1WgtCoverTo = this.AooF1WgtCoverTo,
                         VtoF1Stend = this.VtoF1Stend,
                         VtoF1Cap = this.VtoF1Cap,
                         VtoF1TimeAooVto = this.VtoF1TimeAooVto,
                         IsDateAvoLstF1 = this.IsDateAvoLstF1,
                         DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
                         DateEndAvoLstF1 = this.DateEndAvoLstF1,
                         BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
                         IsBrgAooF1 = this.IsBrgAooF1
                       };

      var sp = new Db.OtkNadavVto();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkNadavVto(Object parameter)
    {
      return true;
    }
    
    public ICommand OtkDistribDefectsOnLengthCommand
    {
      get { return otkDistribDefectsOnLengthCommand ?? (otkDistribDefectsOnLengthCommand = new DelegateCommand<Object>(ExecuteOtkDistribDefectsOnLength, CanExecuteOtkDistribDefectsOnLength)); }
    }

    private void ExecuteOtkDistribDefectsOnLength(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkDistribDefectsOnLengthSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkDistribDefectsOnLengthDest;

      var rptParam = new Db.DistribDefectsOnLengthRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        Defect = this.Defect,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.DistribDefectsOnLength();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkDistribDefectsOnLength(Object parameter)
    {
      return true;
    }

    public ICommand OtkDistribDefectsOnSurfaceCommand
    {
      get { return otkDistribDefectsOnSurfaceCommand ?? (otkDistribDefectsOnSurfaceCommand = new DelegateCommand<Object>(ExecuteOtkDistribDefectsOnSurface, CanExecuteOtkDistribDefectsOnSurface)); }
    }

    private void ExecuteOtkDistribDefectsOnSurface(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkDistribDefectsOnSurfaceSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkDistribDefectsOnSurfaceDest;

      var rptParam = new Db.DistribDefectsOnSurfaceRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        Defect = this.Defect,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.DistribDefectsOnSurface();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkDistribDefectsOnSurface(Object parameter)
    {
      return true;
    }

    public ICommand OtkFreqDistrDefectAvoCommand
    {
      get { return otkFreqDistrDefectAvoCommand ?? (otkFreqDistrDefectAvoCommand = new DelegateCommand<Object>(ExecuteOtkFreqDistrDefectAvo, CanExecuteOtkFreqDistrDefectAvo)); }
    }

    private void ExecuteOtkFreqDistrDefectAvo(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkFreqDistrDefectAvoSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkFreqDistrDefectAvoDest;

      var rptParam = new Db.OtkFreqDistrDefectAvoRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF1,
        ListStendF1 = this.ListStendF1,
        TypeListValueF1 = this.typeListValueF1,
        Defect = this.Defect,
        IsAroF1 = this.IsAroF1,
        Is1200F1 = this.Is1200F1,
        IsApr1F1 = this.IsApr1F1,
        IsAooF1 = this.IsAooF1,
        IsVtoF1 = this.IsVtoF1,
        IsAvoF1 = this.IsAvoF1,
        IsMgOF1 = this.IsMgOF1,
        IsPppF1 = this.IsPppF1,
        IsWgtCoverF1 = this.IsWgtCoverF1,
        IsStVtoF1 = this.IsStVtoF1,
        IsKlpVtoF1 = this.IsKlpVtoF1,
        IsDiskVtoF1 = this.IsDiskVtoF1,
        IsTimeAooVtoF1 = this.IsTimeAooVtoF1,
        IsDateAroF1 = this.IsDateAroF1,
        IsDate1200F1 = this.IsDate1200F1,
        IsDateApr1F1 = this.IsDateApr1F1,
        IsDateAooF1 = this.IsDateAooF1,
        DateBeginAroF1 = this.DateBeginAroF1,
        DateEndAroF1 = this.DateEndAroF1,
        DateBegin1200F1 = this.DateBegin1200F1,
        DateEnd1200F1 = this.DateEnd1200F1,
        DateBeginApr1F1 = this.DateBeginApr1F1,
        DateEndApr1F1 = this.DateEndApr1F1,
        DateBeginAooF1 = this.DateBeginAooF1,
        DateEndAooF1 = this.DateEndAooF1,
        AroF1Item = Convert.ToString(this.SelAroF1Item.Row["StrSql"]),
        Stan1200F1Item = Convert.ToString(this.sel1200F1Item.Row["StrSql"]),
        TolsF1Item = Convert.ToString(this.SelTolsF1Item.Row["StrSql"]),
        BrgApr1F1Item = Convert.ToString(this.SelBrgApr1F1Item.Row["StrSql"]),
        BrgVtoF1Item = Convert.ToString(this.SelBrgVtoF1Item.Row["StrSql"]),
        BrgAvoF1Item = Convert.ToString(this.SelBrgAvoF1Item.Row["StrSql"]),
        AooF1Item = Convert.ToString(this.SelAooF1Item.Row["StrSql"]),
        AvoF1Item = Convert.ToString(this.SelAvoF1Item.Row["StrSql"]),
        ShirApr1F1Item = Convert.ToString(this.SelShirApr1F1Item.Row["StrSql"]),
        DiskVtoF1Item = Convert.ToString(this.SelDiskVtoF1Item.Row["StrSql"]),
        //Apr1F1Width = this.Apr1F1Width,
        AooF1MgOFrom = this.AooF1MgOFrom,
        AooF1MgOTo = this.AooF1MgOTo,
        AooF1PppFrom = this.AooF1PppFrom,
        AooF1PppTo = this.AooF1PppTo,
        AooF1WgtCoverFrom = this.AooF1WgtCoverFrom,
        AooF1WgtCoverTo = this.AooF1WgtCoverTo,
        VtoF1Stend = this.VtoF1Stend,
        VtoF1Cap = this.VtoF1Cap,
        VtoF1TimeAooVto = this.VtoF1TimeAooVto,
        IsDateAvoLstF1 = this.IsDateAvoLstF1,
        DateBeginAvoLstF1 = this.DateBeginAvoLstF1,
        DateEndAvoLstF1 = this.DateEndAvoLstF1,
        BrgAooF1Item = Convert.ToString(this.SelBrgAooF1Item.Row["StrSql"]),
        IsBrgAooF1 = this.IsBrgAooF1
      };

      var sp = new Db.OtkFreqDistrDefectAvo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkFreqDistrDefectAvo(Object parameter)
    {
      return true;
    }

    //--Для отчета Чепова ОТК "Выход металла 1 класса, 1 сорта" 
    public ICommand OtkOutMe1Cls1SrtCommand
    {
      get { return otkOutMe1Cls1SrtCommand ?? (otkOutMe1Cls1SrtCommand = new DelegateCommand<Object>(ExecuteOtkOutMe1Cls1Srt, CanExecuteOtkOutMe1Cls1Srt)); }
    }

    private void ExecuteOtkOutMe1Cls1Srt(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkOutMe1Cls1SrtSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkOutMe1Cls1SrtDest;
      var rptParam = new Db.OtkOutMe1Cls1SrtParam(src, dst)
                       {
                         DateBegin = this.DateBegin,
                         DateEnd = this.DateEnd
                       };

      var sp = new Db.OtkOutMe1Cls1Srt();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkOutMe1Cls1Srt(Object parameter)
    {
      return true;
    }
    
    //--Для отчета Чепова ОТК "Выход металла шириной 1000 мм." 
    public ICommand OtkOutMeWdt1000Command
    {
      get { return otkOutMeWdt1000Command ?? (otkOutMeWdt1000Command = new DelegateCommand<Object>(ExecuteOtkOutMeWdt1000, CanExecuteOtkOutMeWdt1000)); }
    }

    private void ExecuteOtkOutMeWdt1000(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkOutMeWdt1000Source;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkOutMeWdt1000Dest;
      var rptParam = new Db.OutMeWdt1000Param(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.OutMeWdt1000();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkOutMeWdt1000(Object parameter)
    {
      return true;
    }
    
    //--Для отчета Чепова ОТК "Выход металла шириной 1000 мм." 
    public ICommand OtkTo2SortCommand
    {
      get { return otkTo2SortCommand ?? (otkTo2SortCommand = new DelegateCommand<Object>(ExecuteOtkTo2Sort, CanExecuteOtkTo2Sort)); }
    }

    private void ExecuteOtkTo2Sort(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkTo2SortSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkTo2SortDest;
      var rptParam = new Db.To2SortParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        Thickness = Convert.ToString(this.SelTolsItem.Row["StrSql"]),
        FilterThickness = Convert.ToString(this.SelTolsItem.Row["StrDlg"])
      };

      var sp = new Db.To2Sort();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkTo2Sort(Object parameter)
    {
      return true;
    }

    public ICommand OtkInfoShovProryvCommand
    {
      get { return otkInfoShovProryvCommand ?? (otkInfoShovProryvCommand = new DelegateCommand<Object>(ExecuteOtkInfoRkShovProryv, CanExecuteOtkInfoRkShovProryv)); }
    }

    private void ExecuteOtkInfoRkShovProryv(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkInfoShovPoryvSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkInfoShovPoryvDest;
      var rptParam = new Db.InfoShovProryvRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
      };

      var sp = new Db.InfoShovProryv();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkInfoRkShovProryv(Object parameter)
    {
      return true;
    }


    public ICommand OtkSgpDefectsCommand
    {
      get { return otkSgpDefectsCommand ?? (otkSgpDefectsCommand = new DelegateCommand<Object>(ExecuteOtkSgpDefects, CanExecuteOtkSgpDefects)); }
    }

    private void ExecuteOtkSgpDefects(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkSgpDefectsSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkSgpDefectsDest;
      var rptParam = new Db.SgpDefectsRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
      };

      var sp = new Db.SgpDefects();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkSgpDefects(Object parameter)
    {
      return true;
    }

    public ICommand OtkSgpDefectsSort1GostCommand
    {
      get { return otkSgpDefectsSort1GostCommand ?? (otkSgpDefectsSort1GostCommand = new DelegateCommand<Object>(ExecuteOtkSgpDefectsSort1Gost, CanExecuteOtkSgpDefectsSort1Gost)); }
    }

    private void ExecuteOtkSgpDefectsSort1Gost(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkSgpDefectsSort1GostSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkSgpDefectsSort1GostDest;
      var rptParam = new Db.SgpDefectsSort1GostRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
      };

      var sp = new Db.SgpDefectsSort1Gost();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkSgpDefectsSort1Gost(Object parameter)
    {
      return true;
    }

    public ICommand OtkSeqCoilLineAooCommand
    {
      get { return otkSeqCoilLineAooCommand ?? (otkSeqCoilLineAooCommand = new DelegateCommand<Object>(ExecuteOtkSeqCoilLineAoo, CanExecuteOtkSeqCoilLineAoo)); }
    }

    private void ExecuteOtkSeqCoilLineAoo(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkSeqCoilLineAooSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkSeqCoilLineAooDest;
      var rptParam = new Db.SeqCoilLineAooRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
      };

      var sp = new Db.SeqCoilLineAoo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkSeqCoilLineAoo(Object parameter)
    {
      return true;
    }

    public ICommand OtkChratcerListCoilsCommand
    {
      get { return otkChratcerListCoilsCommand ?? (otkChratcerListCoilsCommand = new DelegateCommand<Object>(ExecuteOtkChratcerListCoils, CanExecuteOtkChratcerListCoils)); }
    }

    private void ExecuteOtkChratcerListCoils(Object parameter)
    {
      var src = Smv.Utils.Etc.StartPath + ModuleConst.OtkChratcerListCoilsSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.OtkChratcerListCoilsDest;
      var rptParam = new Db.ChratcerListCoilsRptParam(src, dst)
      {
        ListCoils = ListStendF1
      };

      var sp = new Db.ChratcerListCoils();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkChratcerListCoils(Object parameter)
    {
      return true;
    }
    
    public ICommand OtkFinCutByCatCommand
    {
      get { return otkFinCutByCatCommand ?? (otkFinCutByCatCommand = new DelegateCommand<Object>(ExecuteOtkFinCutByCat, CanExecuteOtkFinCutByCat)); }
    }

    private void ExecuteOtkFinCutByCat(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.FinCutByCatSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.FinCutByCatDest;
      var rptParam = new Db.FinCutByCatRptParam(src, dst)
      {
        DateBegin = DateBegin,
        DateEnd = DateEnd
      };
      
      var sp = new Db.FinCutByCat();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteOtkFinCutByCat(Object parameter)
    {
      return true;
    }


    #endregion


  }
}
