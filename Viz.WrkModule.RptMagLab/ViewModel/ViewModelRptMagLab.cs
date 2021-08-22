using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.LayoutControl;
using Microsoft.Win32;
using Smv.MVVM.Commands;
using Smv.MVVM.ViewModels;
using Smv.Utils;
using Smv.Xls;
using Viz.DbApp.Psi;
using Viz.WrkModule.RptMagLab.Db;
using Viz.WrkModule.RptMagLab.Db.DataSets;

//using Devart.Data.Oracle;

namespace Viz.WrkModule.RptMagLab
{

  internal sealed class ViewModelRptMagLab : ViewModelBase
  {

    #region Fields
    private readonly XlsInstanceBackgroundReport rpt;
    private readonly UserControl usrControl; 
    private readonly Object param;
    private DateTime dateBegin;
    private DateTime dateEnd;
    private string techStep = "STRANN";
    private string alstCo;
    private string alstPco;
    private readonly LayoutGroup lg;
    private readonly DsRptMagLab dsRptMagLab = new DsRptMagLab();

    //--Для отчета Гомзикова ЦЗЛ "Средние эл. маг. х-тики ЭАС"
    private string listTxtValue41;
    private int typeQualityObj;
    private DataRowView selRm1200Item;
    private DataRowView selAroItem;
    private DataRowView selAooItem;
    private DataRowView selAvoItem;
    private DataRowView selAprItem;
    private DataRowView selSortItem;
    private DataRowView selClassPloskItem;
    //--Для отчета Гомзикова ЦЗЛ "Эффективность лазера"
    private decimal p1750023;
    private decimal p1750027;
    private decimal p1750030;
    private decimal b800;
    private decimal kesiAvg;
    private decimal coefVoln;
    private decimal qntShov;
    private string sort;
    private string listTxtValue51;
    private DataRowView selAdgInPrmItem;
    private DataRowView selAdgOutPrmItem;

    //Поля для Фильтра1 Начало
    private int selectedTabIndexF1;

    private Boolean isSortF1;
    private DataRowView selSortF1Item;
    private DataRowView selPlskF1Item;

    private Boolean is1200F1;
    private DataRowView sel1200F1Item;
    private Boolean isDate1200F1;
    private DateTime dateBegin1200F1;
    private DateTime dateEnd1200F1;

    private Boolean isAooF1;
    private DataRowView selAooF1Item;
    private Boolean isDateAooF1;
    private DateTime dateBeginAooF1;
    private DateTime dateEndAooF1;

    private Boolean isAroF1;
    private DataRowView selAroF1Item;
    private Boolean isDateAroF1;
    private DateTime dateBeginAroF1;
    private DateTime dateEndAroF1;

    private Boolean isAvoF1;
    private DataRowView selAvoF1Item;
    private Boolean isDateAvoF1;
    private DateTime dateBeginAvoF1;
    private DateTime dateEndAvoF1;

    private Boolean isVtoF1;
    private string stdVtoF1;

    private Boolean isAprF1;
    private DataRowView selAprF1Item;

    private int typeInclLstF1; //0-из списка, 1-исключая из списка   
    private int typeListValueF1; //0-№ Стенда, 1-№ Стенда ВТО   
    private string listValueF1;
    //Поля для Фильтра1 Конец

    //Поля для Сравнение линий АОО начало
    private string typeLineAoo = "AOO3";
    private int typeMat = 1;
    //Поля для Сравнение линий АОО конец

    #endregion Fields

    #region Public Property
    public DateTime DateBegin
    {
      get{ return dateBegin; }
      set{
        if (value == dateBegin) return;
        dateBegin = value;
        OnPropertyChanged("DateBegin");
      }
    }

    public DateTime DateEnd
    {
      get{ return dateEnd; }
      set{
        if (value == dateEnd) return;
        dateEnd = value;
        OnPropertyChanged("DateEnd");
      }
    }

    public string AlstCo
    {
      get{ return alstCo; }
      set{
        if (value == alstCo) return;
        alstCo = value;
        OnPropertyChanged("AlstCo");
      }
    }

    public string AlstPco
    {
      get{ return alstPco; }
      set{
        if (value == alstPco) return;
        alstPco = value;
        OnPropertyChanged("AlstPco");
      }
    }

    //--Для отчета Гомзикова ЦЗЛ "Средние эл. маг. х-тики ЭАС"
    public DataTable Rm1200Ts
    {
      get { return dsRptMagLab.Rm1200Ts; }
    }

    public DataTable AroTs
    {
      get { return dsRptMagLab.AroTs; }
    }

    public DataTable AooTs
    {
      get { return dsRptMagLab.AooTs; }
    }

    public DataTable AvoTs
    {
      get { return dsRptMagLab.AvoTs; }
    }

    public DataTable AprTs
    {
      get { return dsRptMagLab.AprTs; }
    }

    public DataTable SortTs
    {
      get { return dsRptMagLab.SortTs; }
    }

    public DataTable ClassPloskTs
    {
      get { return dsRptMagLab.ClassPloskTs; }
    }

    public DataRowView SelRm1200Item
    {
      get { return selRm1200Item; }
      set{
        if (value == selRm1200Item) return;
        selRm1200Item = value;
        OnPropertyChanged("SelRm1200Item");
      }
    }

    public DataRowView SelAroItem
    {
      get { return selAroItem; }
      set{
        if (value == selAroItem) return;
        selAroItem = value;
        OnPropertyChanged("SelAroItem");
      }
    }

    public DataRowView SelAooItem
    {
      get { return selAooItem; }
      set{
        if (value == selAooItem) return;
        selAooItem = value;
        OnPropertyChanged("SelAooItem");
      }
    }

    public DataRowView SelAvoItem
    {
      get { return selAvoItem; }
      set{
        if (value == selAvoItem) return;
        selAvoItem = value;
        OnPropertyChanged("SelAvoItem");
      }
    }

    public DataRowView SelAprItem
    {
      get { return selAprItem; }
      set{
        if (value == selAprItem) return;
        selAprItem = value;
        OnPropertyChanged("SelAprItem");
      }
    }

    public DataRowView SelSortItem
    {
      get { return selSortItem; }
      set{
        if (value == selSortItem) return;
        selSortItem = value;
        OnPropertyChanged("SelSortItem");
      }
    }

    //
    public DataRowView SelClassPloskItem
    {
      get { return selClassPloskItem; }
      set{
        if (value == selClassPloskItem) return;
        selClassPloskItem = value;
        OnPropertyChanged("SelClassPloskItem");
      }
    }

    public string ListTxtValue41
    {
      get { return listTxtValue41; }
      set{
        if (value == listTxtValue41) return;
        listTxtValue41 = value;
        OnPropertyChanged("ListTxtValue41");
      }
    }

    //--Для отчета Гомзикова ЦЗЛ "Эффективность лазера"
    public DataTable AdgInPrm
    {
      get { return dsRptMagLab.AdgInPrm; }
    }

    public DataTable AdgOutPrm
    {
      get { return dsRptMagLab.AdgOutPrm; }
    }

    public DataRowView SelAdgInPrmItem
    {
      get { return selAdgInPrmItem; }
      set{
        if (value == SelAdgInPrmItem) return;
        selAdgInPrmItem = value;
        OnPropertyChanged("SelAdgInPrmItem");
      }
    }

    public DataRowView SelAdgOutPrmItem
    {
      get { return selAdgOutPrmItem; }
      set{
        if (value == SelAdgOutPrmItem) return;
        selAdgOutPrmItem = value;
        OnPropertyChanged("SelAdgOutPrmItem");
      }
    }

    public decimal P1750023
    {
      get { return p1750023; }
      set{
        if (value == p1750023) return;
        p1750023 = value;
        OnPropertyChanged("P1750023");
      }
    }

    public decimal P1750027
    {
      get { return p1750027; }
      set{
        if (value == p1750027) return;
        p1750027 = value;
        OnPropertyChanged("P1750027");
      }
    }

    public decimal P1750030
    {
      get { return p1750030; }
      set{
        if (value == p1750030) return;
        p1750030 = value;
        OnPropertyChanged("P1750030");
      }
    }

    public decimal B800
    {
      get { return b800; }
      set{
        if (value == b800) return;
        b800 = value;
        OnPropertyChanged("B800");
      }
    }

    public decimal KesiAvg
    {
      get { return kesiAvg; }
      set{
        if (value == kesiAvg) return;
        kesiAvg = value;
        OnPropertyChanged("KesiAvg");
      }
    }

    public decimal CoefVoln
    {
      get { return coefVoln; }
      set{
        if (value == coefVoln) return;
        coefVoln = value;
        OnPropertyChanged("CoefVoln");
      }
    }

    public decimal QntShov
    {
      get { return qntShov; }
      set{
        if (value == qntShov) return;
        qntShov = value;
        OnPropertyChanged("QntShov");
      }
    }

    public string Sort
    {
      get { return sort; }
      set{
        if (value == sort) return;
        sort = value;
        OnPropertyChanged("Sort");
      }
    }

    public string ListTxtValue51
    {
      get { return listTxtValue51; }
      set{
        if (value == listTxtValue51) return;
        listTxtValue51 = value;
        OnPropertyChanged("ListTxtValue51");
      }
    }

    //Поля для Фильтра1
    public int SelectedTabIndexF1
    {
      get { return selectedTabIndexF1; }
      set{
        if (value == selectedTabIndexF1) return;
        selectedTabIndexF1 = value;
        OnPropertyChanged("SelectedTabIndexF1");
      }
    }

    public Boolean IsSortF1
    {
      get { return isSortF1; }
      set{
        if (value == isSortF1) return;
        isSortF1 = value;
        OnPropertyChanged("IsSortF1");
      }
    }

    public DataRowView SelSortF1Item
    {
      get { return selSortF1Item; }
      set{
        if (value == selSortF1Item) return;
        selSortF1Item = value;
        OnPropertyChanged("SelSortF1Item");
      }
    }

    public DataRowView SelPlskF1Item
    {
      get { return selPlskF1Item; }
      set{
        if (value == selPlskF1Item) return;
        selPlskF1Item = value;
        OnPropertyChanged("SelPlskF1Item");
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

    public DataRowView Sel1200F1Item
    {
      get { return sel1200F1Item; }
      set{
        if (value == sel1200F1Item) return;
        sel1200F1Item = value;
        OnPropertyChanged("Sel1200F1Item");
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

    public Boolean IsAooF1
    {
      get { return isAooF1; }
      set{
        if (value == isAooF1) return;
        isAooF1 = value;
        if (!value) IsDateAooF1 = false;
        OnPropertyChanged("IsAooF1");
      }
    }

    public DataRowView SelAooF1Item
    {
      get { return selAooF1Item; }
      set{
        if (value == selAooF1Item) return;
        selAooF1Item = value;
        OnPropertyChanged("SelAooF1Item");
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

    public DataRowView SelAroF1Item
    {
      get { return selAroF1Item; }
      set{
        if (value == selAroF1Item) return;
        selAroF1Item = value;
        OnPropertyChanged("SelAroF1Item");
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

    public Boolean IsAvoF1
    {
      get { return isAvoF1; }
      set{
        if (value == isAvoF1) return;
        isAvoF1 = value;
        if (!value) IsDateAvoF1 = false;
        OnPropertyChanged("IsAvoF1");
      }
    }

    public DataRowView SelAvoF1Item
    {
      get { return selAvoF1Item; }
      set{
        if (value == selAvoF1Item) return;
        selAvoF1Item = value;
        OnPropertyChanged("SelAvoF1Item");
      }
    }

    public Boolean IsDateAvoF1
    {
      get { return isDateAvoF1; }
      set{
        if (value == isDateAvoF1) return;
        isDateAvoF1 = value;
        OnPropertyChanged("IsDateAvoF1");
      }
    }

    public DateTime DateBeginAvoF1
    {
      get { return dateBeginAvoF1; }
      set{
        if (value == dateBeginAvoF1) return;
        dateBeginAvoF1 = value;
        OnPropertyChanged("DateBeginAvoF1");
      }
    }

    public DateTime DateEndAvoF1
    {
      get { return dateEndAvoF1; }
      set{
        if (value == dateEndAvoF1) return;
        dateEndAvoF1 = value;
        OnPropertyChanged("DateEndAvoF1");
      }
    }

    public Boolean IsVtoF1
    {
      get { return isVtoF1; }
      set{
        if (value == isVtoF1) return;
        isVtoF1 = value;
        OnPropertyChanged("IsVtoF1");
      }
    }

    public string StdVtoF1
    {
      get { return stdVtoF1; }
      set{
        if (value == stdVtoF1) return;
        stdVtoF1 = value;
        OnPropertyChanged("StdVtoF1");
      }
    }

    public Boolean IsAprF1
    {
      get { return isAprF1; }
      set{
        if (value == isAprF1) return;
        isAprF1 = value;
        OnPropertyChanged("IsAprF1");
      }
    }

    public DataRowView SelAprF1Item
    {
      get { return selAprF1Item; }
      set{
        if (value == selAprF1Item) return;
        selAprF1Item = value;
        OnPropertyChanged("SelAprF1Item");
      }
    }

    public string ListValueF1
    {
      get { return listValueF1; }
      set{
        if (value == listValueF1) return;
        listValueF1 = value;
        OnPropertyChanged("ListValueF1");
      }
    }


    #endregion Public Property

    #region Private Method
    private void RunXlsRptCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      //pgbEdit.StyleSettings = new DevExpress.Xpf.Editors.ProgressBarStyleSettings();
      //((param as BarEditItem).EditSettings as ProgressBarEditSettings).StyleSettings = new DevExpress.Xpf.Editors.ProgressBarMarqueeStyleSettings(); 
      if (param is BarEditItem barEditItem) barEditItem.IsVisible = false;
    }

    private void LayoutGroupExpanded(object sender, EventArgs e)
    {
      var layoutGroup = sender as LayoutGroup;
      if (layoutGroup != null){
        var i = Convert.ToInt32(layoutGroup.Tag);
        switch (i){
          case 0:
            break;
          case 1:
            break;
          case 2:
            break;
          case 3:
            break;
          case 4:
            LgExpandedQualitySteel();
            break;
          case 5:
            LgExpandedEfLsr();
            break;
          case 6:
            LgExpandedF1();
            break;
          default:
            break;
        }
      }
    }

    private void LayoutGroupCollapsed(object sender, EventArgs e)
    {
      var layoutGroup = sender as LayoutGroup;
      if (layoutGroup != null){
        var i = Convert.ToInt32(layoutGroup.Tag);
        switch (i)
        {
          case 0:
            break;
          case 1:
            break;
          case 2:
            break;
          case 3:
            break;
          case 4:
            LgCollapsedQualitySteel();
            break;
          default:
            break;
        }
      }
    }

    private void LgExpandedQualitySteel()
    {
      dsRptMagLab.Rm1200Ts.LoadData(1);
      dsRptMagLab.AroTs.LoadData(2);
      dsRptMagLab.AooTs.LoadData(3);
      dsRptMagLab.AvoTs.LoadData(4);
      dsRptMagLab.AprTs.LoadData(5);
      dsRptMagLab.SortTs.LoadData(8);
      dsRptMagLab.ClassPloskTs.LoadData(9);

      for (int i = 0; i < 7; i++){
        var cbe = LogicalTreeHelper.FindLogicalNode(usrControl, "CbeTypeTs" + i) as ComboBoxEdit;
        if (cbe != null)
          cbe.SelectedIndex = 0;
      }       
    }

    private void LgExpandedEfLsr()
    {
      dsRptMagLab.AdgInPrm.LoadData(6);
      dsRptMagLab.AdgOutPrm.LoadData(7);

      var cbe = LogicalTreeHelper.FindLogicalNode(usrControl, "CbeAdgIn_5") as ComboBoxEdit;
      if (cbe != null)
        cbe.SelectedIndex = 0;

      cbe = LogicalTreeHelper.FindLogicalNode(usrControl, "CbeAdgOut_5") as ComboBoxEdit;
      if (cbe != null)
        cbe.SelectedIndex = 0;
    }

    private void LgExpandedF1()
    {
      dsRptMagLab.Rm1200Ts.LoadData(1);
      dsRptMagLab.AroTs.LoadData(2);
      dsRptMagLab.AooTs.LoadData(3);
      dsRptMagLab.AvoTs.LoadData(4);
      dsRptMagLab.AprTs.LoadData(5);
      dsRptMagLab.SortTs.LoadData(8);
      dsRptMagLab.ClassPloskTs.LoadData(9);

      for (int i = 0; i < 7; i++){
        var cbe = LogicalTreeHelper.FindLogicalNode(usrControl, "CbeF1_" + i) as ComboBoxEdit;
        if (cbe != null)
          cbe.SelectedIndex = 0;
      }

      DateBegin1200F1 = DateEnd1200F1 = DateBeginAooF1 = DateEndAooF1 = DateBeginAroF1 = DateEndAroF1
       = DateBeginAvoF1 = DateEndAvoF1 = DateTime.Today;
    }


    private void LgCollapsedQualitySteel()
    {
    }

    private string GetStringFromTxtFile()
    {
      var ofd = new OpenFileDialog {DefaultExt = ".txt", Filter = "text format (.txt)|*.txt"};
      bool? result = ofd.ShowDialog();
      if (result != true) 
        return string.Empty;
      return  File.ReadAllText(ofd.FileName, Encoding.GetEncoding(1251)).Replace(" ", "").Replace("\r\n", " ").Trim().Replace(" ", ",");
    }

    void SelectedTypeFilter(object sender, ValueChangedEventArgs<FrameworkElement> e)
    {
      SelectedTabIndexF1 = (sender as LayoutGroup).SelectedTabIndex;
    }

    #endregion Private Method

    #region Constructor
    internal ViewModelRptMagLab(UserControl control, Object Param)
    {
      param = Param;
      rpt = new XlsInstanceBackgroundReport();
      usrControl = control;
      DateBegin = DateTime.Today;
      DateEnd = DateTime.Today;

      //Группы 1-уровня
      for (int i = ModuleConst.AccGrpMatTechStep; i < ModuleConst.AccGrpCzlPlosk + 1; i++){
        lg = LogicalTreeHelper.FindLogicalNode(usrControl, "Lg" + ModuleConst.ModuleId + "_" + i) as LayoutGroup;

        if (lg != null){
          if (!Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            lg.Visibility = Visibility.Hidden;
            continue;
          }

          lg.Expanded += LayoutGroupExpanded;
          lg.Collapsed += LayoutGroupCollapsed;
        }
      }

      //Группы 2-уровня
      for (int i = ModuleConst.AccG2Plosk; i < ModuleConst.AccG2Plosk + 1; i++){
        var uie = LogicalTreeHelper.FindLogicalNode(usrControl, "G2" + ModuleConst.ModuleId + "_" + i) as UIElement;

        if (uie != null){
          if (!Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            uie.Visibility = Visibility.Hidden;
          }
        }
      }

      //Делаем controls невидимыми
      for (int i = ModuleConst.AccCmdMatTechStep; i < ModuleConst.AccCmdCzlEfLsr9t_NoLst + 1; i++)
      {
        var btn = LogicalTreeHelper.FindLogicalNode(usrControl, "b" + ModuleConst.ModuleId + "_" + i) as UIElement;

        if (btn == null) continue;

        if (!Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
          btn.Visibility = Visibility.Hidden;
        }
      }

      var layoutGroup = LogicalTreeHelper.FindLogicalNode(usrControl, "lgFilter1") as LayoutGroup;
      if (layoutGroup != null)
        layoutGroup.SelectedTabChildChanged += SelectedTypeFilter;


      //DxInfo.ShowDxBoxInfo("Внимание", "Файлы созданных отчетов будут находиться в папке Документы (Мои документы)!", MessageBoxImage.Information);
    }
    #endregion Constructor

    #region Commands
    private DelegateCommand<Object> showListRptCommand;
    private DelegateCommand<Object> selectTechStepCommand;
    private DelegateCommand<Object> selectTypeLineAooCommand;
    private DelegateCommand<Object> selectTypeMatCommand;
    private DelegateCommand<Object> materialStepRptCommand;
    private DelegateCommand<Object> alstIsolRptCommand;
    private DelegateCommand<Object> czlCommand;
    private DelegateCommand<Object> czlLaserCommand;
    private DelegateCommand<Object> czlIsoGoCommand;
    private DelegateCommand<Object> czlFinCutCommand;
    private DelegateCommand<Object> loadFromTxtFileCommand;

    //--Для отчета Гомзикова ЦЗЛ "Средние эл. маг. х-тики ЭАС"
    private DelegateCommand<Object> selectTypeQualObjCommand;
    private DelegateCommand<Object> runQczlReportCommand;
    //--Для отчета Гомзикова ЦЗЛ "Эффективность лазера"
    private DelegateCommand<Object> runCzlEfLsrCommand;
    private DelegateCommand<Object> runCzlEfLsr9tCommand;
   //--Для отчета Гомзикова ЦЗЛ "Выход ЭАС с высоким уровнем магнитных свойств" 
    private DelegateCommand<Object> czlOutP1750Command;
    //--Для отчетов с учетом фильтра1
    private DelegateCommand<Object> selectTypeListValueF1Command;
    private DelegateCommand<Object> selectTypeInclListF1Command;
    private DelegateCommand<Object> czlPloskF1Command;
    private DelegateCommand<Object> czlDefPloskF1Command;
    //--Для отчетов Линии АОО
    private DelegateCommand<Object> czlLineAooCommand;

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
        Source = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.RptMagLab;Component/Images/BarImage.png"))
      };

      ListRpt.ShowListRpt(ModuleConst.ModuleId, im);
    }

    private bool CanExecuteShowListRpt(Object parameter)
    {
      return true;
    } 



   public ICommand SelectTechStepCommand
    {
      get{ return selectTechStepCommand ?? (selectTechStepCommand = new DelegateCommand<Object>(ExecuteSelectTechStep, CanExecuteSelectTechStep));}
    }

    private void ExecuteSelectTechStep(Object parameter)
    {
      techStep = Convert.ToString(parameter);
    }

    private bool CanExecuteSelectTechStep(Object parameter)
    {
      return true;
    }

    public ICommand SelectTypeLineAooCommand
    {
      get { return selectTypeLineAooCommand ?? (selectTypeLineAooCommand = new DelegateCommand<Object>(ExecuteSelectTypeLineAoo, CanExecuteSelectTypeLineAoo)); }
    }

    private void ExecuteSelectTypeLineAoo(Object parameter)
    {
      typeLineAoo = Convert.ToString(parameter);
    }

    private bool CanExecuteSelectTypeLineAoo(Object parameter)
    {
      return true;
    }

    public ICommand SelectTypeMatCommand
    {
      get { return selectTypeMatCommand ?? (selectTypeMatCommand = new DelegateCommand<Object>(ExecuteSelectTypeMat, CanExecuteSelectTypeMat)); }
    }

    private void ExecuteSelectTypeMat(Object parameter)
    {
      typeMat = Convert.ToInt32(parameter);
    }

    private bool CanExecuteSelectTypeMat(Object parameter)
    {
      return true;
    }

    public ICommand MaterialStepRptCommand
    {
      get{ return materialStepRptCommand ?? (materialStepRptCommand = new DelegateCommand<Object>(ExecuteMaterialStepRpt, CanExecuteMaterialStepRpt));}
    }

    private void ExecuteMaterialStepRpt(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.MatTechStepSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.MatTechStepDest;      
      
      string StepIl;
      string StepPj;

      if (techStep == "FINISHED_GOODS"){
        StepIl = "FINISHED_GOODS";
        StepPj = "PACK";
      }
      else{
        StepIl = techStep;
        StepPj = techStep;
      }


      var sp = new MaterialStepRpt();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new RptMagLabParam(src, dst, DateBegin, DateEnd, StepIl, StepPj));
      if (res){
        var barEditItem = param as BarEditItem;
        if (barEditItem != null) barEditItem.IsVisible = true;
      }
    }

    private bool CanExecuteMaterialStepRpt(Object parameter)
    {
      return true;
    }

    public ICommand AlstIsolRptCommand
    {
      get{return alstIsolRptCommand ?? (alstIsolRptCommand = new DelegateCommand<Object>(ExecuteAlstIsolRpt, CanExecuteAlstIsolRpt));}
    }

    private void ExecuteAlstIsolRpt(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.AlstIsolSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.AlstIsolDest;

      var sp = new AlstomIsol();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new AlstIsolRptParam(src, dst, AlstCo, AlstPco));
      if (res)
        (param as BarEditItem).IsVisible = true;
    }

    private bool CanExecuteAlstIsolRpt(Object parameter)
    {
      return true;
    }
    
    public ICommand CzlCommand
    {
      get{return czlCommand ?? (czlCommand = new DelegateCommand<Object>(ExecuteCzl, CanExecuteCzl));}
    }

    private void ExecuteCzl(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.CzlPackSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlPackDest;

      var sp = new CzlPack();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new CzlPackRptParam(src, dst, DateBegin, DateEnd, "FINISHED_GOODS"/*, "PACK"*/));
      if (res)
        (param as BarEditItem).IsVisible = true;
    }

    private bool CanExecuteCzl(Object parameter)
    {
      return true;
    }

    public ICommand CzlLaserCommand
    {
      get{ return czlLaserCommand ?? (czlLaserCommand = new DelegateCommand<Object>(ExecuteCzlLaser, CanExecuteCzlLaser));}
    }

    private void ExecuteCzlLaser(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.CzlLaserSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlLaserDest;

      var sp = new CzlLaser();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new CzlLaserRptParam(src, dst, DateBegin, DateEnd, "LASSCR", "LASSCR"));
      if (res)
        (param as BarEditItem).IsVisible = true;
    }

    private bool CanExecuteCzlLaser(Object parameter)
    {
      return true;
    }

    public ICommand CzlFinCutCommand
    {
      get{return czlFinCutCommand ?? (czlFinCutCommand = new DelegateCommand<Object>(ExecuteCzlFinCut, CanExecuteCzlFinCut));}
    }

    private void ExecuteCzlFinCut(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.CzlFinCutSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlFinCutDest;

      var sp = new CzlFinCut();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new CzlFinCutRptParam(src, dst, DateBegin, DateEnd, "FINCUT", "FINCUT"));
      if (res)
        (param as BarEditItem).IsVisible = true;
    }

    private bool CanExecuteCzlFinCut(Object parameter)
    {
      return true;
    }

    public ICommand CzlIsoGoCommand
    {
      get{return czlIsoGoCommand ?? (czlIsoGoCommand = new DelegateCommand<Object>(ExecuteCzlIsoGo, CanExecuteCzlIsoGo));}
    }

    private void ExecuteCzlIsoGo(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.CzlIsoGoSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlIsoGoDest;

      var sp = new CzlIsoGo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new CzlIsoGoRptParam(src, dst, DateBegin, DateEnd, "ISOGO", "ISOGO"));
      if (res)
        (param as BarEditItem).IsVisible = true;
    }

    private bool CanExecuteCzlIsoGo(Object parameter)
    {
      return true;
    }

    //--Для отчета Гомзикова ЦЗЛ "Средние эл. маг. х-тики ЭАС"
    public ICommand SelectTypeQualObjCommand
    {
      get{return selectTypeQualObjCommand ?? (selectTypeQualObjCommand = new DelegateCommand<Object>(ExecuteSelectTypeQualObj, CanExecuteSelectTypeQualObj));}
    }

    private void ExecuteSelectTypeQualObj(Object parameter)
    {
      typeQualityObj = Convert.ToInt32(parameter);
    }

    private bool CanExecuteSelectTypeQualObj(Object parameter)
    {
      return true;
    }

    public ICommand LoadFromTxtFileCommand
    {
      get{return loadFromTxtFileCommand ?? (loadFromTxtFileCommand = new DelegateCommand<Object>(ExecuteLoadFromTxtFile, CanExecuteLoadFromTxtFile));}
    }

    private void ExecuteLoadFromTxtFile(Object parameter)
    {
      var i = Convert.ToInt32(parameter);
      switch (i){
        case 41:
          ListTxtValue41 = GetStringFromTxtFile();
          break;
        case 51:
          ListTxtValue51 = GetStringFromTxtFile();
          break;
        case 61:
          ListValueF1 = GetStringFromTxtFile();
          break;

        default:
          break;
      }
    }

    private bool CanExecuteLoadFromTxtFile(Object parameter)
    {
      return true;
    }

    public ICommand RunQczlReportCommand
    {
      get{ return runQczlReportCommand ?? (runQczlReportCommand = new DelegateCommand<Object>(ExecuteRunQczlReport, CanExecuteRunQczlReport));}
    }

    private void ExecuteRunQczlReport(Object parameter)
    {
      int typeRpt = Convert.ToInt32(parameter);
      string src = Etc.StartPath + ModuleConst.QczlSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.QczlDest;
      Boolean res = false; 

      if (typeRpt == 0)
      {
        var sp = new QCzl1();


        res = sp.RunXls(rpt, RunXlsRptCompleted, new QCzl1RptParam(src, dst, DateBegin, DateEnd,
                                                                      Convert.ToString(SelRm1200Item.Row["StrSql"]),
                                                                      Convert.ToString(SelAroItem.Row["StrSql"]),
                                                                      Convert.ToString(SelAooItem.Row["StrSql"]),
                                                                      Convert.ToString(SelAvoItem.Row["StrSql"]),
                                                                      Convert.ToString(SelAprItem.Row["StrSql"]),
                                                                      (Convert.ToInt32(SelRm1200Item.Row["Id"]) !=0),
                                                                      (Convert.ToInt32(SelAroItem.Row["Id"]) !=0),
                                                                      (Convert.ToInt32(SelAooItem.Row["Id"]) !=0),
                                                                      (Convert.ToInt32(SelAvoItem.Row["Id"]) !=0),
                                                                      (Convert.ToInt32(SelAprItem.Row["Id"]) !=0)
                                                                     ));
      }
      else{
        var sp = new QCzl2();
        res = sp.RunXls(rpt, RunXlsRptCompleted, new QCzl2RptParam(src, dst, DateBegin, DateEnd,
                                                                      Convert.ToString(SelRm1200Item.Row["StrSql"]),
                                                                      Convert.ToString(SelAroItem.Row["StrSql"]),
                                                                      Convert.ToString(SelAooItem.Row["StrSql"]),
                                                                      Convert.ToString(SelAvoItem.Row["StrSql"]),
                                                                      Convert.ToString(SelAprItem.Row["StrSql"]),
                                                                      (Convert.ToInt32(SelRm1200Item.Row["Id"]) != 0),
                                                                      (Convert.ToInt32(SelAroItem.Row["Id"]) != 0),
                                                                      (Convert.ToInt32(SelAooItem.Row["Id"]) != 0),
                                                                      (Convert.ToInt32(SelAvoItem.Row["Id"]) != 0),
                                                                      (Convert.ToInt32(SelAprItem.Row["Id"]) != 0),
                                                                      (typeRpt == 1),
                                                                      typeQualityObj,
                                                                      ListTxtValue41
                                                                     ));
      } 


      if (res){
        var barEditItem = param as BarEditItem;
        if (barEditItem != null) barEditItem.IsVisible = true;
      }
    }

    private bool CanExecuteRunQczlReport(Object parameter)
    {
      return true;
    }

    //--Для отчета Гомзикова ЦЗЛ "Эффективность лазера"
    public ICommand RunCzlEfLsrCommand
    {
      get{return runCzlEfLsrCommand ?? (runCzlEfLsrCommand = new DelegateCommand<Object>(ExecuteRunCzlEfLsr, CanExecuteRunCzlEfLsr));}
    }

    private void ExecuteRunCzlEfLsr(Object parameter)
    {
      int typeRpt = Convert.ToInt32(parameter);
      string src = Etc.StartPath + ModuleConst.CzlEfLsrSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlEfLsrDest;

      var sp = new CzlEfLsr();
      /*
      {
        IdReport = ModuleConst.AccCmdCzlEfLsr,
        ConnectToTargetDb = DbSelector.ConnectToTargetDb,
        GetCurrentDbAlias = DbSelector.GetCurrentDbAlias
      };
      */
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, 
                              new CzlEfLsrRptParam(src, dst, DateBegin, DateEnd, 
                                                      P1750023, P1750027, P1750030, B800, KesiAvg,
                                                      CoefVoln, QntShov, Sort,
                                                      Convert.ToString(SelAdgInPrmItem.Row["StrSql"]),
                                                      Convert.ToString(SelAdgOutPrmItem.Row["StrSql"]),
                                                      typeRpt, ListTxtValue51,
                                                      Convert.ToString(SelAdgInPrmItem.Row["StrDlg"]),
                                                      Convert.ToString(SelAdgOutPrmItem.Row["StrDlg"])     
                                                     )
                             );
      



      if (res){
        var barEditItem = param as BarEditItem;
        if (barEditItem != null) barEditItem.IsVisible = true;
      }
    }

    private bool CanExecuteRunCzlEfLsr(Object parameter)
    {
      return true;
    }


    public ICommand RunCzlEfLsr9tCommand 
    {
      get { return runCzlEfLsr9tCommand ?? (runCzlEfLsr9tCommand = new DelegateCommand<Object>(ExecuteRunCzlEfLsr9t, CanExecuteRunCzlEfLsr9t)); }
    }

    private void ExecuteRunCzlEfLsr9t(Object parameter)
    {
      int typeRpt = Convert.ToInt32(parameter);
      string src = Etc.StartPath + ModuleConst.CzlEfLsr9tSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlEfLsr9tDest;

      var rptParam = new CzlEfLsr9tRptParam(src, dst)
      {
        DateBegin = DateBegin,
        DateEnd = DateEnd,
        TypeRpt = typeRpt,
        ListVal = ListTxtValue51
      };

      var sp = new CzlEfLsr9t();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (res){
        var barEditItem = param as BarEditItem;
        if (barEditItem != null) barEditItem.IsVisible = true;
      }
    }

    private bool CanExecuteRunCzlEfLsr9t(Object parameter)
    {
      return true;
    }


    //--Для отчета Гомзикова ЦЗЛ "Выход ЭАС с высоким уровнем магнитных свойств" 
    public ICommand CzlOutP1750Command
    {
      get { return czlOutP1750Command ?? (czlOutP1750Command = new DelegateCommand<Object>(ExecuteCzlOutP1750, CanExecuteCzlOutP1750)); }
    }

    private void ExecuteCzlOutP1750(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.CzlOutLowLevelP1750Source;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlOutLowLevelP1750Dest;

      var rptParam = new CzlOutP1750RptParam(src, dst)
                         {
                           PathScriptsDir = Etc.StartPath + ModuleConst.ScriptsFolder,
                           DateBegin = DateBegin,
                           DateEnd = DateEnd,
                           Rm1200 = Convert.ToString(SelRm1200Item.Row["StrSql"]),
                           Aro = Convert.ToString(SelAroItem.Row["StrSql"]),
                           Aoo = Convert.ToString(SelAooItem.Row["StrSql"]),
                           Avo = Convert.ToString(SelAvoItem.Row["StrSql"]),
                           IsRm1200 = (Convert.ToInt32(SelRm1200Item.Row["Id"]) != 0), 
                           IsAro =  (Convert.ToInt32(SelAroItem.Row["Id"]) != 0),
                           IsAoo =  (Convert.ToInt32(SelAooItem.Row["Id"]) != 0),
                           IsAvo =  (Convert.ToInt32(SelAvoItem.Row["Id"]) != 0),
                           TypeRpt = Convert.ToInt32(parameter),
                           TypeList = typeQualityObj,
                           ListVal = ListTxtValue41
                         };

      var sp = new CzlOutP1750();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteCzlOutP1750(Object parameter)
    {
      return true;
    }

    //--Для отчетов учетом фильтра F1
    public ICommand SelectTypeListValueF1Command
    {
      get { return selectTypeListValueF1Command ?? (selectTypeListValueF1Command = new DelegateCommand<Object>(ExecuteSelectTypeListValueF1, CanExecuteSelectTypeListValueF1)); }
    }

    private void ExecuteSelectTypeListValueF1(Object parameter)
    {
      typeListValueF1 = Convert.ToInt32(parameter);
    }

    private bool CanExecuteSelectTypeListValueF1(Object parameter)
    {
      return true;
    }

    public ICommand SelectTypeInclListF1Command
    {
      get { return selectTypeInclListF1Command ?? (selectTypeInclListF1Command = new DelegateCommand<Object>(ExecuteSelectTypeInclListF1, CanExecuteSelectTypeInclListF1)); }
    }

    private void ExecuteSelectTypeInclListF1(Object parameter)
    {
      typeInclLstF1 = Convert.ToInt32(parameter);
    }

    private bool CanExecuteSelectTypeInclListF1(Object parameter)
    {
      return true;
    }

    public ICommand CzlPloskF1Command
    {
      get { return czlPloskF1Command ?? (czlPloskF1Command = new DelegateCommand<Object>(ExecuteCzlPloskF1, CanExecuteCzlPloskF1)); }
    }

    private void ExecuteCzlPloskF1(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.CzlPloskSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlPloskDest;
      var rptParam = new CzlPloskRptParam(src, dst)
                         {
                           PathScriptsDir = Etc.StartPath + ModuleConst.ScriptsFolder,
                           TypeFilter = SelectedTabIndexF1,
                           DateBegin = DateBegin,
                           DateEnd = DateEnd,
                           IsSort = IsSortF1,
                           SelSortItem = SelSortF1Item,
                           SelPlskItem = SelPlskF1Item,
                           Is1200 = Is1200F1,
                           Sel1200Item = Sel1200F1Item,
                           IsDate1200 = IsDate1200F1,
                           DateBegin1200 = DateBegin1200F1,
                           DateEnd1200 = DateEnd1200F1,
                           IsAoo = IsAooF1,
                           SelAooItem = SelAooF1Item,
                           IsDateAoo = IsDateAooF1,
                           DateBeginAoo = DateBeginAooF1,
                           DateEndAoo = DateEndAooF1,
                           IsAro = IsAroF1,
                           SelAroItem = SelAroF1Item,
                           IsDateAro = IsDateAroF1,
                           DateBeginAro = DateBeginAroF1,
                           DateEndAro = DateEndAroF1,
                           IsAvo = IsAvoF1,
                           SelAvoItem = SelAvoF1Item,
                           IsDateAvo = IsDateAvoF1,
                           DateBeginAvo = DateBeginAvoF1,
                           DateEndAvo = DateEndAvoF1,
                           IsVto = IsVtoF1,
                           StdVto = StdVtoF1,
                           IsApr = IsAprF1,
                           SelAprItem = SelAprF1Item,
                           TypeListValue = typeListValueF1,
                           TypeInclList = typeInclLstF1,
                           ListValue = ListValueF1
                         };

      var sp = new CzlPlosk();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteCzlPloskF1(Object parameter)
    {
      return true;
    }

    public ICommand CzlDefPloskF1Command
    {
      get { return czlDefPloskF1Command ?? (czlDefPloskF1Command = new DelegateCommand<Object>(ExecuteCzlDefPloskF1, CanExecuteCzlDefPloskF1)); }
    }

    private void ExecuteCzlDefPloskF1(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.CzlDefPloskSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlDefPloskDest;
      var rptParam = new CzlDefPloskRptParam(src, dst)
      {
        PathScriptsDir = Etc.StartPath + ModuleConst.ScriptsFolder,
        TypeFilter = SelectedTabIndexF1,
        DateBegin = DateBegin,
        DateEnd = DateEnd,
        IsSort = IsSortF1,
        SelSortItem = SelSortF1Item,
        SelPlskItem = SelPlskF1Item,
        Is1200 = Is1200F1,
        Sel1200Item = Sel1200F1Item,
        IsDate1200 = IsDate1200F1,
        DateBegin1200 = DateBegin1200F1,
        DateEnd1200 = DateEnd1200F1,
        IsAoo = IsAooF1,
        SelAooItem = SelAooF1Item,
        IsDateAoo = IsDateAooF1,
        DateBeginAoo = DateBeginAooF1,
        DateEndAoo = DateEndAooF1,
        IsAro = IsAroF1,
        SelAroItem = SelAroF1Item,
        IsDateAro = IsDateAroF1,
        DateBeginAro = DateBeginAroF1,
        DateEndAro = DateEndAroF1,
        IsAvo = IsAvoF1,
        SelAvoItem = SelAvoF1Item,
        IsDateAvo = IsDateAvoF1,
        DateBeginAvo = DateBeginAvoF1,
        DateEndAvo = DateEndAvoF1,
        IsVto = IsVtoF1,
        StdVto = StdVtoF1,
        IsApr = IsAprF1,
        SelAprItem = SelAprF1Item,
        TypeListValue = typeListValueF1,
        TypeInclList = typeInclLstF1,
        ListValue = ListValueF1
      };

      var sp = new CzlDefPlosk();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteCzlDefPloskF1(Object parameter)
    {
      return true;
    }

    public ICommand CzlLineAooCommand
    {
      get { return czlLineAooCommand ?? (czlLineAooCommand = new DelegateCommand<Object>(ExecuteCzlLineAoo, CanExecuteCzlLineAoo)); }
    }

    private void ExecuteCzlLineAoo(Object parameter)
    {
      string src;
      string dst;

      if (typeMat == 1){
        src = Etc.StartPath + ModuleConst.CzlLineAooStendSource;
        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlLineAooStendDest;
      }
      else{
        src = Etc.StartPath + ModuleConst.CzlLineAooCoilSource;
        dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CzlLineAooCoilDest;
      }


      var rptParam = new CzlLineAooRptParam(src, dst)
      {
        DateBegin = DateBegin,
        DateEnd = DateEnd,
        TypeLineAoo = typeLineAoo,
        TypeMat = typeMat
      };

      var sp = new CzlLineAoo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);

      if (res){
        var barEditItem = param as BarEditItem;
        if (barEditItem != null) barEditItem.IsVisible = true;
      }
    }

    private bool CanExecuteCzlLineAoo(Object parameter)
    {
      return true;
    }

    #endregion Commands

  }
}
