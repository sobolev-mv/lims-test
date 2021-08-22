using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Globalization;
using DevExpress.Xpf.Core.HandleDecorator;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using DevExpress.Xpf.Editors.Settings;
using Viz.DbApp.Psi;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.LayoutControl;
using Smv.Utils;
using Viz.WrkModule.RptManager.Db;
using Viz.WrkModule.RptManager.Db.DataSets;

namespace Viz.WrkModule.RptManager
{
  internal sealed class ViewModelRptManager : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields
    private readonly DsDcBlMet dsDcBlMet = new DsDcBlMet();
    private readonly DsList4Filter dsList4Filter = new DsList4Filter();
    private readonly Smv.Xls.XlsInstanceBackgroundReport rpt;
    private readonly System.Windows.Controls.UserControl usrControl;
    private readonly Object param;
    private DateTime dateBegin;
    private DateTime dateEnd;
    private DateTime lastDateDynDef;

    private int selectedTabIndexF1 = 0;
    private Boolean is1200F1;
    private DateTime dateBegin1200F1 = DateTime.Today;
    private DateTime dateEnd1200F1 = DateTime.Today;
    private string listStendF1;
    private string listLocNumF3;

    private int? idThicknessF2;
    private decimal idThicknessF3;

    private DataRowView selThicknessItemF2;

    private Boolean isThicknessF3;
    private int selectedTabIndexF3 = 0;

    private DateTime dateBeginQuart;
    private DateTime dateEndQuart;

    private Boolean  isListStendF4;
    private DateTime dateBegin1F4 = DateTime.Today;
    private DateTime dateEnd1F4 = DateTime.Today;
    private DateTime dateBegin2F4 = DateTime.Today;
    private DateTime dateEnd2F4 = DateTime.Today;
    private DateTime dateBegin3F4 = DateTime.Today;
    private DateTime dateEnd3F4 = DateTime.Today;
    private string listStendF4;

    //Поля для Фильтра качества Начало
    private int selectedTabIndexF5 = 0;
    private string listStendF5;
    private Boolean isAroF5;
    private Boolean is1200F5;
    private Boolean isApr1F5;
    private Boolean isAooF5;
    private Boolean isVtoF5;
    private Boolean isAvoF5;
    private Boolean isDateAroF5;
    private Boolean isDate1200F5;
    private Boolean isDateApr1F5;
    private Boolean isDateAooF5;
    private Boolean isMgOF5;
    private Boolean isPppF5;
    private Boolean isWgtCoverF5;
    private Boolean isStVtoF5;
    private Boolean isKlpVtoF5;
    private Boolean isDiskVtoF5;
    private Boolean isDateAvoLstF5;
    private Boolean isTimeAooVtoF5;
    private Boolean isBrgAooF5;

    private DateTime dateBeginAroF5;
    private DateTime dateEndAroF5;
    private DateTime dateBegin1200F5;
    private DateTime dateEnd1200F5;
    private DateTime dateBeginApr1F5;
    private DateTime dateEndApr1F5;
    private DateTime dateBeginAooF5;
    private DateTime dateEndAooF5;
    private DateTime dateBeginAvoLstF5;
    private DateTime dateEndAvoLstF5;


    private DataRowView selAroF5Item;
    private DataRowView sel1200F5Item;
    private DataRowView selTolsF5Item;
    private DataRowView selBrgApr1F5Item;
    private DataRowView selBrgVtoF5Item;
    private DataRowView selBrgAvoF5Item;
    private DataRowView selAooF5Item;
    private DataRowView selAvoF5Item;
    private DataRowView selShirApr1F5Item;
    private DataRowView selDiskVtoF5Item;
    private DataRowView selBrgAooF5Item;

    private int apr1F5Width;
    private int aooF5MgOFrom;
    private int aooF5MgOTo;
    private decimal aooF5PppFrom;
    private decimal aooF5PppTo;
    private int aooF5WgtCoverFrom;
    private int aooF5WgtCoverTo;
    private string vtoF5Stend;
    private string vtoF5Cap;
    private int typeListValueF5 = 0; //0-№ Стенда, 1-№ Стенда ВТО 
    private int vtoF5TimeAooVto;
    //Поля для Фильтра качества Конец


    #endregion

    #region Public Property
    public DateTime DateBegin
    {
      get { return dateBegin; }
      set
      {
        if (value == dateBegin) return;
        dateBegin = value;
        base.OnPropertyChanged("DateBegin");
      }
    }

    public DateTime DateEnd
    {
      get { return dateEnd; }
      set
      {
        if (value == dateEnd) return;
        dateEnd = value;
        base.OnPropertyChanged("DateEnd");
      }
    }

    public DateTime DateBeginQuart
    {
      get { return dateBeginQuart; }
      set
      {
        if (value == dateBeginQuart) return;
        dateBeginQuart = value;
        base.OnPropertyChanged("DateBeginQuart");
      }
    }

    public DateTime DateEndQuart
    {
      get { return dateEndQuart; }
      set
      {
        if (value == dateEndQuart) return;
        dateEndQuart = value;
        base.OnPropertyChanged("DateEndQuart");
      }
    }



    public DateTime LastDateDynDef
    {
      get { return lastDateDynDef; }
      set
      {
        if (value == lastDateDynDef) return;
        lastDateDynDef = value;
        base.OnPropertyChanged("LastDateDynDef");
      }
    }

    public Boolean Is1200F1
    {
      get { return is1200F1; }
      set
      {
        if (value == is1200F1) return;
        is1200F1 = value;
        OnPropertyChanged("Is1200F1");
      }
    }

    public DateTime DateBegin1200F1
    {
      get { return dateBegin1200F1; }
      set
      {
        if (value == dateBegin1200F1) return;
        dateBegin1200F1 = value;
        OnPropertyChanged("DateBegin1200F1");
      }
    }

    public DateTime DateEnd1200F1
    {
      get { return dateEnd1200F1; }
      set
      {
        if (value == dateEnd1200F1) return;
        dateEnd1200F1 = value;
        OnPropertyChanged("DateEnd1200F1");
      }
    }

    public string ListStendF1
    {
      get { return listStendF1; }
      set
      {
        if (value == listStendF1) return;
        listStendF1 = value;
        OnPropertyChanged("ListStendF1");
      }
    }

    public int SelectedTabIndexF1
    {
      get { return selectedTabIndexF1; }
      set
      {
        if (value == selectedTabIndexF1) return;
        selectedTabIndexF1 = value;
        OnPropertyChanged("SelectedTabIndexF1");
      }
    }


    public Boolean IsThicknessF3
    {
      get { return isThicknessF3; }
      set
      {
        if (value == isThicknessF3) return;
        isThicknessF3 = value;
        OnPropertyChanged("IsThicknessF3");
      }
    }

    public int SelectedTabIndexF3
    {
      get { return selectedTabIndexF3; }
      set
      {
        if (value == selectedTabIndexF3) return;
        selectedTabIndexF3 = value;
        OnPropertyChanged("SelectedTabIndexF3");
      }
    }
    
    public string ListLocNumF3
    {
      get { return listLocNumF3; }
      set
      {
        if (value == listLocNumF3) return;
        listLocNumF3 = value;
        OnPropertyChanged("ListLocNumF3");
      }
    }

    public DataTable ListThicknessF2 => dsDcBlMet.ParamListThickness;
    public DataTable ListThicknessF3 => dsDcBlMet.Thickness;

    public int? IdThicknessF2
    {
      get => idThicknessF2;
      set
      {
        if (value == idThicknessF2) return;
        idThicknessF2 = value;
        OnPropertyChanged("IdThicknessF2");
      }
    }

    public DataRowView SelThicknessItemF2
    {
      get{return selThicknessItemF2;}
      set
      {
        if (Equals(value, selThicknessItemF2)) return;
        selThicknessItemF2 = value;
        OnPropertyChanged("SelThicknessItemF2");
      }
    }

    public decimal IdThicknessF3
    {
      get => idThicknessF3;
      set
      {
        if (value == idThicknessF3) return;
        idThicknessF3 = value;
        OnPropertyChanged("IdThicknessF3");
      }
    }

    public Boolean IsListStendF4
    {
      get { return isListStendF4; }
      set
      {
        if (value == isListStendF4) return;
        isListStendF4 = value;
        OnPropertyChanged("IsListStendF4");
      }
    }

    public DateTime DateBegin1F4
    {
      get => dateBegin1F4;
      set
      {
        if (value == dateBegin1F4) return;
        dateBegin1F4 = value;
        OnPropertyChanged("DateBegin1F4");
      }
    }

    public DateTime DateEnd1F4
    {
      get => dateEnd1F4;
      set
      {
        if (value == dateEnd1F4) return;
        dateEnd1F4 = value;
        OnPropertyChanged("DateEnd1F4");
      }
    }

    public DateTime DateBegin2F4
    {
      get => dateBegin2F4;
      set
      {
        if (value == dateBegin2F4) return;
        dateBegin2F4 = value;
        OnPropertyChanged("DateBegin2F4");
      }
    }

    public DateTime DateEnd2F4
    {
      get => dateEnd2F4;
      set
      {
        if (value == dateEnd2F4) return;
        dateEnd2F4 = value;
        OnPropertyChanged("DateEnd2F4");
      }
    }

    public DateTime DateBegin3F4
    {
      get => dateBegin3F4;
      set
      {
        if (value == dateBegin3F4) return;
        dateBegin3F4 = value;
        OnPropertyChanged("DateBegin3F4");
      }
    }

    public DateTime DateEnd3F4
    {
      get => dateEnd3F4;
      set
      {
        if (value == dateEnd3F4) return;
        dateEnd3F4 = value;
        OnPropertyChanged("DateEnd3F4");
      }
    }

    public string ListStendF4
    {
      get => listStendF4;
      set
      {
        if (value == listStendF4) return;
        listStendF4 = value;
        OnPropertyChanged("ListStendF4");
      }
    }

    //Поля для Фильтра качества
    public string ListStendF5
    {
      get { return listStendF5; }
      set
      {
        if (value == listStendF5) return;
        listStendF5 = value;
        OnPropertyChanged("ListStendF5");
      }
    }

    public int SelectedTabIndexF5
    {
      get { return selectedTabIndexF5; }
      set
      {
        if (value == selectedTabIndexF5) return;
        selectedTabIndexF5 = value;
        OnPropertyChanged("SelectedTabIndexF5");
      }
    }


    public Boolean IsAroF5
    {
      get { return isAroF5; }
      set
      {
        if (value == isAroF5) return;
        isAroF5 = value;
        if (!value) IsDateAroF5 = false;
        OnPropertyChanged("IsAroF5");
      }
    }

    public Boolean Is1200F5
    {
      get { return is1200F5; }
      set
      {
        if (value == is1200F5) return;
        is1200F5 = value;
        if (!value) IsDate1200F5 = false;
        OnPropertyChanged("Is1200F5");
      }
    }

    public Boolean IsApr1F5
    {
      get { return isApr1F5; }
      set
      {
        if (value == isApr1F5) return;
        isApr1F5 = value;
        if (!value) IsDateApr1F5 = false;
        OnPropertyChanged("IsApr1F5");
      }
    }

    public Boolean IsAooF5
    {
      get { return isAooF5; }
      set
      {
        if (value == isAooF5) return;
        isAooF5 = value;
        if (!value)
        {
          IsDateAooF5 = false;
          IsMgOF5 = false;
          IsPppF5 = false;
          IsWgtCoverF5 = false;
          IsBrgAooF5 = false;
        }
        OnPropertyChanged("IsAooF5");
      }
    }

    public Boolean IsVtoF5
    {
      get { return isVtoF5; }
      set
      {
        if (value == isVtoF5) return;
        isVtoF5 = value;
        if (!value)
        {
          IsStVtoF5 = false;
          IsKlpVtoF5 = false;
          IsDiskVtoF5 = false;
          IsTimeAooVtoF5 = false;
        }
        OnPropertyChanged("IsVtoF5");
      }
    }

    public Boolean IsAvoF5
    {
      get { return isAvoF5; }
      set
      {
        if (value == isAvoF5) return;
        isAvoF5 = value;
        OnPropertyChanged("IsAvoF5");
      }
    }

    public Boolean IsDateAroF5
    {
      get { return isDateAroF5; }
      set
      {
        if (value == isDateAroF5) return;
        isDateAroF5 = value;
        OnPropertyChanged("IsDateAroF5");
      }
    }

    public Boolean IsDate1200F5
    {
      get { return isDate1200F5; }
      set
      {
        if (value == isDate1200F5) return;
        isDate1200F5 = value;
        OnPropertyChanged("IsDate1200F5");
      }
    }

    public Boolean IsDateAvoLstF5
    {
      get { return isDateAvoLstF5; }
      set
      {
        if (value == isDateAvoLstF5) return;
        isDateAvoLstF5 = value;
        OnPropertyChanged("IsDateAvoLstF5");
      }
    }

    public Boolean IsDateApr1F5
    {
      get { return isDateApr1F5; }
      set
      {
        if (value == isDateApr1F5) return;
        isDateApr1F5 = value;
        OnPropertyChanged("IsDateApr1F5");
      }
    }

    public Boolean IsDateAooF5
    {
      get { return isDateAooF5; }
      set
      {
        if (value == isDateAooF5) return;
        isDateAooF5 = value;
        OnPropertyChanged("IsDateAooF5");
      }
    }

    public Boolean IsMgOF5
    {
      get { return isMgOF5; }
      set
      {
        if (value == isMgOF5) return;
        isMgOF5 = value;
        OnPropertyChanged("IsMgOF5");
      }
    }

    public Boolean IsPppF5
    {
      get { return isPppF5; }
      set
      {
        if (value == isPppF5) return;
        isPppF5 = value;
        OnPropertyChanged("IsPppF5");
      }
    }

    public Boolean IsWgtCoverF5
    {
      get { return isWgtCoverF5; }
      set
      {
        if (value == isWgtCoverF5) return;
        isWgtCoverF5 = value;
        OnPropertyChanged("IsWgtCoverF5");
      }
    }

    public Boolean IsStVtoF5
    {
      get { return isStVtoF5; }
      set
      {
        if (value == isStVtoF5) return;
        isStVtoF5 = value;
        OnPropertyChanged("IsStVtoF5");
      }
    }

    public Boolean IsKlpVtoF5
    {
      get { return isKlpVtoF5; }
      set
      {
        if (value == isKlpVtoF5) return;
        isKlpVtoF5 = value;
        OnPropertyChanged("IsKlpVtoF5");
      }
    }

    public Boolean IsDiskVtoF5
    {
      get { return isDiskVtoF5; }
      set
      {
        if (value == isDiskVtoF5) return;
        isDiskVtoF5 = value;
        OnPropertyChanged("IsDiskVtoF5");
      }
    }

    public Boolean IsTimeAooVtoF5
    {
      get { return isTimeAooVtoF5; }
      set
      {
        if (value == isTimeAooVtoF5) return;
        isTimeAooVtoF5 = value;
        OnPropertyChanged("IsTimeAooVtoF5");
      }
    }

    public Boolean IsBrgAooF5
    {
      get { return isBrgAooF5; }
      set
      {
        if (value == isBrgAooF5) return;
        isBrgAooF5 = value;
        OnPropertyChanged("IsBrgAooF5");
      }
    }

    public DateTime DateBeginAroF5
    {
      get { return dateBeginAroF5; }
      set
      {
        if (value == dateBeginAroF5) return;
        dateBeginAroF5 = value;
        OnPropertyChanged("DateBeginAroF5");
      }
    }

    public DateTime DateEndAroF5
    {
      get { return dateEndAroF5; }
      set
      {
        if (value == dateEndAroF5) return;
        dateEndAroF5 = value;
        OnPropertyChanged("DateEndAroF5");
      }
    }

    public DateTime DateBegin1200F5
    {
      get { return dateBegin1200F5; }
      set
      {
        if (value == dateBegin1200F5) return;
        dateBegin1200F5 = value;
        OnPropertyChanged("DateBegin1200F5");
      }
    }

    public DateTime DateEnd1200F5
    {
      get { return dateEnd1200F5; }
      set
      {
        if (value == dateEnd1200F5) return;
        dateEnd1200F5 = value;
        OnPropertyChanged("DateEnd1200F5");
      }
    }

    public DateTime DateBeginApr1F5
    {
      get { return dateBeginApr1F5; }
      set
      {
        if (value == dateBeginApr1F5) return;
        dateBeginApr1F5 = value;
        OnPropertyChanged("DateBeginApr1F5");
      }
    }

    public DateTime DateEndApr1F5
    {
      get { return dateEndApr1F5; }
      set
      {
        if (value == dateEndApr1F5) return;
        dateEndApr1F5 = value;
        OnPropertyChanged("DateEndApr1F5");
      }
    }

    public DateTime DateBeginAooF5
    {
      get { return dateBeginAooF5; }
      set
      {
        if (value == dateBeginAooF5) return;
        dateBeginAooF5 = value;
        OnPropertyChanged("DateBeginAooF5");
      }
    }

    public DateTime DateEndAooF5
    {
      get { return dateEndAooF5; }
      set
      {
        if (value == dateEndAooF5) return;
        dateEndAooF5 = value;
        OnPropertyChanged("DateEndAooF5");
      }
    }

    public DateTime DateBeginAvoLstF5
    {
      get { return dateBeginAvoLstF5; }
      set
      {
        if (value == dateBeginAvoLstF5) return;
        dateBeginAvoLstF5 = value;
        OnPropertyChanged("DateBeginAvoLstF5");
      }
    }

    public DateTime DateEndAvoLstF5
    {
      get { return dateEndAvoLstF5; }
      set
      {
        if (value == dateEndAvoLstF5) return;
        dateEndAvoLstF5 = value;
        OnPropertyChanged("DateEndAvoLstF5");
      }
    }


    public DataRowView SelAroF5Item
    {
      get { return selAroF5Item; }
      set
      {
        if (Equals(value, selAroF5Item)) return;
        selAroF5Item = value;
        OnPropertyChanged("SelAroF5Item");
      }
    }

    
    public DataTable AroTs
    {
      get { return dsList4Filter.AroTs; }
    }
    

    public DataRowView Sel1200F5Item
    {
      get { return sel1200F5Item; }
      set
      {
        if (Equals(value, sel1200F5Item)) return;
        sel1200F5Item = value;
        OnPropertyChanged("Sel1200F5Item");
      }
    }

    
    public DataTable Rm1200Ts
    {
      get { return dsList4Filter.Rm1200Ts; }
    }
    

    public DataRowView SelTolsF5Item
    {
      get { return selTolsF5Item; }
      set
      {
        if (Equals(value, selTolsF5Item)) return;
        selTolsF5Item = value;
        OnPropertyChanged("SelTolsF5Item");
      }
    }
    
    public DataTable Tols
    {
      get { return dsList4Filter.Tols; }
    }

    public DataRowView SelBrgApr1F5Item
    {
      get { return selBrgApr1F5Item; }
      set
      {
        if (Equals(value, selBrgApr1F5Item)) return;
        selBrgApr1F5Item = value;
        OnPropertyChanged("SelBrgApr1F5Item");
      }
    }

   
    public DataTable ShirApr1
    {
      get { return dsList4Filter.ShirApr1; }
    }
   

    public DataRowView SelShirApr1F5Item
    {
      get { return selShirApr1F5Item; }
      set
      {
        if (Equals(value, selShirApr1F5Item)) return;
        selShirApr1F5Item = value;
        OnPropertyChanged("SelShirApr1F5Item");
      }
    }


    public DataRowView SelBrgVtoF5Item
    {
      get { return selBrgVtoF5Item; }
      set
      {
        if (Equals(value, selBrgVtoF5Item)) return;
        selBrgVtoF5Item = value;
        OnPropertyChanged("SelBrgVtoF5Item");
      }
    }

    public DataRowView SelBrgAvoF5Item
    {
      get { return selBrgAvoF5Item; }
      set
      {
        if (Equals(value, selBrgAvoF5Item)) return;
        selBrgAvoF5Item = value;
        OnPropertyChanged("SelBrgAvoF5Item");
      }
    }

    public DataRowView SelDiskVtoF5Item
    {
      get { return selDiskVtoF5Item; }
      set
      {
        if (Equals(value, selDiskVtoF5Item)) return;
        selDiskVtoF5Item = value;
        OnPropertyChanged("SelDiskVtoF5Item");
      }
    }

    public DataRowView SelBrgAooF5Item
    {
      get { return selBrgAooF5Item; }
      set
      {
        if (Equals(value, selBrgAooF5Item)) return;
        selBrgAooF5Item = value;
        OnPropertyChanged("SelBrgAooF5Item");
      }
    }
    public DataTable Brg
    {
      get { return dsList4Filter.Brg; }
    }

    public DataRowView SelAooF5Item
    {
      get { return selAooF5Item; }
      set
      {
        if (Equals(value, selAooF5Item)) return;
        selAooF5Item = value;
        OnPropertyChanged("SelAooF5Item");
      }
    }
    public DataTable AooTs
    {
      get { return dsList4Filter.AooTs; }
    }
    

    public DataRowView SelAvoF5Item
    {
      get { return selAvoF5Item; }
      set
      {
        if (Equals(value, selAvoF5Item)) return;
        selAvoF5Item = value;
        OnPropertyChanged("SelAvoF5Item");
      }
    }

    public DataTable AvoTs
    {
      get { return dsList4Filter.AvoTs; }
    }
    

    public DataTable DiskVtoTs
    {
      get { return dsList4Filter.DiskVto; }
    }

    public int Apr1F5Width
    {
      get { return apr1F5Width; }
      set
      {
        if (value == apr1F5Width) return;
        apr1F5Width = value;
        OnPropertyChanged("Apr1F5Width");
      }
    }

    public int AooF5MgOFrom
    {
      get { return aooF5MgOFrom; }
      set
      {
        if (value == aooF5MgOFrom) return;
        aooF5MgOFrom = value;
        OnPropertyChanged("AooF5MgOFrom");
      }
    }

    public int AooF5MgOTo
    {
      get { return aooF5MgOTo; }
      set
      {
        if (value == aooF5MgOTo) return;
        aooF5MgOTo = value;
        OnPropertyChanged("AooF5MgOTo");
      }
    }

    public decimal AooF5PppFrom
    {
      get { return aooF5PppFrom; }
      set
      {
        if (value == aooF5PppFrom) return;
        aooF5PppFrom = value;
        OnPropertyChanged("AooF5PppFrom");
      }
    }

    public decimal AooF5PppTo
    {
      get { return aooF5PppTo; }
      set
      {
        if (value == aooF5PppTo) return;
        aooF5PppTo = value;
        OnPropertyChanged("AooF5PppTo");
      }
    }

    public int AooF5WgtCoverFrom
    {
      get { return aooF5WgtCoverFrom; }
      set
      {
        if (value == aooF5WgtCoverFrom) return;
        aooF5WgtCoverFrom = value;
        OnPropertyChanged("AooF5WgtCoverFrom");
      }
    }

    public int AooF5WgtCoverTo
    {
      get { return aooF5WgtCoverTo; }
      set
      {
        if (value == aooF5WgtCoverTo) return;
        aooF5WgtCoverTo = value;
        OnPropertyChanged("AooF5WgtCoverTo");
      }
    }

    public string VtoF5Stend
    {
      get { return vtoF5Stend; }
      set
      {
        if (value == vtoF5Stend) return;
        vtoF5Stend = value;
        OnPropertyChanged("VtoF5Stend");
      }
    }

    public string VtoF5Cap
    {
      get { return vtoF5Cap; }
      set
      {
        if (value == vtoF5Cap) return;
        vtoF5Cap = value;
        OnPropertyChanged("VtoF5Cap");
      }
    }

    public int VtoF5TimeAooVto
    {
      get { return vtoF5TimeAooVto; }
      set
      {
        if (value == vtoF5TimeAooVto) return;
        vtoF5TimeAooVto = value;
        OnPropertyChanged("VtoF5TimeAooVto");
      }
    }
    //конец Поля для фильтра качество


    #endregion

    #region Private Method
    private void RunXlsRptCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      //Обновляем дату последнего расчета
      LastDateDynDef = Db.DbUtils.GetLastDateDynDef();


      var barEditItem = param as BarEditItem;
      if (barEditItem != null)
        barEditItem.IsVisible = false;
    }

    private void LayoutGroupExpanded(object sender, EventArgs e)
    {
      var layoutGroup = sender as DevExpress.Xpf.LayoutControl.LayoutGroup;
      if (layoutGroup != null)
      {

        if (layoutGroup.Tag == null)
          return;

        var i = Convert.ToInt32(layoutGroup.Tag);
        switch (i)
        {
          case 0:
            break;
          case 1:
            break;
          case 2:
            LastDateDynDef = Db.DbUtils.GetLastDateDynDef();
            break;
          case 3:
            break;
          case 4:
            IdThicknessF2 = 0;
            LgExpandedF5();
            break;
          case 5:
            break;
          case 6:
            DateBeginQuart = DbUtils.GetDateBeginQuart();
            DateEndQuart = DbUtils.GetDateEndQuart();
            break;
          default:
            break;
        }
      }
    }

    private void LgExpandedF5()
    {
      if (dsList4Filter.AroTs.Rows.Count == 0)
        dsList4Filter.AroTs.LoadData(2);

      if (dsList4Filter.Rm1200Ts.Rows.Count == 0)
        dsList4Filter.Rm1200Ts.LoadData(1);

      if (dsList4Filter.Tols.Rows.Count == 0)
        dsList4Filter.Tols.LoadData(12);

      if (dsList4Filter.Brg.Rows.Count == 0)
        dsList4Filter.Brg.LoadData(13);

      if (dsList4Filter.AooTs.Rows.Count == 0)
        dsList4Filter.AooTs.LoadData(3);

      if (dsList4Filter.AvoTs.Rows.Count == 0)
        dsList4Filter.AvoTs.LoadData(4);

      if (dsList4Filter.ShirApr1.Rows.Count == 0)
        dsList4Filter.ShirApr1.LoadData(16);

      if (dsList4Filter.DiskVto.Rows.Count == 0)
        dsList4Filter.DiskVto.LoadData(17);

      for (int i = 0; i < 11; i++)
      {
        var cbe = LogicalTreeHelper.FindLogicalNode(this.usrControl, "CbeTypeF1Ts" + i.ToString(CultureInfo.InvariantCulture)) as DevExpress.Xpf.Editors.ComboBoxEdit;
        if (cbe != null)
          cbe.SelectedIndex = 0;
      }
    }

    void SelectedTypeFilter(object sender, DevExpress.Xpf.Core.ValueChangedEventArgs<FrameworkElement> e)
    {
      var layoutGroup = sender as LayoutGroup;

      if ((layoutGroup != null) && (layoutGroup.Name == "lgFilter1"))
        this.SelectedTabIndexF1 = layoutGroup.SelectedTabIndex;

      if ((layoutGroup != null) && (layoutGroup.Name == "lgFilter3"))
        this.SelectedTabIndexF3 = layoutGroup.SelectedTabIndex;

      if ((layoutGroup != null) && (layoutGroup.Name == "lgFilter5"))
        this.SelectedTabIndexF5 = layoutGroup.SelectedTabIndex;
    }
    #endregion

    #region Constructor
    internal ViewModelRptManager(System.Windows.Controls.UserControl control, Object Param)
    {
      param = Param;
      rpt = new Smv.Xls.XlsInstanceBackgroundReport();
      usrControl = control;
      DateBegin = DateEnd = DateBeginAroF5 = DateEndAroF5 = DateBegin1200F5 = DateEnd1200F5 = DateBeginApr1F5 = DateEndApr1F5 = DateBeginAooF5 = DateEndAooF5 = 
      DateBeginAvoLstF5 = DateEndAvoLstF5 = DateTime.Today;

      dsDcBlMet.ParamListThickness.LoadData(23);
 
      //Группы 1-уровня
      for (int i = ModuleConst.AccGrpRk; i < ModuleConst.AccGrpKpaRolling + 1; i++){
        var lg = LogicalTreeHelper.FindLogicalNode(this.usrControl, "Lg" + ModuleConst.ModuleId + "_" + i.ToString()) as DevExpress.Xpf.LayoutControl.LayoutGroup;

        if (lg != null){
          if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            lg.Visibility = Visibility.Hidden;
            continue;
          }

          lg.Expanded += LayoutGroupExpanded;
          //this.lg.Collapsed += LayoutGroupCollapsed;
        }
      }

      //Группы 2-уровня
      for (int i = ModuleConst.AccGrpDynDefUtl; i < ModuleConst.AccGrpDynDefUtl + 1; i++){
        var uie = LogicalTreeHelper.FindLogicalNode(this.usrControl, "G2" + ModuleConst.ModuleId + "_" + i) as UIElement;

        if (uie != null){
          if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            uie.Visibility = Visibility.Hidden;
            continue;
          }
        }
      }


      //Делаем controls невидимыми
      for (int i = ModuleConst.AccCmdRkSko; i < ModuleConst.AccCmdDefects1StRoll + 1; i++){
        var btn = LogicalTreeHelper.FindLogicalNode(this.usrControl, "b" + ModuleConst.ModuleId + "_" + i.ToString()) as UIElement;

        if (btn == null) continue;

        if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
          btn.Visibility = Visibility.Hidden;
          continue;
        }
      }

      var layoutGroup = LogicalTreeHelper.FindLogicalNode(this.usrControl, "lgFilter1") as LayoutGroup;
      if (layoutGroup != null)
        layoutGroup.SelectedTabChildChanged += SelectedTypeFilter;

      layoutGroup = LogicalTreeHelper.FindLogicalNode(this.usrControl, "lgFilter3") as LayoutGroup;
      if (layoutGroup != null)
        layoutGroup.SelectedTabChildChanged += SelectedTypeFilter;

      layoutGroup = LogicalTreeHelper.FindLogicalNode(this.usrControl, "lgFilter5") as LayoutGroup;
      if (layoutGroup != null)
        layoutGroup.SelectedTabChildChanged += SelectedTypeFilter;


      //DxInfo.ShowDxBoxInfo("Внимание", "Файлы созданных отчетов будут находиться в папке Документы (Мои документы)!", MessageBoxImage.Information);

      IdThicknessF3 = new decimal(0.23);
    }
    #endregion Constructor

    #region Commands
    private DelegateCommand<Object> showListRptCommand;
    private DelegateCommand<Object> loadFromTxtFileCommand;
    private DelegateCommand<Object> rkSkoCommand;
    private DelegateCommand<Object> rkDinamCommand;
    private DelegateCommand<Object> lst1300Command;
    private DelegateCommand<Object> pdbSclCommand;
    private DelegateCommand<Object> pdbSclDcCommand;
    private DelegateCommand<Object> sgpStateCommand;
    private DelegateCommand<Object> appHCommand;
    private DelegateCommand<Object> w313c2CoCommand;
    private DelegateCommand<Object> dynDefect12Cat1SortCommand;
    private DelegateCommand<Object> defectEdgeCommand;
    private DelegateCommand<Object> rkBrgSt1200Command;
    private DelegateCommand<Object> partOf1SortCommand;
    private DelegateCommand<Object> dlgDcBlMetCommand;
    private DelegateCommand<Object> qualityFinCutUoCommand;
    private DelegateCommand<Object> qualityFinCutUoMCommand;
    private DelegateCommand<Object> qntWeldOn2ndRollCommand;
    private DelegateCommand<Object> pdbScl313cCommand;
    private DelegateCommand<Object> kesiAvoCommand;
    private DelegateCommand<Object> dynDefect2CatCommand;
    private DelegateCommand<Object> sgpTo2CatCommand;
    private DelegateCommand<Object> sgpTo3CatCommand;
    private DelegateCommand<Object> currentQualityCommand;
    private DelegateCommand<Object> sgpTo2SortCommand;
    private DelegateCommand<Object> sgpTo2SortFinCutCommand;
    private DelegateCommand<Object> sgpTo2CatFinCutCommand;
    private DelegateCommand<Object> sgpTo3CatFinCutCommand;
    private DelegateCommand<Object> lasScrAfterFinCutCommand;
    private DelegateCommand<Object> reasonOfStripBreakageRm1300Command;
    private DelegateCommand<Object> loadTblExceptCommand;
    private DelegateCommand<Object> loadProdTargetsCommand;
    private DelegateCommand<Object> balanceWrkTimeCommand;
    private DelegateCommand<Object> monitorDef2CatCommand;
    private DelegateCommand<Object> kpaRollingCommand;
    private DelegateCommand<Object> monitorDefCommand;
    private DelegateCommand<Object> lider2CatCommand;
    private DelegateCommand<Object> defects1StRollCommand;

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
        Source = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.RptManager;Component/Images/BarImage.png"))
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
      switch (Convert.ToInt32(parameter))
      {
        case 61:
          this.ListStendF1 = Etc.GetStringWithDelimFromTxtFile(Encoding.GetEncoding("windows-1251"), ","); 
          break;
        case 62:
          this.ListStendF4 = Etc.GetStringWithDelimFromTxtFile(Encoding.GetEncoding("windows-1251"), ",");
          break;
        default:
          break;
      }
    }

    private bool CanExecuteLoadFromTxtFile(Object parameter)
    {
      return true;
    }
    
    public ICommand RkSkoCommand
    {
      get { return rkSkoCommand ?? (rkSkoCommand = new DelegateCommand<Object>(ExecuteRkSko, CanExecuteRkSko)); }
    }

    private void ExecuteRkSko(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.RkSkoSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.RkSkoDest;

      var sp = new Db.RkSko();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.RkSkoRptParam(src, dst, this.DateBegin, this.DateEnd));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteRkSko(Object parameter)
    {
      return true;
    }

    public ICommand RkDinamCommand
    {
      get { return rkDinamCommand ?? (rkDinamCommand = new DelegateCommand<Object>(ExecuteRkDinam, CanExecuteRkDinam)); }
    }

    private void ExecuteRkDinam(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.RkDinamSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.RkDinamDest;

      var sp = new Db.RkDinam();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.RkDinamRptParam(src, dst, this.DateBegin, this.DateEnd));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteRkDinam(Object parameter)
    {
      return true;
    }


    public ICommand Lst1300Command
    {
      get { return lst1300Command ?? (lst1300Command = new DelegateCommand<Object>(ExecuteLst1300, CanExecuteLst1300)); }
    }

    private void ExecuteLst1300(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.Lst1300Source;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.Lst1300Dest;

      var sp = new Db.Lst1300();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.Lst1300RptParam(src, dst, this.DateBegin, this.DateEnd));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteLst1300(Object parameter)
    {
      return true;
    }

    public ICommand PdbSclCommand
    {
      get { return pdbSclCommand ?? (pdbSclCommand = new DelegateCommand<Object>(ExecutePdbScl, CanExecutePdbScl)); }
    }

    private void ExecutePdbScl(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.PdbSclSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.PdbSclDest;

      var sp = new Db.PdbScl();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.PdbSclRptParam(src, dst));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecutePdbScl(Object parameter)
    {
      return true;
    }

    public ICommand PdbSclDcCommand => pdbSclDcCommand ?? (pdbSclDcCommand = new DelegateCommand<Object>(ExecutePdbSclDc, CanExecutePdbSclDc));

    private void ExecutePdbSclDc(Object parameter)
    {
      string src = Etc.StartPath + ModuleConst.PdbSclDcSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.PdbSclDcDest;

      var sp = new Db.PdbSclDc();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.PdbSclDcRptParam(src, dst));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecutePdbSclDc(Object parameter)
    {
      return true;
    }
    
    public ICommand SgpStateCommand
    {
      get { return sgpStateCommand ?? (sgpStateCommand = new DelegateCommand<Object>(ExecuteSgpState, CanExecuteSgpState)); }
    }

    private void ExecuteSgpState(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.SgpStateSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpStateDest;

      var rptParam = new Db.SgpStateRptParam(src, dst)
      {
        DateBegin = this.DateBegin
      };

      var sp = new Db.SgpState();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteSgpState(Object parameter)
    {
      return true;
    }

    public ICommand AppHCommand
    {
      get { return appHCommand ?? (appHCommand = new DelegateCommand<Object>(ExecuteAppH, CanExecuteAppH)); }
    }

    private void ExecuteAppH(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.AppHSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.AppHDest;

      var rptParam = new Db.AppHRptParam(src, dst)
      {
        DateBegin = this.DateBegin
      };

      var sp = new Db.AppH();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteAppH(Object parameter)
    {
      return true;
    }


    public ICommand W313c2CoCommand
    {
      get { return w313c2CoCommand ?? (w313c2CoCommand = new DelegateCommand<Object>(ExecuteW313c2Co, CanExecuteW313c2Co)); }
    }

    private void ExecuteW313c2Co(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.W313c2CoSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.W313c2CoDest;

      var sp = new Db.W313c2Co();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.W313c2CoRptParam(src, dst));
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteW313c2Co(Object parameter)
    {
      return true;
    }
    
    public ICommand PdbScl313cCommand
    {
      get { return pdbScl313cCommand ?? (pdbScl313cCommand = new DelegateCommand<Object>(ExecutePdbScl313c, CanExecutePdbScl313c)); }
    }

    private void ExecutePdbScl313c(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.PdbScl313cSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.PdbScl313cDest;

      var rptParam = new Db.PdbScl313cRptParam(src, dst)
      {
        DateBegin = this.DateBegin
      };

      var sp = new Db.PdbScl313c();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecutePdbScl313c(Object parameter)
    {
      return true;
    }


    public ICommand DynDefect12Cat1SortCommand
    {
      get { return dynDefect12Cat1SortCommand ?? (dynDefect12Cat1SortCommand = new DelegateCommand<Object>(ExecuteDynDefect12Cat1Sort, CanExecuteDynDefect12Cat1Sort)); }
    }

    private void ExecuteDynDefect12Cat1Sort(Object parameter)
    {
      int prm = Convert.ToInt32(parameter);

      if (prm == 0)
      {
        string src = Smv.Utils.Etc.StartPath + ModuleConst.DynDefect12Cat1SortSource;
        string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.DynDefect12Cat1SortDest;

        var rptParam = new Db.DynDefect12Cat1SortRptParam(src, dst)
        {
          DateBegin = this.DateBegin,
          DateEnd = this.DateEnd
        };

        var sp = new Db.DynDefect12Cat1Sort();
        Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);


        if (!res) return;
      }
      else
      {
        //Здесь делаем проверку даты
        if (prm == 1)
          if (this.DateBegin != LastDateDynDef.AddDays(1))
          {
            Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка выбора периода", "Дата начала диапазона должна быть: " + string.Format("{0:dd.MM.yyyy}", LastDateDynDef.AddDays(1)), MessageBoxImage.Error);
            return;
          }

        var rptParam = new Db.DynDefect12Cat1SortUtlRptParam(null, null)
        {
          TypeAction = prm,
          DateBegin = this.DateBegin,
          DateEnd = this.DateEnd
        };

        var sp = new Db.DynDefect12Cat1SortUtl();
        Boolean res = sp.Run(rpt, RunXlsRptCompleted, rptParam);


        if (prm == 2)
          Smv.Utils.DxInfo.ShowDxBoxInfo("Дата отката", "Был выполнен откат за дату: " + string.Format("{0:dd.MM.yyyy}", LastDateDynDef), MessageBoxImage.Information);

        if (!res) return;
      }


      //Обновляем дату последнего расчета
      LastDateDynDef = Db.DbUtils.GetLastDateDynDef();

      var barEditItem = param as BarEditItem;

      if (prm != 2)
        if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteDynDefect12Cat1Sort(Object parameter)
    {
      return true;
    }

    public ICommand DefectEdgeCommand
    {
      get { return defectEdgeCommand ?? (defectEdgeCommand = new DelegateCommand<Object>(ExecuteDefectEdge, CanExecuteDefectEdge)); }
    }

    private void ExecuteDefectEdge(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.DefefectEdgeSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.DefectEdgeDest;

      var rptParam = new Db.DefectEdgeRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.DefectEdge();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteDefectEdge(Object parameter)
    {
      return true;
    }

    public ICommand RkBrgSt1200Command
    {
      get { return rkBrgSt1200Command ?? (rkBrgSt1200Command = new DelegateCommand<Object>(ExecuteRkBrgSt1200, CanExecuteRkBrgSt1200)); }
    }

    private void ExecuteRkBrgSt1200(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.RkBrgSt1200Source;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.RkBrgSt1200Dest;

      var rptParam = new Db.RkBrgSt1200RptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilterF1 = SelectedTabIndexF1,
        Is1200F1 = Is1200F1,
        DateBegin1200F1 = DateBegin1200F1,
        DateEnd1200F1 = DateEnd1200F1,
        ListStendF1 = ListStendF1
      };

      var sp = new Db.RkBrgSt1200();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteRkBrgSt1200(Object parameter)
    {
      return true;
    }
    

    public ICommand PartOf1SortCommand
    {
      get { return partOf1SortCommand ?? (partOf1SortCommand = new DelegateCommand<Object>(ExecutePartOf1Sort, CanExecutePartOf1Sort)); }
    }

    private void ExecutePartOf1Sort(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.PartOf1SortSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.PartOf1SortDest;

      var rptParam = new Db.PartOf1SortRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilterF3 = SelectedTabIndexF3,
        IsThicknessF3 = IsThicknessF3,
        ThicknessF3 = IdThicknessF3,
        ListLocNumF3 = ListLocNumF3
      };

      var sp = new Db.PartOf1Sort();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecutePartOf1Sort(Object parameter)
    {
      return true;
    }


    public ICommand DlgDcBlMetCommand
    {
      get { return dlgDcBlMetCommand ?? (dlgDcBlMetCommand = new DelegateCommand<Object>(ExecuteDlgDcBlMet, CanExecuteDlgDcBlMet)); }
    }

    private void ExecuteDlgDcBlMet(Object parameter)
    {
      var dlg = new ViewDlgDcBlMet();
      dlg.ShowDialog();
    }

    private bool CanExecuteDlgDcBlMet(Object parameter)
    {
      return true;
    }

    public ICommand QualityFinCutUoCommand
    {
      get { return qualityFinCutUoCommand ?? (qualityFinCutUoCommand = new DelegateCommand<Object>(ExecuteQualityFinCutUo, CanExecuteQualityFinCutUo)); }
    }

    private void ExecuteQualityFinCutUo(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.QualityFinCutUoSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.QualityFinCutUoDest;

      var rptParam = new Db.QualityFinCutUoRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.QualityFinCutUo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteQualityFinCutUo(Object parameter)
    {
      return true;
    }

    public ICommand QualityFinCutUoMCommand => qualityFinCutUoMCommand ?? (qualityFinCutUoMCommand = new DelegateCommand<Object>(ExecuteQualityFinCutUoM, CanExecuteQualityFinCutUoM));

    private void ExecuteQualityFinCutUoM(Object parameter)
    {
      string src = Smv.Utils.Etc.StartPath + ModuleConst.QualityFinCutUoMSource;
      string dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.QualityFinCutUoMDest;

      var rptParam = new Db.QualityFinCutUoMRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.QualityFinCutUoM();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteQualityFinCutUoM(Object parameter)
    {
      return true;
    }
    
    public ICommand QntWeldOn2ndRollCommand => qntWeldOn2ndRollCommand ?? (qntWeldOn2ndRollCommand = new DelegateCommand<Object>(ExecuteQntWeldOn2ndRoll, CanExecuteQntWeldOn2ndRoll));

    private void ExecuteQntWeldOn2ndRoll(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.WeldOn2ndRollSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.WeldOn2ndRollDest;

      var rptParam = new Db.QntWeldOn2ndRollRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.QntWeldOn2ndRoll();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteQntWeldOn2ndRoll(Object parameter)
    {
      return true;
    }
    
    public ICommand KesiAvoCommand => kesiAvoCommand ?? (kesiAvoCommand = new DelegateCommand<Object>(ExecuteKesiAvoCommand, CanExecuteKesiAvoCommand));

    private void ExecuteKesiAvoCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.KesiAvoSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.KesiAvoDest;

      var rptParam = new Db.KesiAvoRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.KesiAvo();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteKesiAvoCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand DynDefect2CatCommand => dynDefect2CatCommand ?? (dynDefect2CatCommand = new DelegateCommand<Object>(ExecuteDynDefect2CatCommand, CanExecuteDynDefect2CatCommand));

    private void ExecuteDynDefect2CatCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.DynDefect2CatSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.DynDefect2CatDest;

      var rptParam = new Db.DynDefect2CatRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.DynDefect2Cat();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteDynDefect2CatCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand SgpTo2CatCommand => sgpTo2CatCommand ?? (sgpTo2CatCommand = new DelegateCommand<Object>(ExecuteSgpTo2CatCommand, CanExecuteSgpTo2CatCommand));

    private void ExecuteSgpTo2CatCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.SgpTo2CatSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpTo2CatDest;

      var rptParam = new Db.SgpTo2CatRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.SgpTo2Cat();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteSgpTo2CatCommand(Object parameter)
    {
      return true;
    }

    public ICommand SgpTo3CatCommand => sgpTo3CatCommand ?? (sgpTo3CatCommand = new DelegateCommand<Object>(ExecuteSgpTo3CatCommand, CanExecuteSgpTo3CatCommand));

    private void ExecuteSgpTo3CatCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.SgpTo3CatSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpTo3CatDest;

      var rptParam = new Db.SgpTo3CatRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.SgpTo3Cat();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteSgpTo3CatCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand CurrentQualityCommand => currentQualityCommand ?? (currentQualityCommand = new DelegateCommand<Object>(ExecuteCurrentQualityCommand, CanExecuteCurrentQualityCommand));

    private void ExecuteCurrentQualityCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.CurrentQualitySource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.CurrentQualityDest;

      var rptParam = new Db.CurrentQualityRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.CurrentQuality();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteCurrentQualityCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand SgpTo2SortCommand => sgpTo2SortCommand ?? (sgpTo2SortCommand = new DelegateCommand<Object>(ExecuteSgpTo2SortCommand, CanExecuteSgpTo2SortCommand));

    private void ExecuteSgpTo2SortCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.SgpTo2SortSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpTo2SortDest;

      var rptParam = new Db.SgpTo2SortRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.SgpTo2Sort();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteSgpTo2SortCommand(Object parameter)
    {
      return true;
    }
    //
    public ICommand SgpTo2SortFinCutCommand => sgpTo2SortFinCutCommand ?? (sgpTo2SortFinCutCommand = new DelegateCommand<Object>(ExecuteSgpTo2SortFinCutCommand, CanExecuteSgpTo2SortFinCutCommand));

    private void ExecuteSgpTo2SortFinCutCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.SgpTo2SortFinCutSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpTo2SortFinCutDest;

      var rptParam = new Db.SgpTo2SortFinCutRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF5,
        ListStendF5 = this.ListStendF5,
        TypeListValueF5 = this.typeListValueF5,
        IsAroF5 = this.IsAroF5,
        Is1200F5 = this.Is1200F5,
        IsApr1F5 = this.IsApr1F5,
        IsAooF5 = this.IsAooF5,
        IsVtoF5 = this.IsVtoF5,
        IsAvoF5 = this.IsAvoF5,
        IsMgOF5 = this.IsMgOF5,
        IsPppF5 = this.IsPppF5,
        IsWgtCoverF5 = this.IsWgtCoverF5,
        IsStVtoF5 = this.IsStVtoF5,
        IsKlpVtoF5 = this.IsKlpVtoF5,
        IsDiskVtoF5 = this.IsDiskVtoF5,
        IsTimeAooVtoF5 = this.IsTimeAooVtoF5,
        IsDateAroF5 = this.IsDateAroF5,
        IsDate1200F5 = this.IsDate1200F5,
        IsDateApr1F5 = this.IsDateApr1F5,
        IsDateAooF5 = this.IsDateAooF5,
        DateBeginAroF5 = this.DateBeginAroF5,
        DateEndAroF5 = this.DateEndAroF5,
        DateBegin1200F5 = this.DateBegin1200F5,
        DateEnd1200F5 = this.DateEnd1200F5,
        DateBeginApr1F5 = this.DateBeginApr1F5,
        DateEndApr1F5 = this.DateEndApr1F5,
        DateBeginAooF5 = this.DateBeginAooF5,
        DateEndAooF5 = this.DateEndAooF5,
        AroF5Item = Convert.ToString(this.SelAroF5Item.Row["StrSql"]),
        Stan1200F5Item = Convert.ToString(this.sel1200F5Item.Row["StrSql"]),
        TolsF5Item = Convert.ToString(this.SelTolsF5Item.Row["StrSql"]),
        BrgApr1F5Item = Convert.ToString(this.SelBrgApr1F5Item.Row["StrSql"]),
        BrgVtoF5Item = Convert.ToString(this.SelBrgVtoF5Item.Row["StrSql"]),
        BrgAvoF5Item = Convert.ToString(this.SelBrgAvoF5Item.Row["StrSql"]),
        AooF5Item = Convert.ToString(this.SelAooF5Item.Row["StrSql"]),
        AvoF5Item = Convert.ToString(this.SelAvoF5Item.Row["StrSql"]),
        ShirApr1F5Item = Convert.ToString(this.SelShirApr1F5Item.Row["StrSql"]),
        DiskVtoF5Item = Convert.ToString(this.SelDiskVtoF5Item.Row["StrSql"]),
        //Apr1F5Width = this.Apr1F5Width,
        AooF5MgOFrom = this.AooF5MgOFrom,
        AooF5MgOTo = this.AooF5MgOTo,
        AooF5PppFrom = this.AooF5PppFrom,
        AooF5PppTo = this.AooF5PppTo,
        AooF5WgtCoverFrom = this.AooF5WgtCoverFrom,
        AooF5WgtCoverTo = this.AooF5WgtCoverTo,
        VtoF5Stend = this.VtoF5Stend,
        VtoF5Cap = this.VtoF5Cap,
        VtoF5TimeAooVto = this.VtoF5TimeAooVto,
        IsDateAvoLstF5 = this.IsDateAvoLstF5,
        DateBeginAvoLstF5 = this.DateBeginAvoLstF5,
        DateEndAvoLstF5 = this.DateEndAvoLstF5,
        BrgAooF5Item = Convert.ToString(this.SelBrgAooF5Item.Row["StrSql"]),
        IsBrgAooF5 = this.IsBrgAooF5,
        StrThicknessSql = Convert.ToString(this.SelThicknessItemF2.Row["StrSql"])
      };

      var sp = new Db.SgpTo2SortFinCut();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    
    private bool CanExecuteSgpTo2SortFinCutCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand SgpTo2CatFinCutCommand => sgpTo2CatFinCutCommand ?? (sgpTo2CatFinCutCommand = new DelegateCommand<Object>(ExecuteSgpTo2CatFinCutCommand, CanExecuteSgpTo2CatFinCutCommand));

    private void ExecuteSgpTo2CatFinCutCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.SgpTo2CatFinCutSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpTo2CatFinCutDest;

      var rptParam = new Db.SgpTo2CatFinCutRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF5,
        ListStendF5 = this.ListStendF5,
        TypeListValueF5 = this.typeListValueF5,
        IsAroF5 = this.IsAroF5,
        Is1200F5 = this.Is1200F5,
        IsApr1F5 = this.IsApr1F5,
        IsAooF5 = this.IsAooF5,
        IsVtoF5 = this.IsVtoF5,
        IsAvoF5 = this.IsAvoF5,
        IsMgOF5 = this.IsMgOF5,
        IsPppF5 = this.IsPppF5,
        IsWgtCoverF5 = this.IsWgtCoverF5,
        IsStVtoF5 = this.IsStVtoF5,
        IsKlpVtoF5 = this.IsKlpVtoF5,
        IsDiskVtoF5 = this.IsDiskVtoF5,
        IsTimeAooVtoF5 = this.IsTimeAooVtoF5,
        IsDateAroF5 = this.IsDateAroF5,
        IsDate1200F5 = this.IsDate1200F5,
        IsDateApr1F5 = this.IsDateApr1F5,
        IsDateAooF5 = this.IsDateAooF5,
        DateBeginAroF5 = this.DateBeginAroF5,
        DateEndAroF5 = this.DateEndAroF5,
        DateBegin1200F5 = this.DateBegin1200F5,
        DateEnd1200F5 = this.DateEnd1200F5,
        DateBeginApr1F5 = this.DateBeginApr1F5,
        DateEndApr1F5 = this.DateEndApr1F5,
        DateBeginAooF5 = this.DateBeginAooF5,
        DateEndAooF5 = this.DateEndAooF5,
        AroF5Item = Convert.ToString(this.SelAroF5Item.Row["StrSql"]),
        Stan1200F5Item = Convert.ToString(this.sel1200F5Item.Row["StrSql"]),
        TolsF5Item = Convert.ToString(this.SelTolsF5Item.Row["StrSql"]),
        BrgApr1F5Item = Convert.ToString(this.SelBrgApr1F5Item.Row["StrSql"]),
        BrgVtoF5Item = Convert.ToString(this.SelBrgVtoF5Item.Row["StrSql"]),
        BrgAvoF5Item = Convert.ToString(this.SelBrgAvoF5Item.Row["StrSql"]),
        AooF5Item = Convert.ToString(this.SelAooF5Item.Row["StrSql"]),
        AvoF5Item = Convert.ToString(this.SelAvoF5Item.Row["StrSql"]),
        ShirApr1F5Item = Convert.ToString(this.SelShirApr1F5Item.Row["StrSql"]),
        DiskVtoF5Item = Convert.ToString(this.SelDiskVtoF5Item.Row["StrSql"]),
        //Apr1F5Width = this.Apr1F5Width,
        AooF5MgOFrom = this.AooF5MgOFrom,
        AooF5MgOTo = this.AooF5MgOTo,
        AooF5PppFrom = this.AooF5PppFrom,
        AooF5PppTo = this.AooF5PppTo,
        AooF5WgtCoverFrom = this.AooF5WgtCoverFrom,
        AooF5WgtCoverTo = this.AooF5WgtCoverTo,
        VtoF5Stend = this.VtoF5Stend,
        VtoF5Cap = this.VtoF5Cap,
        VtoF5TimeAooVto = this.VtoF5TimeAooVto,
        IsDateAvoLstF5 = this.IsDateAvoLstF5,
        DateBeginAvoLstF5 = this.DateBeginAvoLstF5,
        DateEndAvoLstF5 = this.DateEndAvoLstF5,
        BrgAooF5Item = Convert.ToString(this.SelBrgAooF5Item.Row["StrSql"]),
        IsBrgAooF5 = this.IsBrgAooF5,
        StrThicknessSql = Convert.ToString(this.SelThicknessItemF2.Row["StrSql"])
      };

      var sp = new Db.SgpTo2CatFinCut();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteSgpTo2CatFinCutCommand(Object parameter)
    {
      return true;
    }

    public ICommand SgpTo3CatFinCutCommand => sgpTo3CatFinCutCommand ?? (sgpTo3CatFinCutCommand = new DelegateCommand<Object>(ExecuteSgpTo3CatFinCutCommand, CanExecuteSgpTo3CatFinCutCommand));

    private void ExecuteSgpTo3CatFinCutCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.SgpTo3CatFinCutSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.SgpTo3CatFinCutDest;

      var rptParam = new Db.SgpTo3CatFinCutRptParam(src, dst)
      {
        PathScriptsDir = Smv.Utils.Etc.StartPath + ModuleConst.ScriptsFolder,
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd,
        TypeFilter = this.SelectedTabIndexF5,
        ListStendF5 = this.ListStendF5,
        TypeListValueF5 = this.typeListValueF5,
        IsAroF5 = this.IsAroF5,
        Is1200F5 = this.Is1200F5,
        IsApr1F5 = this.IsApr1F5,
        IsAooF5 = this.IsAooF5,
        IsVtoF5 = this.IsVtoF5,
        IsAvoF5 = this.IsAvoF5,
        IsMgOF5 = this.IsMgOF5,
        IsPppF5 = this.IsPppF5,
        IsWgtCoverF5 = this.IsWgtCoverF5,
        IsStVtoF5 = this.IsStVtoF5,
        IsKlpVtoF5 = this.IsKlpVtoF5,
        IsDiskVtoF5 = this.IsDiskVtoF5,
        IsTimeAooVtoF5 = this.IsTimeAooVtoF5,
        IsDateAroF5 = this.IsDateAroF5,
        IsDate1200F5 = this.IsDate1200F5,
        IsDateApr1F5 = this.IsDateApr1F5,
        IsDateAooF5 = this.IsDateAooF5,
        DateBeginAroF5 = this.DateBeginAroF5,
        DateEndAroF5 = this.DateEndAroF5,
        DateBegin1200F5 = this.DateBegin1200F5,
        DateEnd1200F5 = this.DateEnd1200F5,
        DateBeginApr1F5 = this.DateBeginApr1F5,
        DateEndApr1F5 = this.DateEndApr1F5,
        DateBeginAooF5 = this.DateBeginAooF5,
        DateEndAooF5 = this.DateEndAooF5,
        AroF5Item = Convert.ToString(this.SelAroF5Item.Row["StrSql"]),
        Stan1200F5Item = Convert.ToString(this.sel1200F5Item.Row["StrSql"]),
        TolsF5Item = Convert.ToString(this.SelTolsF5Item.Row["StrSql"]),
        BrgApr1F5Item = Convert.ToString(this.SelBrgApr1F5Item.Row["StrSql"]),
        BrgVtoF5Item = Convert.ToString(this.SelBrgVtoF5Item.Row["StrSql"]),
        BrgAvoF5Item = Convert.ToString(this.SelBrgAvoF5Item.Row["StrSql"]),
        AooF5Item = Convert.ToString(this.SelAooF5Item.Row["StrSql"]),
        AvoF5Item = Convert.ToString(this.SelAvoF5Item.Row["StrSql"]),
        ShirApr1F5Item = Convert.ToString(this.SelShirApr1F5Item.Row["StrSql"]),
        DiskVtoF5Item = Convert.ToString(this.SelDiskVtoF5Item.Row["StrSql"]),
        //Apr1F5Width = this.Apr1F5Width,
        AooF5MgOFrom = this.AooF5MgOFrom,
        AooF5MgOTo = this.AooF5MgOTo,
        AooF5PppFrom = this.AooF5PppFrom,
        AooF5PppTo = this.AooF5PppTo,
        AooF5WgtCoverFrom = this.AooF5WgtCoverFrom,
        AooF5WgtCoverTo = this.AooF5WgtCoverTo,
        VtoF5Stend = this.VtoF5Stend,
        VtoF5Cap = this.VtoF5Cap,
        VtoF5TimeAooVto = this.VtoF5TimeAooVto,
        IsDateAvoLstF5 = this.IsDateAvoLstF5,
        DateBeginAvoLstF5 = this.DateBeginAvoLstF5,
        DateEndAvoLstF5 = this.DateEndAvoLstF5,
        BrgAooF5Item = Convert.ToString(this.SelBrgAooF5Item.Row["StrSql"]),
        IsBrgAooF5 = this.IsBrgAooF5,
        StrThicknessSql = Convert.ToString(this.SelThicknessItemF2.Row["StrSql"])
      };

      var sp = new Db.SgpTo3CatFinCut();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }

    private bool CanExecuteSgpTo3CatFinCutCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand LasScrAfterFinCutCommand => lasScrAfterFinCutCommand ?? (lasScrAfterFinCutCommand = new DelegateCommand<Object>(ExecuteLasScrAfterFinCutCommand, CanExecuteLasScrAfterFinCutCommand));

    private void ExecuteLasScrAfterFinCutCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.LasScrAfterFinCutSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.LasScrAfterFinCutDest;

      var rptParam = new Db.LasScrAfterFinCutRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.LasScrAfterFinCut();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteLasScrAfterFinCutCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand ReasonOfStripBreakageRm1300Command => reasonOfStripBreakageRm1300Command ?? (reasonOfStripBreakageRm1300Command = new DelegateCommand<Object>(ExecuteReasonOfStripBreakageRm1300Command, CanExecuteReasonOfStripBreakageRm1300Command));
    private void ExecuteReasonOfStripBreakageRm1300Command(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.ReasonOfStripBreakageRm1300Source;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.ReasonOfStripBreakageRm1300Dest;

      var rptParam = new Db.ReasonOfStripBreakageRm1300RptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.ReasonOfStripBreakageRm1300();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteReasonOfStripBreakageRm1300Command(Object parameter)
    {
      return true;
    }
    
    public ICommand BalanceWrkTimeCommand => balanceWrkTimeCommand ?? (balanceWrkTimeCommand = new DelegateCommand<Object>(ExecuteBalanceWrkTimeCommand, CanExecuteBalanceWrkTimeCommand));
    private void ExecuteBalanceWrkTimeCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.BalanceWrkTimeSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.BalanceWrkTimeDest;

      var rptParam = new Db.BalanceWrkTimeRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Db.BalanceWrkTime();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteBalanceWrkTimeCommand(Object parameter)
    {
      return true;
    }


    public ICommand LoadTblExceptCommand => loadTblExceptCommand ?? (loadTblExceptCommand = new DelegateCommand<Object>(ExecuteLoadTblExceptCommand, CanExecuteLoadTblExceptCommand));
    private void ExecuteLoadTblExceptCommand(Object parameter)
    {
      DbUtils.LoadTblExcept();
    }
    private bool CanExecuteLoadTblExceptCommand(Object parameter)
    {
      return true;
    }
    public ICommand LoadProdTargetsCommand => loadProdTargetsCommand ?? (loadProdTargetsCommand = new DelegateCommand<Object>(ExecuteLoadProdTargetsCommand, CanExecuteLoadProdTargetsCommand));
    private void ExecuteLoadProdTargetsCommand(Object parameter)
    {
      var rptParam = new LoadProdTargetsRptParam(null, null)
      {
        DateBegin = this.DateBegin,
        CfgFile = ModuleConst.LoadProdTargetsConfig
      };

      var sp = new LoadProdTargets();
      Boolean res = sp.Run(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteLoadProdTargetsCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand MonitorDef2CatCommand => monitorDef2CatCommand ?? (monitorDef2CatCommand = new DelegateCommand<Object>(ExecuteMonitorDef2CatCommand, CanExecuteMonitorDef2CatCommand));
    private void ExecuteMonitorDef2CatCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.MonitorDef2CatSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.MonitorDef2CatDest;

      var rptParam = new MonitorDef2CatRptParam(src, dst)
      {
        DateBegin = this.DateBeginQuart,
        DateEnd = this.DateEndQuart
      };

      DbUtils.SaveDateQuart(this.DateBeginQuart, this.DateEndQuart);

      var sp = new MonitorDef2Cat();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;
      
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteMonitorDef2CatCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand KpaRollingCommand => kpaRollingCommand ?? (kpaRollingCommand = new DelegateCommand<Object>(ExecuteKpaRollingCommand, CanExecuteKpaRollingCommand));
    private void ExecuteKpaRollingCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.KpaRollingSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.KpaRollingDest;

      var rptParam = new KpaRollingRptParam(src, dst)
      {
        IsListStendF4 = this.IsListStendF4,
        DateBegin1F4 = this.DateBegin1F4,
        DateEnd1F4 = this.DateEnd1F4,
        DateBegin2F4 = this.DateBegin2F4,
        DateEnd2F4 = this.DateEnd2F4,
        DateBegin3F4 = this.DateBegin3F4,
        DateEnd3F4 = this.DateEnd3F4,
        ListStendF4 = this.ListStendF4
      };
      
      var sp = new KpaRolling();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;

      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteKpaRollingCommand(Object parameter)
    {
      return true;
    }

    public ICommand MonitorDefCommand => monitorDefCommand ?? (monitorDefCommand = new DelegateCommand<Object>(ExecuteMonitorDefCommand, CanExecuteMonitorDefCommand));
    private void ExecuteMonitorDefCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.MonitorDefSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.MonitorDefDest;

      var rptParam = new MonitorDefRptParam(src, dst)
      {
        DateBegin = this.DateBeginQuart,
        DateEnd = this.DateEndQuart
      };

      DbUtils.SaveDateQuart(this.DateBeginQuart, this.DateEndQuart);

      var sp = new MonitorDef();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;

      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteMonitorDefCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand Lider2CatCommand => lider2CatCommand ?? (lider2CatCommand = new DelegateCommand<Object>(ExecuteLider2CatCommand, CanExecuteLider2CatCommand));
    private void ExecuteLider2CatCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.Lider2CatSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.Lider2CatDest;

      var rptParam = new Lider2CatRptParam(src, dst)
      {
        DateBegin = this.DateBeginQuart,
        DateEnd = this.DateEndQuart
      };

      var sp = new Lider2Cat();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;

      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteLider2CatCommand(Object parameter)
    {
      return true;
    }
    
    public ICommand Defects1StRollCommand => defects1StRollCommand ?? (defects1StRollCommand = new DelegateCommand<Object>(ExecuteDefects1StRollCommand, CanExecuteDefects1StRollCommand));
    private void ExecuteDefects1StRollCommand(Object parameter)
    {
      var src = Etc.StartPath + ModuleConst.Defects1StRollSource;
      var dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.Defects1StRollDest;

      var rptParam = new Defects1StRollRptParam(src, dst)
      {
        DateBegin = this.DateBegin,
        DateEnd = this.DateEnd
      };

      var sp = new Defects1StRoll();
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, rptParam);
      if (!res) return;

      var barEditItem = param as BarEditItem;
      if (barEditItem != null) barEditItem.IsVisible = true;
    }
    private bool CanExecuteDefects1StRollCommand(Object parameter)
    {
      return true;
    }


    #endregion Commands
  }
}
