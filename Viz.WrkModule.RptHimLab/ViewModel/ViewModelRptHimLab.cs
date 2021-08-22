using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Controls;
using Microsoft.Win32;
using DevExpress.Xpf.Editors.Settings;
using DevExpress.Xpf.Bars;
using Smv.Utils;
using Viz.DbApp.Psi;


namespace Viz.WrkModule.RptHimLab
{
  internal sealed class ViewModelRptHimLab : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields
    private readonly Smv.Xls.XlsInstanceBackgroundReport rpt;
    private readonly System.Windows.Controls.UserControl usrControl;
    private readonly Object param;
    private DateTime dateBegin;
    private DateTime dateEnd;
    private readonly DevExpress.Xpf.LayoutControl.LayoutGroup lg;

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

    #endregion 

    #region Private Method
    private void RunXlsRptCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      var barEditItem = param as BarEditItem;
      if (barEditItem != null) 
        barEditItem.IsVisible = false;
    }

    private void LayoutGroupExpanded(object sender, EventArgs e)
    {
      var layoutGroup = sender as DevExpress.Xpf.LayoutControl.LayoutGroup;
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
            break;
          case 5:
            break;
          default:
            break;
        }
      }
    }

    private void LayoutGroupCollapsed(object sender, EventArgs e)
    {
      var layoutGroup = sender as DevExpress.Xpf.LayoutControl.LayoutGroup;
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
            break;
          default:
            break;
        }
      }
    }

    #endregion

    #region Constructor
    internal ViewModelRptHimLab(System.Windows.Controls.UserControl control, Object Param)
    {
      param = Param;
      rpt = new Smv.Xls.XlsInstanceBackgroundReport();
      usrControl = control;
      DateBegin = DateTime.Today;
      DateEnd = DateTime.Today;

      //Группы 1-уровня
      for (int i = ModuleConst.AccGrpIsolProp; i < ModuleConst.AccGrpIsolProp + 1; i++){
        this.lg = LogicalTreeHelper.FindLogicalNode(this.usrControl, "Lg" + ModuleConst.ModuleId + "_" + i) as DevExpress.Xpf.LayoutControl.LayoutGroup;

        if (this.lg != null){
          if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
            lg.Visibility = Visibility.Hidden;
            continue;
          }

          this.lg.Expanded += LayoutGroupExpanded;
          this.lg.Collapsed += LayoutGroupCollapsed;
        }
      }

      /*
      //Группы 2-уровня
      for (int i = ModuleConst.AccG2Plosk; i < ModuleConst.AccG2Plosk + 1; i++)
      {
        var uie = LogicalTreeHelper.FindLogicalNode(this.usrControl, "G2" + ModuleConst.ModuleId + "_" + i) as UIElement;

        if (uie != null)
        {
          if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId))
          {
            uie.Visibility = Visibility.Hidden;
            continue;
          }
        }
      }
      */
 
      //Делаем controls невидимыми
      for (int i = ModuleConst.AccCmdIsolProp01; i < ModuleConst.AccCmdIsolProp03 + 1; i++){
        var btn = LogicalTreeHelper.FindLogicalNode(this.usrControl, "b" + ModuleConst.ModuleId + "_" + i) as UIElement;

        if (btn == null) continue;

        if (!DbApp.Psi.Permission.GetPermissionForModuleUif(i, ModuleConst.ModuleId)){
          btn.Visibility = Visibility.Hidden;
          continue;
        }
      }

      //DxInfo.ShowDxBoxInfo("Внимание", "Файлы созданных отчетов будут находиться в папке Документы (Мои документы)!", MessageBoxImage.Information);
    }
    #endregion

    #region Commands
    //--Для отчета Карениной ХЛ "Диэлектрич. свойство покрытия"
    private DelegateCommand<Object> showListRptCommand;
    private DelegateCommand<Object> himLabIsolPropCommand;

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

    public ICommand HimLabIsolPropCommand
    {
      get { return himLabIsolPropCommand ?? (himLabIsolPropCommand = new DelegateCommand<Object>(ExecuteHimLabIsolProp, CanExecuteHimLabIsolProp)); }
    }

    private void ExecuteHimLabIsolProp(Object parameter)
    {
      string src = null;
      string dst = null;

      int subRpt = Convert.ToInt32(parameter);
      switch (subRpt){
        case 1:
          src = Smv.Utils.Etc.StartPath + ModuleConst.HimLabIsolProp01Source;
          dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.HimLabIsolProp01Dest;
          break;
        case 2:
          src = Smv.Utils.Etc.StartPath + ModuleConst.HimLabIsolProp02Source;
          dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.HimLabIsolProp02Dest;
          break;
        case 3:
          src = Smv.Utils.Etc.StartPath + ModuleConst.HimLabIsolProp03Source;
          dst = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + ModuleConst.HimLabIsolProp03Dest;
          break;
        default:
          Console.WriteLine("Default case");
          break;
      }

      var sp = new Db.HimLabIsolProp();
      var res = sp.RunXls(rpt, RunXlsRptCompleted, new Db.HimLabIsolPropRptParam(src, dst, this.DateBegin, this.DateEnd, subRpt));
      if (res){
        var barEditItem = param as BarEditItem;
        if (barEditItem != null) barEditItem.IsVisible = true;
      }      
    }

    private bool CanExecuteHimLabIsolProp(Object parameter)
    {
      int subRpt = Convert.ToInt32(parameter);
      switch (subRpt){
        case 1:
          return true; 
        case 2:
          return true; 
        case 3:
          return true; 
        default:
          return false;
      }
    }    


    #endregion


  }
}
