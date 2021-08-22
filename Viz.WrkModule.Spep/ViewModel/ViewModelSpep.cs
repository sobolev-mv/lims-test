using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.ComponentModel;
using System.Collections.ObjectModel;
using Smv.MVVM.Commands;
using DevExpress.Xpf.Bars;
using Viz.WrkModule.Spep.Db;


namespace Viz.WrkModule.Spep
{
  internal sealed class ViewModelSpep : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields
    private Object param; 
    private Smv.Xls.XlsInstanceBackgroundReport rpt;
    private System.Windows.Controls.UserControl usrControl;
    private Boolean isSendTo = true;
    private Boolean isAutomat = true;
    private DateTime spepDate;
    private ObservableCollection<SpepStageResult> spepStageResultCollect;
    #endregion Fields

    #region Public Property
    public System.DateTime SpepDate
    {
      get { return spepDate; }
      set
      {
        if (value == spepDate) return;
        spepDate = value;
        base.OnPropertyChanged("SpepDate");
      }
    }

    public Boolean IsSendTo
    {
      get { return isSendTo; }
      set
      {
        if (value == isSendTo) return;
        isSendTo = value;
        base.OnPropertyChanged("IsSendTo");
      }
    }

    public Boolean IsAutomat
        {
      get { return isAutomat; }
      set
      {
        if (value == isAutomat) return;
        isAutomat = value;
        base.OnPropertyChanged("IsAutomat");
      }
    }

    public ObservableCollection<SpepStageResult> SpepStageResultCollect
    {
      get { return spepStageResultCollect; }
      set
      {
        if (value == spepStageResultCollect) return;
        spepStageResultCollect = value;
        base.OnPropertyChanged("SpepStageResultCollect");
      }
    }


    #endregion Public Property

    #region Private Method
    private void RunXlsRptCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      GC.Collect();
      (param as BarEditItem).IsVisible = false;
    }
    #endregion

    #region Constructor
    internal ViewModelSpep(System.Windows.Controls.UserControl control, Object Param)
    {
      param = Param;
      rpt = new Smv.Xls.XlsInstanceBackgroundReport();
      usrControl = control;
      //pgbEdit = LogicalTreeHelper.FindLogicalNode(this.usrControl, "pgbRpt") as DevExpress.Xpf.Editors.ProgressBarEdit;
      spepDate = DateTime.Today;
      spepStageResultCollect = new ObservableCollection<SpepStageResult>();
    }
    #endregion Constructor

    #region Commands
    private DelegateCommand<Object> runSpepCommand;


    public ICommand RunSpepCommand
    {
      get{return runSpepCommand ?? (runSpepCommand = new DelegateCommand<Object>(ExecuteRunSpep, CanExecuteRunSpep));}
    }

    private void ExecuteRunSpep(Object parameter)
    {
      string cfg = Smv.Utils.Etc.StartPath + ModuleConst.SpepConfig;
      string src = Smv.Utils.Etc.StartPath + ModuleConst.SpepSource;

      Db.SpepRpt sp = new Db.SpepRpt();
      SpepRptParam spPram = new Db.SpepRptParam(src, cfg, spepDate, isSendTo, spepStageResultCollect)
      {
        IsAutomat = this.IsAutomat
      };
      Boolean res = sp.RunXls(rpt, RunXlsRptCompleted, spPram);


      if (res){
        (param as BarEditItem).IsVisible = true;
      }
    }

    private bool CanExecuteRunSpep(Object parameter)
    {
      return true;
    }    

    #endregion


  }
}
