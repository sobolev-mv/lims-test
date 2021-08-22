using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;


namespace Viz.MagLab.MeasureUnits
{
  internal sealed class ViewModelMeasureIsol : Smv.MVVM.ViewModels.ViewModelBase
  {

    #region Fields
    private uint mpCount = 10;
    private Window win = null;
    private List<decimal?> lstMeasureVal = null;
    private int typeMeasureUnit = -1;
    //private int realIndexMeasurevalue = 0;
    private IMeasureIsolUnit iMeasureUnit = null;
    public ObservableCollection<decimal?> ivalue = new ObservableCollection<decimal?> {null, null, null, null, null, null, null, null, null, null};
    public ObservableCollection<Visibility> visMeasPoint = new ObservableCollection<Visibility> { Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden, Visibility.Hidden};      
    //public System.Windows.Controls.RadioButton rbBrokhaus = null;    

    #endregion Fields

    #region Public Property

    public ObservableCollection<decimal?> Ivalue
    {
      get { return ivalue; }
      set
      {
        if (value == ivalue) return;
        ivalue = value;
        base.OnPropertyChanged("Ivalue");
      }
    }


    public ObservableCollection<Visibility> VisMeasPoint
    {
      get { return visMeasPoint; }
      set
      {
        if (value == visMeasPoint) return;
        visMeasPoint = value;
        base.OnPropertyChanged("VisMeasPoint");
      }
    }

    #endregion Public Property

    #region Private Method
    private void MeasureEventHandler(Object sender, MeasureEventArgs e)
    {
      //MessageBox.Show(idx.ToString());
      if (this.visMeasPoint[e.IndexMeasureValue] == Visibility.Visible){ 
        this.Ivalue[e.IndexMeasureValue] = e.MeasureValue;
      }
      else{
        int idx = e.IndexMeasureValue;

        while (this.VisMeasPoint[idx] == Visibility.Hidden){ 
          idx++;
          if (idx == this.mpCount - 1)
            idx = 0;
        }

        this.iMeasureUnit.IndexMeasureValue = idx;
        this.Ivalue[idx] = e.MeasureValue;
      }
    }

    private void ClearMeasure()
    {
      this.iMeasureUnit.StopMeasure();
      this.iMeasureUnit.IndexMeasureValue = 0;
      for (int i = 0; i <= this.ivalue.Count - 1; i++) ivalue[i] = null;
    }

    private void WinClosed(object sender, EventArgs e)
    {
      if (this.iMeasureUnit != null){
        this.ClearMeasure();
        this.iMeasureUnit.Close();
        this.iMeasureUnit.MeasuredValue -= MeasureEventHandler;
        this.iMeasureUnit = null;
        GC.Collect();
      }

      win.Closed -= WinClosed;      
    }

    #endregion Private Method
     
    #region Constructor
    internal ViewModelMeasureIsol(Window winDlg, List<decimal?> MeasureVal, List<Boolean> lstVisibleMeasurePoint)
    {
      
      if (lstVisibleMeasurePoint.Count != 10)
        throw new ArgumentOutOfRangeException(lstVisibleMeasurePoint.Count.ToString(), "Кол-во точек измерений тока не верно!");
      
      win = winDlg;
      lstMeasureVal = MeasureVal;
      win.Closed += this.WinClosed;

      
      for (int i = 0; i <= lstVisibleMeasurePoint.Count  - 1; i++){
        this.Ivalue[i] = lstMeasureVal[i];
        if (lstVisibleMeasurePoint[i])
          this.visMeasPoint[i] = Visibility.Visible;
        else
          this.visMeasPoint[i] = Visibility.Hidden;
      }

      //rbBrokhaus = LogicalTreeHelper.FindLogicalNode(this.win, "rbBrokhaus") as System.Windows.Controls.RadioButton;
      //rbBrokhaus.IsChecked = true;
      SelectUnitCommand.Execute(2);

    }
    #endregion Constructor

    #region Commands
    private DelegateCommand<Object> selectUnitCommand;
    private DelegateCommand<Object> clearMeasureCommand;
    private DelegateCommand<Object> okCommand;
    private DelegateCommand<Object> selMeasPosCommand;

    public ICommand SelectUnitCommand
    {
      get{return selectUnitCommand ?? (selectUnitCommand = new DelegateCommand<Object>(ExecuteSelectUnit, CanExecuteSelectUnit));}
    }

    private void ExecuteSelectUnit(Object parameter)
    {
      this.typeMeasureUnit = Convert.ToInt32(parameter);

      if (this.iMeasureUnit != null){
        this.ClearMeasure(); 
        this.iMeasureUnit.Close();
        this.iMeasureUnit.MeasuredValue -= MeasureEventHandler;
        this.iMeasureUnit = null;
        GC.Collect();
      }    

      if (this.typeMeasureUnit == 1)
        this.iMeasureUnit = new JapanIsolUnit(Smv.Utils.Etc.StartPath + ModuleConst.JapanIsolMeasureUnitConfig, this.mpCount);
      else
        this.iMeasureUnit = new BrokhausIsolUnit(Smv.Utils.Etc.StartPath + ModuleConst.BrokhausIsolMeasureUnitConfig, this.mpCount);

        this.iMeasureUnit.MeasuredValue += MeasureEventHandler;
    }

    private bool CanExecuteSelectUnit(Object parameter)
    {
      return true;
    }


    public ICommand ClearMeasureCommand
    {
      get{return clearMeasureCommand ?? (clearMeasureCommand = new DelegateCommand<Object>(ExecuteClearMeasure, CanExecuteClearMeasure));}
    }

    private void ExecuteClearMeasure(Object parameter)
    {
      this.ClearMeasure();
      this.iMeasureUnit.StartMeasure();
    }

    private bool CanExecuteClearMeasure(Object parameter)
    {
      return ((this.iMeasureUnit != null) && (!this.iMeasureUnit.IsError));
    }

    public ICommand OkCommand
    {
      get{return okCommand ?? (okCommand = new DelegateCommand<Object>(ExecuteOk, CanExecuteOk));}
    }

    private void ExecuteOk(Object parameter)
    {
      //this.iMeasureUnit.StopMeasure();
      //this.iMeasureUnit.MeasuredValue -= MeasureEventHandler;
      //this.iMeasureUnit = null;

      for (int i = 0; i <= this.visMeasPoint.Count - 1; i++)
        if (this.visMeasPoint[i] == Visibility.Visible) lstMeasureVal[i] = this.ivalue[i]; 

      (parameter as Window).DialogResult = true;
      (parameter as Window).Close();
    }

    private bool CanExecuteOk(Object parameter)
    {
      return ((this.iMeasureUnit != null) && (!this.iMeasureUnit.IsError));
    }


    public ICommand SelMeasPosCommand
    {
      get{return selMeasPosCommand ?? (selMeasPosCommand = new DelegateCommand<Object>(ExecuteSelMeasPos, CanExecuteSelMeasPos));}
    }

    private void ExecuteSelMeasPos(Object parameter)
    {
      iMeasureUnit.IndexMeasureValue = Convert.ToInt32((parameter as DevExpress.Xpf.Editors.TextEdit).Tag);
      (parameter as DevExpress.Xpf.Editors.TextEdit).Clear();
    }

    private bool CanExecuteSelMeasPos(Object parameter)
    {
      return ((iMeasureUnit != null) && (!iMeasureUnit.IsError));
    }

    #endregion Commands

  }
}
