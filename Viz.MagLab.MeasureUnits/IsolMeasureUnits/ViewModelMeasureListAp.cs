using System;
using System.Data;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Grid;

namespace Viz.MagLab.MeasureUnits
{

  class ViewModelMeasureListAp : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields
    private readonly int uType;
    private readonly Window view; 
    private readonly string host;
    private readonly int port;
    private readonly ProgressBarEdit pgb;
    private readonly GridControl gcData;
    private Boolean IsOkEnabled;
    private readonly Boolean isSignalPcSpeaker;
    private readonly Boolean isSignalAudio;

    private readonly System.Windows.Controls.RadioButton rbList;
    private readonly System.Windows.Controls.RadioButton rbAp;
    private readonly DataTable paramData;    

    private decimal? sMass;
    private decimal? sLen;
    private decimal? sWid;
    private decimal? sDens; 
    private readonly Mk4a mk4a;
    private readonly string mD;
    private readonly int mesDevice;
    private readonly DataTable crcftData;
    #endregion

    #region Public Property
    public System.Decimal? Massa
    {
      get{ return sMass; }
      set{
        if (value == sMass) return;
        sMass = value;
        base.OnPropertyChanged("Massa");
      }
    }

    public System.Decimal? Length
    {
      get{ return sLen; }
      set{
        if (value == sLen) return;
        sLen = value;
        base.OnPropertyChanged("Length");
      }
    }

    public System.Decimal? Width
    {
      get{ return sWid; }
      set{
        if (value == sWid) return;
        sWid = value;
        base.OnPropertyChanged("Width");
      }
    }

    public System.Decimal? Density
    {
      get{ return sDens; }
      set{
        if (value == sDens) return;
        sDens = value;
        base.OnPropertyChanged("Density");
      }
    }


    public DataTable ParamData
    {
      get { return this.paramData; }
    }

    #endregion

    #region Private Method
    private void MeasureCompleteEventHandler(Object sender, EventArgs e)
    {
      UpdateErrors();
      this.IsOkEnabled = true;
      CommandManager.InvalidateRequerySuggested();

      if (this.isSignalPcSpeaker){
        Smv.Utils.Sound.Beep(2000, 250);
        Smv.Utils.Sound.Beep(1700, 200);
      }

      if (this.isSignalAudio)
        Smv.Utils.Sound.PlaySoundFile(Smv.Utils.Etc.StartPath + ModuleConst.EndWorkFileSound);
    }

    private void ClosedView(Object sender, EventArgs e)  
    {
      this.paramData.ColumnChanging -= Column_OutValDataChanging;
      paramData.AcceptChanges();
      mk4a.Close();  
    }

    private void ValidateRow(object sender, GridRowValidationEventArgs e)
    {
      //DataRow row = (e.Row as DataRowView).Row;
    }

    private void InvalidRowException(object sender, InvalidRowExceptionEventArgs e) 
    {
      //e.ExceptionMode = ExceptionMode.NoAction;
    }

    private void UpdateErrors()
    {
      gcData.BeginDataUpdate();
      try{
        foreach (DataRow row in (gcData.ItemsSource as DataTable).Rows){
          row.BeginEdit();
          this.UpdateColumnError(row);
          row.EndEdit();
        }
      }
      finally{
        gcData.EndDataUpdate();
      }
    }

    private void UpdateColumnError(DataRow row)
    {
      if ((row["IsValidate"] == DBNull.Value) | (Convert.ToInt32(row["IsValidate"]) == 0) | (row["OutVal"] == DBNull.Value)){
        row.SetColumnError("OutVal", string.Empty);
        return;
      }

      decimal val = Convert.ToDecimal(row["OutVal"]);
      decimal minVal = Convert.ToDecimal(row["MinValue"]);
      decimal maxVal = Convert.ToDecimal(row["MaxValue"]);

      Boolean rez = (val >= minVal) && (val <= maxVal);

      if (!rez)
        row.SetColumnError("OutVal", "Значение параметра должно быть в диапазоне от " + minVal.ToString() + " до " + maxVal.ToString());
      else 
        row.SetColumnError("OutVal", string.Empty);
    }


    private void CellChanged(object sender, CellValueChangedEventArgs e)
    {
      if (e.Column.FieldName == "IsActive")
        return;

      UpdateColumnError((e.Row as DataRowView).Row);
    }

    private void Column_OutValDataChanging(object sender, DataColumnChangeEventArgs e)
    {
      if (e.Column.ColumnName != "OutVal")
        return;

      if (e.ProposedValue == null)
        e.ProposedValue = DBNull.Value;

      //Здесь происходит корректировка измеренных значений
      var fldName = Convert.ToString(e.Row["MeasMl"]);
      crcftData.DefaultView.ApplyDefaultSort = true;
      int i = crcftData.DefaultView.Find(new Object[] {this.mD, fldName, this.uType, this.mesDevice});

      if ((i != -1) && (Convert.ToChar(crcftData.DefaultView[i]["TypCor"]) == 'D')){

        /*Ввод чистого значения и корректирующего коэфф.
        if (fldName == "P1750")
          MessageBox.Show(Convert.ToDecimal(e.ProposedValue).ToString() + " / " + Convert.ToDecimal(crcftData.DefaultView[i]["Corr"]).ToString());
        */

        e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + Convert.ToDecimal(crcftData.DefaultView[i]["Corr"]);
      }
    }

    #endregion

    #region Constructor
    internal ViewModelMeasureListAp(Window winDlg, int uType, DataTable Data, decimal? Mass, decimal? sLen, decimal? sWid, decimal? sDens, string mD, int mesDevice, DataTable crcftData)
    {
      this.uType = uType;
      this.view = winDlg;
      this.paramData = Data;
      this.sLen = sLen;
      this.sWid = sWid;
      this.sDens = sDens;  
      this.IsOkEnabled = false;
      this.mD = mD;
      this.mesDevice = mesDevice;
      this.crcftData = crcftData;

      this.view.Closed += ClosedView;

      this.host = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Smv.Utils.Etc.StartPath + ModuleConst.Mk4aMeasureUnitConfig, "Host");
      this.port = Convert.ToInt32(Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Smv.Utils.Etc.StartPath + ModuleConst.Mk4aMeasureUnitConfig, "Port"));
      this.isSignalPcSpeaker = (Convert.ToInt32(Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Smv.Utils.Etc.StartPath + ModuleConst.Mk4aMeasureUnitConfig, "IsSignalPcSpeaker")) >= 1);
      this.isSignalAudio = (Convert.ToInt32(Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Smv.Utils.Etc.StartPath + ModuleConst.Mk4aMeasureUnitConfig, "IsSignalAudio")) >= 1);

      this.rbList = LogicalTreeHelper.FindLogicalNode(this.view, "rbList") as System.Windows.Controls.RadioButton;
      this.rbAp = LogicalTreeHelper.FindLogicalNode(this.view, "rbAp") as System.Windows.Controls.RadioButton;
      this.pgb = LogicalTreeHelper.FindLogicalNode(this.view, "PgbMeasure") as ProgressBarEdit;
      this.gcData = LogicalTreeHelper.FindLogicalNode(this.view, "GcData") as GridControl;

      this.mk4a = new Mk4a(1024, this.uType, this.view.Dispatcher, pgb, this.host, this.port, this.paramData);
      this.mk4a.MeasureComplete += this.MeasureCompleteEventHandler;
      (this.gcData.View as TableView).ValidateRow += ValidateRow;
      (this.gcData.View as TableView).InvalidRowException += InvalidRowException;
      (this.gcData.View as TableView).CellValueChanged += CellChanged;
      this.paramData.ColumnChanging += Column_OutValDataChanging;

      switch (this.uType){
        case 1:
          this.rbList.IsChecked = true;
          break;
        case 2:
          this.rbAp.IsChecked = true;
          break;
        default:
          
          break;
      }

    }
    #endregion

    #region Commands
    private DelegateCommand<Object> startMeasureCommand;
    private DelegateCommand<Object> okCommand;

    public ICommand StartMeasureCommand
    {
      get{return startMeasureCommand ?? (startMeasureCommand = new DelegateCommand<Object>(ExecuteStartMeasure, CanExecuteStartMeasure));}
    }

    private void ExecuteStartMeasure(Object parameter)
    {
      IsOkEnabled = false;
      this.view.Tag = this.sMass;
      (gcData.ItemsSource as DataTable).AcceptChanges();
      this.mk4a.StartMeasure(this.sMass, this.sLen, this.sWid, this.sDens); 
    }

    private bool CanExecuteStartMeasure(Object parameter)
    {
      return (this.sMass != null) && (this.sMass != 0);
    }


    public ICommand OkCommand
    {
      get{return okCommand ?? (okCommand = new DelegateCommand<Object>(ExecuteOk, CanExecuteOk));}
    }

    private void ExecuteOk(Object parameter)
    {
      paramData.AcceptChanges();
      this.view.DialogResult = true;
      this.view.Close();
    }

    private bool CanExecuteOk(Object parameter)
    {
      return IsOkEnabled;
    }


    #endregion

  }
}
