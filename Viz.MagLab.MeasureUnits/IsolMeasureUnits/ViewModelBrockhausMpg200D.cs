using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Editors;
using Smv.App.Config;
using Smv.Utils;

namespace Viz.MagLab.MeasureUnits
{


  public class ViewModelBrockhausMpg200D
  {
    #region Fields
    private readonly DataTable tblCoileSystem;
    private readonly int uType;
    private readonly decimal thicknessNominal;
    private readonly string sampNum;
    private readonly Window view;
    private readonly string host;
    private readonly int port;
    private readonly int readTimeout;
    private readonly Boolean isSignalPcSpeaker;
    private readonly Boolean isSignalAudio;
    private Boolean isOkEnabled;
    private Boolean isStartMeasureEnabled = true;

    private readonly RadioButton rbList;
    private readonly RadioButton rbAp;
    private readonly ProgressBarEdit pgb;
    private readonly ComboBoxEdit cbeCoileName;

    private readonly BrockhausMpg200D mpg200D;
    private Dictionary<string, decimal> resData;
    #endregion

    #region Public Property 
    public DataTable MlMpg200D => tblCoileSystem;
    public virtual DataRowView CoileNameItem { get; set; }
    public virtual decimal? SampWeight { get; set; }
    public virtual int SampLength { get; set; }
    public virtual int SampWidth { get; set; }
    public virtual decimal SampDensity { get; set; }
    public virtual int SampQuantity { get; set; }
    public virtual decimal? B100 { get; set; }
    public virtual decimal? B800 { get; set; }
    public virtual decimal? B2500 { get; set; }
    public virtual decimal? P1550 { get; set; }
    public virtual decimal? P1750 { get; set; }
    public virtual string State1 { get; set; }
    public virtual string State2 { get; set; }
    public virtual string State3 { get; set; }
    public virtual string State4 { get; set; }
    public virtual string State5 { get; set; }
    public virtual string ErrorMsg { get; set; }
    #endregion

    #region Private Method 
    private void ClosedView(Object sender, EventArgs e)
    {
      mpg200D.Close();
    }

    private void ActivatedView(Object sender, EventArgs e)
    {
      cbeCoileName.SelectedIndex = 0;
      SampLength = Convert.ToInt32(CoileNameItem.Row["LengthSmp"]);
      SampWidth = Convert.ToInt32(CoileNameItem.Row["WidthSmp"]);
      SampDensity = Convert.ToDecimal(CoileNameItem.Row["Density"]);
      SampQuantity = Convert.ToInt32(CoileNameItem.Row["Quantity"]);
    }

    private void ClearState()
    {
      State1 = State2 = State3 = State4 = State5 = ErrorMsg = String.Empty;
      B100 = B800 = B2500 = P1550 = P1750 = null;
    }
    #endregion

    #region Constructor
    public ViewModelBrockhausMpg200D(Window winDlg, int uType, decimal thicknessNominal, string sampNum, DataTable tblCoileSystem, Dictionary<string, decimal> resData)
    {
      this.tblCoileSystem = tblCoileSystem;
      this.uType = uType;
      view = winDlg;
      this.thicknessNominal = thicknessNominal;
      this.sampNum = sampNum;
      this.resData = resData;

      view.Closed += ClosedView;
      view.Activated += ActivatedView;

      host = ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.Mpg200DMeasureUnitConfig, "Host");
      port = Convert.ToInt32(ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.Mpg200DMeasureUnitConfig, "Port"));
      readTimeout = Convert.ToInt32(ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.Mpg200DMeasureUnitConfig, "ReadTimeout"));
      isSignalPcSpeaker = (Convert.ToInt32(ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.Mk4aMeasureUnitConfig, "IsSignalPcSpeaker")) >= 1);
      isSignalAudio = (Convert.ToInt32(ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.Mk4aMeasureUnitConfig, "IsSignalAudio")) >= 1);

      rbList = LogicalTreeHelper.FindLogicalNode(view, "rbList") as RadioButton;
      rbAp = LogicalTreeHelper.FindLogicalNode(view, "rbAp") as RadioButton;
      pgb = LogicalTreeHelper.FindLogicalNode(view, "PgbMeasure") as ProgressBarEdit;
      cbeCoileName = LogicalTreeHelper.FindLogicalNode(view, "cbeCoileName") as ComboBoxEdit;

      switch (this.uType){
        case 1:
          rbList.IsChecked = true;
          break;
        case 2:
          rbAp.IsChecked = true;
          break;
      }

      mpg200D = new BrockhausMpg200D(host, port, readTimeout, new MessageInfo(view), 2000);
    }

        #endregion

    #region Command 
    public async void StartMeasure()
    {

      Task<BrockhausMpg200D.MeasurementResult> taskRes = null;
      BrockhausMpg200D.SampleTypeMeasure typeMeasure;

      ClearState();

      if (uType == 1)
        typeMeasure = BrockhausMpg200D.SampleTypeMeasure.Sheet;
      else if (uType == 2)
        typeMeasure = BrockhausMpg200D.SampleTypeMeasure.Epstein;
      else{
        DXMessageBox.Show(view, "Тип измерения не определен!", "Ошибка типа измерения", MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }

      var coileSystem = Convert.ToString(CoileNameItem.Row["CoilName"]);

      isOkEnabled = isStartMeasureEnabled = false;
      
      mpg200D.Close();

      if (!mpg200D.Connect()){
        DXMessageBox.Show(view, mpg200D.LastError, "Ошибка соединения", MessageBoxButton.OK, MessageBoxImage.Error);
        ErrorMsg = mpg200D.LastError;
        isStartMeasureEnabled = true;
        return;
      }

      pgb.Visibility = Visibility.Visible;
      pgb.StyleSettings = new ProgressBarMarqueeStyleSettings();

      taskRes = mpg200D.RunMeasurement(sampNum, typeMeasure, SampWeight.Value / 1000, SampDensity, SampLength, SampWidth, thicknessNominal, SampQuantity, coileSystem);
      await taskRes;

      pgb.StyleSettings = new ProgressBarStyleSettings();
      pgb.Visibility = Visibility.Hidden;
      mpg200D.Close();

      if (this.isSignalPcSpeaker){
        Sound.Beep(2000, 250);
        Sound.Beep(1700, 200);
      }

      if (this.isSignalAudio)
        Sound.PlaySoundFile(Etc.StartPath + ModuleConst.EndWorkFileSound);


      ErrorMsg = taskRes.Result.CodeError + ":";
      B100 = taskRes.Result.B100;
      B800 = taskRes.Result.B800;
      B2500 = taskRes.Result.B2500;
      P1550 = taskRes.Result.P1550;
      P1750 = taskRes.Result.P1750;
      State1 = taskRes.Result.State1;
      State2 = taskRes.Result.State2;
      State3 = taskRes.Result.State3;
      State4 = taskRes.Result.State4;
      State5 = taskRes.Result.State5;
      
      isOkEnabled = (taskRes.Result.CodeError == 0);
      isStartMeasureEnabled = true;
      CommandManager.InvalidateRequerySuggested();
    }

    public bool CanStartMeasure()
    {
      return isStartMeasureEnabled;
    }

    public void ReturnResultMeasurement()
    {
      resData.Clear();
      resData.Add("B100", B100.Value);
      resData.Add("B800", B800.Value);
      resData.Add("B2500", B2500.Value);
      resData.Add("P1550", P1550.Value);
      resData.Add("P1750", P1750.Value);
      resData.Add("Weight", SampWeight.Value);

      view.DialogResult = true;
      view.Close();
    }

    public bool CanReturnResultMeasurement()
    {
      return isOkEnabled;
    }
    #endregion

    }


}
