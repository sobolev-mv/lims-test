using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Viz.WrkModule.MagLab
{

  public enum MlMeasureDevice{
    Ui5099 = 1,
    Mk4A = 2,
    Mpg200D = 3
  };

  public enum MlTypeCorrect
  {
    General = 'G',
    Dievice = 'D'
  };

  public static class ModuleConst
  {

    public const string ModuleId = "165001";
    
    public const int AccCmdViewSamples = 2;
    public const int AccCmdFindSamples = 3;
    public const int AccCmdSaveChangedData = 4;
    public const int AccCmdMeasureSample = 5;
    public const int AccCmdSampleToMes = 6;
    public const int AccCmdSampleToMeasure = 7;
    public const int AccCmdCopyProp = 8;
    public const int AccCmdViewPropProbe = 9;
    public const int AccCmdEditSample = 10;
    public const int AccCmdEditProbe = 11;
    public const int AccCmdDeleteSample = 12;
    public const int AccCmdValidateSample = 13;
    public const int AccCmdChangeStatFlag = 14;
    public const int AccCmdCheckS2L = 15;
    public const int AccCmdUnCheckS2L = 16;
    public const int AccCmdCopyPropS2L = 17;
    public const int AccCmdFr20Import = 18;
    public const int AccCmdStZap = 19;
    public const int AccCmdMesurCof = 20;
    public const int AccCmdCopyApstProp = 21;
    public const int AccCmdSiemensSample = 22;
    public const int AccCmdAdgRpt = 23;
    public const int AccCmdCopyApstCmpCoil = 24;

    public const string MagLabConfig = "\\Config\\MagLab.config";
    public const string Fr20UsbIsolMeasureUnitConfig = "\\Config\\Fr20UsbIsolMeasureUnit.config";  

    public const string AdgRptScript1 = "\\Scripts\\AdgRptScript1.sql";
    public const string AdgRptScript2 = "\\Scripts\\AdgRptScript2.sql";
    public const string AdgRptScript3 = "\\Scripts\\AdgRptScript3.sql";
  }


}
