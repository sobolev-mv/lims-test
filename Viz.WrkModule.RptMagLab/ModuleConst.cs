using System;
using System.Collections.Generic;
using System.Text;

namespace Viz.WrkModule.RptMagLab
{
  public static class ModuleConst
  {

    public const string ModuleId = "166001";

    public const int AccCmdMatTechStep = 101;
    public const int AccCmdAlstIsol = 102;
    public const int AccCmdCzlPack = 103;
    public const int AccCmdCzlLaser = 104;
    public const int AccCmdCzlIsoGo = 105;
    public const int AccCmdCzlFinCut = 106;
    //Группа "Средние эл.маг. св-ва и распределение по маркам ЭАС; Доля ЭАС с высок. МС"
    public const int AccCmdQCzl = 107;
    public const int AccCmdCzlOutLowLevelP1750 = 108;
    public const int AccCmdQCzl_Lst = 112;
    public const int AccCmdCzlOutLowLevelP1750_Lst = 113;
    public const int AccCmdQCzl_NoLst = 114;
    public const int AccCmdCzlOutLowLevelP1750_NoLst = 115;

    //Группа "Эффективность лазерного комплекса"
    public const int AccCmdCzlEfLsr = 109;
    public const int AccCmdCzlEfLsr_Lst = 116;
    public const int AccCmdCzlEfLsr_NoLst = 117;
    public const int AccCmdCzlEfLsr9t = 118;
    public const int AccCmdCzlEfLsr9t_Lst = 119;
    public const int AccCmdCzlEfLsr9t_NoLst = 120;

    //Группа плоскостность, дефекты плоскастности
    public const int AccCmdCzlPlosk = 110;
    public const int AccCmdCzlDefPlosk = 111;


    //Группы 1 уровня
    public const int AccGrpMatTechStep = 200;
    public const int AccGrpAlstIsol = 201;
    public const int AccGrpCzlPackEtc = 202;
    public const int AccGrpLineAoo = 203;
    public const int AccGrpQCzl = 204;
    public const int AccGrpCzlEfLsr = 205;
    public const int AccGrpCzlPlosk = 206;

    //Группы 2 уровня
    public const int AccG2Plosk = 260;

    
    public const string ScriptsFolder = "\\Scripts";
    public const string MatTechStepSource = "\\Xlt\\Viz.WrkModule.RptMagLab-MatTechStep.xltx";
    public const string MatTechStepDest = "\\Viz.WrkModule.RptMagLab-MatTechStep.xlsx";

    public const string AlstIsolSource = "\\Xlt\\Viz.WrkModule.RptMagLab-Alstom.xltx";
    public const string AlstIsolDest = "\\Viz.WrkModule.RptMagLab-Alstom.xlsx";

    public const string CzlPackSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlPack.xltx";
    public const string CzlPackDest = "\\Viz.WrkModule.RptMagLab-CzlPack.xlsx";

    public const string CzlFinCutSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlFinCut.xltx";
    public const string CzlFinCutDest = "\\Viz.WrkModule.RptMagLab-CzlFinCut.xlsx";

    public const string CzlLaserSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlLaser.xltx";
    public const string CzlLaserDest = "\\Viz.WrkModule.RptMagLab-CzlLaser.xlsx";

    public const string CzlIsoGoSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlIsoGo.xltx";
    public const string CzlIsoGoDest = "\\Viz.WrkModule.RptMagLab-CzlIsoGo.xlsx";

    public const string QczlSource = "\\Xlt\\Viz.WrkModule.RptMagLab-QCzl.xltx";
    public const string QczlDest = "\\Viz.WrkModule.RptMagLab-QCzl.xlsx";

    public const string CzlDefPloskSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlDefPlosk.xltx";
    public const string CzlDefPloskDest = "\\Viz.WrkModule.RptMagLab-CzlDefPlosk.xlsx";

    public const string CzlEfLsrSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlEfLsr.xltx";
    public const string CzlEfLsrDest = "\\Viz.WrkModule.RptMagLab-CzlEfLsr.xlsx";

    public const string CzlPloskSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlPlosk.xltx";
    public const string CzlPloskDest = "\\Viz.WrkModule.RptMagLab-CzlPlosk.xlsx";

    public const string CzlOutLowLevelP1750Source = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlOutLowLevelP1750.xltx";
    public const string CzlOutLowLevelP1750Dest = "\\Viz.WrkModule.RptMagLab-CzlOutLowLevelP1750.xlsx";

    public const string CzlEfLsr9tSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlEfLsr9t.xltx";
    public const string CzlEfLsr9tDest = "\\Viz.WrkModule.RptMagLab-CzlEfLsr9t.xlsx";

    public const string CzlLineAooStendSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlLineAooStend.xltx";
    public const string CzlLineAooStendDest = "\\Viz.WrkModule.RptMagLab-CzlLineAooStend.xlsx";

    public const string CzlLineAooCoilSource = "\\Xlt\\Viz.WrkModule.RptMagLab-CzlLineAooCoil.xltx";
    public const string CzlLineAooCoilDest = "\\Viz.WrkModule.RptMagLab-CzlLineAooCoil.xlsx";


  }

                                             
}
