namespace Viz.WrkModule.Isc
{
  public enum TypeTypeMnf
  {
    Viz = 1,
    Nlmk = 2
  };

  public static class ModuleConst
  {
    public const string ModuleId = "175001";

    //Группы 1ур.
    internal enum AccL1Gr
    {
      ShiftRptUo = 10001
    }

    //Группы 2ур.
    internal enum AccL2Gr
    {
      ShiftRptUo = 13000
    }

    //Кнопки запуска отчетов
    internal enum AccRunControl
    {
      ShiftRptUo = 16000
    }

    public const string ScriptsFolder = "\\Scripts";
    public const string ShiftRptSlRuSource = "\\Xlt\\Viz.WrkModule.Isc-Sl-Ru.xltx";
    public const string ShiftRptSlRuDest = "\\Viz.WrkModule.Isc-Sl-Ru.xlsx";
    public const string ShiftRptSlEnSource = "\\Xlt\\Viz.WrkModule.Isc-Sl-En.xltx";
    public const string ShiftRptSlEnDest = "\\Viz.WrkModule.Isc-Sl-En.xlsx";

    public const string ShiftRptCtlRuSource = "\\Xlt\\Viz.WrkModule.Isc-Ctl-Ru.xltx";
    public const string ShiftRptCtlRuDest = "\\Viz.WrkModule.Isc-Ctl-Ru.xlsx";
    public const string ShiftRptCtlEnSource = "\\Xlt\\Viz.WrkModule.Isc-Ctl-En.xltx";
    public const string ShiftRptCtlEnDest = "\\Viz.WrkModule.Isc-Ctl-En.xlsx";

    public const string VizPruductSource = "\\Xlt\\Viz.WrkModule.Isc-VizProduct.xltx";
    public const string VizPruductDest = "\\Viz.WrkModule.Isc-VizProduct.xlsx";

    public const string IscParamConfig = "\\Config\\IscParam.config";

  }
}
