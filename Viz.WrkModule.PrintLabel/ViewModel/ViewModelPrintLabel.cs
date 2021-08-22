using System;
using System.Collections.Generic;
using System.Data;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Core;
using Viz.WrkModule.PrintLabel.Db.DataSets;
using DevExpress.Xpf.Grid;
using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Ribbon;
using Smv.Print.RawData;
using Smv.Utils;
using Viz.WrkModule.PrintLabel.Db;

namespace Viz.WrkModule.PrintLabel
{
  public class ViewModelPrintLabel
  {
    #region Fields
    const string cfieldNameLocNum = "Bezeichnung";
    private const int idListApr = 22;
    private const string cApr12 = "APR12";
    private readonly UserControl usrControl;
    private readonly DsPrintLabel dsPrintLabel = new DsPrintLabel();

    private readonly string labelPrinterName;
    private readonly string apr12LabelFileName;
    private readonly string otherAprLabelFileName;
    private readonly int timeRefresh;
    private readonly Boolean isReadPrintConfig;

    private readonly GridControl dbgMaterial;
    private DataRow currentMatDataRow = null;
    private System.Timers.Timer refreshTimer;
     
    #endregion

    #region Public Property
    public DataTable LstFinishApr => this.dsPrintLabel.LstFinishApr;
    public DataTable AprMat => this.dsPrintLabel.AprMat;
    public virtual string FinishApr { get; set; }
    public virtual Boolean IsRefresh { get; set; }
    public virtual Boolean IsPrintBlankWgtOnStripe { get; set; }
    public virtual Boolean IsEnablePrintBlankWgtOnStripe { get; set; }
    public virtual Boolean IsSelectMany { get; set; }
    #endregion

    #region Protected Method
    protected void OnFinishAprChanged()
    {
      IsEnablePrintBlankWgtOnStripe = string.Equals(FinishApr, cApr12, StringComparison.Ordinal);
      
      if (!string.Equals(FinishApr, cApr12, StringComparison.Ordinal))
        IsPrintBlankWgtOnStripe = false;

      this.dsPrintLabel.AprMat.LoadData(FinishApr);
    }

    protected void OnIsRefreshChanged()
    {
      if (IsRefresh)
      {
        IsSelectMany = false;
        refreshTimer.Start();
      }
      else {
        refreshTimer.Stop();
      }
    }

    protected void OnIsSelectManyChanged()
    {
      dbgMaterial.SelectionMode = IsSelectMany ? MultiSelectMode.MultipleRow : MultiSelectMode.Row;

      if (IsSelectMany)
        IsRefresh = !IsSelectMany;
    }
    #endregion

    #region Private Method
    private void CurrentItemGridChanged(object sender, CurrentItemChangedEventArgs args)
    {
      currentMatDataRow = (args.NewItem as DataRowView)?.Row;
    }

    private Boolean PrintLabel4Apr12(DataRow dtRow, Boolean isPrintBlankWgtOnStripe = false)
    {
      string str2Printer;
      string strFmt = System.IO.File.ReadAllText(Etc.StartPath + "\\Scripts\\" + apr12LabelFileName, Encoding.GetEncoding(1251));

      str2Printer = isPrintBlankWgtOnStripe ? string.Format(strFmt, Convert.ToDecimal(dtRow["Dicke"]), Convert.ToDecimal(dtRow["Breite"]), " ", Convert.ToString(dtRow[cfieldNameLocNum]), Convert.ToString(dtRow[cfieldNameLocNum])) : string.Format(strFmt, Convert.ToDecimal(dtRow["Dicke"]), Convert.ToDecimal(dtRow["Breite"]), Convert.ToInt32(dtRow["Gew"]), Convert.ToString(dtRow[cfieldNameLocNum]), Convert.ToString(dtRow[cfieldNameLocNum]));
      return RawPrinterHelper.SendStringToPrinter(labelPrinterName, str2Printer);

    }
    private void PrintLabel4OtherApr()
    {
      string strFmt = System.IO.File.ReadAllText(Etc.StartPath + "\\Scripts\\" + otherAprLabelFileName, Encoding.GetEncoding(1251));
      string str2Printer = string.Format(strFmt, Convert.ToString(currentMatDataRow[cfieldNameLocNum]), Convert.ToString(currentMatDataRow[cfieldNameLocNum]));
      Boolean res = RawPrinterHelper.SendStringToPrinter(labelPrinterName, str2Printer);

      if (res)
        DXMessageBox.Show(Application.Current.Windows[0], "Задание на печать отправлено в очередь принтера успешно.", "Печать этикетки", MessageBoxButton.OK, MessageBoxImage.Information);
      else
        DXMessageBox.Show(Application.Current.Windows[0], "Во время печати возникла ошибка!", "Ошибка печати этикетки", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void NavigateRowGrid(GridControl grid, string fieldName, string fieldValue)
    {
      //Здесь точное позиционирование на запись в случае поиска по нoмеру образца
      if (grid.VisibleRowCount > 0){
        int rowHandle = grid.View.FocusedRowHandle + 1;

        while (Convert.ToString(grid.GetCellValue(rowHandle, fieldName)) != fieldValue && grid.IsValidRowHandle(rowHandle))
          rowHandle++;

        if (grid.IsValidRowHandle(rowHandle)){
          //this.gcSamples.View.FocusedColumn = grid.Columns["UnitPrice"];
          grid.View.FocusedRowHandle = rowHandle;
        }
      }
    }

    private void RefreshMat()
    {
      string fieldValue = null;

      Boolean isPrevStateEmpty = (dsPrintLabel.AprMat.Rows.Count == 0) | (currentMatDataRow == null);

      if (!isPrevStateEmpty)
        fieldValue = Convert.ToString(currentMatDataRow[cfieldNameLocNum]);

      this.dsPrintLabel.AprMat.LoadData(FinishApr);
      
      if (!isPrevStateEmpty)
        NavigateRowGrid(dbgMaterial, cfieldNameLocNum, fieldValue);
    }

    private void OnTimedEvent(object source, ElapsedEventArgs e)
    {
      if (string.IsNullOrEmpty(FinishApr))
        return;


      usrControl.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => RefreshMat()));
    }

    private void Print()
    {
      Boolean res = false;

      if ((dbgMaterial.SelectedItems.Count == 0) && (currentMatDataRow != null))
        res = PrintLabel4Apr12(currentMatDataRow, IsPrintBlankWgtOnStripe);
      else
        foreach (var item  in dbgMaterial.SelectedItems)
          res = PrintLabel4Apr12((item as DataRowView)?.Row, IsPrintBlankWgtOnStripe);

      if (res)
        DXMessageBox.Show(Application.Current.Windows[0], "Задание на печать отправлено в очередь принтера успешно.", "Печать этикетки", MessageBoxButton.OK, MessageBoxImage.Information);
      else
        DXMessageBox.Show(Application.Current.Windows[0], "Во время печати возникла ошибка!", "Ошибка печати этикетки", MessageBoxButton.OK, MessageBoxImage.Error);
    }
    #endregion

    #region Constructor
    public ViewModelPrintLabel(UserControl control, Object mainWindow)
    {
      usrControl = control;
      this.dsPrintLabel.LstFinishApr.LoadData(idListApr);

      //Читаем параметры конфиг файла для настройки параметров печати этикеток
      try{
        labelPrinterName = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.PrintLabelParamConfig, "LabelPrinterName");
        apr12LabelFileName = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.PrintLabelParamConfig, "Apr12LabelFileName");
        otherAprLabelFileName = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.PrintLabelParamConfig, "OtherAprLabelFileName");
        timeRefresh = Convert.ToInt32(Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.PrintLabelParamConfig, "TimeRefresh"));

        isReadPrintConfig = true;
      }
      catch (Exception){
        DXMessageBox.Show(Application.Current.Windows[0], "Ошибка при чтении конфигурационных параметров печати!", "Ошибка конфигурации", MessageBoxButton.OK, MessageBoxImage.Error);
      }

      dbgMaterial = LogicalTreeHelper.FindLogicalNode(this.usrControl, "GcMat") as GridControl;
      if (this.dbgMaterial != null)
        this.dbgMaterial.CurrentItemChanged += CurrentItemGridChanged;

      refreshTimer = new System.Timers.Timer(timeRefresh);
      refreshTimer.Elapsed += OnTimedEvent;
      refreshTimer.Start();
      this.IsRefresh = true;
      this.IsPrintBlankWgtOnStripe = this.IsEnablePrintBlankWgtOnStripe = false;
    }
    #endregion

    #region Command
    public void ShowMat()
    {
      if (IsRefresh)
        refreshTimer.Stop();

      this.dsPrintLabel.AprMat.LoadData(FinishApr);

      if (IsRefresh)
        refreshTimer.Start();
    }

    public bool CanShowMat()
    {
      return (!string.IsNullOrEmpty(FinishApr)) && !IsRefresh;
    }

    public void PrintLabel()
    {

      //var view = new ViewDlgWgWd();
      //view.ShowDialog();
      Print();

      
      /*
      if (string.Equals(FinishApr, cApr12, StringComparison.Ordinal))
        PrintLabel4Apr12();
      else
        PrintLabel4OtherApr();
       */
    }

    public bool CanPrintLabel()
    {
      return (dsPrintLabel.AprMat.Rows.Count != 0) && isReadPrintConfig;
    }


    #endregion
  }
}
