using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptMagLab.Db
{
  public sealed class CzlFinCutRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string TechStepInspLot { get; set; }
    public string TechStepPrjJornal { get; set; }

    public CzlFinCutRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd, string RptTechStepIl, string RptTechStepPj)
           : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.TechStepInspLot = RptTechStepIl;
      this.TechStepPrjJornal = RptTechStepPj;
    }
  }

  public sealed class CzlFinCut : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      CzlFinCutRptParam prm = (e.Argument as CzlFinCutRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        //Здесь формирование самого отчета
        //wrkSheet.Range("A1").Value = prm.ExcelApp.Version;
        //wrkSheet.Range("A2").Value = "asdadsdgsfgsfsg";

        //Здесь визуализация Экселя
        //prm.ExcelApp.ScreenUpdating = true;
        //prm.ExcelApp.Visible = true;
        this.SaveResult(prm);
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally
      {
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        //Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(CzlFinCutRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean result = false;

      int colorRow = 49407;
      const string sqlStmt = "SELECT * FROM VIZ_PRN.CZL_FINCUT ORDER BY MLOCID, TSDATE";

      try{

        DbVar.SetString(prm.TechStepInspLot, prm.TechStepPrjJornal);
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);
        odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, null);

        if (odr == null) return false;

        CurrentWrkSheet.Cells[2, 3].Value2 = $"{dtBegin:dd.MM.yyyy HH:mm:ss}" + " - " + $"{dtEnd:dd.MM.yyyy HH:mm:ss}";

        int flds = odr.FieldCount;
        int row = 7;

        string prevLocId = null;
        const int firstExcelColumn = 1;
        const int lastExcelColumn = 189;

        while (odr.Read()){
          var curLocId = Convert.ToString(odr.GetValue("MLOCID"));
          CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, firstExcelColumn], CurrentWrkSheet.Cells[row + 1, lastExcelColumn]]);

          if (curLocId == prevLocId){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Interior.Pattern = 1;//xlSolid
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Interior.Color = colorRow;
            //===================
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row - 1, firstExcelColumn], CurrentWrkSheet.Cells[row - 1, lastExcelColumn]].Interior.Pattern = 1;//xlSolid
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row - 1, firstExcelColumn], CurrentWrkSheet.Cells[row - 1, lastExcelColumn]].Interior.Color = colorRow;
            colorRow -= 100;
          }

          prevLocId = curLocId;

          for (int i = 0; i < flds; i++)
            CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);

          row++;
        }

        CurrentWrkSheet.Cells[1, 1].Select();
        result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
        result = false;
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


  }






}

