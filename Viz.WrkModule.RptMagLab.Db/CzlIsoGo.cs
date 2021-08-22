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
  public sealed class CzlIsoGoRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string TechStepInspLot { get; set; }
    public string TechStepPrjJornal { get; set; }

    public CzlIsoGoRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd, string RptTechStepIl, string RptTechStepPj)
      : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.TechStepInspLot = RptTechStepIl;
      this.TechStepPrjJornal = RptTechStepPj;
    }
  }

  public sealed class CzlIsoGo : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      CzlIsoGoRptParam prm = (e.Argument as CzlIsoGoRptParam);
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

    private Boolean RunRpt(CzlIsoGoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      string SqlStmt = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{

        SqlStmt = "SELECT * FROM VIZ_PRN.CZL_ISOGO";
        DbVar.SetString(prm.TechStepInspLot, prm.TechStepPrjJornal);
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef);

        if (odr == null) return false;

        CurrentWrkSheet.Cells[2, 7].Value = string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " - " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        int flds = odr.FieldCount;
        int row = 7;

        while (odr.Read()){
          CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 24]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 24]]);

          for (int i = 0; i < flds; i++)
            CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

          row++;
        }

        CurrentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      finally{
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return Result;
    }


  }






}
