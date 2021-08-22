using System;
using System.Collections.Generic;
using System.Data;
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

namespace Viz.WrkModule.RptOpr.Db
{
  public sealed class CuttingMatScrapUoRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public CuttingMatScrapUoRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class CuttingMatScrapUo : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as CuttingMatScrapUoRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        //Здесь визуализация Экселя
        //prm.ExcelApp.ScreenUpdating = true;
        //prm.ExcelApp.Visible = true;
        this.SaveResult(prm);
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
      }
      finally{
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

    private Boolean RunRpt(CuttingMatScrapUoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);

        //CurrentWrkSheet.Range["H1", "L1"].ClearContents();
        //CurrentWrkSheet.Range["H1:L1"].ClearContents();
        CurrentWrkSheet.Cells[1, 4].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";
        
        const string sqlStmt0 = "VIZ_PRN.Schrott_UO.preSchrott_UO";
        Odac.ExecuteNonQuery(sqlStmt0, CommandType.StoredProcedure, false, null);


        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.UO_SCRAB";
        odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);
       
        if (odr != null){
          int flds = odr.FieldCount;
          
          int inRow1 = 5;
          int inRowInsert1 = 7;

          const int firstExcelColumn = 1;
          const int lastExcelColumn = 17;

          while (odr.Read()){

            if (inRow1 == inRowInsert1){
              CurrentWrkSheet.Rows[inRow1].Insert();
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[inRow1 - 1, firstExcelColumn], CurrentWrkSheet.Cells[inRow1 - 1, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[inRow1, firstExcelColumn], CurrentWrkSheet.Cells[inRow1, lastExcelColumn]]);
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[inRow1, firstExcelColumn], CurrentWrkSheet.Cells[inRow1, lastExcelColumn]].ClearContents();
              inRowInsert1++;
            }
            
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[inRow1, i + 1].Value = odr.GetValue(i);

            inRow1++;

          }
        }

        CurrentWrkSheet.Range["A4:Q4"].AutoFilter();

        CurrentWrkSheet.Cells[2, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return Result;
    }


  }






}


