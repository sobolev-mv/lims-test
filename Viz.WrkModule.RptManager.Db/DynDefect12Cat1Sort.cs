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

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class DynDefect12Cat1SortRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public DynDefect12Cat1SortRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class DynDefect12Cat1Sort : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DynDefect12Cat1SortRptParam);
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

    private Boolean RunRpt(DynDefect12Cat1SortRptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;

      DateTime dtTmpBegin;
      DateTime dtTmpEnd; 

      try
      {
        dtTmpBegin = new DateTime(prm.DateBegin.AddMonths(-1).Year, prm.DateBegin.AddMonths(-1).Month, 20);
        dtTmpEnd = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        CurrentWrkSheet.Cells[2, 1].Value = string.Format("за период с {0:dd.MM.yyyy} по {1:dd.MM.yyyy}", dtTmpBegin, dtTmpEnd);
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("VIZ_PRN.OTK_DINAMIKA.OTK_DINAM", CommandType.StoredProcedure, false, false, null); }));

        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;
        
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null){
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }

        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.OTK_DINAMIKA_ALL";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt1, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 6;

          while (odr.Read()){
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            row++;
          }
        }

        odr.Close();
        odr.Dispose();

        prm.ExcelApp.ActiveWorkbook.WorkSheets[5].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[2, 1].Value = string.Format("за период с {0:dd.MM.yyyy} по {1:dd.MM.yyyy}", dtTmpBegin, dtTmpEnd);

        const string sqlStmt2 = "SELECT * FROM VIZ_PRN.OTK_DINAMIKA_ALL_SGP";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt2, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null)
        {
          int flds = odr.FieldCount;
          int row = 7;

          while (odr.Read())
          {
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            row++;
          }
        }

        odr.Close();
        odr.Dispose();

        prm.ExcelApp.ActiveWorkbook.WorkSheets[7].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

        dtTmpBegin = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
        dtTmpEnd = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(dtTmpBegin, dtTmpEnd, 1)));
        CurrentWrkSheet.Cells[2, 1].Value = string.Format("за период с {0:dd.MM.yyyy} по {1:dd.MM.yyyy}", dtTmpBegin, dtTmpEnd);

        const string sqlStmt3 = "SELECT * FROM VIZ_PRN.OTK_DINAMIKA_SGP_1SORT";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt3, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null)
        {
          int flds = odr.FieldCount;
          int row = 7;

          while (odr.Read())
          {
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            row++;
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 1].Select();
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


