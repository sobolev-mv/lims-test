using System;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptHimLab.Db
{
  public sealed class HimLabIsolPropRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int SubRpt { get; set; } //какой из подотчетов будет использоваться

    public HimLabIsolPropRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd, int SubRpt)
           : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.SubRpt = SubRpt;
    }
  }

  public sealed class HimLabIsolProp : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as HimLabIsolPropRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
       
        switch (prm.SubRpt){
          case 1:
            this.RunRpt01(prm, wrkSheet);
            break;
          case 2:
            this.RunRpt02(prm, wrkSheet);
            break;
          case 3:
            this.RunRpt03(prm, wrkSheet);
            break;
          default:
            break;
        }
        this.SaveResult(prm);
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally{
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        Marshal.ReleaseComObject(wrkSheet);
        Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
        //GC.WaitForPendingFinalizers();
        //GC.Collect();
      }
    }
  
    //За кажды сутки в течении месяца
    private Boolean RunRpt01(HimLabIsolPropRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      var dt1 = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
      var dt2 = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));

      try{
        const string SqlStmt = "SELECT * FROM VIZ_PRN.CZL_TOK_FRANKL_DAYS";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(dt1, dt2, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[1, 13].Value = string.Format(" с {0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
        
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 4;
          var flds = odr.FieldCount;

          while (odr.Read()){
            for (int i = 0; i < flds; i++){
              if (i != 1)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            }
            row++;
          }
        }

        //выбираем лист 2
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("за период с {0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        //выбираем лист 1 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception){
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

    //За каждую неделю в течении года
    private Boolean RunRpt02(HimLabIsolPropRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      var dt1 = new DateTime(prm.DateBegin.Year, 1, 1);
      var dt2 = new DateTime(prm.DateBegin.Year, 12, 31);

      try{
        const string SqlStmt = "SELECT * FROM VIZ_PRN.CZL_TOK_FRANKL_NED";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(dt1, dt2, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[1, 13].Value = string.Format(" с {0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 4;
          var flds = odr.FieldCount;

          while (odr.Read()){
            for (int i = 0; i < flds; i++){
              if (i != 1)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            }
            row++;
          }
        }

        //выбираем лист 2
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("за период с {0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        //выбираем лист 1 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception){
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

    //За каждую неделю в течении года
    private Boolean RunRpt03(HimLabIsolPropRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      var dt1 = new DateTime(prm.DateBegin.Year, 1, 1);
      var dt2 = new DateTime(prm.DateBegin.Year, 12, 31);

      try{
        const string SqlStmt = "SELECT * FROM VIZ_PRN.CZL_TOK_FRANKL_ALL";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(dt1, dt2, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[1, 5].Value = string.Format(" с {0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 5;
          var flds = odr.FieldCount;

          while (odr.Read()){
            for (int i = 0; i < flds; i++){
              if (i != 0)
                CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
            }
            row++;
          }
        }
      }
      catch (Exception){
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
