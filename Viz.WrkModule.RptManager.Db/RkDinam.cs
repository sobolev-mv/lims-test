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

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class RkDinamRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public RkDinamRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd) : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = new DateTime(RptDateBegin.Year, 01, 01);
      this.DateEnd = new DateTime(RptDateBegin.Year, 12, 31);
    }
  }

  public sealed class RkDinam : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as RkDinamRptParam);
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
      finally{
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(RkDinamRptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;
      var oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        const string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_PK_DIN_PRN";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        CurrentWrkSheet.Cells[2, 4].Value = "с " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 5;

          while (odr.Read()){
            var month = odr.GetInt32("mes");
            var str = odr.GetInt32("str");
            
            if (str == 1){
              switch (month){
                case 1:
                  row = 5;
                  break;
                case 2:
                  row = 6;
                  break;
                case 3:
                  row = 7;
                  break;
                case 4:
                  row = 8;
                  break;
                case 5:
                  row = 9;
                  break;
                case 6:
                  row = 10;
                  break;
                case 7:
                  row = 11;
                  break;
                case 8:
                  row = 12;
                  break;
                case 9:
                  row = 13;
                  break;
                case 10:
                  row = 14;
                  break;
                case 11:
                  row = 15;
                  break;
                case 12:
                  row = 16;
                  break;
              }
            }
            else{
              switch (month){
                case 1:
                  row = 17;
                  break;
                case 2:
                  row = 18;
                  break;
                case 3:
                  row = 19;
                  break;
                case 4:
                  row = 20;
                  break;
                case 5:
                  row = 21;
                  break;
                case 6:
                  row = 22;
                  break;
                case 7:
                  row = 23;
                  break;
                case 8:
                  row = 24;
                  break;
                case 9:
                  row = 25;
                  break;
                case 10:
                  row = 26;
                  break;
                case 11:
                  row = 27;
                  break;
                case 12:
                  row = 28;
                  break;
              }
            }
 
            for (int i = 2; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
          }
        }

        CurrentWrkSheet.Cells[1, 1].Select();
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


  }






}

