using System;
using System.Data;
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
  public sealed class SgpStateRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public SgpStateRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class SgpState : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      SgpStateRptParam prm = (e.Argument as SgpStateRptParam);
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
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка выполнения", ex.Message, MessageBoxImage.Stop)));
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

    private Boolean RunRpt(SgpStateRptParam prm, dynamic currentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleCommand oracleCommand = null; 
      Boolean Result = false;
      //OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      DateTime dTmp;       

      try{
        string SqlStmt = "VIZ_PRN.SGP_STATE.PrepareData";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, DateTime.Now, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString("3131")));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        currentWrkSheet.Cells[2, 7].Value = string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin);
        dTmp = Convert.ToDateTime(dtBegin);

        //Готовим данные во временной таблице
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(SqlStmt, CommandType.StoredProcedure, false, false, null); }));
        if (iar != null) 
         iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null){
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }

        //Заполняем первую страницу (весь металл)
        dtBegin = dTmp.AddYears(-30); 
        dtEnd = dTmp;
        CreateDataSheet(prm, currentWrkSheet, dtBegin, dtEnd);

        //Заполняем вторую страницу (год и более)
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select(); //выбираем лист
        currentWrkSheet = prm.ExcelApp.ActiveSheet;

        dtBegin = dTmp.AddYears(-30);
        dtEnd = dTmp.AddYears(-1); 
        CreateDataSheet(prm, currentWrkSheet, dtBegin, dtEnd);

        //Заполняем третью страницу (от полугода до года)
        prm.ExcelApp.ActiveWorkbook.WorkSheets[3].Select(); //выбираем лист
        currentWrkSheet = prm.ExcelApp.ActiveSheet;

        dtBegin = dTmp.AddYears(-1).AddSeconds(1);
        dtEnd = dTmp.AddMonths(-6);
        CreateDataSheet(prm, currentWrkSheet, dtBegin, dtEnd);

        //Заполняем третью страницу (менее полугода)
        prm.ExcelApp.ActiveWorkbook.WorkSheets[4].Select(); //выбираем лист
        currentWrkSheet = prm.ExcelApp.ActiveSheet;

        dtBegin = dTmp.AddMonths(-6).AddSeconds(1);
        dtEnd = dTmp;
        CreateDataSheet(prm, currentWrkSheet, dtBegin, dtEnd);


        //возвращаемся на первый лист
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        currentWrkSheet = prm.ExcelApp.ActiveSheet;
        currentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка выполнения", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      return Result;
    }

    private void CreateDataSheet(SgpStateRptParam prm, dynamic currentWrkSheet, DateTime? dtBegin, DateTime? dtEnd)
    {
      string SqlStmt;
      IAsyncResult iar = null;
      OracleDataReader odr = null;

      int cntIter = 0;
      int oldThick = 0;
      int curThick = 0;
      int cntIns = 0;
      int rowCur = 0;
      int rowIns = 0;
      int r023 = 5;
      int r027 = 8;
      int r030 = 11;
      int r035 = 14;
      int r050 = 17;
      SqlStmt = "SELECT * FROM VIZ_PRN.PDB_SGP_STATE_DATA";

      try{
        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart) (() => DbVar.SetRangeDate(Convert.ToDateTime(dtBegin), Convert.ToDateTime(dtEnd), 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart) (() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;

        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;

          while (odr.Read()){
            curThick = Convert.ToInt32(odr.GetDecimal("TLS") * 100);

            //Определяем с какой позиции надо делать вставку строки    
            switch (curThick){
              case 23:
                rowIns = r023 + 1 + cntIns;
                break;
              case 27:
                rowIns = r027 + 1 + cntIns;
                break;
              case 30:
                rowIns = r030 + 1 + cntIns;
                break;
              case 35:
                rowIns = r035 + 1 + cntIns;
                break;
              case 50:
                rowIns = r050 + 1 + cntIns;
                break;
            }

            //в случае, если выборка только началась
            if (cntIter == 0)
              oldThick = curThick;

            //Начинается новая толщина
            if ((oldThick != curThick) || (cntIter == 0))
              switch (curThick){
                case 23:
                  rowCur = r023 + cntIns;
                  break;
                case 27:
                  rowCur = r027 + cntIns;
                  break;
                case 30:
                  rowCur = r030 + cntIns;
                  break;
                case 35:
                  rowCur = r035 + cntIns;
                  break;
                case 50:
                  rowCur = r050 + cntIns;
                  break;
              }
            else{

              //Здесь продолжается текущая толщина
              //Определяем, надо ли зделать вставку пустой строки
              if (rowCur >= rowIns){
                currentWrkSheet.Range[currentWrkSheet.Cells[rowCur, 2], currentWrkSheet.Cells[rowCur, 2]].EntireRow
                  .Insert();
                currentWrkSheet.Range[currentWrkSheet.Cells[rowCur, 2], currentWrkSheet.Cells[rowCur, 15]].Copy(
                  currentWrkSheet.Range[currentWrkSheet.Cells[rowCur + 1, 2], currentWrkSheet.Cells[rowCur + 1, 15]]);
                cntIns++;
              }
            }

            //заполняем данными ячейки 
            for (int i = 1; i < flds; i++)
              currentWrkSheet.Cells[rowCur, i + 1].Value = odr.GetValue(i);

            cntIter++;
            rowCur++;
            oldThick = curThick;
          }
        }
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      } 

    }







  }
}
