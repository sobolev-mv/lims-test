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
  public sealed class AppHRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public AppHRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class AppH : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      AppHRptParam prm = (e.Argument as AppHRptParam);
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

    private Boolean RunRpt(AppHRptParam prm, dynamic currentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      DateTime dTmp;       

      try{
        const string sqlStmtPrepare = "VIZ_PRN.SGP_STATE.PrepareData";

        DbVar.SetRangeDate(prm.DateBegin, DateTime.Now, 1);
        DbVar.SetString("3131");
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        //currentWrkSheet.Cells[8, 7].Value = "от " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin);
        dTmp = Convert.ToDateTime(dtBegin);

        //Готовим данные во временной таблице
        Odac.ExecuteNonQuery(sqlStmtPrepare, CommandType.StoredProcedure, false, null);

        //Заполняем первую таблицу (от 3 до 9 месяцев)
        dtBegin = dTmp.AddMonths(-9);
        dtEnd = dTmp.AddMonths(-3);

        const string sqlStmt = "SELECT * FROM VIZ_PRN.PDB_SGP_APPH";
        DbVar.SetRangeDate(Convert.ToDateTime(dtBegin), Convert.ToDateTime(dtEnd), 1);

        int curRow = 14;
        int insRow = 15;
        int cntIns = 0;

        odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

        if (odr != null){

          int flds = odr.FieldCount;

          while (odr.Read()){

            if (curRow >= insRow){
              currentWrkSheet.Range[currentWrkSheet.Cells[curRow, 1], currentWrkSheet.Cells[curRow, 1]].EntireRow.Insert();
              currentWrkSheet.Range[currentWrkSheet.Cells[curRow, 1], currentWrkSheet.Cells[curRow, 17]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[curRow + 1, 1], currentWrkSheet.Cells[curRow + 1, 17]]);
              cntIns++;
            }

            //заполняем данными ячейки 
            for (int i = 0; i < flds; i++)
              currentWrkSheet.Cells[curRow, i + 1].Value = odr.GetValue(i);

            curRow++;
          }

          odr.Close();
          odr.Dispose();

          currentWrkSheet.Range[currentWrkSheet.Cells[curRow, 1], currentWrkSheet.Cells[curRow, 1]].EntireRow.Delete();
          cntIns--;
        }
        

        //Заполняем вторую таблицу (более 9 месяцев)
        dtBegin = dTmp.AddYears(-30);
        dtEnd = dTmp.AddMonths(-9);
        DbVar.SetRangeDate(Convert.ToDateTime(dtBegin), Convert.ToDateTime(dtEnd), 1);

        curRow = 21 + cntIns;
        insRow = 22 + cntIns;

        odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

        if (odr != null){

          int flds = odr.FieldCount;

          while (odr.Read()){

            if (curRow >= insRow){
              currentWrkSheet.Range[currentWrkSheet.Cells[curRow, 1], currentWrkSheet.Cells[curRow, 1]].EntireRow.Insert();
              currentWrkSheet.Range[currentWrkSheet.Cells[curRow, 1], currentWrkSheet.Cells[curRow, 17]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[curRow + 1, 1], currentWrkSheet.Cells[curRow + 1, 17]]);
              cntIns++;
            }

            //заполняем данными ячейки 
            for (int i = 0; i < flds; i++)
              currentWrkSheet.Cells[curRow, i + 1].Value = odr.GetValue(i);

            curRow++;
          }

          odr.Close();
          odr.Dispose();

          currentWrkSheet.Range[currentWrkSheet.Cells[curRow, 1], currentWrkSheet.Cells[curRow, 1]].EntireRow.Delete();
        }

        currentWrkSheet.Cells[1, 1].Select();

        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка выполнения", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      return Result;
    }

    private void CreateDataSheet(AppHRptParam prm, dynamic currentWrkSheet, DateTime? dtBegin, DateTime? dtEnd)
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
        prm.Disp.Invoke(DispatcherPriority.Normal,(ThreadStart) (() => DbVar.SetRangeDate(Convert.ToDateTime(dtBegin), Convert.ToDateTime(dtEnd), 1)));
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
