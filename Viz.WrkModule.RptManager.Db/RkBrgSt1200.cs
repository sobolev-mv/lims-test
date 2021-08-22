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
  public sealed class RkBrgSt1200RptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int TypeFilterF1 { get; set; }
    public Boolean  Is1200F1 { get; set; }
    public DateTime DateBegin1200F1 { get; set; }
    public DateTime DateEnd1200F1 { get; set; }
    public string   ListStendF1 { get; set; }
    public RkBrgSt1200RptParam(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class RkBrgSt1200 : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as RkBrgSt1200RptParam);
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

    private Boolean RunRpt(RkBrgSt1200RptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;

      try{

        if ((prm.TypeFilterF1 == 0) && (!prm.Is1200F1)){

          CurrentWrkSheet.Cells[2, 1].Value = "за период с " + string.Format("{0:dd.MM.yyyy}", prm.DateBegin) + " по " + string.Format("{0:dd.MM.yyyy}", prm.DateEnd);
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 0)));
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("DELETE FROM VIZ_PRN.TMP_PU_FILTR_CORE", CommandType.Text, false, true, null); }));
          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          var oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("INSERT INTO VIZ_PRN.TMP_PU_FILTR_CORE SELECT * FROM VIZ_PRN.PU_RK_APR8_CORE", CommandType.Text, false, true, null); }));
          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          string sqlStmt = "";
          int row = 0;
          for (int j = 1; j < 5; j++){

            sqlStmt = "SELECT * FROM VIZ_PRN.PU_RK_APR8 WHERE NBRIG = " + j.ToString(System.Globalization.CultureInfo.InvariantCulture);
            prm.Disp.Invoke(DispatcherPriority.Normal,
            (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt, CommandType.Text, false, null, null); }));
            oracleCommand = iar.AsyncState as OracleCommand;
            if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

            if (odr != null){
              int flds = odr.FieldCount;

              switch (j){
                case 1:
                  row = 7;
                  break;
                case 2:
                  row = 13;
                  break;
                case 3:
                  row = 19;
                  break;
                case 4:
                  row = 25;
                  break;
                default:
                  break;
              }

              while (odr.Read()){
                for (int i = 1; i < flds; i++)
                  CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
                row++;
              }
              odr.Close();
              odr.Dispose();              
            }
          }

        }
        else if ((prm.TypeFilterF1 == 0) && (prm.Is1200F1)){
          CurrentWrkSheet.Cells[2, 1].Value = "за период с " + string.Format("{0:dd.MM.yyyy}", prm.DateBegin) + " по " + string.Format("{0:dd.MM.yyyy}", prm.DateEnd);
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 0)));
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("DELETE FROM VIZ_PRN.TMP_PU_FILTR_CORE", CommandType.Text, false, true, null); }));
          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          var oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("INSERT INTO VIZ_PRN.TMP_PU_FILTR_CORE SELECT * FROM VIZ_PRN.PU_RK_APR8_CORE", CommandType.Text, false, true, null); }));
          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          string sqlStmt = "";
          int row = 0;
          for (int j = 1; j < 5; j++){

            sqlStmt = "SELECT * FROM VIZ_PRN.PU_RK_APR8 WHERE NBRIG = " + j.ToString(System.Globalization.CultureInfo.InvariantCulture);
            prm.Disp.Invoke(DispatcherPriority.Normal,
            (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt, CommandType.Text, false, null, null); }));
            oracleCommand = iar.AsyncState as OracleCommand;
            if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

            if (odr != null)
            {
              int flds = odr.FieldCount;

              switch (j){
                case 1:
                  row = 7;
                  break;
                case 2:
                  row = 13;
                  break;
                case 3:
                  row = 19;
                  break;
                case 4:
                  row = 25;
                  break;
                default:
                  break;
              }

              while (odr.Read()){
                for (int i = 1; i < flds; i++)
                  CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
                row++;
              }
              odr.Close();
              odr.Dispose();
            }
          }

          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin1200F1, prm.DateEnd1200F1, 0)));
          prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select(); //выбираем лист
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[2, 1].Value = "за период с " + string.Format("{0:dd.MM.yyyy}", prm.DateBegin1200F1) + " по " + string.Format("{0:dd.MM.yyyy}", prm.DateEnd1200F1);

          for (int j = 1; j < 5; j++){

            sqlStmt = "SELECT * FROM VIZ_PRN.PU_RK_APR8_1200 WHERE NBRIG = " + j.ToString(System.Globalization.CultureInfo.InvariantCulture);
            prm.Disp.Invoke(DispatcherPriority.Normal,
            (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt, CommandType.Text, false, null, null); }));
            oracleCommand = iar.AsyncState as OracleCommand;
            if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

            if (odr != null){
              int flds = odr.FieldCount;

              switch (j){
                case 1:
                  row = 7;
                  break;
                case 2:
                  row = 13;
                  break;
                case 3:
                  row = 19;
                  break;
                case 4:
                  row = 25;
                  break;
                default:
                  break;
              }

              while (odr.Read()){
                for (int i = 1; i < flds; i++)
                  CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
                row++;
              }
              odr.Close();
              odr.Dispose();
            }
          }
          prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

        } else if (prm.TypeFilterF1 == 1){

          CurrentWrkSheet.Cells[2, 1].Value = "стенды: " + prm.ListStendF1;
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.ListStendF1)));
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("DELETE FROM VIZ_PRN.TMP_PU_FILTR_CORE", CommandType.Text, false, true, null); }));
          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          var oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("INSERT INTO VIZ_PRN.TMP_PU_FILTR_CORE SELECT * FROM VIZ_PRN.PU_RK_APR8_STEND_CORE", CommandType.Text, false, true, null); }));
          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          string sqlStmt = "";
          int row = 0;
          for (int j = 1; j < 5; j++){

            sqlStmt = "SELECT * FROM VIZ_PRN.PU_RK_APR8 WHERE NBRIG = " + j.ToString(System.Globalization.CultureInfo.InvariantCulture);
            prm.Disp.Invoke(DispatcherPriority.Normal,
            (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt, CommandType.Text, false, null, null); }));
            oracleCommand = iar.AsyncState as OracleCommand;
            if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

            if (odr != null){
              int flds = odr.FieldCount;

              switch (j){
                case 1:
                  row = 7;
                  break;
                case 2:
                  row = 13;
                  break;
                case 3:
                  row = 19;
                  break;
                case 4:
                  row = 25;
                  break;
                default:
                  break;
              }

              while (odr.Read()){
                for (int i = 1; i < flds; i++)
                  CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
                row++;
              }
              odr.Close();
              odr.Dispose();
            }
          }
        }


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

