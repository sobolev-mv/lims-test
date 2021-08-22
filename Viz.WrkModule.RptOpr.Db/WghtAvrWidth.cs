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
  public sealed class WghtAvrWidthRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public Boolean IsFill1StSheet { get; set; }
    public Boolean IsTypeProdF3 { get; set; }
    public string TypeProdSqlStrF3 { get; set; }
    public Boolean IsThicknessF3 { get; set; }
    public string ThicknessSqlStrF3 { get; set; }
    public Boolean IsSortF3 { get; set; }
    public string SortSqlStrF3 { get; set; }

    public WghtAvrWidthRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class WghtAvrWidth : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as WghtAvrWidthRptParam);
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

    private Boolean RunRpt(WghtAvrWidthRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;

      try{

        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);

        //Stopwatch stopWatch = new Stopwatch();
        //TimeSpan ts = stopWatch.Elapsed;
        //string elapsedTime = null;

        CurrentWrkSheet.Cells[1, 1].Value = $"Период с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";
        Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.TMP_SGP", CommandType.Text, false, null);
        //stopWatch.Start();

        const string sqlStmtFilter = "INSERT INTO VIZ_PRN.TMP_SGP " +
                                     "SELECT * FROM VIZ_PRN.SGP_FILTR_CORE " +
                                     "WHERE " +
                                     "((MATART IN (SELECT VL_STRING FROM TABLE(VIZ_PRN.VAR_RPT.GetTabOfStrDelim(:TYPPROD, ',')))) OR (:FTYPPROD = 0)) " +
                                     "AND ((TOLS IN (SELECT TO_NUMBER(VL_STRING) FROM TABLE(VIZ_PRN.VAR_RPT.GetTabOfStrDelim(:THKS, ',')))) OR (:FTHKS = 0)) " +
                                     "AND ((SORT IN (SELECT VL_STRING FROM TABLE(VIZ_PRN.VAR_RPT.GetTabOfStrDelim(:SRT, ',')))) OR (:FSRT = 0))";

        var lstParam = new List<OracleParameter>();
        var param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "TYPPROD",
          Value = prm.TypeProdSqlStrF3,
          Size = prm.TypeProdSqlStrF3.Length
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "FTYPPROD",
          Value = prm.IsTypeProdF3 ? 1 : 0
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "THKS",
          Value = prm.ThicknessSqlStrF3,
          Size = prm.ThicknessSqlStrF3.Length
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "FTHKS",
          Value = prm.IsThicknessF3 ? 1 : 0
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "SRT",
          Value = prm.SortSqlStrF3,
          Size = prm.SortSqlStrF3.Length
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "FSRT",
          Value = prm.IsSortF3 ? 1 : 0
        };
        lstParam.Add(param);

        Odac.ExecuteNonQuery(sqlStmtFilter, CommandType.Text, false, lstParam);
        //stopWatch.Stop();
        //ts = stopWatch.Elapsed;
        //elapsedTime = $"INSERT INTO VIZ_PRN.TMP_SGP... -  {ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        //MessageBox.Show(elapsedTime);

        //stopWatch.Restart();
        CurrentWrkSheet.Cells[3, 2].Value = Odac.ExecuteScalar("SELECT SUM(VES) FROM VIZ_PRN.TMP_SGP", CommandType.Text, false, null);
        //stopWatch.Stop();
        //ts = stopWatch.Elapsed;
        //elapsedTime = $"SELECT SUM(VES)... -  {ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        //MessageBox.Show(elapsedTime);
        
        if (prm.IsFill1StSheet){
          const string sqlStmt1 = "SELECT * FROM VIZ_PRN.SGP_SRV_SHIR_L1";
          odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);

          if (odr != null){
            int flds = odr.FieldCount;
            int row = 7;

            const int firstExcelColumn = 1;
            const int lastExcelColumn = 10;

            while (odr.Read()){
              CurrentWrkSheet.Range[
                CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, firstExcelColumn], CurrentWrkSheet.Cells[row + 1, lastExcelColumn]]);

              for (int i = 0; i < flds; i++)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

              row++;
            }

            odr.Close();
            odr.Dispose();
            CurrentWrkSheet.Cells[4, 2].Formula = "=SUBTOTAL(9,J7:J" + Convert.ToString(row - 1) + ")";
          }

        }

        //stopWatch.Restart();
        int itogRow = 0;

        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        const string sqlStmt2 = "SELECT * FROM VIZ_PRN.SGP_SRV_SHIR_L2";
        odr = Odac.GetOracleReader(sqlStmt2, CommandType.Text, false, null, null);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 4;

          const int firstExcelColumn = 1;
          const int lastExcelColumn = 6;

          while (odr.Read()){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, firstExcelColumn], CurrentWrkSheet.Cells[row + 1, lastExcelColumn]]);

            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }

          itogRow = row;

          odr.Close();
          odr.Dispose();
        }

        //stopWatch.Stop();
        //ts = stopWatch.Elapsed;
        //elapsedTime = $"SGP_SRV_SHIR_L2... -  {ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        //MessageBox.Show(elapsedTime);

        //stopWatch.Restart();

        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        const string sqlStmt3 = "SELECT * FROM VIZ_PRN.SGP_SRV_SHIR_L2ITOG";
        odr = Odac.GetOracleReader(sqlStmt3, CommandType.Text, false, null, null);

        if (odr != null){
          int flds = odr.FieldCount;
          
          //const int firstExcelColumn = 1;
          //const int lastExcelColumn = 6;

          while (odr.Read()){
            //CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, firstExcelColumn], CurrentWrkSheet.Cells[row + 1, lastExcelColumn]]);

            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[itogRow, i + 1].Value = odr.GetValue(i);

            break;
          }
        }

        //stopWatch.Stop();
        //ts = stopWatch.Elapsed;
        //elapsedTime = $"SGP_SRV_SHIR_L2ITOG... -  {ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        //MessageBox.Show(elapsedTime);


        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
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


