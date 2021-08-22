using System;
using System.Collections.Generic;
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

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class DevDayKesiRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public String   AvoAgr { get; set; }

    public DevDayKesiRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd, String AvoAgr)
           : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.AvoAgr = AvoAgr;
    }
  }

  public sealed class DevDayKesi : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DevDayKesiRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);
      }
      catch (Exception ex){
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

    private decimal? GetSumAll(DevDayKesiRptParam prm)
    {
      decimal? rez;
      Object rz = null;
      string stmt = "SELECT * FROM VIZ_PRN.OTK_SK_KESI_ALL";
      prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => {rz = Odac.ExecuteScalar(stmt, System.Data.CommandType.Text, false, null); }));
      rez = Convert.ToDecimal(rz);
      return rez;
    }


    private Boolean RunRpt(DevDayKesiRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      string SqlStmt = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_SK_KESI";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.AvoAgr)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;

        CurrentWrkSheet.Cells[2, 2].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        int flds = odr.FieldCount;
        int row = 5;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
          }


          row++;
        }

        CurrentWrkSheet.Cells[36, 12].Value = GetSumAll(prm);
        odr.Close();
        odr.Dispose();

        //АВО3
        prm.AvoAgr = "AVO3";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { DbVar.SetString(prm.AvoAgr); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;
        flds = odr.FieldCount;
        row = 41;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
          }
          row++;
        }

        CurrentWrkSheet.Cells[72, 12].Value = GetSumAll(prm);
        odr.Close();
        odr.Dispose();

        //АВО4
        prm.AvoAgr = "AVO4";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { DbVar.SetString(prm.AvoAgr); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;
        flds = odr.FieldCount;
        row = 77;

        while (odr.Read()){
          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
          }
          row++;
        }

        CurrentWrkSheet.Cells[108, 12].Value = GetSumAll(prm);
        odr.Close();
        odr.Dispose();

        //АВО5
        prm.AvoAgr = "AVO5";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { DbVar.SetString(prm.AvoAgr); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;
        flds = odr.FieldCount;
        row = 113;

        while (odr.Read()){
          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
          }
          row++;
        }

        CurrentWrkSheet.Cells[144, 12].Value = GetSumAll(prm);
        odr.Close();
        odr.Dispose();

        //АВО7
        prm.AvoAgr = "AVO7";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { DbVar.SetString(prm.AvoAgr); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;
        flds = odr.FieldCount;
        row = 149;

        while (odr.Read()){
          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 2].Value2 = odr.GetValue(i);
          }
          row++;
        }

        CurrentWrkSheet.Cells[180, 12].Value = GetSumAll(prm);

        //Возвращаемся на первую страницу
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
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


