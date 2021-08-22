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

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class CatBrAooRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public CatBrAooRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd)
           : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
    }
  }

  public sealed class CatBrAoo : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      CatBrAooRptParam prm = (e.Argument as CatBrAooRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
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

        //Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(CatBrAooRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      var oef = new OdacErrorInfo();
      IAsyncResult iar = null;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AOO_IT WHERE FF1FFBRG = '1'";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[2, 1].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;

        int flds = odr.FieldCount;
        int row = 7;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }


          row++;
        }
        odr.Close();
        odr.Dispose();

        //Бригада 2
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AOO_IT WHERE FF1FFBRG = '2'";
        row = 12;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }
        odr.Close();
        odr.Dispose();

        //Бригада 3
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AOO_IT WHERE FF1FFBRG = '3'";
        row = 17;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }
        odr.Close();
        odr.Dispose();

        //Бригада 4
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AOO_IT WHERE FF1FFBRG = '4'";
        row = 22;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }

        odr.Close();
        odr.Dispose();

        //Лист2         
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[2, 2].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_VTO";
        row = 6;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);
 
        if (odr == null) return false;
        flds = odr.FieldCount;

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

        odr.Close();
        odr.Dispose();

        //Лист3  
        prm.ExcelApp.ActiveWorkbook.WorkSheets[3].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[2, 1].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        //Лист3-Бригада 1       
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AVO WHERE FF1FFBRG = '1'";
        row = 8;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        flds = odr.FieldCount;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }
        odr.Close();
        odr.Dispose();

        //Лист3-Бригада 2       
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AVO WHERE FF1FFBRG = '2'";
        row = 13;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        flds = odr.FieldCount;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }
        odr.Close();
        odr.Dispose();

        //Лист3-Бригада 3       
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AVO WHERE FF1FFBRG = '3'";
        row = 18;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        flds = odr.FieldCount;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }
        odr.Close();
        odr.Dispose();

        //Лист3-Бригада 4       
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_AVO WHERE FF1FFBRG = '4'";
        row = 23;
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        flds = odr.FieldCount;

        while (odr.Read()){

          for (int i = 0; i < flds; i++){
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }
        odr.Close();
        odr.Dispose();

        //Лист 4
        prm.ExcelApp.ActiveWorkbook.WorkSheets[4].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[2, 2].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        //Лист 4 Категорийность ЭАС по бригадам УО
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_UO ORDER BY BRG";
        row = 6;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        flds = odr.FieldCount;

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
        odr.Close();
        odr.Dispose();

        //Лист 5
        prm.ExcelApp.ActiveWorkbook.WorkSheets[5].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        

        //Лист 4 Рулоны
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_KAT_NOT";
        row = 2;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr == null) return false;
        flds = odr.FieldCount;

        while (odr.Read()){
          for (int i = 0; i < flds; i++)
          {
            string fn = odr.GetName(i);

            if (fn.Length < 6)
              CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
            else
              if (fn.Substring(0, 5) != "FF1FF")
                CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }




        //Возвращаемся на первую страницу
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

        CurrentWrkSheet.Cells[1, 1].Select();
        CurrentWrkSheet.Protect("xxx***xxx");
        prm.WorkBook.Protect("xxx***xxx");
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

