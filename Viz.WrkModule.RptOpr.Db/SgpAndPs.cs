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
  public sealed class SgpAndPsRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int TypeRpt { get; set; }
    public string TypeShiftFinishApr { get; set; }
    public SgpAndPsRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class SgpAndPs : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as SgpAndPsRptParam);
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

    private void ToTypeRptSgp(SgpAndPsRptParam rptPrm)
    {
      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Date,
        Value = rptPrm.DateBegin
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = rptPrm.TypeShiftFinishApr.Length,
        Value = rptPrm.TypeShiftFinishApr
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery("VIZ_PRN.SGP_InStock.preSGP_Shift ", CommandType.StoredProcedure, false, lstPrm);

    }

    private void CreateData4TwoWrkSheet(dynamic CurrentWrkSheet, int prm4View)
    {
      OracleDataReader odr = null;

      DbVar.SetNum(prm4View);

      const string sqlStmt1 = "SELECT * FROM VIZ_PRN.V_SGP_INSTOCK";
      odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);

      if (odr != null)
      {
        int flds = odr.FieldCount;
        int inRow1 = 5;
        int inRowInsert1 = 8;

        while (odr.Read())
        {

          if (inRow1 == inRowInsert1)
          {

            CurrentWrkSheet.Rows[inRow1].Insert();
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[inRow1 - 1, 1], CurrentWrkSheet.Cells[inRow1 - 1, 40]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[inRow1, 1], CurrentWrkSheet.Cells[inRow1, 40]]);
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[inRow1, 1], CurrentWrkSheet.Cells[inRow1, 40]].ClearContents();
            inRowInsert1++;
          }

          for (int i = 0; i < flds; i++)
            CurrentWrkSheet.Cells[inRow1, i + 1].Value = odr.GetValue(i);

          inRow1++;

        }
        odr.Close();
        odr.Dispose();
      }
    }

    private void CreateData4LastWrkSheetSort(dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      odr = Odac.GetOracleReader("SELECT * FROM VIZ_PRN.V_SGP_INSTOCK_SORT", CommandType.Text, false, null, null);

      if (odr != null){
        int flds = odr.FieldCount;

        while (odr.Read()){

          if (odr.GetDouble(0) == 0.23)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[4, i + 1].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 0.27)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[5, i + 1].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 0.30)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[6, i + 1].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 0.35)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[7, i + 1].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 0.5)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[8, i + 1].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 98)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[9, i + 1].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 99)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[10, i + 1].Value = odr.GetValue(i);
        }

        odr.Close();
        odr.Dispose();
      }
    }

    private void CreateData4LastWrkSheetLaser(dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      odr = Odac.GetOracleReader("SELECT * FROM VIZ_PRN.V_SGP_INSTOCK_LK", CommandType.Text, false, null, null);

      if (odr != null){
        int flds = odr.FieldCount;

        while (odr.Read()){

          if (odr.GetDouble(0) == 0.23)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[4, i + 9].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 0.27)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[5, i + 9].Value = odr.GetValue(i);
          else if (odr.GetDouble(0) == 0.30)
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[6, i + 9].Value = odr.GetValue(i);
        }

        odr.Close();
        odr.Dispose();
      }
    }

    private void CreateData4LastWrkSheetShup(dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      odr = Odac.GetOracleReader("SELECT * FROM VIZ_PRN.V_SGP_INSTOCK_SHUP", CommandType.Text, false, null, null);

      if (odr != null)
      {
        int flds = odr.FieldCount;

        while (odr.Read())
        {

          if (odr.GetString(0) == "О1")
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[9, i + 9].Value = odr.GetValue(i);
          else if (odr.GetString(0) == "О2")
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[10, i + 9].Value = odr.GetValue(i);
          else if (odr.GetString(0) == "О7")
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[11, i + 9].Value = odr.GetValue(i);
          else if (odr.GetString(0) == "О8")
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[12, i + 9].Value = odr.GetValue(i);
          else if (odr.GetString(0) == "О9")
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[13, i + 9].Value = odr.GetValue(i);


        }

        odr.Close();
        odr.Dispose();
      }
    }

    private void CreateData4LastWrkSheetOtgruzka(dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      odr = Odac.GetOracleReader("SELECT * FROM VIZ_PRN.V_SGP_INSTOCK_SHIPPING", CommandType.Text, false, null, null);

      if (odr != null) {
        int rowCurr = 2;
        int rowEnd = 52;
        int flds = odr.FieldCount;
        int cntCol = 1;

        while (odr.Read()){

          for (int i = 0; i < flds; i++)
            CurrentWrkSheet.Cells[rowCurr, cntCol + i].Value = odr.GetValue(i);

          if (rowCurr >= rowEnd) {
            cntCol += 2;
            rowCurr = 2;
          }
          else {
            rowCurr++;
          }
        }

        odr.Close();
        odr.Dispose();
      }
    }



    //VIZ_PRN.V_SGP_INSTOCK_BRIG
    private string GetBrigPack()
    {
      string res = Convert.ToString(Odac.ExecuteScalar("SELECT BRIG_PACK FROM VIZ_PRN.V_SGP_INSTOCK_BRIG", CommandType.Text, false, null));
      return res;
    }

      private Boolean RunRpt(SgpAndPsRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin;
      DateTime? dtEnd;

      try{

        //за период, кнопка «Сдача ГП»
        if (prm.TypeRpt == 1) {
          DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
          Odac.ExecuteNonQuery("begin VIZ_PRN.SGP_InStock.preSGP(VIZ_PRN.VAR_RPT.GetDateBegin(0), VIZ_PRN.VAR_RPT.GetDateEnd(0)); end;", CommandType.Text, false, null);
          dtBegin = DbVar.GetDateBeginEnd(true, true);
          dtEnd = DbVar.GetDateBeginEnd(false, true);

          prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[1, 1].Value = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}";
          CurrentWrkSheet.Cells[1, 4].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

          prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[1, 1].Value = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}";
          CurrentWrkSheet.Cells[1, 4].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";
        }
        else {
          ToTypeRptSgp(prm);
          prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[1, 1].Value = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}";
          CurrentWrkSheet.Cells[1, 4].Value = $"за {prm.DateBegin:dd.MM.yyyy}, смена {prm.TypeShiftFinishApr}";

          prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[1, 1].Value = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}";
          CurrentWrkSheet.Cells[1, 4].Value = $"за {prm.DateBegin:dd.MM.yyyy}, смена {prm.TypeShiftFinishApr}";
        }

        //Формируем 1-лист.
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CreateData4TwoWrkSheet(CurrentWrkSheet, 0);

        //Формируем 2-лист.
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CreateData4TwoWrkSheet(CurrentWrkSheet, 1);

        //Формируем 3-лист. для второй кнопки
        if (prm.TypeRpt == 2){
          prm.ExcelApp.ActiveWorkbook.WorkSheets[3].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[1, 8].Value = prm.DateBegin;
          DbVar.SetNum(0);

          CreateData4LastWrkSheetSort(CurrentWrkSheet);
          CreateData4LastWrkSheetLaser(CurrentWrkSheet);
          CreateData4LastWrkSheetShup(CurrentWrkSheet);

          CurrentWrkSheet.Cells[1, 10].Value = GetBrigPack();

          //формируем 4-й лист "Отгрузка"
          prm.ExcelApp.ActiveWorkbook.WorkSheets[4].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CreateData4LastWrkSheetOtgruzka(CurrentWrkSheet);

        }



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


