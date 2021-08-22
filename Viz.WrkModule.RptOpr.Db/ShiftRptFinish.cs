using System;
using System.Data;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptOpr.Db
{
  internal enum TypeFinishApr
  {
    FinishApr12,
    FinishAprLaser,
    FinishAprOther,
    NotExistsApr
  }

  /*Расположение листов АПР в файле*/
  internal enum SheetsOfApr
  {
    Apr3 =  1,
    Apr4 =  2,
    Apr5 =  3,
    Apr6 =  4,
    Apr9 =  5,
    Apr10 = 6,
    Apr14 = 7,
    Apr12 = 8
  }

  internal struct RwPoColor
  {
    public string LocNo;
    public string Sort;
    public string Category;
    public RwPoColor(string locNo, string sort, string category)
    {
      LocNo = locNo;
      Sort = sort;
      Category = category;
    }

    public bool Equals(RwPoColor other)
    {
      return string.Equals(LocNo, other.LocNo) && string.Equals(Sort, other.Sort) && string.Equals(Category, other.Category);
    }

    public override bool Equals(object obj)
    {
      if (ReferenceEquals(null, obj)) return false;
      return obj is RwPoColor && Equals((RwPoColor) obj);
    }

    public override int GetHashCode()
    {
      unchecked
      {
        var hashCode = (LocNo != null ? LocNo.GetHashCode() : 0);
        hashCode = (hashCode * 397) ^ (Sort != null ? Sort.GetHashCode() : 0);
        hashCode = (hashCode * 397) ^ (Category != null ? Category.GetHashCode() : 0);
        return hashCode;
      }
    }

    public static Boolean operator ==(RwPoColor a, RwPoColor b)
    {
      return string.Equals(a.LocNo, b.LocNo) && string.Equals(a.Sort, b.Sort) && string.Equals(a.Category, b.Category);
    }

    public static bool operator !=(RwPoColor a, RwPoColor b)
    {
      return !(string.Equals(a.LocNo, b.LocNo) && string.Equals(a.Sort, b.Sort) && string.Equals(a.Category, b.Category));
    }
  }

  public sealed class ShiftRptFinishRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public string FinishApr { get; set; }
    public string FinishAprLabel { get; set; }
    public string TeamFinishApr { get; set; }
    public string ShiftMasterFinishApr { get; set; }
    public string TopWorkerFinishApr { get; set; }
    public Boolean IsApr12 { get; set; }
    public string TypeShiftFinishApr { get; set; }
    public Boolean IsLogInfo { get; set; }
    public ShiftRptFinishRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class ShiftRptFinish : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as ShiftRptFinishRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);

        //вызывается в случае перключения целевой БД
        base.DoWorkXls(sender, e);
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

        if (wrkSheet != null)
          Marshal.ReleaseComObject(prm.ExcelApp);

        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private TypeFinishApr GetTypeApr(string psiNameApr)
    {
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR3"))
        return TypeFinishApr.FinishAprOther;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR4"))
        return TypeFinishApr.FinishAprOther;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR5"))
        return TypeFinishApr.FinishAprOther;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR6"))
        return TypeFinishApr.FinishAprOther;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR9"))
        return TypeFinishApr.FinishAprOther;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR10"))
        return TypeFinishApr.FinishAprLaser;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR14"))
        return TypeFinishApr.FinishAprLaser;
      if ((psiNameApr != null) && string.Equals(psiNameApr.ToUpper(CultureInfo.InvariantCulture), "APR12"))
        return TypeFinishApr.FinishApr12;
      else
        return TypeFinishApr.NotExistsApr;
    }

    private int GetOneVisibleSheet(ShiftRptFinishRptParam prm)
    {
      if (prm.FinishApr.IndexOf(",", StringComparison.Ordinal) > -1)
        return 0;
      
      int idxActiveSheet = 1;

      foreach (int idxSheet in Enum.GetValues(typeof(SheetsOfApr)))
      {

        var name = Enum.GetName(typeof(SheetsOfApr), idxSheet);
        Boolean isVisible = (name != null) && (string.Equals(name.ToUpper(CultureInfo.InvariantCulture), prm.FinishApr.ToUpper(CultureInfo.InvariantCulture), StringComparison.Ordinal));

        if (!isVisible){
          prm.ExcelApp.ActiveWorkbook.WorkSheets[idxSheet].Select();
          prm.ExcelApp.ActiveSheet.Visible = false;
        }
        else
          idxActiveSheet = idxSheet;
      }

      prm.ExcelApp.ActiveWorkbook.WorkSheets[idxActiveSheet].Select();
      return idxActiveSheet;
    }

    private void FinishApr12(ShiftRptFinishRptParam prm, DateTime? dtBegin, int idxSheet)
    {
      OracleDataReader odr = null;
      prm.ExcelApp.ActiveWorkbook.WorkSheets[idxSheet].Select();
      dynamic currentWrkSheet = prm.ExcelApp.ActiveSheet;
      
      currentWrkSheet.Cells[2, 7].Value = $"{dtBegin:dd.MM.yyyy}";
      currentWrkSheet.Cells[2, 13].Value = prm.TeamFinishApr;
      currentWrkSheet.Cells[2, 14].Value = prm.TypeShiftFinishApr;
      currentWrkSheet.Cells[5, 4].Value = "См.мастер: " + prm.ShiftMasterFinishApr;
      currentWrkSheet.Cells[5, 12].Value = "Ст.рабочий: " + prm.TopWorkerFinishApr;

      int qntInsert = 0;

      var sqlStmtApr121 = prm.IsLogInfo ? "SELECT * FROM VIZ_PRN.SM_RAPORT_VH_12_ALL" : "SELECT * FROM VIZ_PRN.SM_RAPORT_VH_12_LALL";
      odr = Odac.GetOracleReader(sqlStmtApr121, CommandType.Text, false, null, null);
    
      if (odr != null){

        int inRow1 = 13;
        int inRowInsert1 = 15;

        while (odr.Read()){

          if (inRow1 == inRowInsert1){

            currentWrkSheet.Rows[inRow1].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow1 - 1, 1], currentWrkSheet.Cells[inRow1 - 1, 17]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 17]]);
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 17]].ClearContents();
            inRowInsert1++;
            qntInsert++;

          }

          currentWrkSheet.Cells[inRow1, 1].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRow1, 2].Value = odr.GetValue(1);
          currentWrkSheet.Cells[inRow1, 3].Value = odr.GetValue(2);
          currentWrkSheet.Cells[inRow1, 4].Value = odr.GetValue(3);
          currentWrkSheet.Cells[inRow1, 5].Value = odr.GetValue(4);
          currentWrkSheet.Cells[inRow1, 6].Value = odr.GetValue(5);
          currentWrkSheet.Cells[inRow1, 7].Value = odr.GetValue(6);
          currentWrkSheet.Cells[inRow1, 8].Value = odr.GetValue(7);
          currentWrkSheet.Cells[inRow1, 13].Value = odr.GetValue(8);
          currentWrkSheet.Cells[inRow1, 14].Value = odr.GetValue(9);
          currentWrkSheet.Cells[inRow1, 16].Value = odr.GetValue(10);
          currentWrkSheet.Cells[inRow1, 17].Value = odr.GetValue(11);

          if (prm.IsLogInfo){
            currentWrkSheet.Cells[inRow1, 53].Value = odr.GetValue(12);
            currentWrkSheet.Cells[inRow1, 54].Value = odr.GetValue(13);
            currentWrkSheet.Cells[inRow1, 55].Value = odr.GetValue(14);
          }

          inRow1++;

        }
        odr.Close();
        odr.Dispose();
      }

      decimal? res = GetItogFinishAprOther(prm.FinishAprLabel, "S1");
      if (res != null) 
        currentWrkSheet.Cells[20 + qntInsert, 17].Value = res;

      res = GetItogFinishAprOther(prm.FinishAprLabel, "S2");
      if (res != null)
        currentWrkSheet.Cells[21 + qntInsert, 17].Value = res;
      
      var sqlStmtApr122 = prm.IsLogInfo ? "SELECT * FROM VIZ_PRN.SM_RAPORT_ISH_12_ALL" : "SELECT * FROM VIZ_PRN.SM_RAPORT_ISH_12_LALL";
      odr = Odac.GetOracleReader(sqlStmtApr122, CommandType.Text, false, null, null);

      if (odr != null){

        int inRow2 = 28 + qntInsert;
        int inRowInsert2 = 30 + qntInsert;

        //int colorRow = 49407;

        while (odr.Read()){

          if (inRow2 == inRowInsert2){

            currentWrkSheet.Rows[inRow2].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2 - 1, 1], currentWrkSheet.Cells[inRow2 - 1, 19]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 19]]);
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 19]].ClearContents();

            inRowInsert2++;
            qntInsert++;
          }

          if ((odr.GetValue("PR_COLOR") != DBNull.Value) && (odr.GetInt32("PR_COLOR") > 1)){
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 18]].Interior.Pattern = 1; //xlSolid
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 18]].Interior.PatternColorIndex = -4105; //xlAutomatic;
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 18]].Interior.ThemeColor = 6; //xlThemeColorAccent2
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 18]].Interior.TintAndShade = 0.799981688894314;
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 18]].Interior.PatternTintAndShade = 0;
          }
          else
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 18]].Interior.Pattern = -4142; //xlNone

          currentWrkSheet.Cells[inRow2, 1].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRow2, 2].Value = odr.GetValue(1);
          currentWrkSheet.Cells[inRow2, 3].Value = odr.GetValue(2);
          currentWrkSheet.Cells[inRow2, 5].Value = odr.GetValue(3);
          currentWrkSheet.Cells[inRow2, 7].Value = odr.GetValue(4);

          currentWrkSheet.Cells[inRow2, 9].Value = odr.GetValue(5);
          currentWrkSheet.Cells[inRow2, 11].Value = odr.GetValue(6);
          currentWrkSheet.Cells[inRow2, 12].Value = odr.GetValue(7);
          currentWrkSheet.Cells[inRow2, 13].Value = odr.GetValue(8);
          currentWrkSheet.Cells[inRow2, 14].Value = odr.GetValue(9);
          currentWrkSheet.Cells[inRow2, 15].Value = odr.GetValue(10);
          currentWrkSheet.Cells[inRow2, 17].Value = odr.GetValue(11);
          currentWrkSheet.Cells[inRow2, 19].Value = odr.GetValue(12);

          if (prm.IsLogInfo){
            currentWrkSheet.Cells[inRow2, 55].Value = odr.GetValue(13);
            currentWrkSheet.Cells[inRow2, 56].Value = odr.GetValue(14);
          }

          inRow2++;

        }
        odr.Close();
        odr.Dispose();
      }


      const string sqlStmtApr12Futer2 = "SELECT * FROM VIZ_PRN.SM_RAPORT_PROSTOI_ALL";
      odr = Odac.GetOracleReader(sqlStmtApr12Futer2, CommandType.Text, false, null, null);

      if (odr != null){

        int inRowFuter2 = 37 + qntInsert;
        int inRowInsertFuter2 = 43 + qntInsert;

        while (odr.Read()){

          if (inRowFuter2 == inRowInsertFuter2){
            currentWrkSheet.Rows[inRowFuter2].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter2 + 1, 1], currentWrkSheet.Cells[inRowFuter2 + 1, 18]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter2, 1], currentWrkSheet.Cells[inRowFuter2, 18]]);
            inRowInsertFuter2++;
            //qntInsert++;
          }

          currentWrkSheet.Cells[inRowFuter2, 1].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRowFuter2, 2].Value = odr.GetValue(1);
          currentWrkSheet.Cells[inRowFuter2, 3].Value = odr.GetValue(2);
          currentWrkSheet.Cells[inRowFuter2, 5].Value = odr.GetValue(3);
          inRowFuter2++;
        }
        odr.Close();
        odr.Dispose();
      }

      const string sqlStmtApr12Futer3 = "SELECT * FROM VIZ_PRN.SM_RAPORT_FIO";
      odr = Odac.GetOracleReader(sqlStmtApr12Futer3, CommandType.Text, false, null, null);

      if (odr != null){

        int inRowFuter3 = 37 + qntInsert;

        while (odr.Read()){

          currentWrkSheet.Cells[inRowFuter3, 10].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRowFuter3, 12].Value = odr.GetValue(1);
          inRowFuter3++;
        }
        odr.Close();
        odr.Dispose();
      }

    }

    private void FinishAprLaser(ShiftRptFinishRptParam prm, DateTime? dtBegin, int idxSheet)
    {
      OracleDataReader odr = null;
      prm.ExcelApp.ActiveWorkbook.WorkSheets[idxSheet].Select();
      dynamic currentWrkSheet = prm.ExcelApp.ActiveSheet;

      currentWrkSheet.Cells[2, 12].Value = $"{dtBegin:dd.MM.yyyy}";
      currentWrkSheet.Cells[2, 14].Value = prm.TeamFinishApr;
      currentWrkSheet.Cells[2, 16].Value = prm.TypeShiftFinishApr;
      currentWrkSheet.Cells[2, 17].Value = prm.FinishAprLabel;
      currentWrkSheet.Cells[5, 8].Value = "См.мастер: " + prm.ShiftMasterFinishApr;
      currentWrkSheet.Cells[5, 15].Value = "Ст.рабочий: " + prm.TopWorkerFinishApr;

      int qntInsert = 0;

      var sqlStmtApr121 = prm.IsLogInfo ? "SELECT * FROM VIZ_PRN.SM_RAPORT_LSR_ALL" : "SELECT * FROM VIZ_PRN.SM_RAPORT_LSR_LALL";
      odr = Odac.GetOracleReader(sqlStmtApr121, CommandType.Text, false, null, null);

      if (odr != null){

        int inRow1 = 12;
        int inRowInsert1 = 14;

        while (odr.Read()){

          if (inRow1 == inRowInsert1){
            currentWrkSheet.Rows[inRow1].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow1 - 1, 1], currentWrkSheet.Cells[inRow1 - 1, 20]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 20]]);
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 20]].ClearContents();
            inRowInsert1++;
            qntInsert++;
          }

          currentWrkSheet.Cells[inRow1, 1].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRow1, 2].Value = odr.GetValue(1);
          currentWrkSheet.Cells[inRow1, 3].Value = odr.GetValue(2);
          currentWrkSheet.Cells[inRow1, 4].Value = odr.GetValue(3);
          currentWrkSheet.Cells[inRow1, 5].Value = odr.GetValue(4);
          currentWrkSheet.Cells[inRow1, 6].Value = odr.GetValue(5);
          currentWrkSheet.Cells[inRow1, 8].Value = odr.GetValue(6);
          currentWrkSheet.Cells[inRow1, 9].Value = odr.GetValue(7);
          currentWrkSheet.Cells[inRow1, 10].Value = odr.GetValue(8);
          currentWrkSheet.Cells[inRow1, 11].Value = odr.GetValue(9);
          currentWrkSheet.Cells[inRow1, 13].Value = odr.GetValue(10);
          currentWrkSheet.Cells[inRow1, 14].Value = odr.GetValue(11);
          currentWrkSheet.Cells[inRow1, 15].Value = odr.GetValue(12);
          currentWrkSheet.Cells[inRow1, 16].Value = odr.GetValue(13);
          currentWrkSheet.Cells[inRow1, 17].Value = odr.GetValue(14);
          currentWrkSheet.Cells[inRow1, 18].Value = odr.GetValue(15);
          currentWrkSheet.Cells[inRow1, 19].Value = odr.GetValue(16);
          currentWrkSheet.Cells[inRow1, 21].Value = odr.GetValue(17);

          if (prm.IsLogInfo){
            currentWrkSheet.Cells[inRow1, 53].Value = odr.GetValue(18);
            currentWrkSheet.Cells[inRow1, 54].Value = odr.GetValue(19);
            currentWrkSheet.Cells[inRow1, 55].Value = odr.GetValue(20);
            currentWrkSheet.Cells[inRow1, 56].Value = odr.GetValue(21);
          }
          inRow1++;
        }
        odr.Close();
        odr.Dispose();
      }

      DbVar.SetRangeDate(prm.DateBegin, prm.DateBegin, 1);
      DbVar.SetString(prm.FinishAprLabel, prm.TeamFinishApr);

      const string sqlStmtApr12Futer2 = "SELECT * FROM VIZ_PRN.SM_RAPORT_PROSTOI_ALL";
      odr = Odac.GetOracleReader(sqlStmtApr12Futer2, CommandType.Text, false, null, null);

      if (odr != null){

        int inRowFuter2 = 24 + qntInsert;
        int inRowInsertFuter2 = 28 + qntInsert;

        while (odr.Read()){

          if (inRowFuter2 == inRowInsertFuter2){
            currentWrkSheet.Rows[inRowFuter2].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter2 + 1, 1], currentWrkSheet.Cells[inRowFuter2 + 1, 20]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter2, 1], currentWrkSheet.Cells[inRowFuter2, 20]]);
            inRowInsertFuter2++;
            //qntInsert++;
          }

          currentWrkSheet.Cells[inRowFuter2, 1].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRowFuter2, 3].Value = odr.GetValue(1);
          currentWrkSheet.Cells[inRowFuter2, 4].Value = odr.GetValue(2);
          currentWrkSheet.Cells[inRowFuter2, 6].Value = odr.GetValue(3);
          inRowFuter2++;
        }
        odr.Close();
        odr.Dispose();
      }

      const string sqlStmtApr12Futer3 = "SELECT * FROM VIZ_PRN.SM_RAPORT_FIO";
      odr = Odac.GetOracleReader(sqlStmtApr12Futer3, CommandType.Text, false, null, null);

      if (odr != null){
        int inRowFuter3 = 24 + qntInsert;

        while (odr.Read()){
          currentWrkSheet.Cells[inRowFuter3, 11].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRowFuter3, 12].Value = odr.GetValue(1);
          inRowFuter3++;
        }
        odr.Close();
        odr.Dispose();
      }
    }

    private void FinishAprOther(ShiftRptFinishRptParam prm, DateTime? dtBegin, int idxSheet)
    {
      OracleDataReader odr = null;
      prm.ExcelApp.ActiveWorkbook.WorkSheets[idxSheet].Select();
      dynamic currentWrkSheet = prm.ExcelApp.ActiveSheet;

      currentWrkSheet.Cells[2, 11].Value = $"{dtBegin:dd.MM.yyyy}";
      currentWrkSheet.Cells[2, 13].Value = prm.TeamFinishApr;
      currentWrkSheet.Cells[2, 14].Value = prm.TypeShiftFinishApr;
      currentWrkSheet.Cells[2, 16].Value = prm.FinishAprLabel;
      currentWrkSheet.Cells[5, 8].Value = "См.мастер: " + prm.ShiftMasterFinishApr;
      currentWrkSheet.Cells[5, 14].Value = "Ст.рабочий: " + prm.TopWorkerFinishApr;

      int qntInsert = 0;
      
      var sqlStmtAprOther1 = prm.IsLogInfo ? "SELECT * FROM VIZ_PRN.SM_RAPORT_APR_ALL" : "SELECT * FROM VIZ_PRN.SM_RAPORT_APR_LALL";
      odr = Odac.GetOracleReader(sqlStmtAprOther1, CommandType.Text, false, null, null);

      if (odr != null){

        int inRow1 = 13;
        int inRowInsert1 = 15;
        string oldLocNo = "";
        string currLocNo = "";


        while (odr.Read()){
          currLocNo = odr.GetString(2);

          if (inRow1 == inRowInsert1){
            currentWrkSheet.Rows[inRow1].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow1 - 1, 1], currentWrkSheet.Cells[inRow1 - 1, 24]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 24]]);
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 23]].ClearContents();
            inRowInsert1++;
            qntInsert++;
          }

          if (oldLocNo != currLocNo){
            currentWrkSheet.Cells[inRow1, 1].Value = odr.GetValue(0);
            currentWrkSheet.Cells[inRow1, 2].Value = odr.GetValue(1);
            currentWrkSheet.Cells[inRow1, 3].Value = odr.GetValue(2);
            currentWrkSheet.Cells[inRow1, 4].Value = odr.GetValue(3);
            currentWrkSheet.Cells[inRow1, 5].Value = odr.GetValue(4);
            currentWrkSheet.Cells[inRow1, 6].Value = odr.GetValue(5);
            currentWrkSheet.Cells[inRow1, 7].Value = odr.GetValue(6);
            currentWrkSheet.Cells[inRow1, 8].Value = odr.GetValue(7);
            currentWrkSheet.Cells[inRow1, 9].Value = odr.GetValue(8);
            currentWrkSheet.Cells[inRow1, 10].Value = odr.GetValue(9);

          }
          currentWrkSheet.Cells[inRow1, 11].Value = odr.GetValue(10);
          currentWrkSheet.Cells[inRow1, 13].Value = odr.GetValue(11);
          currentWrkSheet.Cells[inRow1, 14].Value = odr.GetValue(12);
          currentWrkSheet.Cells[inRow1, 15].Value = odr.GetValue(13);
          currentWrkSheet.Cells[inRow1, 16].Value = odr.GetValue(14);
          currentWrkSheet.Cells[inRow1, 17].Value = odr.GetValue(15);
          currentWrkSheet.Cells[inRow1, 18].Value = odr.GetValue(16);
          currentWrkSheet.Cells[inRow1, 21].Value = odr.GetValue(17);
          currentWrkSheet.Cells[inRow1, 22].Value = odr.GetValue(18);
          currentWrkSheet.Cells[inRow1, 23].Value = odr.GetValue(19);
          currentWrkSheet.Cells[inRow1, 24].Value = odr.GetValue(20);
          currentWrkSheet.Cells[inRow1, 25].Value = odr.GetValue(21);

          if (prm.IsLogInfo){
            currentWrkSheet.Cells[inRow1, 53].Value = odr.GetValue(22);
            currentWrkSheet.Cells[inRow1, 54].Value = odr.GetValue(23);
            currentWrkSheet.Cells[inRow1, 55].Value = odr.GetValue(24);
            currentWrkSheet.Cells[inRow1, 56].Value = odr.GetValue(25);
          }

          inRow1++;
          oldLocNo = currLocNo;
        }

        odr.Close();
        odr.Dispose();
      }


      var sqlStmtAprOther2 = prm.IsLogInfo ? "SELECT * FROM VIZ_PRN.SM_RAPORT_APR_2_ALL" : "SELECT * FROM VIZ_PRN.SM_RAPORT_APR_2_LALL";
      odr = Odac.GetOracleReader(sqlStmtAprOther2, CommandType.Text, false, null, null);

      if (odr != null){

        int inRow2 = 26 + qntInsert;
        int inRowInsert2 = 28 + qntInsert;

        string oldLocNo = "";
        string currLocNoOut = "";
        string currLocNo = "";
        decimal oldVesZ = 0;
        decimal currVesZ = 0;

        //int colorRow = 49407;

        while (odr.Read()){
          currLocNo = odr.GetString(2);
          currLocNoOut = odr.GetString(12);

          currVesZ = odr.GetDecimal(5);
          

          if (inRow2 == inRowInsert2){
            currentWrkSheet.Rows[inRow2].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2 - 1, 1], currentWrkSheet.Cells[inRow2 - 1, 24]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]]);
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 27]].ClearContents();
            
            inRowInsert2++;
            qntInsert++;
          }
          
          if (odr.GetValue("PR_COLOR") != DBNull.Value && odr.GetInt32("PR_COLOR") > 1){
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]].Interior.Pattern = 1; //xlSolid
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]].Interior.PatternColorIndex = -4105; //xlAutomatic;
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]].Interior.ThemeColor = 6; //xlThemeColorAccent2
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]].Interior.TintAndShade = 0.799981688894314;
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]].Interior.PatternTintAndShade = 0;
          }
          else
            currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 24]].Interior.Pattern = -4142; //xlNone
            
          
          if (
              (oldLocNo != currLocNo) || 
              ((oldLocNo == currLocNo) && (oldVesZ != currVesZ)) ||
              ((oldLocNo == currLocNo) && (currLocNo == currLocNoOut))
             )
          {
            currentWrkSheet.Cells[inRow2, 1].Value = odr.GetValue(0);
            currentWrkSheet.Cells[inRow2, 2].Value = odr.GetValue(1);
            currentWrkSheet.Cells[inRow2, 3].Value = odr.GetValue(2);
            currentWrkSheet.Cells[inRow2, 4].Value = odr.GetValue(3);
            currentWrkSheet.Cells[inRow2, 5].Value = odr.GetValue(4);
            currentWrkSheet.Cells[inRow2, 6].Value = odr.GetValue(5);
            currentWrkSheet.Cells[inRow2, 7].Value = odr.GetValue(6);
            currentWrkSheet.Cells[inRow2, 8].Value = odr.GetValue(7);
            currentWrkSheet.Cells[inRow2, 9].Value = odr.GetValue(8);
            currentWrkSheet.Cells[inRow2, 10].Value = odr.GetValue(9);

          }

          currentWrkSheet.Cells[inRow2, 11].Value = odr.GetValue(10);
          currentWrkSheet.Cells[inRow2, 12].Value = odr.GetValue(11);
          currentWrkSheet.Cells[inRow2, 13].Value = odr.GetValue(12);
          currentWrkSheet.Cells[inRow2, 15].Value = odr.GetValue(13);
          currentWrkSheet.Cells[inRow2, 16].Value = odr.GetValue(14);
          currentWrkSheet.Cells[inRow2, 17].Value = odr.GetValue(15);
          currentWrkSheet.Cells[inRow2, 18].Value = odr.GetValue(16);
          currentWrkSheet.Cells[inRow2, 19].Value = odr.GetValue(17);
          currentWrkSheet.Cells[inRow2, 21].Value = odr.GetValue(18);
          currentWrkSheet.Cells[inRow2, 22].Value = odr.GetValue(19);
          currentWrkSheet.Cells[inRow2, 23].Value = odr.GetValue(20);
          currentWrkSheet.Cells[inRow2, 24].Value = odr.GetValue(21);
          currentWrkSheet.Cells[inRow2, 25].Value = odr.GetValue(22);

          if (prm.IsLogInfo){
            currentWrkSheet.Cells[inRow2, 53].Value = odr.GetValue(23);
            currentWrkSheet.Cells[inRow2, 54].Value = odr.GetValue(24);
            currentWrkSheet.Cells[inRow2, 55].Value = odr.GetValue(25);
            currentWrkSheet.Cells[inRow2, 56].Value = odr.GetValue(26);
            currentWrkSheet.Cells[inRow2, 57].Value = odr.GetValue(27);
            currentWrkSheet.Cells[inRow2, 58].Value = odr.GetValue(28);
          }

          inRow2++;
          oldLocNo = currLocNo;
          oldVesZ = currVesZ;


        }
        odr.Close();
        odr.Dispose();
      }

      decimal? res = GetItogFinishAprOther(prm.FinishAprLabel, "S1");
      if (res !=null)
        currentWrkSheet.Cells[30 + qntInsert, 15].Value = res;

      res = GetItogFinishAprOther(prm.FinishAprLabel, "SD_S1");
      if (res != null)
        currentWrkSheet.Cells[35 + qntInsert, 15].Value = res;

      res = GetItogFinishAprOther(prm.FinishAprLabel, "SD_S2");
      if (res != null)
        currentWrkSheet.Cells[36 + qntInsert, 15].Value = res;
      
      //здесь вытаскиваем производительность
      const string sqlStmtAprOtherProizv = "SELECT PROIZV FROM VIZ_PRN.SM_RAPORT_PROIZV_ALL";
      odr = Odac.GetOracleReader(sqlStmtAprOtherProizv, CommandType.Text, false, null, null);

      if (odr != null){
        odr.Read();
        currentWrkSheet.Cells[33 + qntInsert, 21].Value = odr.GetValue(0);
        odr.Close();
        odr.Dispose();
      }


      //
      const string sqlStmtAprOtherFuter1 = "SELECT * FROM VIZ_PRN.SM_RAPORT_PROSTOI_ALL";
      odr = Odac.GetOracleReader(sqlStmtAprOtherFuter1, CommandType.Text, false, null, null);

      int qntInsertAll = qntInsert;

      if (odr != null){

        int inRowFuter1 = 41 + qntInsert;
        int inRowInsertFuter1 = 47 + qntInsert;


        while (odr.Read()){
          if (inRowFuter1 == inRowInsertFuter1){
            currentWrkSheet.Rows[inRowFuter1].Insert();
            currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter1 - 1, 1], currentWrkSheet.Cells[inRowFuter1 - 1, 23]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter1, 1], currentWrkSheet.Cells[inRowFuter1, 23]]);
            inRowInsertFuter1++;
            //qntInsert++;
            qntInsertAll++;
          }

          currentWrkSheet.Cells[inRowFuter1, 1].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRowFuter1, 3].Value = odr.GetValue(1);
          currentWrkSheet.Cells[inRowFuter1, 4].Value = odr.GetValue(2);
          currentWrkSheet.Cells[inRowFuter1, 6].Value = odr.GetValue(3);
          inRowFuter1++;
        }
        odr.Close();
        odr.Dispose();
      }

      //
      const string sqlStmtAprOtherFuter2 = "SELECT * FROM VIZ_PRN.SM_RAPORT_FIO";
      odr = Odac.GetOracleReader(sqlStmtAprOtherFuter2, CommandType.Text, false, null, null);

      if (odr != null){
        int inRowFuter2 = 41 + qntInsert;

        while (odr.Read()){
          currentWrkSheet.Cells[inRowFuter2, 11].Value = odr.GetValue(0);
          currentWrkSheet.Cells[inRowFuter2, 12].Value = odr.GetValue(1);
          inRowFuter2++;
        }
        odr.Close();
        odr.Dispose();
      }

      currentWrkSheet.PageSetup.PrintArea = "$A$1:$U$" + (50 + qntInsertAll).ToString();
  }


  private decimal? GetItogFinishAprOther(string aprNameRus, string fieldName)
  {
    string stmt = "SELECT " + fieldName  + " FROM VIZ_PRN.SM_RAPORT_SORT WHERE (ANLAGE = :PANLAGE)";

    var lstPrm = new List<OracleParameter>();
    var prm = new OracleParameter
    {
      ParameterName = "PANLAGE",
      DbType = DbType.String,
      Direction = ParameterDirection.Input,
      OracleDbType = OracleDbType.VarChar,
      Size = aprNameRus.Length,
      Value = aprNameRus
    };
    lstPrm.Add(prm);

    object res = Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm);

    if (res != DBNull.Value)
    
      //MessageBox.Show("NULL");
      return Convert.ToDecimal(res);
    else
      return null;
  }


  private Boolean RunRpt(ShiftRptFinishRptParam prm, dynamic currentWrkSheet)
  {
      OracleDataReader odr = null;
      Boolean result = false;
      
      try{
        var idxActiveSheet = GetOneVisibleSheet(prm);

        DbVar.SetRangeDate(prm.DateBegin, prm.DateBegin, 0);
        DbVar.SetString(prm.FinishAprLabel, prm.TeamFinishApr);
        var dtBegin = DbVar.GetDateBeginEnd(true, false);
        DbVar.GetDateBeginEnd(false, false);

        const string sqlStmt = "VIZ_PRN.SMEN_RAPORT_APRUO_ALL.preSM_Raport_ALL";
        var lstOraPrm = new List<OracleParameter>()
        {
          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value = dtBegin
          },

          new OracleParameter()
          {
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = string.IsNullOrEmpty(prm.FinishAprLabel) ? 0 : prm.FinishAprLabel.Length,
            Value = prm.FinishAprLabel
          },
           
          new OracleParameter()
          {
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = string.IsNullOrEmpty(prm.TeamFinishApr) ? 0 : prm.TeamFinishApr.Length,
            Value = prm.TeamFinishApr
          },

          new OracleParameter()
          {
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = string.IsNullOrEmpty(prm.TeamFinishApr) ? 0 : prm.FinishApr.Length,
            Value = prm.FinishApr
          },

          new OracleParameter()
          {
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = string.IsNullOrEmpty(prm.TypeShiftFinishApr) ? 0 : prm.TypeShiftFinishApr.Length,
            Value = prm.TypeShiftFinishApr
          },
        };

        Odac.ExecuteNonQuery(sqlStmt, CommandType.StoredProcedure, false, lstOraPrm);

        if (idxActiveSheet == 0){
          //здесь идет заполнение всех листов (для всех АПР)

          foreach (int idxSheet in Enum.GetValues(typeof(SheetsOfApr))){

            var name = Enum.GetName(typeof(SheetsOfApr), idxSheet);

            if (String.IsNullOrEmpty(name))
              continue;
            
            prm.FinishApr = name.ToUpper(CultureInfo.InvariantCulture);
            prm.FinishAprLabel = "АПР" + prm.FinishApr.Substring(3);

            //MessageBox.Show(prm.FinishApr);
            //MessageBox.Show(prm.FinishAprLabel);

            //Выставляем переменные среды
            DbVar.SetString(prm.FinishAprLabel, prm.TeamFinishApr);
            
            if (GetTypeApr(name.ToUpper(CultureInfo.InvariantCulture)) == TypeFinishApr.FinishApr12)
              FinishApr12(prm, dtBegin, idxSheet);
            else if (GetTypeApr(name.ToUpper(CultureInfo.InvariantCulture)) == TypeFinishApr.FinishAprLaser)
              FinishAprLaser(prm, dtBegin, idxSheet);
            else if (GetTypeApr(name.ToUpper(CultureInfo.InvariantCulture)) == TypeFinishApr.FinishAprOther)
              FinishAprOther(prm, dtBegin, idxSheet);
            else
              prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart) (() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "Шаблона под выбранный агрегат не существует!", MessageBoxImage.Stop)));
          }

          prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        }
        else{
          if (GetTypeApr(prm.FinishApr) == TypeFinishApr.FinishApr12)
            FinishApr12(prm, dtBegin, idxActiveSheet);
          else if (GetTypeApr(prm.FinishApr) == TypeFinishApr.FinishAprLaser)
            FinishAprLaser(prm, dtBegin, idxActiveSheet);
          else if (GetTypeApr(prm.FinishApr) == TypeFinishApr.FinishAprOther)
            FinishAprOther(prm, dtBegin, idxActiveSheet);
          else
            prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "Шаблона под выбранный агрегат не существует!", MessageBoxImage.Stop)));
        }
        
        result = true;
      }
      
      catch (Exception e){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.Message, MessageBoxImage.Stop)));
        result = false;
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


  }






}
