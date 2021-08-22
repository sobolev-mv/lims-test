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

namespace Viz.WrkModule.RptMagLab.Db
{
  public sealed class CzlEfLsrRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd   { get; set; }
    public decimal  P1750023  { get; set; }
    public decimal  P1750027  { get; set; }
    public decimal  P1750030  { get; set; }
    public decimal  B800      { get; set; }
    public decimal  KesiAvg   { get; set; }
    public decimal  CoefVoln  { get; set; }
    public decimal  QntShov   { get; set; }
    public string Sort        { get; set; }
    public string AdgIn       { get; set; }
    public string AdgOut      { get; set; }
    public int TypeRpt        { get; set; }
    public string ListVal     { get; set; }
    public string AdgInFlt    { get; set; }
    public string AdgOutFlt   { get; set; }   


    public CzlEfLsrRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd,
                            decimal P1750023, decimal P1750027, decimal P1750030, decimal B800, decimal KesiAvg,
                            decimal CoefVoln, decimal QntShov, string Sort, string AdgIn, string AdgOut, int TypeRpt, string ListVal,
                            string AdgInFlt, string AdgOutFlt)
      : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.P1750023 = P1750023;
      this.P1750027 = P1750027;
      this.P1750030 = P1750030;
      this.B800 = B800;
      this.KesiAvg = KesiAvg;
      this.CoefVoln = CoefVoln;
      this.QntShov = QntShov;
      this.Sort = Sort;
      this.AdgIn = AdgIn;
      this.AdgOut = AdgOut;
      this.TypeRpt = TypeRpt;
      this.ListVal = ListVal;
      this.AdgInFlt = AdgInFlt;
      this.AdgOutFlt = AdgOutFlt;
    }
  }

  public sealed class CzlEfLsr : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as CzlEfLsrRptParam);
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

        //Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }

      //вызывается в случае перключения целевой БД
      //base.DoWorkXls(sender, e);
    }

    private void ListFilterInfoToExcel(CzlEfLsrRptParam prm)
    {
      dynamic wrkSheet = null;
      //выбираем лист
      prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
      wrkSheet = prm.ExcelApp.ActiveSheet;
      const int row = 3;
      string[] strArr = prm.ListVal.Split(new char[] { ',' });
      for (int i = 0; i < strArr.Length; i++) wrkSheet.Cells[row + i, 1].Value = strArr[i];
    }


    private Boolean RunRpt(CzlEfLsrRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      int row = 0;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[4, 3].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
        
        string strFlt = "P1.7/50-0.23<=" + prm.P1750023.ToString("n2") + "; " +
                        "P1.7/50-0.27<=" + prm.P1750027.ToString("n2") + "; " +
                        "P1.7/50-0.30<=" + prm.P1750027.ToString("n2") + "; " +
                        "B800>=" + prm.B800.ToString("n2") + "; " +
                        "Сорт=" + prm.Sort + "; " +
                        "КЭСИ ср.>=" + prm.KesiAvg.ToString("n0") + "; " +
                        "Адг.Внутр<=" + prm.AdgInFlt + "; " +
                        "Адг.Внеш=" + prm.AdgOutFlt + "; " + 
                        "Коэфф.Волны<=" + prm.CoefVoln.ToString("n2") + "; " +
                        "Кол-во швов<=" + prm.QntShov.ToString("n0");

        CurrentWrkSheet.Cells[6, 3].Value = strFlt;
        switch (prm.TypeRpt){
          case 0:
            //Общий отчет
            string SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_10 ORDER BY 1";
            DbVar.SetString(prm.Sort, prm.AdgIn, prm.AdgOut, string.Empty);
            DbVar.SetNum(prm.P1750023, prm.P1750027, prm.P1750030, prm.B800, prm.KesiAvg, prm.CoefVoln, prm.QntShov);

            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
            

            if (odr == null) return false;
            row = 12;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B100");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("B100_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 16].Value = odr.GetValue("EF_LSR_T");
              } else{
                CurrentWrkSheet.Cells[15, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[15, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[15, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[15, 7].Value = odr.GetValue("B100");
                CurrentWrkSheet.Cells[15, 8].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[15, 9].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[15, 10].Value = odr.GetValue("B100_T");
                CurrentWrkSheet.Cells[15, 11].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[15, 12].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[15, 13].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[15, 14].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[15, 15].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[15, 16].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            odr.Close();
            odr.Dispose();
            
            //----------------
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_14 ORDER BY 1";
            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 16;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B100");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("B100_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 16].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[19, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[19, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[19, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[19, 7].Value = odr.GetValue("B100");
                CurrentWrkSheet.Cells[19, 8].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[19, 9].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[19, 10].Value = odr.GetValue("B100_T");
                CurrentWrkSheet.Cells[19, 11].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[19, 12].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[19, 13].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[19, 14].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[19, 15].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[19, 16].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            odr.Close();
            odr.Dispose();
            
            //----------------
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP ORDER BY 1";
            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 20;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B100");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("B100_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 16].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[23, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[23, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[23, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[23, 7].Value = odr.GetValue("B100");
                CurrentWrkSheet.Cells[23, 8].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[23, 9].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[23, 10].Value = odr.GetValue("B100_T");
                CurrentWrkSheet.Cells[23, 11].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[23, 12].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[23, 13].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[23, 14].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[23, 15].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[23, 16].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            break;
          case 1:
            ListFilterInfoToExcel(prm);
            prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
            CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

            //По списку стендовых партий 
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_10_SPIS ORDER BY 1";
            DbVar.SetString(prm.Sort, prm.AdgIn, prm.AdgOut, prm.ListVal);
            DbVar.SetNum(prm.P1750023, prm.P1750027, prm.P1750030, prm.B800, prm.KesiAvg, prm.CoefVoln, prm.QntShov);

            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 12;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[15, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[15, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[15, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[15, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[15, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[15, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[15, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[15, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[15, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[15, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[15, 14].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            odr.Close();
            odr.Dispose();

            //----------------
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_14_SPIS ORDER BY 1";
            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 16;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[19, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[19, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[19, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[19, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[19, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[19, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[19, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[19, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[19, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[19, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[19, 14].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            odr.Close();
            odr.Dispose();

            //----------------
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_SPIS ORDER BY 1";
            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 20;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[23, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[23, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[23, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[23, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[23, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[23, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[23, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[23, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[23, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[23, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[23, 14].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }

            break;
          case 2:
            ListFilterInfoToExcel(prm);
            prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
            CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

            //Без учета списка стендовых партий 
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_10_NSPIS ORDER BY 1";
            DbVar.SetString(prm.Sort, prm.AdgIn, prm.AdgOut, prm.ListVal);
            DbVar.SetNum(prm.P1750023, prm.P1750027, prm.P1750030, prm.B800, prm.KesiAvg, prm.CoefVoln, prm.QntShov);

            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 12;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[15, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[15, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[15, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[15, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[15, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[15, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[15, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[15, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[15, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[15, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[15, 14].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            odr.Close();
            odr.Dispose();

            //----------------
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_14_NSPIS ORDER BY 1";
            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 16;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP"); 
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[19, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[19, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[19, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[19, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[19, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[19, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[19, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[19, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[19, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[19, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[19, 14].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }
            odr.Close();
            odr.Dispose();

            //----------------
            SqlStmt = "SELECT * FROM VIZ_PRN.CZL_LSR_TXP_NSPIS ORDER BY 1";
            odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);

            if (odr == null) return false;

            row = 20;
            while (odr.Read()){
              if (Convert.ToInt32(Convert.ToInt32(odr.GetValue("TOLS"))) != 99){
                CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
                CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("EF_LSR_T");
              }
              else{
                CurrentWrkSheet.Cells[23, 4].Value = odr.GetValue("VES");
                CurrentWrkSheet.Cells[23, 5].Value = odr.GetValue("VES_T");
                CurrentWrkSheet.Cells[23, 6].Value = odr.GetValue("PROC_TXP");
                CurrentWrkSheet.Cells[23, 7].Value = odr.GetValue("B800");
                CurrentWrkSheet.Cells[23, 8].Value = odr.GetValue("P1750AP");
                CurrentWrkSheet.Cells[23, 9].Value = odr.GetValue("B800_T");
                CurrentWrkSheet.Cells[23, 10].Value = odr.GetValue("P1750AP_T");
                CurrentWrkSheet.Cells[23, 11].Value = odr.GetValue("VES_1");
                CurrentWrkSheet.Cells[23, 12].Value = odr.GetValue("EF_LSR");
                CurrentWrkSheet.Cells[23, 13].Value = odr.GetValue("VES_2");
                CurrentWrkSheet.Cells[23, 14].Value = odr.GetValue("EF_LSR_T");
              }
              row++;
            }

            break;
          default:
            break;
        }

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




