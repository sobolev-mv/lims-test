using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class KpaRollingRptParam : Smv.Xls.XlsInstanceParam
  {
    public Boolean IsListStendF4 { get; set; }
    public DateTime DateBegin1F4 { get; set; }
    public DateTime DateEnd1F4 { get; set; }
    public DateTime DateBegin2F4 { get; set; }
    public DateTime DateEnd2F4 { get; set; }
    public DateTime DateBegin3F4 { get; set; }
    public DateTime DateEnd3F4 { get; set; }
    public string ListStendF4 { get; set; }

    public KpaRollingRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class KpaRolling : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as KpaRollingRptParam);
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

        //Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private void Prepare(int typePriod, DateTime dtBegin, DateTime dtEnd, string agTyp)
    {
      const string sqlStmt1 = "VIZ_PRN.PU.PREPU";

      List<OracleParameter> lstPrm = new List<OracleParameter>();
      OracleParameter oraParam = new OracleParameter
      {
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        Value = typePriod
      };
      lstPrm.Add(oraParam);

      oraParam = new OracleParameter
      {
        DbType = DbType.DateTime,
        OracleDbType = OracleDbType.Date,
        Direction = ParameterDirection.Input,
        Value = dtBegin
      };
      lstPrm.Add(oraParam);

      oraParam = new OracleParameter
      {
        DbType = DbType.DateTime,
        OracleDbType = OracleDbType.Date,
        Direction = ParameterDirection.Input,
        Value = dtEnd
      };
      lstPrm.Add(oraParam);

      oraParam = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        Size = 60, 
        Value = agTyp
      };
      lstPrm.Add(oraParam);
      
      Odac.ExecuteNonQuery(sqlStmt1, CommandType.StoredProcedure, false, lstPrm);
    }

    private void FillTableAgr(string nameView, List<string> agrLst, List<int> rowLst, int col, dynamic wrkSheet)
    {
      OracleDataReader odr = null;
      string sqlStmt = "SELECT * FROM VIZ_PRN." + nameView;
      odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

      if (odr != null){
        
        while (odr.Read()){
          int idx = agrLst.IndexOf(Convert.ToString(odr.GetValue(0)));

          if (idx == -1)
            continue;

          int row = rowLst[idx];
          
          for (int i = 0; i < 4; i++)
            wrkSheet.Cells[row, col + i].Value = odr.GetValue(i + 1);
        }
        odr.Close();
        odr.Dispose();
      }
    }
    private void FillTableDefAndAgr(string nameView, List<string> agrLst, List<int> rowLst, int col, dynamic wrkSheet)
    {
      OracleDataReader odr = null;
      string sqlStmt = "SELECT * FROM VIZ_PRN." + nameView;
      odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

      if (odr != null){

        while (odr.Read()){
          int idx = agrLst.IndexOf(Convert.ToString(odr.GetValue(0)) + Convert.ToString(odr.GetValue(1)));

          if (idx == -1)
            continue;

          int row = rowLst[idx];

          for (int i = 0; i < 4; i++)
            wrkSheet.Cells[row, col + i].Value = odr.GetValue(i + 2);
        }
        odr.Close();
        odr.Dispose();
      }
    }

    private Boolean RunRpt(KpaRollingRptParam prm, dynamic CurrentWrkSheet)
    {
      Boolean Result = false;

      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      
      try{
        //Debug.WriteLine("Start: "  + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.TMP_PU", CommandType.Text, false, null);
        DbVar.SetRangeDate(prm.DateBegin1F4 , prm.DateEnd1F4, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[3, 3].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        DbVar.SetRangeDate(prm.DateBegin2F4, prm.DateEnd2F4, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[4, 3].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        DbVar.SetRangeDate(prm.DateBegin3F4, prm.DateEnd3F4, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[5, 3].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        //Debug.WriteLine("Before Prepare: " + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        Prepare(1, prm.DateBegin1F4, prm.DateEnd1F4, "1STROLL,1STCUT,2NDROLL,2NDCUT,DECARB,ISOGO");
        //Debug.WriteLine("Prepare1: " + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        Prepare(2, prm.DateBegin2F4, prm.DateEnd2F4, "1STCUT,2NDCUT");
        //Debug.WriteLine("Prepare2: " + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        Prepare(3, prm.DateBegin3F4, prm.DateEnd3F4, "2NDROLL");
        //Debug.WriteLine("Prepare3: " + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        Odac.ExecuteNonQuery("VIZ_PRN.PU.PU_RASCHET", CommandType.StoredProcedure, false, null);
        //Debug.WriteLine("PU_RASCHET: " + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        if (prm.IsListStendF4){
          DbVar.SetString(prm.ListStendF4);
          Odac.ExecuteNonQuery("VIZ_PRN.PU.PU_FILTR", CommandType.StoredProcedure, false, null);
          CurrentWrkSheet.Cells[7, 3].Value = prm.ListStendF4;
        }

        var agrLst = new List<string>{"RM1300","APR1","RM12001","RM12002","RRM","APR8"};
        var rowLst = new List<int>{11, 14, 19, 23, 27, 31};
        var col = 4;
        FillTableAgr("V_PU_PROIZV", agrLst, rowLst, col, CurrentWrkSheet);

        rowLst.Clear();
        rowLst.InsertRange(0,  new int[]{13, 15, 22, 26, 30, 33});
        FillTableAgr("V_PU_RK", agrLst, rowLst, col, CurrentWrkSheet);

        rowLst.Clear();
        agrLst.Clear();
        agrLst.InsertRange(0, new string[] { "248RM1300", "248RM12001", "247RM12001", "248RM12002", "247RM12002", "248RRM", "247RRM" });
        rowLst.InsertRange(0, new int[] { 12, 20, 21, 24, 25, 28, 29 });
        FillTableDefAndAgr("V_PU_247_248", agrLst, rowLst, col, CurrentWrkSheet);

        rowLst.Clear();
        agrLst.Clear();
        agrLst.InsertRange(0, new string[] { "APR1" });
        rowLst.InsertRange(0, new int[] { 18 });
        FillTableAgr("V_PU_TN_DEF", agrLst, rowLst, col, CurrentWrkSheet);

        rowLst.Clear();
        agrLst.Clear();
        agrLst.InsertRange(0, new string[] { "1APR1", "2APR1", "3APR8" });
        rowLst.InsertRange(0, new int[] { 16, 17, 32 });
        FillTableDefAndAgr("V_PU_230", agrLst, rowLst, col, CurrentWrkSheet);

        //Debug.WriteLine("FillTable: " + $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}");

        CurrentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      finally{
        ;
      }

      return Result;
    }


  }

}

