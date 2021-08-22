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

namespace Viz.WrkModule.RptOpr.Db
{
  public sealed class ProcLaserAndAprRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public Boolean IsGroupParam1F1 { get; set; }
    public decimal P1750023LstF1 { get; set; }
    public decimal P1750027LstF1 { get; set; }
    public decimal P1750030LstF1{ get; set; }
    public decimal B800LstF1 { get; set; }
    public int QntWeldsF1 { get; set; }
    public int KesiAvgF1 { get; set; }
    public string Cat1F1 { get; set; }
    public string Cat2F1 { get; set; }
    public string Cat3F1 { get; set; }
    public string CatWcF1 { get; set; }
    public string AdgInF1 { get; set; }
    public string AdgOutF1 { get; set; }
    public decimal CoffWave1F1 { get; set; }
    public decimal CoffWave2F1 { get; set; }
    public decimal HeightWave1F1 { get; set; }
    public decimal HeightWave2F1 { get; set; }
    public string ClsNoPloskF1 { get; set; }
    public string TargetNextProcF1 { get; set; }
    //public int IdtargetNextProcF1 { get; set; }
    public ProcLaserAndAprRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class ProcLaserAndApr : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as ProcLaserAndAprRptParam);
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
        prm.WorkBook.Close();
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        if (prm.ExcelApp != null)
          Marshal.ReleaseComObject(prm.ExcelApp);

        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(ProcLaserAndAprRptParam prm, dynamic currentWrkSheet)
    {
      
      var result = false;
      OracleDataReader odr = null;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);
        currentWrkSheet.Cells[2, 6].Value = "за период с " + $"{dtBegin:dd.MM.yyyy HH:mm:ss}" + " по " + $"{dtEnd:dd.MM.yyyy HH:mm:ss}";

        const string stmtSql1 = "DELETE FROM VIZ_PRN.TMP_LK_APR";
        Odac.ExecuteNonQuery(stmtSql1, CommandType.Text, false, null);
        const string stmtSql2 = "INSERT INTO VIZ_PRN.TMP_LK_APR SELECT * FROM VIZ_PRN.V_LK_APR_CORE";
        Odac.ExecuteNonQuery(stmtSql2, CommandType.Text, false, null);
        /*
        const string stmtSqlZ = "SELECT COUNT(*) FROM VIZ_PRN.TMP_LK_APR";
        object o = Odac.ExecuteScalar(stmtSqlZ, CommandType.Text, false, null);
        MessageBox.Show(Convert.ToString(o), "");
        */
        const string stmtSql3 = "DELETE FROM VIZ_PRN.TMP_LK_APR_TXP";
        Odac.ExecuteNonQuery(stmtSql3, CommandType.Text, false, null);

        if (prm.IsGroupParam1F1){
          
          const string stmtSql4 = "INSERT INTO VIZ_PRN.TMP_LK_APR_TXP" +
                                  "(VP1750_023, VP1750_027, VP1750_030, VB800, VKESI, VKAT1, VKAT2, VKAT3, VKAT4, VSHOV, VADIN, VADOUT, VKOEF_VOLN_1, VVYSOTAVOLN_1, VKOEF_VOLN_2, VVYSOTAVOLN_2, VSTATUS_MET, VCLASSNEPL)" +
                                  "VALUES" +
                                  "(:PVP1750_023, :PVP1750_027, :PVP1750_030, :PVB800, :PVKESI, :PVKAT1, :PVKAT2, :PVKAT3, :PVKAT4, :PVSHOV, :PVADIN, :PVADOUT, :PVKOEF_VOLN_1, :PVVYSOTAVOLN_1, :PVKOEF_VOLN_2, :PVVYSOTAVOLN_2, :PVSTATUS_MET, :PVCLASSNEPL)";

          var lstParam = new List<OracleParameter>();

          var param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 2,
            ParameterName = "PVP1750_023",
            Value = prm.P1750023LstF1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 2,
            ParameterName = "PVP1750_027",
            Value = prm.P1750027LstF1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 2,
            ParameterName = "PVP1750_030",
            Value = prm.P1750030LstF1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 2,
            ParameterName = "PVB800",
            Value = prm.B800LstF1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "PVKESI",
            Value = prm.KesiAvgF1
          };
          lstParam.Add(param);
          
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVKAT1",
            Value = prm.Cat1F1 == "Y" ? "1" : null,
            Size = prm.Cat1F1 == "Y" ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVKAT2",
            Value = prm.Cat2F1 == "Y" ? "2" : null,
            Size = prm.Cat2F1 == "Y" ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVKAT3",
            Value = prm.Cat3F1 == "Y" ? "3" : null,
            Size = prm.Cat3F1 == "Y" ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVKAT4",
            Value = prm.CatWcF1 == "Y" ? "Б/К" : null,
            Size = prm.CatWcF1 == "Y" ? 3 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "PVSHOV",
            Value = prm.QntWeldsF1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVADIN",
            Value = prm.AdgInF1,
            Size = prm.AdgInF1.Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVADOUT",
            Value = prm.AdgOutF1,
            Size = prm.AdgOutF1.Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 1,
            ParameterName = "PVKOEF_VOLN_1",
            Value = prm.CoffWave1F1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 1,
            ParameterName = "PVVYSOTAVOLN_1",
            Value = prm.HeightWave1F1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 1,
            ParameterName = "PVKOEF_VOLN_2",
            Value = prm.CoffWave2F1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Precision = 4,
            Scale = 1,
            ParameterName = "PVVYSOTAVOLN_2",
            Value = prm.HeightWave2F1
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVSTATUS_MET",
            Value = prm.TargetNextProcF1,
            Size = prm.TargetNextProcF1.Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PVCLASSNEPL",
            Value = prm.ClsNoPloskF1,
            Size = prm.ClsNoPloskF1.Length
          };
          lstParam.Add(param);

          Odac.ExecuteNonQuery(stmtSql4, CommandType.Text, false, lstParam);
          Odac.ExecuteNonQuery("VIZ_PRN.LK_APR.preLK_APR", CommandType.StoredProcedure, false, null);
        }

        const string stmtSql5 = "SELECT * FROM VIZ_PRN.V_LK_APR";
        odr = Odac.GetOracleReader(stmtSql5, CommandType.Text, false, null, null);
       
        if (odr != null){
          const int columnPart1 = 35;
          const int firstExcelColumn = 1;
          const int lastExcelColumn = 75;
          var row = 6;
          while (odr.Read()){

            currentWrkSheet.Range[currentWrkSheet.Cells[row, firstExcelColumn], currentWrkSheet.Cells[row, lastExcelColumn]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[row + 1, firstExcelColumn], currentWrkSheet.Cells[row + 1, lastExcelColumn]]);

            for (int i = 0; i < columnPart1; i++)
              currentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            for (int i = columnPart1; i < odr.FieldCount; i++)
              currentWrkSheet.Cells[row, i + 3].Value = odr.GetValue(i);

            row++;
          }

          odr.Close();
          odr.Dispose();
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        currentWrkSheet = prm.ExcelApp.ActiveSheet;
        const string stmtSql6 = "SELECT * FROM VIZ_PRN.V_LK_APR_PROC";
        odr = Odac.GetOracleReader(stmtSql6, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 6;
          while (odr.Read()){
            for (int i = 0; i < odr.FieldCount; i++)
              currentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);

            row++;
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        result = true;
      }
      catch (Exception e){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.Message, MessageBoxImage.Stop)));
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
