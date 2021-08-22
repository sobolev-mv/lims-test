using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using System.Data;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptMagLab.Db
{
  public sealed class CzlOutP1750RptParam : Smv.Xls.XlsInstanceParam
  {
    public string PathScriptsDir { get; set; }
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string Rm1200 { get; set; }
    public string Aro { get; set; }
    public string Aoo { get; set; }
    public string Avo { get; set; }
    public Boolean IsRm1200 { get; set; }
    public Boolean IsAro { get; set; }
    public Boolean IsAoo { get; set; }
    public Boolean IsAvo { get; set; }
    public int TypeRpt { get; set; }
    public int TypeList { get; set; }
    public string ListVal { get; set; }

    public CzlOutP1750RptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }




  public sealed class CzlOutP1750 : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as CzlOutP1750RptParam);
      dynamic wrkSheet = null;

      try{

        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);
      }
      catch (Exception ex)
      {
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally
      {
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

    private Boolean RunScriptType0(CzlOutP1750RptParam prm)
    {
      var sqlStmt = System.IO.File.ReadAllText(prm.PathScriptsDir + "\\CzlFilterCoreP1750.sql", Encoding.GetEncoding(1251)).Replace("\r", " ");
      IAsyncResult iar = null;
      List<OracleParameter> lstParam = null;

      lstParam = new List<OracleParameter>();

      //ST1200
      var param = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "S1200",
        Value = prm.Rm1200,
        Size = prm.Rm1200.Length
      };
      lstParam.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        ParameterName = "F1200",
        Value = prm.IsRm1200 ? 1 : 0
      };
      lstParam.Add(param);
     
      //ARO
      param = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "ARO",
        Value = prm.Aro,
        Size = prm.Aro.Length
      };
      lstParam.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        ParameterName = "FARO",
        Value = prm.IsAro ? 1 : 0
      };
      lstParam.Add(param);

      //АОО
      param = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "AOO",
        Value = prm.Aoo,
        Size = prm.Aoo.Length
      };
      lstParam.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        ParameterName = "FAOO",
        Value = prm.IsAoo ? 1 : 0
      };
      lstParam.Add(param);

      //АВО
      param = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "AVO",
        Value = prm.Avo,
        Size = prm.Avo.Length
      };
      lstParam.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        ParameterName = "FAVO",
        Value = prm.IsAvo ? 1 : 0
      };
      lstParam.Add(param);

      prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(sqlStmt, CommandType.Text, false, true, lstParam); }));
      if (iar != null)
        iar.AsyncWaitHandle.WaitOne();
      else
        return false;

      var oracleCommand = iar.AsyncState as OracleCommand;
      if (oracleCommand != null){
        oracleCommand.EndExecuteNonQuery(iar);
        iar = null;
      }

      return true;
    }

    private Boolean RunScriptType1(CzlOutP1750RptParam prm)
    {
      const string sqlStmt = @"BEGIN
                        DELETE FROM VIZ_PRN.TMP_CZL_FILTR_CORE;
                        INSERT INTO VIZ_PRN.TMP_CZL_FILTR_CORE 
                        SELECT * 
                        FROM VIZ_PRN.CZL_P1750_CORE
                        WHERE $$$2$$$ IN (SELECT VL_STRING FROM TABLE(VIZ_PRN.VAR_RPT.GetTabOfStrDelim(:VALLIST, ',')));
                     END;";

      string sqlStmtVar = null;

      switch (prm.TypeList){
        case 0:
          sqlStmtVar = sqlStmt.Replace("$$$2$$$", "STEND");
          break;
        case 1:
          sqlStmtVar = sqlStmt.Replace("$$$2$$$", "ST_VTO");
          break;
        case 2:
          sqlStmtVar = sqlStmt.Replace("$$$2$$$", "PLAVKA");
          break;
        default:
          break;
      }

      IAsyncResult iar = null;
      var lstParam = new List<OracleParameter>();

      var param = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "VALLIST",
        Value = prm.ListVal,
        Size = prm.ListVal.Length
      };
      lstParam.Add(param);

      prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(sqlStmtVar, CommandType.Text, false, true, lstParam); }));
      if (iar != null)
        iar.AsyncWaitHandle.WaitOne();
      else
        return false;

      var oracleCommand = iar.AsyncState as OracleCommand;
      if (oracleCommand != null){
        oracleCommand.EndExecuteNonQuery(iar);
        iar = null;
      }

      return true;
    }

    private Boolean RunScriptType2(CzlOutP1750RptParam prm)
    {
      const string sqlStmt = @"BEGIN
                        DELETE FROM VIZ_PRN.TMP_CZL_FILTR_CORE;
                        INSERT INTO VIZ_PRN.TMP_CZL_FILTR_CORE 
                        SELECT * 
                        FROM VIZ_PRN.CZL_P1750_CORE
                        WHERE $$$2$$$ NOT IN (SELECT VL_STRING FROM TABLE(VIZ_PRN.VAR_RPT.GetTabOfStrDelim(:VALLIST, ',')));
                     END;";

      string sqlStmtVar = null;

      switch (prm.TypeList)
      {
        case 0:
          sqlStmtVar = sqlStmt.Replace("$$$2$$$", "STEND");
          break;
        case 1:
          sqlStmtVar = sqlStmt.Replace("$$$2$$$", "ST_VTO");
          break;
        case 2:
          sqlStmtVar = sqlStmt.Replace("$$$2$$$", "PLAVKA");
          break;
        default:
          break;
      }

      IAsyncResult iar = null;
      var lstParam = new List<OracleParameter>();

      var param = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "VALLIST",
        Value = prm.ListVal,
        Size = prm.ListVal.Length
      };
      lstParam.Add(param);

      prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(sqlStmtVar, CommandType.Text, false, true, lstParam); }));
      if (iar != null)
        iar.AsyncWaitHandle.WaitOne();
      else
        return false;

      var oracleCommand = iar.AsyncState as OracleCommand;
      if (oracleCommand != null)
      {
        oracleCommand.EndExecuteNonQuery(iar);
        iar = null;
      }

      return true;
    }

    private string GetFilterStringType0(CzlOutP1750RptParam prm)
    {
      string strFlt = "";

      if (prm.TypeRpt == 0){
        if (prm.IsRm1200)
          strFlt += ":Ст1200 =" + prm.Rm1200;
        if (prm.IsAro)
          strFlt += ":АРО=" + prm.Aro;
        if (prm.IsAoo)
          strFlt += ":АОО=" + prm.Aoo;
        if (prm.IsAvo)
          strFlt += ":АВО=" + prm.Avo;
      }
      else{
        switch (prm.TypeList){
          case 0:
            strFlt = "Список стендов: " + prm.ListVal;
            break;
          case 1:
            strFlt = "Список стендов ВТО: " + prm.ListVal;
            break;
          case 2:
            strFlt = "Список плавок: " + prm.ListVal;
            break;
          default:
            break;
        }
      }

      return strFlt;
    }



    private Boolean RunRpt(CzlOutP1750RptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      int row = 0;

      try{
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        CurrentWrkSheet.Cells[2, 3].Value = "c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        switch (prm.TypeRpt){
          case 0:
            RunScriptType0(prm);
            break;
          case 1:
             RunScriptType1(prm);
            break;
          case 2:
             RunScriptType2(prm);
            break;
          default:
            break;
        }
        CurrentWrkSheet.Cells[3, 3].Value = GetFilterStringType0(prm);

        const string SqlStmt1 = "SELECT * FROM VIZ_PRN.CZL_P1750"; 
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString("APR")));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt1, System.Data.CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){

          while (odr.Read()){
            decimal thickness = odr.GetDecimal("TOLS");

            if (thickness == 0.23m)
              row = 7;
            else if (thickness == 0.27m)
              row = 8;
            else if (thickness == 0.30m)
              row = 9;
            else if (thickness == 0.35m)
              row = 10;

            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_ALL");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }


        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString("SGP")));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt1, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          while (odr.Read()){

            decimal thickness = odr.GetDecimal("TOLS");

            if (thickness == 0.23m)
              row = 15;
            else if (thickness == 0.27m)
              row = 16;
            else if (thickness == 0.30m)
              row = 17;
            else if (thickness == 0.35m)
              row = 18;

            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_ALL");
            row++;
          }
        }

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

