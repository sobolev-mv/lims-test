using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Smv.Utils;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptManager.Db
{
  public static class DbUtils
  {
    public static DateTime GetLastDateDynDef()
    {
      const string stmt = "SELECT TRUNC(MAX(DATA), 'DD') FROM VIZ_PRN.OTK_DINAMIKA_KAT_SORT";
      return Convert.ToDateTime(Odac.ExecuteScalar(stmt, CommandType.Text, false, null));
    }

    public static Boolean AddNewDateRange(DateTime dateFrom)
    {
      const string stmtSql = "VIZ_PRN.DG_DCBLMETAL.AddNewDateRange";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      OracleParameter prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        OracleDbType = OracleDbType.Date,
        Direction = ParameterDirection.Input,
        Size = 64,
        Value = dateFrom
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean DeleteLastDateRange()
    {
      const string stmtSql = "VIZ_PRN.DG_DCBLMETAL.DeleteLastDateRange";
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, null);
    }

    public static Boolean LoadTblExcept()
    {
      const string delim = ",";
      const string sqlStmtRowCnt = "SELECT COUNT(*) FROM VIZ_PRN.OTK_DEF_ISKL";

      int rowCnt1 = Convert.ToInt32(Odac.ExecuteScalar(sqlStmtRowCnt, CommandType.Text, false, null));

      var strLst = Etc.GetStringWithDelimFromTxtFile(Encoding.GetEncoding("windows-1251"), delim);
      DbVar.SetStringList(strLst, delim);
      var sqlStmt = File.ReadAllText(Etc.StartPath + "\\Scripts\\LoadTblExcept.sql", Encoding.GetEncoding(1251)).Replace("\r", " ");

      Boolean res = Odac.ExecuteNonQuery(sqlStmt, CommandType.Text, false, null);

      if (res){
        var rowInsert = Convert.ToInt32(Odac.ExecuteScalar(sqlStmtRowCnt, CommandType.Text, false, null)) - rowCnt1;
        DxInfo.ShowDxBoxInfo("Загрузка лок. номеров", "Загружено: " + rowInsert + " записей", MessageBoxImage.Information);
      }

      return res;
    }

    public static DateTime GetDateBeginQuart()
    {
      const string stmtSql = "select DTBEGIN from VIZ_PRN.DG_QUARTILE where ID = 1";
      return Convert.ToDateTime(Odac.ExecuteScalar(stmtSql, CommandType.Text, false, null));
    }

    public static DateTime GetDateEndQuart()
    {
      const string stmtSql = "select DTEND from VIZ_PRN.DG_QUARTILE where ID = 1";
      return Convert.ToDateTime(Odac.ExecuteScalar(stmtSql, CommandType.Text, false, null));
    }

    public static void SaveDateQuart(DateTime dateBegin, DateTime dateEnd)
    {
      const string stmtSql = "UPDATE VIZ_PRN.DG_QUARTILE SET DTBEGIN = TRUNC(:PDTBEGIN, 'MM'), DTEND = TRUNC(:PDTEND, 'MM'), DTUPDT = SYSDATE WHERE ID = 1";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      OracleParameter prm = new OracleParameter
      {
        ParameterName = "PDTBEGIN",
        DbType = DbType.DateTime,
        OracleDbType = OracleDbType.Date,
        Direction = ParameterDirection.Input,
        Value = dateBegin
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        ParameterName = "PDTEND",
        DbType = DbType.DateTime,
        OracleDbType = OracleDbType.Date,
        Direction = ParameterDirection.Input,
        Value = dateEnd
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.Text, false, lstPrm, true);
    }

  }
}
