using System;
using System.Collections.Generic;
using System.Data;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.WrkModule.Isc.Db
{
  public static class IscAction
  {
    public static int IsRewindAfterLasScr(string meId)
    {
      const string stmtSql = "VIZ_PRN.ISC.IsRwdAfterLasScr";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(meId))
        len = meId.Length;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = len,
        Value = meId
      };
      lstPrm.Add(prm);

      var prmRetVal = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.ReturnValue,
        OracleDbType = OracleDbType.Integer,
      };
      lstPrm.Add(prmRetVal);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);

      return Convert.ToInt32(prmRetVal.Value);
    }
    


  }

}