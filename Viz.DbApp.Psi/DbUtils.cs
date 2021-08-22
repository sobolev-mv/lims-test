using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Navigation;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.DbApp.Psi
{
  public static class Permission
  {
    public static Boolean GetPermissionForModuleUif(int FuncId, string ModuleId)
    {
      int rezVal = 0;
      const string stmtSql = "LIMS.AcsUserModuleFunc";
      OracleParameter prm = null;
      OracleParameter rezPrm = null;

      List<OracleParameter> lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Integer,
        Value = FuncId
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = ModuleId.Length,
        Value = ModuleId
      };
      lstPrm.Add(prm);

      rezPrm = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.ReturnValue,
        OracleDbType = OracleDbType.Integer
      };
      lstPrm.Add(rezPrm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
      rezVal = Convert.ToInt32(rezPrm.Value);

      return (rezVal != 0);
    }

    public static Boolean GetPermissionForModuleUif2(int FuncId, string ModuleId)
    {
      int rezVal = 0;
      const string stmtSql = "LIMS.AcsUserModuleFunc2";
      OracleParameter prm = null;
      OracleParameter rezPrm = null;

      List<OracleParameter> lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Integer,
        Value = FuncId
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = ModuleId.Length,
        Value = ModuleId
      };
      lstPrm.Add(prm);

      rezPrm = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.ReturnValue,
        OracleDbType = OracleDbType.Integer
      };
      lstPrm.Add(rezPrm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
      rezVal = Convert.ToInt32(rezPrm.Value);

      return (rezVal != 0);
    }
  }

  public static class ModuleInfo
  {
    public static string GetModuleDescription(string ModuleId)
    {
      const string stmt = "SELECT DESCR FROM LIMS.V_MODULES WHERE (ID = :MID)";
      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "MID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Value = ModuleId
      };

      if (!String.IsNullOrEmpty(ModuleId))
        prm.Size = ModuleId.Length;
     
      lstPrm.Add(prm);
      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }
  }

  public static class DbVar
  {
    public static void SetRangeDate(DateTime DtBegin, DateTime DtEnd, int TypeDayCorr)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetDate";
      OracleParameter prm = null;
      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.DateTime,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Date,
                Value = DtBegin
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.DateTime,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Date,
                Value = DtEnd
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.Int32,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Integer,
                Value = TypeDayCorr
              };
      lstPrm.Add(prm);
      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr1";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
                Value = str1
              };
      lstPrm.Add(prm);
      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1, string str2)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr2";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
                Value = str1
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str2) ? 0 : str2.Length,
                Value = str2
              };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1, string str2, string str3)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr3";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
                Value = str1
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str2) ? 0 : str2.Length,
                Value = str2
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str3) ? 0 : str3.Length,
                Value = str3
              };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1, string str2, string str3, string str4)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr4";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
        Value = str1
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str2) ? 0 : str2.Length,
        Value = str2
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str3) ? 0 : str3.Length,
        Value = str3
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str4) ? 0 : str4.Length,
        Value = str4
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1, string str2, string str3, string str4, string str5)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr5";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
        Value = str1
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str2) ? 0 : str2.Length,
        Value = str2
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str3) ? 0 : str3.Length,
        Value = str3
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str4) ? 0 : str4.Length,
        Value = str4
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str5) ? 0 : str5.Length,
        Value = str5
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1, string str2, string str3, string str4, string str5, string str6, string str7)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr7";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
                Value = str1
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str2) ? 0 : str2.Length,
                Value = str2
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = string.IsNullOrEmpty(str3) ? 0 : str3.Length,
                Value = str3
              };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str4) ? 0 : str4.Length,
        Value = str4
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str5) ? 0 : str5.Length,
        Value = str5
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str6) ? 0 : str6.Length,
        Value = str6
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str7) ? 0 : str7.Length,
        Value = str7
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetString(string str1, string str2, string str3, string str4, string str5, string str6, string str7, string str8, string str9, string str10)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetStr10";
      OracleParameter prm = null;

      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str1) ? 0 : str1.Length,
        Value = str1
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str2) ? 0 : str2.Length,
        Value = str2
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str3) ? 0 : str3.Length,
        Value = str3
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str4) ? 0 : str4.Length,
        Value = str4
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str5) ? 0 : str5.Length,
        Value = str5
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str6) ? 0 : str6.Length,
        Value = str6
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str7) ? 0 : str7.Length,
        Value = str7
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str8) ? 0 : str8.Length,
        Value = str8
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str9) ? 0 : str9.Length,
        Value = str9
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = string.IsNullOrEmpty(str10) ? 0 : str10.Length,
        Value = str10
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetStringList(string str1, string Delim)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetListStr1";
      OracleParameter prm = null;
      var lstPrm = new List<OracleParameter>();

      string[] dlmString;
      string[] stringSeparators = {Delim};
      dlmString = str1.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        ArrayLength = dlmString.Length,
        Value = dlmString
      };
      lstPrm.Add(prm);
      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void SetNum(decimal num1)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetNum1";
      OracleParameter prm = null;
      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.Decimal,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Number,
                Precision = 17,
                Scale = 4,
                Value = num1
              };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm); 
   }

    public static void SetNum(decimal num1, decimal num2)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetNum2";
      OracleParameter prm = null;
      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Precision = 17,
        Scale = 4,
        Value = num1
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Precision = 17,
        Scale = 4,
        Value = num2
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }



    public static void SetNum(decimal num1, decimal num2, decimal num3, decimal num4, decimal num5, decimal num6, decimal num7)
    {
      const string stmtSql = "VIZ_PRN.VAR_RPT.SetNum7";
      OracleParameter prm = null;
      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num1
               };
      lstPrm.Add(prm);

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num2
               };
      lstPrm.Add(prm);

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num3
               };
      lstPrm.Add(prm);

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num4
               };
      lstPrm.Add(prm);

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num5
               };
      lstPrm.Add(prm);

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num6
               };
      lstPrm.Add(prm);

      prm = new OracleParameter
               {
                 DbType = DbType.Decimal,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Number,
                 Precision = 17,
                 Scale = 4,
                 Value = num7
               };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm); 
    }

    public static DateTime? GetDateBeginEnd(Boolean IsDateBegin, Boolean IsCorrect)
    {
      DateTime? rezVal = null;
      String stmtSql = IsDateBegin ? "VIZ_PRN.VAR_RPT.GetDateBegin" : "VIZ_PRN.VAR_RPT.GetDateEnd";
      OracleParameter rezPrm = null;
      var lstPrm = new List<OracleParameter>();

      rezPrm = new OracleParameter
                 {
                   DbType = DbType.Int32,
                   Direction = ParameterDirection.Input,
                   OracleDbType = OracleDbType.Integer,
                   Value = IsCorrect ? 1 : 0
                 };
      lstPrm.Add(rezPrm);

      rezPrm = new OracleParameter
                 {
                   DbType = DbType.DateTime,
                   Direction = ParameterDirection.ReturnValue,
                   OracleDbType = OracleDbType.Date
                 };
      lstPrm.Add(rezPrm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
      rezVal = Convert.ToDateTime(rezPrm.Value);
      return rezVal;
    }


  }

  public static class DbSelector
  {
    private static string GetDbAlias4Report(int? idReport)
    {
      var stmtSql = $"SELECT UPPER(DBALIAS) FROM LIMS.MODULES_FUNC WHERE ID = {Convert.ToString(idReport)}";
      var res = Odac.ExecuteScalar(stmtSql, CommandType.Text, false, null);

      return res != null ? Convert.ToString(res) : null;
    }

    public static void ConnectToTargetDb(int? idReport, string dbAlias = null)
    {
      if ((idReport == null) && string.IsNullOrEmpty(dbAlias))
        return;
      

      if ((idReport != null) && string.IsNullOrEmpty(dbAlias)){

        string rptAlias = GetDbAlias4Report(idReport);
        Odac.DbConnection.Close();
        Odac.DbConnection.Server = rptAlias.ToUpper(CultureInfo.InvariantCulture);
        //MessageBox.Show(rptAlias.ToUpper(CultureInfo.InvariantCulture));
        Odac.DbConnection.Open();
        Odac.DbConnection.Close();
        return;
      }

      if ((idReport == null) && !string.IsNullOrEmpty(dbAlias)){

        Odac.DbConnection.Close();
        Odac.DbConnection.Server = dbAlias.ToUpper(CultureInfo.InvariantCulture);
        //MessageBox.Show(dbAlias.ToUpper(CultureInfo.InvariantCulture));
        Odac.DbConnection.Open();
        Odac.DbConnection.Close();
      }
    }

    public static string GetCurrentDbAlias()
    {
      return Odac.GetDbAlias();
    }


  }

}
