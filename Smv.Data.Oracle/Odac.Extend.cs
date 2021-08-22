using System;
using System.Data;
using System.Collections.Generic;
using Devart.Data.Oracle;

namespace Smv.Data.Oracle
{
  partial class Odac
  {
    public static bool Connect(Boolean Pooling)
    {
      return _Odac.OvrConnect(Pooling);
    }

    public static void Disconnect(Boolean IsDispose)
    {
      _Odac.OvrDisconnect(IsDispose);
    }

    public static string GetServerVersion()
    {
      return _Odac.OvrGetServerVersion();
    }

    public static string GetClientVersion()
    {
      return _Odac.OvrGetClientVersion();
    }
    public static string GetDbAlias()
    {
      return _Odac.OvrGetDbAlias();
    }

    public static int LoadDataTable(OracleDataTable table, Boolean ClearBeforeLoad, List<Object> lstParam)
    {
      return _Odac.OvrLoadDataTable(table, ClearBeforeLoad, lstParam);
    }
   
    public static int LoadDataTable(DataTable table, OracleDataAdapter adp, Boolean ClearBeforeLoad, List<Object> lstParamValue)
    {
      return _Odac.OvrLoadDataTable(table, adp, ClearBeforeLoad, lstParamValue);
    }
 
    public static int SaveChangedData(OracleDataTable dt)
    {
      return _Odac.OvrSaveChangedData(dt);
    }
   
    public static int SaveChangedData(DataTable dt, OracleDataAdapter adp)
    {
      return _Odac.OvrSaveChangedData(dt, adp);
    } 

    public static bool ExecuteNonQuery(String SqlStmt, CommandType cmdType, Boolean ParamCheck, List<OracleParameter> lstParam, Boolean IsPassParametersByName = false)
    {
      return _Odac.OvrExecuteNonQuery(SqlStmt, cmdType, ParamCheck, lstParam, IsPassParametersByName);
    }

    public static IAsyncResult ExecuteNonQueryAsync(String SqlStmt, CommandType cmdType, Boolean ParamCheck, Boolean IsPassParametersByName, List<OracleParameter> lstParam)
    {
      return _Odac.OvrExecuteNonQueryAsync(SqlStmt, cmdType, ParamCheck, IsPassParametersByName, lstParam);
    }  

    public static OracleDataReader GetOracleReader(String SqlStmt, CommandType cmdType, Boolean ParamCheck, List<OracleParameter> lstParam, OdacErrorInfo ef)
    {
      return _Odac.OvrGetOracleReader(SqlStmt, cmdType, ParamCheck, lstParam, ef);
    }

    public static IAsyncResult GetOracleReaderAsync(String SqlStmt, CommandType cmdType, Boolean ParamCheck, IEnumerable<OracleParameter> lstParam, OdacErrorInfo ef)
    {
      return _Odac.OvrGetOracleReaderAsync(SqlStmt, cmdType, ParamCheck, lstParam, ef);
    }

    public static Object ExecuteScalar(String SqlStmt, CommandType cmdType, Boolean ParamCheck, List<OracleParameter> lstParam)
    {
      return _Odac.OvrExecuteScalar(SqlStmt, cmdType, ParamCheck, lstParam); 
    }
     
    /*
    public static Object ExecuteScalar(String SqlStmt, CommandType cmdType, TrnType trnType, List<FbParameter> lstParam)
    {
      return _Fdac.OvrExecuteScalar(SqlStmt, cmdType, trnType, lstParam);
    }
    */
  }
}
