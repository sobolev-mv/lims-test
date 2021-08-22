using System;
using System.Text;
using System.Data;
using System.Collections.Generic;
using System.Data.Common;
using System.Windows;
using Devart.Data.Oracle;

namespace Smv.Data.Oracle
{

  public interface IOdacUtils
  {
    void ShowErrorInfo(string errorsTitle, string errorsMsg);
    Boolean GetLogonInfo(ref string Login, ref string DbName, ref string Pass, ref Boolean isUnicode);
    string LoginName{get; set;}
    string DbName{get; set;}
    Boolean IsUnicode { get; set; }
  }
  
  public sealed class OdacErrorInfo
  {
    public string ErrorMsg {get; set;}
  }

  public enum SqlStmtType
  {
    NativeSql = 1,
    MarsSql = 2
  }

  public enum CmdType
  {
    CmdSelect = 1,
    CmdInsert = 2,
    CmdUpdate = 3,
    CmdDelete = 4
  }

  public enum TrnType
  {
    TrnRead = 1,
    TrnWrite = 2
  }

  //Singleton pattern
  public sealed partial class Odac
  {
    //private EntityConnection _EntityConnection = null;
        
    private OracleConnection _DbConnection = null;
    private IOdacUtils iutls;

    private static Odac _Odac = null; //Singleton
    private static Boolean _isShowError = true;


    public static Boolean IsShowError
    {
      get {return _isShowError;}
    }

    private static Boolean _isShowSqlError = true;
    public static Boolean IsShowSqlError
    {
      get {return _isShowSqlError;}
      set {_isShowSqlError = value;}
    }
    
    public static OracleConnection DbConnection
    {
      get { return _Odac._DbConnection; }
    }
    
    private Odac(IOdacUtils iutls)
    {
      this.iutls = iutls;    
    }


    public static void Init(IOdacUtils iutls)
    {
      if (_Odac == null) _Odac = new Odac(iutls);
    }
        
    public static Boolean IsDataTableHasChanged(DataTable dtTable)
    {
      return ((dtTable.Select(null, null, DataViewRowState.Added).Length != 0) ||
              (dtTable.Select(null, null, DataViewRowState.Deleted).Length != 0) ||
              (dtTable.Select(null, null, DataViewRowState.ModifiedOriginal).Length != 0)
             ); 
    }

    private void ShowSqlError(DbException ex)
    {
      if (_isShowSqlError) MessageBox.Show(ex.Message, "Oracle Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void ShowSqlError(OracleException ex)
    {
      if (_isShowSqlError) MessageBox.Show(ex.Message, "Oracle Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void ShowSqlError(Exception ex)
    {
      System.String errMsg = ex.Message;
      while (ex.InnerException != null){
        ex = ex.InnerException;
        errMsg = errMsg + "\r\n" + ex.Message;
      }
      if (_isShowSqlError) MessageBox.Show(errMsg, "Oracle Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void ShowError(Exception ex)
    {
      String errMsg = ex.Message + "\r\n" + ex.Source;
      if (_isShowError) MessageBox.Show(errMsg, "Program Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }
        
    private bool OvrConnect(Boolean Pooling)
    {
      Boolean Rez = false;
      string strLoginName = this.iutls.LoginName;
      string strDbName = this.iutls.DbName;
      string strPass = "";

      Boolean isUnicode = string.IsNullOrEmpty(strLoginName) && string.IsNullOrEmpty(strDbName) || this.iutls.IsUnicode;

      Boolean wndDialogResult = this.iutls.GetLogonInfo(ref strLoginName, ref strDbName, ref strPass, ref isUnicode);
      if (!wndDialogResult)
        return false;  

      //Build the FireBirdConnection connection string.
      var connStrBuilder = new OracleConnectionStringBuilder
                             {
                               Pooling = Pooling,
                               UserId = strLoginName,
                               Password = strPass,
                               Server = strDbName,
                               Direct = false,
                               Unicode = isUnicode
      };

      try{
        if (_Odac._DbConnection != null) _Odac._DbConnection.Dispose();
        _Odac._DbConnection = new OracleConnection(connStrBuilder.ConnectionString); 
        _Odac._DbConnection.Open();

        //if (cWin.tbLogin.Text != this.iutls.LoginName)
        this.iutls.LoginName = strLoginName;
        
        //if (cWin.tbBase.Text != this.iutls.DbName)
        this.iutls.DbName = strDbName;

        this.iutls.IsUnicode = isUnicode;

        Rez = true;
      } 
      catch (OracleException ex){
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Oracle Error", ex.Message);

        Rez = false;
      }
      catch (Exception ex){
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
        Rez = false;
      }
      finally{
        if (_Odac._DbConnection != null) _Odac._DbConnection.Close();
      }
      return Rez;
    }
    
            
    private void OvrDisconnect(Boolean IsDispose)
    {
      if (_Odac._DbConnection != null){
        if (_Odac._DbConnection.State == ConnectionState.Open) _Odac._DbConnection.Close(); 
        if (IsDispose) _Odac._DbConnection.Dispose();
      }
    }

    private string OvrGetServerVersion()
    {
      string ver = null;
      
      if (_Odac._DbConnection != null){
        if (_Odac._DbConnection.State == ConnectionState.Closed){
          _Odac._DbConnection.Open();
          ver = _Odac._DbConnection.ServerVersion; 
          _Odac._DbConnection.Close();
        }
        else
          ver = _Odac._DbConnection.ServerVersion; 
      }
      return ver;
    }

    private string OvrGetClientVersion()
    {
      string ver = null;

      if (_Odac._DbConnection != null){
        if (_Odac._DbConnection.State == ConnectionState.Closed){
          _Odac._DbConnection.Open();
          ver = _Odac._DbConnection.ClientVersion;
          _Odac._DbConnection.Close();
        }
        else
          ver = _Odac._DbConnection.ClientVersion;
      }
      return ver;
    }

    private string OvrGetDbAlias()
    {
      string ver = null;

      if (_Odac._DbConnection != null){
        if (_Odac._DbConnection.State == ConnectionState.Closed){
          _Odac._DbConnection.Open();
          ver = _Odac._DbConnection.Server;
          _Odac._DbConnection.Close();
        }
        else
          ver = _Odac._DbConnection.Server;
      }
      return ver;
    }

    private int OvrLoadDataTable(OracleDataTable table, Boolean ClearBeforeLoad, List<Object> lstParamValue)
    {
      int rw = 0;

      try{
        if (ClearBeforeLoad) 
          table.Clear();

        if (lstParamValue != null)
          for (int i = 0; i < lstParamValue.Count; i++)
            table.SelectCommand.Parameters[i].Value = lstParamValue[i];

        rw = table.Fill();
      }
      catch (Exception ex){
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }
      return rw;
    }

    private int OvrLoadDataTable(DataTable table, OracleDataAdapter adp, Boolean ClearBeforeLoad, List<Object> lstParamValue)
    {
      int rw = 0;

      try
      {
        if (ClearBeforeLoad)
          table.Clear();

        if (lstParamValue != null)
          for (int i = 0; i < lstParamValue.Count; i++)
            adp.SelectCommand.Parameters[i].Value = lstParamValue[i];

        rw = adp.Fill(table);
      }
      catch (Exception ex)
      {
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }
      return rw;
    }

    
    private int OvrSaveChangedData(OracleDataTable dt)
    {
      int rw = 0;
      OracleTransaction trx = null;

      if (!IsDataTableHasChanged(dt)) return 0;

      try{
        foreach (DataRow errRow in dt.GetErrors()) errRow.ClearErrors();

        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        trx = _Odac._DbConnection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        rw = dt.Update();
        trx.Commit();
      }
      catch (Exception ex){
        trx.Rollback();
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }
      finally{
        if (trx != null) trx.Dispose();
        _Odac._DbConnection.Close();
      }

      foreach (DataRow errRow in dt.GetErrors())
        if (errRow.RowState == DataRowState.Deleted) errRow.RejectChanges();

      //DataRow[] errRows = dt.GetErrors();
      //if (errRows.Length > 0)
      //MessageBox.Show("Во время сохранения возникли ошибки!\r\n" + errRows[0].RowError, "Fb Server Error", MessageBoxButton.OK, MessageBoxImage.Error);

      return rw;
    }

    private int OvrSaveChangedData(DataTable dt, OracleDataAdapter adp)
    {
      int rw = 0;
      OracleTransaction trx = null;

      if (!IsDataTableHasChanged(dt)) return 0;

      try
      {
        foreach (DataRow errRow in dt.GetErrors()) errRow.ClearErrors();

        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        trx = _Odac._DbConnection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        rw = adp.Update(dt);
        trx.Commit();
      }
      catch (Exception ex)
      {
        trx.Rollback();
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }
      finally
      {
        if (trx != null) trx.Dispose();
        _Odac._DbConnection.Close();
      }

      foreach (DataRow errRow in dt.GetErrors())
        if (errRow.RowState == DataRowState.Deleted) errRow.RejectChanges();

      //DataRow[] errRows = dt.GetErrors();
      //if (errRows.Length > 0)
      //MessageBox.Show("Во время сохранения возникли ошибки!\r\n" + errRows[0].RowError, "Fb Server Error", MessageBoxButton.OK, MessageBoxImage.Error);

      return rw;
    }

    private bool OvrExecuteNonQuery(String SqlStmt, CommandType cmdType, Boolean ParamCheck,  List<OracleParameter> lstParam, Boolean IsPassParametersByName = false)
    {
      bool rez = true;
      OracleTransaction trx = null;
      OracleCommand cmd = null; 

      try{

        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        cmd = _Odac._DbConnection.CreateCommand();
        cmd.Connection = _Odac._DbConnection; 
        cmd.ParameterCheck = ParamCheck;
        cmd.PassParametersByName = IsPassParametersByName;
        cmd.CommandType = cmdType;
        cmd.CommandText = SqlStmt;

        //ListFbParameters!!!!
        if (lstParam != null)
          foreach (OracleParameter prm in lstParam) cmd.Parameters.Add(prm);

        trx = _Odac._DbConnection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        cmd.ExecuteNonQuery();
        trx.Commit(); 
      }
      catch (Exception ex){
        if (trx != null) 
          trx.Rollback();
        rez = false;
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }
      finally{
         if (trx != null) trx.Dispose();
         if (cmd != null) cmd.Dispose();
        _Odac._DbConnection.Close();
      }
      return rez;
    }


    private IAsyncResult OvrExecuteNonQueryAsync(String SqlStmt, CommandType cmdType, Boolean ParamCheck, Boolean IsPassParametersByName, List<OracleParameter> lstParam)
    {
      IAsyncResult rez = null;       

      try{
        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        OracleCommand cmd = _Odac._DbConnection.CreateCommand();
        cmd.ParameterCheck = ParamCheck;
        cmd.PassParametersByName = IsPassParametersByName;
        cmd.CommandType = cmdType;
        cmd.CommandText = SqlStmt;

        //ListFbParameters!!!!
        if (lstParam != null)
          foreach (OracleParameter prm in lstParam) cmd.Parameters.Add(prm);

        rez = cmd.BeginExecuteNonQuery(null, cmd);
      }
      catch (Exception ex){
        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }

      return rez;
    }


    private OracleDataReader OvrGetOracleReader(String SqlStmt, CommandType cmdType, Boolean ParamCheck, IEnumerable<OracleParameter> lstParam, OdacErrorInfo ef)
    {
     
      OracleCommand cmd = null;

      try{
        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        cmd = _Odac._DbConnection.CreateCommand();
        cmd.PassParametersByName = true;
        cmd.ParameterCheck = ParamCheck;
        cmd.CommandType = cmdType;
        cmd.CommandText = SqlStmt;

        //ListFbParameters!!!!
        if (lstParam != null)
          foreach (OracleParameter prm in lstParam) cmd.Parameters.Add(prm);

        return cmd.ExecuteReader(CommandBehavior.CloseConnection);
      }
      catch (Exception ex){
        if (ef == null){
          if (this.iutls == null)
            ShowSqlError(ex);
          else
            this.iutls.ShowErrorInfo("Odac Error", ex.Message);
        }
        else
          ef.ErrorMsg = ex.Message;
          
        return null;
      }
    }

    private IAsyncResult OvrGetOracleReaderAsync(String SqlStmt, CommandType cmdType, Boolean ParamCheck, IEnumerable<OracleParameter> lstParam, OdacErrorInfo ef)
    {
      try{
        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        var cmd = _Odac._DbConnection.CreateCommand();
        cmd.PassParametersByName = true;
        cmd.ParameterCheck = ParamCheck;
        cmd.CommandType = cmdType;
        cmd.CommandText = SqlStmt;

        //ListFbParameters!!!!
        if (lstParam != null)
          foreach (OracleParameter prm in lstParam) cmd.Parameters.Add(prm);

        return cmd.BeginExecuteReader(null, cmd, CommandBehavior.CloseConnection);
      }
      catch (Exception ex){
        if (ef == null){
          if (this.iutls == null)
            ShowSqlError(ex);
          else
            this.iutls.ShowErrorInfo("Odac Error", ex.Message);
        }
        else
          ef.ErrorMsg = ex.Message;

        return null;
      }
    }


    private Object OvrExecuteScalar(String SqlStmt, CommandType cmdType, Boolean ParamCheck, List<OracleParameter> lstParam)
    {
      Object rez = null;
      OracleTransaction trx = null;
      OracleCommand cmd = null;

      try{

        if (_Odac._DbConnection.State == ConnectionState.Closed) _Odac._DbConnection.Open();
        cmd = _Odac._DbConnection.CreateCommand();
        cmd.ParameterCheck = ParamCheck;
        cmd.PassParametersByName = false;
        cmd.CommandType = cmdType;
        cmd.CommandText = SqlStmt;

        //ListFbParameters!!!!
        if (lstParam != null)
          foreach (OracleParameter prm in lstParam) cmd.Parameters.Add(prm);

        trx = _Odac._DbConnection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        rez = cmd.ExecuteScalar();
        trx.Commit();
      }
      catch (Exception ex){
        trx.Rollback();

        if (this.iutls == null)
          ShowSqlError(ex);
        else
          this.iutls.ShowErrorInfo("Odac Error", ex.Message);
      }
      finally{
        if (trx != null) trx.Dispose();
        if (cmd != null) cmd.Dispose();
        _Odac._DbConnection.Close();
      }
      return rez;
    }

       
    /*
    private Object OvrExecuteScalar(String SqlStmt, CommandType cmdType, TrnType trnType, List<FbParameter> lstParam)
    {
      Object rez = null;
      FbCommand cmd = _Fdac._DbConnection.CreateCommand();
      cmd.CommandType = cmdType;
      cmd.CommandText = SqlStmt;

      if (trx != null){
        _Fdac.OvrCommitTrn(trx);
        trx.Dispose();
        trx = null;
      }

      try{
        
        //ListFbParameters!!!!
        if (lstParam != null)
          foreach (FbParameter prm in lstParam) cmd.Parameters.Add(prm);

        if (trnType == TrnType.TrnRead)
          trx = _Fdac.OvrStartReadTrn();
        else
          trx = _Fdac.OvrStartWriteTrn();

        cmd.Transaction = trx;
        rez = cmd.ExecuteScalar();
        _Fdac.OvrCommitTrn(trx);
      }
      catch (Exception ex){
        _Fdac.OvrRollbackTrn(trx);
        MessageBox.Show(ex.Message, "Fb Server Error", MessageBoxButton.OK, MessageBoxImage.Error);
      }

      if (trx != null){
        trx.Dispose();
        trx = null;
      }

      return rez;
    }


    */



  }
}
