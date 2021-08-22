using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using Devart.Data.Oracle;
using Microsoft.Win32;
using Smv.App.Config;
using Smv.Data.Oracle;
using Smv.Utils;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.MagLab.Db
{
  public enum SampleState
  {
    Deleted = -1,
    Edited = 0,
    InMes = 10,
    Closed = 20
  }

  public static class LabAction
  {
    public static Boolean MatToMes(string SampleId, string Tstep)
    {
      String stmtSql = "LIMS.MagLab.MatToMes";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      OracleParameter prm = new OracleParameter();
      prm.DbType = DbType.String;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 64;
      prm.Value = SampleId;
      lstPrm.Add(prm);
      /* 
      prm = new OracleParameter();
      prm.DbType = System.Data.DbType.String;
      prm.Direction = System.Data.ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 20;
      prm.Value = Tstep;
      lstPrm.Add(prm); 
      */
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean SampleToMes(string SampleId, string Tstep)
    {
      String stmtSql = "LIMS.MagLab.SampleToMes";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      OracleParameter prm = new OracleParameter();
      prm.DbType = DbType.String;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 64;
      prm.Value = SampleId;
      lstPrm.Add(prm);
       
      prm = new OracleParameter();
      prm.DbType = DbType.String;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 20;
      prm.Value = Tstep;
      lstPrm.Add(prm); 
      
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }


    public static Boolean SampleToEdit(string SampleId)
    {
      String stmtSql = "LIMS.MagLab.SampleToEdit";

      OracleParameter prm = new OracleParameter();
      prm.DbType = DbType.String;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 64;
      prm.Value = SampleId;

      List<OracleParameter> lstPrm = new List<OracleParameter>();
      lstPrm.Add(prm);
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static List<Boolean> GetVisibleElements(string SampleId, int Utp, int IsSmp, uint ListSize)
    {
      String stmtSql = "LIMS.MagLab.GetVisibleEditElement";
      OracleParameter prm = null;
      OracleParameter prmTable = null;
      OracleParameter prmCount = null;
      List<Boolean> rezList = null;

      List<OracleParameter> lstPrm = new List<OracleParameter>();

      prm = new OracleParameter();
      prm.DbType = DbType.String;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 64;
      prm.Value = SampleId;
      lstPrm.Add(prm);

      prm = new OracleParameter();
      prm.DbType = DbType.Int32;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.Integer;
      prm.Value = Utp;
      lstPrm.Add(prm);

      prm = new OracleParameter();
      prm.DbType = DbType.Int32;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.Integer;
      prm.Value = IsSmp;
      lstPrm.Add(prm);

      prmCount = new OracleParameter();
      prmCount.DbType = DbType.Int32;
      prmCount.Direction = ParameterDirection.Output;
      prmCount.OracleDbType = OracleDbType.Integer;
      lstPrm.Add(prmCount);

      prmTable = new OracleParameter();
      prmTable.DbType = DbType.Int32;
      prmTable.Direction = ParameterDirection.ReturnValue;
      prmTable.OracleDbType = OracleDbType.Integer;
      prmTable.ArrayLength = (int)ListSize;
      lstPrm.Add(prmTable);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
      int LstCount = Convert.ToInt32(prmCount.Value);

      if (LstCount == 0) return null;
      rezList = new List<bool>();
      for (int i = 0; i <= LstCount - 1; i++) rezList.Add((Convert.ToInt32(prmTable[i]) != 0));
      return rezList;
    }

    public static void FillSampleInfo(string SampleId, DataTable dt)
    {
      DbVar.SetString(SampleId);

      OracleDataAdapter adapter = new OracleDataAdapter(); 
      adapter.SelectCommand = new OracleCommand();
      adapter.SelectCommand.Connection = Odac.DbConnection;
      //String SqlStmt = System.IO.File.ReadAllText(Smv.Utils.Etc.StartPath + "\\Sql\\New1.sql", System.Text.Encoding.GetEncoding(1251)).Replace("\r", " ");
      String SqlStmt = "SELECT * FROM VIZ_PRN.VIEW_ML_PROP_MAT1";
      adapter.SelectCommand.CommandType = CommandType.Text;
      adapter.SelectCommand.CommandText = SqlStmt; 
      Odac.LoadDataTable(dt,adapter,true,null);
    }

    public static void FillProbeInfo(string MatLocalId, string Usg,  string InspLotStatus, DataTable dt)
    {
      DbVar.SetString(InspLotStatus, Usg, MatLocalId);

      var adapter = new OracleDataAdapter();
      adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
      //String SqlStmt = System.IO.File.ReadAllText(Smv.Utils.Etc.StartPath + "\\Sql\\New2b.sql", System.Text.Encoding.GetEncoding(1251)).Replace("\r", " ");
      const string SqlStmt = "SELECT * FROM VIZ_PRN.VIEW_ML_PROP_MAT2";
      adapter.SelectCommand.CommandType = CommandType.Text;
      adapter.SelectCommand.CommandText = SqlStmt;
      /*
      List<Object> lstPrm = new List<Object>();
      OracleParameter SqlPrm = new OracleParameter();
      SqlPrm.DbType = System.Data.DbType.String;
      SqlPrm.Direction = System.Data.ParameterDirection.Input;
      SqlPrm.OracleDbType = OracleDbType.VarChar;
      SqlPrm.ParameterName = "MLOCID";
      SqlPrm.Size = 20;
      adapter.SelectCommand.Parameters.Add(SqlPrm);

      SqlPrm = new OracleParameter();
      SqlPrm.DbType = System.Data.DbType.String;
      SqlPrm.Direction = System.Data.ParameterDirection.Input;
      SqlPrm.OracleDbType = OracleDbType.VarChar;
      SqlPrm.ParameterName = "USG";
      SqlPrm.Size = 20;
      adapter.SelectCommand.Parameters.Add(SqlPrm);

      SqlPrm = new OracleParameter();
      SqlPrm.DbType = System.Data.DbType.String;
      SqlPrm.Direction = System.Data.ParameterDirection.Input;
      SqlPrm.OracleDbType = OracleDbType.VarChar;
      SqlPrm.ParameterName = "ILS";
      SqlPrm.Size = 20;
      adapter.SelectCommand.Parameters.Add(SqlPrm);

      lstPrm.Add(MatLocalId);
      lstPrm.Add(Usg);
      lstPrm.Add(InspLotStatus);
      */

      Odac.LoadDataTable(dt, adapter, true, null);
    }

    public static Boolean CopyPropToAnotherSample(string SampleId1, string SampleId2)
    {
      const string stmtSql = "LIMS.MagLab.CopyPropToAnoterSample";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
                  {
                    DbType = DbType.String,
                    Direction = ParameterDirection.Input,
                    OracleDbType = OracleDbType.VarChar,
                    Size = 60,
                    Value = SampleId1
                  };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = 60,
                Value = SampleId2
              };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean CopyApstPropToAnotherSample(string SampleId1, string SampleId2)
    {
      const string stmtSql = "LIMS.MagLab.CopyApstPropToAnoterSample";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = 60,
        Value = SampleId1
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = 60,
        Value = SampleId2
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static int GetCntSamples(string matLocNum, string techStep)
    {
      const string stmtSql = "select count(*) from LIMS.SAMPLEMEAS where MATLOCALNUMBER = :PLOCNUM and TSTEP = :PTSTEP and state >= 0";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        ParameterName = "PLOCNUM",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocNum.Length,
        Value = matLocNum
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        ParameterName = "PTSTEP",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = techStep.Length,
        Value = techStep
      };
      lstPrm.Add(prm);
      
      return Convert.ToInt32(Odac.ExecuteScalar(stmtSql,CommandType.Text, false, lstPrm));
    }

    public static Boolean CopyApstProp4CmpCoil(string sampleId)
    {
      const string stmtSql = "LIMS.MagLab.CopyApstProp4CmpCoil";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = sampleId.Length,
        Value = sampleId
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static List<int?> GetKesi(string SampleId)
    {
      const string stmtSql = "LIMS.MagLab.GetKESI";
      OracleParameter prm = null;
      OracleParameter prmTable = null;
      //OracleParameter prmCount = null;
      
      var lstPrm = new List<OracleParameter>();

      prm = new OracleParameter
              {
                DbType = DbType.String,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.VarChar,
                Size = 64,
                Value = SampleId
              };
      lstPrm.Add(prm);

      prmTable = new OracleParameter
                   {
                     DbType = DbType.Int32,
                     Direction = ParameterDirection.ReturnValue,
                     OracleDbType = OracleDbType.Integer,
                     ArrayLength = 6
                   };
      lstPrm.Add(prmTable);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
      var rezList = new List<int?>();

      for (int i = 0; i <= 5; i++){ 
        if (prmTable[i] != DBNull.Value)
          rezList.Add(Convert.ToInt32(prmTable[i]));
        else
          rezList.Add(null);
      }
      return rezList;
    }

    public static Boolean DeleteSample(string SampleId)
    {
      String stmtSql = "LIMS.MagLab.DeleteSample";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      OracleParameter prm = new OracleParameter();
      prm.DbType = DbType.String;
      prm.Direction = ParameterDirection.Input;
      prm.OracleDbType = OracleDbType.VarChar;
      prm.Size = 60;
      prm.Value = SampleId;
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean ChangeSimensSampleState(Int64 id, SampleState stateSample)
    {
      String stmtSql = "UPDATE LIMS.ML_SAMPLE_SIEMENS SET STATE = :PSTATE WHERE ID = :PID";
      //String stmtSql = "UPDATE LIMS.ML_SAMPLE_SIEMENS SET STATE = -1 WHERE ID = 22";
      var lstPrm = new List<OracleParameter>();
      
      var prm = new OracleParameter
      {
        ParameterName = "PSTATE",
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        Value = stateSample
      };
      lstPrm.Add(prm);
      

      prm = new OracleParameter
      {
        ParameterName = "PID",
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        Value = id
      };
      lstPrm.Add(prm);
      
      return Odac.ExecuteNonQuery(stmtSql, CommandType.Text, false, lstPrm);
    }

    public static int?  GetSimensSampleDpp1750(Int64 SampId)
    {
      String stmtUpdateSql = "LIMS.MagLab.Calc4SiemensDpp1750";
      String stmtSelectSql = "SELECT DPP1750 FROM LIMS.ML_SAMPLE_SIEMENS WHERE ID = :PID";

      var lstPrm1 = new List<OracleParameter>();
      var lstPrm2 = new List<OracleParameter>();

      var prm1 = new OracleParameter
      {
        ParameterName = "PID",
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        Value = SampId
      };

      var prm2 = new OracleParameter
      {
        ParameterName = "PID",
        DbType = DbType.Int32,
        OracleDbType = OracleDbType.Integer,
        Direction = ParameterDirection.Input,
        Value = SampId
      };

      lstPrm1.Add(prm1);
      lstPrm2.Add(prm2);

      Odac.ExecuteNonQuery(stmtUpdateSql, CommandType.StoredProcedure, false, lstPrm1);
      var oTmp = Odac.ExecuteScalar(stmtSelectSql, CommandType.Text, false, lstPrm2);

      if (oTmp == Convert.DBNull)
        return null;

      return Convert.ToInt32(oTmp);
    }



    public static void SaveSimensRpt(DateTime dtBegin, DateTime dtEnd)
    {
      var sfd = new SaveFileDialog
      {
        OverwritePrompt = false,
        AddExtension = true,
        DefaultExt = ".csv",
        Filter = "csv file (.csv)|*.csv"
      };
      
      if (sfd.ShowDialog().GetValueOrDefault() != true)
        return;

      if (File.Exists(sfd.FileName)){
        DxInfo.ShowDxBoxInfo("Файл", "Файл: " + sfd.FileName + " уже существует!", MessageBoxImage.Error);
        return;
      }


      OracleDataReader odr = null;
      DbVar.SetRangeDate(dtBegin, dtEnd, 1);
      DbVar.SetString("%");

      const string sqlStmt1 = "SELECT * FROM LIMS.V_SIEMENS_RPT ORDER BY 1";
      odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, "Все рулоны", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName,"Толщина, мм;Количество измерений (Nизм);Количество измерений с ?Р1,7/50 ?5% (N?Р1,7/50 ?5%);?5%Р1,7/50, %",Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();

      DbVar.SetString("%/1");
      odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "1-ые рулоны", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "Толщина, мм;Количество измерений (Nизм);Количество измерений с ?Р1,7/50 ?5% (N?Р1,7/50 ?5%);?5%Р1,7/50, %", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();

      DbVar.SetString("%/2");
      odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "2-ые рулоны", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "Толщина, мм;Количество измерений (Nизм);Количество измерений с ?Р1,7/50 ?5% (N?Р1,7/50 ?5%);?5%Р1,7/50, %", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();
      /***********************************************************************************/
      const string sqlStmt2 = "SELECT * FROM LIMS.V_SIEMENS_STDDEV_P1750_THICK";
      odr = Odac.GetOracleReader(sqlStmt2, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "Толщина, мм;СКО", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();

      /***********************************************************************************/
      const string sqlStmt3 = "SELECT * FROM LIMS.V_SIEMENS_STDDEV_P1750_STEND";
      odr = Odac.GetOracleReader(sqlStmt3, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "Ст. партия;СКО", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();

      DxInfo.ShowDxBoxInfo("Файл", "Файл: " + sfd.FileName + " сформирован!", MessageBoxImage.Information);
    }

    public static Boolean ChangeValidationMode(string SampleId, int IsValidate)
    {
      String stmtSql = "LIMS.MagLab.SampleValidate";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
                  {
                    DbType = DbType.String,
                    Direction = ParameterDirection.Input,
                    OracleDbType = OracleDbType.VarChar,
                    Size = 60,
                    Value = SampleId
                  };
      lstPrm.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.Int32,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Integer,
                Value = IsValidate
              };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean ChangeStatFlag(string MatLocalId)
    {
      const string stmtSql = "LIMS.MagLab.EditStatFlag";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(MatLocalId))
        len = MatLocalId.Length;

      var prm = new OracleParameter
                  {
                    DbType = DbType.String,
                    Direction = ParameterDirection.Input,
                    OracleDbType = OracleDbType.VarChar,
                    Size = len,
                    Value = MatLocalId
                  };
      lstPrm.Add(prm);
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean CheckS2L(string MatLocalId)
    {
      const string stmtSql = "LIMS.ML_S2L.CheckS2LMaterial";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(MatLocalId))
        len = MatLocalId.Length;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = len,
        Value = MatLocalId
      };
      lstPrm.Add(prm);
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean UnCheckS2L(string MatLocalId)
    {
      const string stmtSql = "LIMS.ML_S2L.UnCheckS2LMaterial";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(MatLocalId))
        len = MatLocalId.Length;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = len,
        Value = MatLocalId
      };
      lstPrm.Add(prm);
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean CopyPropS2L(string MatLocalId)
    {
      const string stmtSql = "LIMS.ML_S2L.CopyPropS2LMaterial";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(MatLocalId))
        len = MatLocalId.Length;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = len,
        Value = MatLocalId
      };
      lstPrm.Add(prm);
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void Fr20DataImport(string CfgFile)
    {
      string connStr = ConfigParam.ReadConnectionStringParamValue(CfgFile, "ConStrMeasureBase");
      SqlConnection sqlCon = new SqlConnection(connStr);
      SqlDataAdapter sqlAdapter = new SqlDataAdapter();
      sqlAdapter.SelectCommand = new SqlCommand { Connection = sqlCon };
      sqlAdapter.UpdateCommand = new SqlCommand { Connection = sqlCon };
      OracleLoader loader = new OracleLoader();

      //создаем таблицу для импорта измеренных токов
      DataTable dtImport = new DataTable("Import");
      DataColumn col = null;

      col = new DataColumn("Id", typeof(Int64), null, MappingType.Element) { AllowDBNull = false };
      dtImport.Columns.Add(col);

      col = new DataColumn("DataSet", typeof(string), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("Samplename", typeof(string), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("Operator", typeof(string), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("Charge", typeof(string), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("Type", typeof(string), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("Condition", typeof(string), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("SideFlag", typeof(Int32), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("TotalCurrent", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("TotalCoefficent", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode1", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode2", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode3", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode4", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode5", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode6", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode7", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode8", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode9", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      col = new DataColumn("I_Electrode10", typeof(decimal), null, MappingType.Element);
      dtImport.Columns.Add(col);

      dtImport.Constraints.Add(new UniqueConstraint("Pk_Import", new[] { dtImport.Columns["Id"] }, true));
      dtImport.Columns["Id"].Unique = true;

      sqlAdapter.TableMappings.Clear();
      var dtm = new DataTableMapping("FrResults", "Import");
      dtm.ColumnMappings.Add("ID", "Id");
      dtm.ColumnMappings.Add("DataSet", "DataSet");
      dtm.ColumnMappings.Add("Samplename", "Samplename");
      dtm.ColumnMappings.Add("Operator", "Operator");
      dtm.ColumnMappings.Add("Charge", "Charge");
      dtm.ColumnMappings.Add("Type", "Type");
      dtm.ColumnMappings.Add("Condition", "Condition");
      dtm.ColumnMappings.Add("SideFlag", "SideFlag");
      dtm.ColumnMappings.Add("TotalCurrent", "TotalCurrent");
      dtm.ColumnMappings.Add("TotalCoefficent", "TotalCoefficent");
      dtm.ColumnMappings.Add("I_Electrode1", "I_Electrode1");
      dtm.ColumnMappings.Add("I_Electrode2", "I_Electrode2");
      dtm.ColumnMappings.Add("I_Electrode3", "I_Electrode3");
      dtm.ColumnMappings.Add("I_Electrode4", "I_Electrode4");
      dtm.ColumnMappings.Add("I_Electrode5", "I_Electrode5");
      dtm.ColumnMappings.Add("I_Electrode6", "I_Electrode6");
      dtm.ColumnMappings.Add("I_Electrode7", "I_Electrode7");
      dtm.ColumnMappings.Add("I_Electrode8", "I_Electrode8");
      dtm.ColumnMappings.Add("I_Electrode9", "I_Electrode9");
      dtm.ColumnMappings.Add("I_Electrode10", "I_Electrode10");
      sqlAdapter.TableMappings.Add(dtm);

      try{
        sqlCon.Open();
        
        //Заполняем таблицу dtImport
        sqlAdapter.SelectCommand.CommandText =
          "Select ID, DataSet, Samplename, Operator, Charge, Type, Condition, SideFlag, " +
          "TotalCurrent, TotalCoefficent, I_Electrode1, I_Electrode2, I_Electrode3, I_Electrode4, " +
          "I_Electrode5, I_Electrode6, I_Electrode7, I_Electrode8, I_Electrode9, I_Electrode10 from dbo.FrResults  where IsNew = 'Y'";


        sqlAdapter.SelectCommand.CommandType = CommandType.Text;
        sqlAdapter.Fill(dtImport);

        if (dtImport.Rows.Count == 0){
          DxInfo.ShowDxBoxInfo("Внимание", "Данные с Франклина-20 отсутствуют!", MessageBoxImage.Warning);
          return;
        }
        //чистим таблицу
        if (!Odac.ExecuteNonQuery("DELETE FROM LIMS.ML_FR20USB", CommandType.Text, false, null))
          throw new ApplicationException(string.Format(CultureInfo.InvariantCulture, "Ошибка очистки таблицы LIMS.ML_FR20USB", 1));

        //грузим в таблицу
        loader.Connection = Odac.DbConnection;
        loader.TableName = "LIMS.ML_FR20USB";
        Odac.DbConnection.Open();
        loader.CreateColumns();
        loader.Open();
        foreach (DataRow row in dtImport.Rows){
          for (int i = 0; i < dtImport.Columns.Count; i++)
            loader.SetValue(i, row[i]);

          loader.NextRow();
        }
        loader.Close();

        
        //обрабатываем в Lims
        if (!Odac.ExecuteNonQuery("LIMS.MagLab.Fr20Usb", CommandType.StoredProcedure, false, null))
          throw new ApplicationException(string.Format(CultureInfo.InvariantCulture, "Ошибка обработки данных из таблицы LIMS.ML_FR20USB"));
        

        //изменяем обработанные данные
        sqlAdapter.UpdateCommand.CommandText = "update dbo.FrResults set IsNew = 'N' where IsNew = 'Y'";
        sqlAdapter.UpdateCommand.ExecuteNonQuery();

        DxInfo.ShowDxBoxInfo("Обработка данных", "Данные с устройства Франклин успешно обработаны!", MessageBoxImage.Information);
      }
      catch (Exception e){
        DxInfo.ShowDxBoxInfo("Ошибка","Ошибка: " + e.Message, MessageBoxImage.Error);
      }
      finally{
        loader.Close();
        loader.Dispose();
        Odac.DbConnection.Close();
        sqlAdapter.SelectCommand.Dispose();
        sqlAdapter.UpdateCommand.Dispose();
        sqlAdapter.Dispose();
        sqlCon.Close();
        sqlCon.Dispose();
        dtImport.Clear();
      }
      
    }

    public static void SetOstNaprForStend(string sampleId)
    {
      const string stmtSql = "LIMS.MagLab.OstNaprFromStend";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(sampleId))
        len = sampleId.Length;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = len,
        Value = sampleId
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static string CheckThickness(string sampleId, Boolean allState)
    {
      const string stmtSql = "LIMS.MagLab.CheckThickness";
      var lstPrm = new List<OracleParameter>();
      int len = 0;

      if (!String.IsNullOrEmpty(sampleId))
        len = sampleId.Length;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = len,
        Value = sampleId
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = 1,
        Value = allState ? "Y" : "N"
      };
      lstPrm.Add(prm);

      var prmRetVal = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.ReturnValue,
        OracleDbType = OracleDbType.VarChar,
        Size = 1,
      };
      lstPrm.Add(prmRetVal);
      
      Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);

      return Convert.ToString(prmRetVal.Value);
    }

    public static Boolean AddNewDateRangeStZap(DateTime dateFrom)
    {
      const string stmtSql = "LIMS.MagLab.AddNewDateRangeStZap";
      List<OracleParameter> lstPrm = new List<OracleParameter>();

      OracleParameter prm = new OracleParameter()
      {
        DbType = DbType.DateTime,
        OracleDbType = OracleDbType.Date,
        Direction = ParameterDirection.Input,
        Value = dateFrom
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static Boolean DeleteLastDateRangeStZap()
    {
      const string stmtSql = "LIMS.MagLab.DeleteLastDateRangeStZap";
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, null);
    }

    public static Boolean CreateSiemensSamples(string locNum)
    {
      const string stmtSql = "LIMS.MagLab.CreateSiemensSamples";

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = 64,
        Value = locNum
      };

      List<OracleParameter> lstPrm = new List<OracleParameter> {prm};
      return Odac.ExecuteNonQuery(stmtSql, CommandType.StoredProcedure, false, lstPrm);
    }

    public static void AdgRptRpt(DateTime dtBegin, DateTime dtEnd, string adgRptScript1, string adgRptScript2, string adgRptScript3)
    {
      var sfd = new SaveFileDialog
      {
        OverwritePrompt = false,
        AddExtension = true,
        DefaultExt = ".csv",
        Filter = "csv file (.csv)|*.csv"
      };

      if (sfd.ShowDialog().GetValueOrDefault() != true)
        return;

      if (File.Exists(sfd.FileName)){
        DxInfo.ShowDxBoxInfo("Файл", "Файл: " + sfd.FileName + " уже существует!", MessageBoxImage.Error);
        return;
      }

      OracleDataReader odr = null;
      DbVar.SetRangeDate(dtBegin, dtEnd, 1);

      Etc.WriteToEndTxtFile(sfd.FileName, "Период:", Encoding.GetEncoding("windows-1251"));
      string dtStr = $" c {DbVar.GetDateBeginEnd(true, true):dd.MM.yyyy HH:mm:ss}" + " по " + $"{DbVar.GetDateBeginEnd(false, true):dd.MM.yyyy HH:mm:ss}";
      Etc.WriteToEndTxtFile(sfd.FileName, dtStr, Encoding.GetEncoding("windows-1251"));
      Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));

      string sqlStmt = File.ReadAllText(adgRptScript1, Encoding.GetEncoding("windows-1251")); 
      odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, "Кол-во рулонов;% Cеред Внутр ДА;% Cеред Внутр НЕТ;% Cеред Внеш ДА;% Cеред Внеш НЕТ", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();
      /***********************************************************************************/
      sqlStmt = File.ReadAllText(adgRptScript2, Encoding.GetEncoding("windows-1251")); 
      odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "Кол-во рулонов;% Cеред Внутр O;% Cеред Внутр A;% Cеред Внутр B;% Cеред Внутр C;% Cеред Внутр D;% Cеред Внутр E;% Cеред Внутр F;% Cеред Внеш G;% Cеред Внеш X", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i).ToString() + ";";

          Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }

      }

      odr?.Close();
      odr?.Dispose();
      /***********************************************************************************/
      sqlStmt = File.ReadAllText(adgRptScript3, Encoding.GetEncoding("windows-1251"));
      odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

      if (odr != null){

        int flds = odr.FieldCount;
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, " ", Encoding.GetEncoding("windows-1251"));
        Etc.WriteToEndTxtFile(sfd.FileName, "Кол-во рулонов;% Внутр O;% Внутр A;% Внутр B;% Внутр C;% Внутр D;% Внутр E;% Внутр F;% Внеш G;% Внеш X", Encoding.GetEncoding("windows-1251"));

        while (odr.Read()){
          
          var strTmp = "";

          for (int i = 0; i < flds; i++)
            strTmp += odr.GetValue(i: i).ToString() + ";";
          

        Etc.WriteToEndTxtFile(sfd.FileName, strTmp, Encoding.GetEncoding("windows-1251"));
        }
        
      }

      odr?.Close();
      odr?.Dispose();

      DxInfo.ShowDxBoxInfo("Файл", "Файл: " + sfd.FileName + " сформирован!", MessageBoxImage.Information);
    }


  }
}
