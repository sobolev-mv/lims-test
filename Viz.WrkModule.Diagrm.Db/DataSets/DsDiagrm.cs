using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.WrkModule.Diagrm.Db.DataSets
{
  public enum AxisExtremes
  {
    Min,
    Max
  }

  public sealed class DsDiagrm : DataSet
  {
    private DgTypeDiagrammDataTable dgTypeDiagramm;
    private DgMeasDataDataTable     dgMeasData;
    private DgChoiceTypeDataTable   dgChoiceType;
    private DgGroupDataTable        dgGroup;

    public DgTypeDiagrammDataTable DgTypeDiagramm
    {
      get { return dgTypeDiagramm; }
    }

    public DgMeasDataDataTable DgMeasData
    {
      get { return dgMeasData; }
    }

    public DgChoiceTypeDataTable DgChoiceType
    {
      get { return dgChoiceType; }
    }

    public DgGroupDataTable DgGroup
    {
      get { return dgGroup; }
    }

    public DsDiagrm()
    {
       DataSetName = "DsDiagrm";
       dgTypeDiagramm = new DgTypeDiagrammDataTable();
       Tables.Add(dgTypeDiagramm);

       dgMeasData = new DgMeasDataDataTable();
       Tables.Add(dgMeasData);

       dgChoiceType = new DgChoiceTypeDataTable();
       Tables.Add(dgChoiceType);

       dgGroup = new DgGroupDataTable();
       Tables.Add(dgGroup); 
    }

    public sealed class DgTypeDiagrammDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public DgTypeDiagrammDataTable()
      {
        TableName = "DgTypeDiagramm";
        adapter = new OracleDataAdapter();

        DataColumn col = null;

        col = new DataColumn("PvName", typeof(string), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Desc", typeof(string), null, MappingType.Element);
        Columns.Add(col);

        col = new DataColumn("DgXtype", typeof(string), null, MappingType.Element);
        Columns.Add(col);

        Constraints.Add(new UniqueConstraint("Pk_DgTypeDiagramm", new []{ Columns["PvName"] }, true));
        Columns["PvName"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new DataTableMapping("SourceTable1", "DgTypeDiagramm");
        dtm.ColumnMappings.Add("PV_NAME", "PvName");
        dtm.ColumnMappings.Add("DESCRIPTION", "Desc");
        dtm.ColumnMappings.Add("XTYPE", "DgXtype");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
      }

      public int LoadData(int GroupId)
      {
        const string SqlStmt = "SELECT S.PV_NAME, NVL(S.ALT_DESC, PD.DESCRIPTION) DESCRIPTION, S.XTYPE " +
                               "FROM VIZ_PRN.DG_SERIES S " +
                               "LEFT JOIN VIZ.PAM_DDDESCRIPTION PD ON (PD.PV_NAME = S.PV_NAME) AND (PD.LNGCODE = 'ru') " +   
                               "WHERE (S.GROUP_ID = :GID) " +
                               "ORDER BY S.SRT";

        adapter.SelectCommand.Parameters.Clear();
        adapter.SelectCommand.CommandText = SqlStmt;
        adapter.SelectCommand.CommandType = CommandType.Text;

        var prm = new OracleParameter
                    {
                      DbType = DbType.Int32,
                      Direction = ParameterDirection.Input,
                      OracleDbType = OracleDbType.Integer,
                      ParameterName = "GID"
                    };
        adapter.SelectCommand.Parameters.Add(prm);

        var lstPrmValue = new List<Object> {GroupId};
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

      public string GetXType(string dgName)
      {
        DefaultView.Sort = "PvName";
        var i = DefaultView.Find(dgName);
        return (i == -1) ? null : Convert.ToString(DefaultView[i]["DgXtype"]);
      }
      
    }

    public sealed class DgChoiceTypeDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public DgChoiceTypeDataTable()
      {
        TableName = "DgChoiceType";
        adapter = new OracleDataAdapter();

        DataColumn col = null;

        col = new DataColumn("PvName", typeof(string), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Name", typeof(string), null, MappingType.Element);
        Columns.Add(col);

        col = new DataColumn("GroupName", typeof(string), null, MappingType.Element);
        Columns.Add(col);

        Constraints.Add(new UniqueConstraint("Pk_DgChoiceType", new[] { Columns["PvName"] }, true));
        Columns["PvName"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new DataTableMapping("SourceTable1", "DgChoiceType");
        dtm.ColumnMappings.Add("PV_NAME", "PvName");
        dtm.ColumnMappings.Add("NAME", "Name");
        dtm.ColumnMappings.Add("GROUP_NAME", "GroupName");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
      }

      public int LoadData()
      {
        const string SqlStmt = "SELECT PD.PV_NAME, PD.DESCRIPTION NAME, G.NAME GROUP_NAME " + 
                               "FROM VIZ.PAM_DDDESCRIPTION PD " +
                               "INNER JOIN VIZ_PRN.DG_SERIES S ON (PD.PV_NAME = S.PV_NAME) " +
                               "INNER JOIN VIZ_PRN.DG_GROUP G ON (S.GROUP_ID = G.ID) " +
                               "WHERE PD.LNGCODE = 'ru' " +
                               "ORDER BY G.NAME";

        adapter.SelectCommand.Parameters.Clear();
        adapter.SelectCommand.CommandText = SqlStmt;
        adapter.SelectCommand.CommandType = CommandType.Text;

        return Odac.LoadDataTable(this, adapter, true, null);
      }

    }

    public sealed class DgGroupDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public DgGroupDataTable()
      {
        TableName = "DgGroup";
        adapter = new OracleDataAdapter();

        DataColumn col = null;

        col = new DataColumn("Id", typeof(Int32), null, MappingType.Element)
                {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Name", typeof(string), null, MappingType.Element);
        Columns.Add(col);

        Constraints.Add(new UniqueConstraint("Pk_DgGroup", new[] { Columns["Id"] }, true));
        Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new DataTableMapping("SourceTable1", "DgGroup");
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("NAME", "Name");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
      }

      public int LoadData()
      {
        const string SqlStmt = "SELECT ID, NAME FROM VIZ_PRN.DG_GROUP ORDER BY ID";
        adapter.SelectCommand.Parameters.Clear();
        adapter.SelectCommand.CommandText = SqlStmt;
        adapter.SelectCommand.CommandType = CommandType.Text;
        return Odac.LoadDataTable(this, adapter, true, null);
      }

    }

    public sealed class DgMeasDataDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public DgMeasDataDataTable()
      {
        TableName = "DgMeasDataData";
        adapter = new OracleDataAdapter();

        DataColumn col = null;

        col = new DataColumn("Len", typeof(Decimal), null, MappingType.Element);
        Columns.Add(col);

        col = new DataColumn("DgDateTime", typeof(DateTime), null, MappingType.Element);
        Columns.Add(col);

        col = new DataColumn("Value", typeof(Decimal), null, MappingType.Element);
        Columns.Add(col);

        //this.Constraints.Add(new System.Data.UniqueConstraint("Pk_DgMeasDataData", new System.Data.DataColumn[] { this.Columns["Len"] }, true));
        //this.Columns["Len"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new DataTableMapping("SourceTable1", "DgMeasDataData");
        dtm.ColumnMappings.Add("LENGTH", "Len");
        dtm.ColumnMappings.Add("TS_GMT", "DgDateTime");
        dtm.ColumnMappings.Add("VALUE", "Value");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
      }

      public int LoadData(string mlocId, string typeDiagrm, string xAxisType)
      {
        const string SqlStmt = "SELECT PD.LENGTH / 1000 LENGTH, PD.TS_GMT, PD.VALUE " +
                               "FROM VIZ.PAM_MV_HEADER PH " +
                               "INNER JOIN VIZ.PAM_MV_DATA PD ON (PH.MV_ID = PD.MV_ID) " +
                               "WHERE PH.ME_ID = (SELECT ME_ID FROM VIZ.MAT WHERE BEZEICHNUNG = :LID) " +
                               "AND (PH.MV_NAME = :TDG)";
        string sqlStmtOdr;

        if (xAxisType == "NUM")
          sqlStmtOdr = SqlStmt + " ORDER BY 1";
        else
          sqlStmtOdr = SqlStmt + " ORDER BY 2";

        adapter.SelectCommand.Parameters.Clear();
        adapter.SelectCommand.CommandText = sqlStmtOdr;
        adapter.SelectCommand.CommandType = CommandType.Text;

        var prm = new OracleParameter
                    {
                      DbType = DbType.String,
                      Direction = ParameterDirection.Input,
                      OracleDbType = OracleDbType.VarChar,
                      Size = 64,
                      ParameterName = "LID"
                    };
        adapter.SelectCommand.Parameters.Add(prm);

        prm = new OracleParameter
                {
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  OracleDbType = OracleDbType.VarChar,
                  Size = typeDiagrm.Length,
                  ParameterName = "TDG"
                };
        adapter.SelectCommand.Parameters.Add(prm);

        var lstPrmValue = new List<Object> {mlocId, typeDiagrm};
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

      public object GetAxisXExtremesVal(string mlocId, string typeDiagrm, string xType, AxisExtremes axisExtr)
      {
        string sqlStmtPart2 = "FROM VIZ.PAM_MV_HEADER PH " +
                         "INNER JOIN VIZ.PAM_MV_DATA PD ON (PH.MV_ID = PD.MV_ID) " +
                         "WHERE PH.ME_ID = (SELECT ME_ID FROM VIZ.MAT WHERE BEZEICHNUNG = :LID) " +
                         "AND (PH.MV_NAME = :TDG)";
        string sqlStmtPart1;

        if (string.Equals(xType, "NUM", StringComparison.Ordinal) && axisExtr == AxisExtremes.Min)
          sqlStmtPart1 = "SELECT MIN(PD.LENGTH) / 1000 ";
        else if (string.Equals(xType, "NUM", StringComparison.Ordinal) && axisExtr == AxisExtremes.Max)
          sqlStmtPart1 = "SELECT MAX(PD.LENGTH) / 1000 ";
        else if (string.Equals(xType, "DTM", StringComparison.Ordinal) && axisExtr == AxisExtremes.Min)
          sqlStmtPart1 = "SELECT MIN(PD.TS_GMT) ";
        else if (string.Equals(xType, "DTM", StringComparison.Ordinal) && axisExtr == AxisExtremes.Max)
          sqlStmtPart1 = "SELECT MAX(PD.TS_GMT) ";
        else
          sqlStmtPart1 = "";

        var lstPrmValue = new List<OracleParameter>();

        var prm = new OracleParameter
        {
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = 64,
          ParameterName = "LID",
          Value = mlocId
        };
        lstPrmValue.Add(prm);

        prm = new OracleParameter
        {
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = typeDiagrm.Length,
          ParameterName = "TDG",
          Value = typeDiagrm
        };
        lstPrmValue.Add(prm);
        
        return Odac.ExecuteScalar(sqlStmtPart1 + sqlStmtPart2, CommandType.Text, false, lstPrmValue);
      }

      public object GetAxisYExtremesVal(string mlocId, string typeDiagrm, AxisExtremes axisExtr)
      {
        string sqlStmtPart2 = "FROM VIZ.PAM_MV_HEADER PH " +
                              "INNER JOIN VIZ.PAM_MV_DATA PD ON (PH.MV_ID = PD.MV_ID) " +
                              "WHERE PH.ME_ID = (SELECT ME_ID FROM VIZ.MAT WHERE BEZEICHNUNG = :LID) " +
                              "AND (PH.MV_NAME = :TDG)";
        string sqlStmtPart1;

        if (axisExtr == AxisExtremes.Min)
          sqlStmtPart1 = "SELECT MIN(PD.VALUE) ";
        else if (axisExtr == AxisExtremes.Max)
          sqlStmtPart1 = "SELECT MAX(PD.VALUE) ";
        else
          sqlStmtPart1 = "";

        var lstPrmValue = new List<OracleParameter>();

        var prm = new OracleParameter
        {
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = 64,
          ParameterName = "LID",
          Value = mlocId
        };
        lstPrmValue.Add(prm);

        prm = new OracleParameter
        {
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = typeDiagrm.Length,
          ParameterName = "TDG",
          Value = typeDiagrm
        };
        lstPrmValue.Add(prm);

        return Odac.ExecuteScalar(sqlStmtPart1 + sqlStmtPart2, CommandType.Text, false, lstPrmValue);
      }


    }





  }
  
}
