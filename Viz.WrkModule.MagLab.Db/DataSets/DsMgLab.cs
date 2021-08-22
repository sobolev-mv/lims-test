using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.MagLab.Db.DataSets {

  public sealed class DsMgLab : DataSet
  {
    private MlSamplesDataTable     mlSamples;
    private MlDataDataTable        mlData;
    private MlDataProbeDataTable   mlDataProbe;
    private MlUsetDataTable        mlUset;
    private MlValDataDataTable     mlValData;
    private MlUtypeInfoDataTable   mlUtypeInfoDataTable;
    private MlListApInfoDataTable  mlListApInfoDataTable;
    private MlMatGnlDataTable      mlMatGnlDataTable;
    private MlMk4auDataTable       mlMk4auDataTable; 
    private MlMk4apDataTable       mlMk4apDataTable;
    private MlStZapDataTable       mlStZapDataTable;
    private MlMesurCofDataTable    mlMesurCofDataTable;
    private MlShiftDataTable       mlShift;
    private MlSiemensSmpDataTable  mlSiemensSmp;
    private MlDeviceLstDataTable   mlDeviceLst;
    private MlMpg200dDataTable     mlMpg200d;

    public MlMpg200dDataTable MlMpg200d
    {
      get { return this.mlMpg200d; }
    }

    public MlDeviceLstDataTable MlDeviceLst
    {
      get { return this.mlDeviceLst; }
    }
    
    public MlSiemensSmpDataTable MlSiemensSmp
    {
      get { return this.mlSiemensSmp; }
    }

    public MlShiftDataTable MlShift
    {
      get { return this.mlShift; }
    }
    
    public MlSamplesDataTable MlSamples
    {
      get { return this.mlSamples; }
    }
    
    public MlDataDataTable MlData
    {
      get { return this.mlData; }
    }

    
    public MlDataProbeDataTable MlDataProbe
    {
      get { return this.mlDataProbe; }
    }

    
    public MlUsetDataTable MlUset
    {
      get { return this.mlUset; }
    }

    public MlValDataDataTable MlValData
    {
      get { return this.mlValData; }
    }

    public MlUtypeInfoDataTable MlUtypeInfo
    {
      get { return this.mlUtypeInfoDataTable; }
    }

    public MlListApInfoDataTable MlListApInfo
    {
      get { return this.mlListApInfoDataTable; }
    }

    public MlMatGnlDataTable MlMatGnl
    {
      get { return this.mlMatGnlDataTable; }
    }

    public MlMk4auDataTable MlMk4au
    { 
      get { return this.mlMk4auDataTable;}
    }

    public MlMk4apDataTable MlMk4ap
    {
      get { return this.mlMk4apDataTable; }
    }

    public MlStZapDataTable MlStZap
    {
      get { return this.mlStZapDataTable; }
    }

    public MlMesurCofDataTable MlMesurCof
    {
      get { return this.mlMesurCofDataTable; }
    }

    /*возвращает DateTime начала и конца смены, если shift = 1 или 2 либо просто диапазон, если shift = 3*/
    internal void GetShiftDate(DateTime dateStartIn, DateTime dateEndIn, int shift, out DateTime dateStart, out DateTime dateEnd)
    {
      switch (shift){
        case 1:
          dateStart = dateStartIn.AddHours(8);
          dateEnd = dateStart.AddHours(11).AddMinutes(59).AddSeconds(59);
          break;
        case 2:
          dateStart = dateStartIn.AddHours(20);
          dateEnd = dateStart.AddHours(11).AddMinutes(59).AddSeconds(59);
          break;
        default:
          dateStart = dateStartIn;
          dateEnd = dateEndIn;
          break;
      }
    }

    public DsMgLab() : base()
    {
      this.DataSetName = "DsMgLab";

      this.mlSiemensSmp = new MlSiemensSmpDataTable();
      this.Tables.Add(this.mlSiemensSmp);
      this.mlShift = new MlShiftDataTable();
      this.Tables.Add(this.mlShift);
      this.mlSamples = new MlSamplesDataTable();
      this.Tables.Add(this.mlSamples);
      this.mlData = new MlDataDataTable();
      this.Tables.Add(this.mlData);
      this.mlDataProbe = new MlDataProbeDataTable();
      this.Tables.Add(this.mlDataProbe);
      this.mlUset = new MlUsetDataTable();
      this.Tables.Add(this.mlUset);
      this.mlValData = new MlValDataDataTable();
      this.Tables.Add(this.mlValData);
      this.mlUtypeInfoDataTable = new MlUtypeInfoDataTable();
      this.Tables.Add(this.mlUtypeInfoDataTable);
      this.mlListApInfoDataTable = new MlListApInfoDataTable();
      this.Tables.Add(this.mlListApInfoDataTable);
      this.mlMatGnlDataTable = new MlMatGnlDataTable();
      this.Tables.Add(this.mlMatGnlDataTable);
      this.mlMk4auDataTable = new MlMk4auDataTable();
      this.Tables.Add(this.mlMk4auDataTable);
      this.mlMk4apDataTable = new MlMk4apDataTable();
      this.Tables.Add(this.mlMk4apDataTable);
      this.mlStZapDataTable = new MlStZapDataTable("MlStZap");
      this.Tables.Add(this.mlStZapDataTable);
      this.mlMesurCofDataTable = new MlMesurCofDataTable();
      this.Tables.Add(this.mlMesurCofDataTable);
      this.mlDeviceLst = new MlDeviceLstDataTable();
      this.Tables.Add(this.mlDeviceLst);
      this.mlMpg200d = new MlMpg200dDataTable();
      this.Tables.Add(this.mlMpg200d);

    }
  }

  /******************************************************/
  public sealed class MlMpg200dDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlMpg200dDataTable() : base()
    {
      TableName = "MlMpg200d";
      adapter = new OracleDataAdapter();

      DataColumn col = null;

      col = new DataColumn("Utype", typeof(Int32), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);
      
      col = new DataColumn("CoilName", typeof(string), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("Density", typeof(double), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("LengthSmp", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("WidthSmp", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Quantity", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("MlMpg200d_pk", new[] { this.Columns["Utype"], this.Columns["CoilName"] }, true));
      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("LIMS.ML_MPG200D", "MlMpg200d");
      dtm.ColumnMappings.Add("UTYPE", "Utype");
      dtm.ColumnMappings.Add("COILNAME", "CoilName");
      dtm.ColumnMappings.Add("DENSITY", "Density");
      dtm.ColumnMappings.Add("LENGTHSMP", "LengthSmp");
      dtm.ColumnMappings.Add("WIDTHSMP", "WidthSmp");
      dtm.ColumnMappings.Add("QUANTITY", "Quantity");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand { Connection = Odac.DbConnection };
      adapter.SelectCommand.CommandText = "SELECT * FROM LIMS.ML_MPG200D WHERE UTYPE = :PUTYPE ORDER BY 1, 2";
      adapter.SelectCommand.CommandType = CommandType.Text;

      var prm = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Integer,
        ParameterName = "PUTYPE"
      };
      adapter.SelectCommand.Parameters.Add(prm);
    }

    public int LoadData(Int32 uType)
    {
      var lstPrmValue = new List<Object> { uType };
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

  }

  public sealed class MlDeviceLstDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlDeviceLstDataTable() : base()
    {
      TableName = "MlDeviceLst";
      adapter = new OracleDataAdapter();

      DataColumn col = null;

      col = new DataColumn("Id", typeof(Int32), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("Name", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("MlDeviceLst_pk", new[] { this.Columns["Id"] }, true));
      this.Columns["Id"].Unique = true;

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("LIMS.ML_LSTDEVICE", "MlDeviceLst");
      dtm.ColumnMappings.Add("ID", "Id");
      dtm.ColumnMappings.Add("NAME", "Name");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand { Connection = Odac.DbConnection };
      adapter.SelectCommand.CommandText = "SELECT ID, NAME FROM LIMS.V_MESURE_DEVICE WHERE UTYPE = :PUTYPE ORDER BY 1";
      adapter.SelectCommand.CommandType = CommandType.Text;
      var prm = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Integer,
        ParameterName = "PUTYPE"
      };
      adapter.SelectCommand.Parameters.Add(prm);
    }

    public int LoadData(Int32 uType)
    {
      var lstPrmValue = new List<Object> { uType };
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }
    
  }


  public sealed class MlShiftDataTable : DataTable
  {
    
    public MlShiftDataTable() : base()
    {
      TableName = "MlShift";
      Columns.Add("Id", typeof(int));
      Columns.Add("NameShift", typeof(string));

      DataRow row = this.NewRow();
      row[0] = 1;
      row[1] = "1 Смена";
      Rows.Add(row);

      row = NewRow();
      row[0] = 2;
      row[1] = "2 Смена";
      Rows.Add(row);

      row = NewRow();
      row[0] = 3;
      row[1] = "Без смен";
      Rows.Add(row);

      AcceptChanges();

    }
  }

  public sealed class MlSiemensSmpDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlSiemensSmpDataTable() : base()
    {
      this.TableName = "MlSiemensSmp";
      adapter = new OracleDataAdapter();

      DataColumn col = null;

      col = new DataColumn("Id", typeof(Int64), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("TestType", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("SteelType", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("ThickNessNominal", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Line", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("LaserFlag", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("SamplePos", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("SamplePos2", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MatLocalNumber", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MatMarkingInfo", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("DtBegin", typeof(DateTime), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("DtEnd", typeof(DateTime), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("State", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Md", typeof(String), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Tstep", typeof(String), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("SlFlg", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("UsrIns", typeof(String), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("UsrUpd", typeof(String), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("B100", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("B800", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("B2500", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("P1550", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("P1750", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Dpp1750", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);


      this.Constraints.Add(new UniqueConstraint("MlSiemensSmp_pk", new[] { this.Columns["Id"] }, true));
      this.Columns["Id"].Unique = true;

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("V_SIEMENSSMP", "MlSiemensSmp");
      dtm.ColumnMappings.Add("ID", "Id");
      dtm.ColumnMappings.Add("TESTTYPE", "TestType");
      dtm.ColumnMappings.Add("STEELTYPE", "SteelType");
      dtm.ColumnMappings.Add("THICNESSNOMINAL", "ThickNessNominal");
      dtm.ColumnMappings.Add("LINE", "Line");
      dtm.ColumnMappings.Add("LASERFLAG", "LaserFlag");
      dtm.ColumnMappings.Add("SAMPLEPOS", "SamplePos");
      dtm.ColumnMappings.Add("SAMPLEPOS2", "SamplePos2");
      dtm.ColumnMappings.Add("MATLOCALNUMBER", "MatLocalNumber");
      dtm.ColumnMappings.Add("MATMARKINGINFO", "MatMarkingInfo");
      dtm.ColumnMappings.Add("DTBEGIN", "DtBegin");
      dtm.ColumnMappings.Add("DTEND", "DtEnd");
      dtm.ColumnMappings.Add("STATE", "State");
      dtm.ColumnMappings.Add("MD", "Md");
      dtm.ColumnMappings.Add("TSTEP", "Tstep");
      dtm.ColumnMappings.Add("SL_FLG", "SlFlg");
      dtm.ColumnMappings.Add("USR_INS", "UsrIns");
      dtm.ColumnMappings.Add("USR_UPD", "UsrUpd");
      dtm.ColumnMappings.Add("B100", "B100");
      dtm.ColumnMappings.Add("B800", "B800");
      dtm.ColumnMappings.Add("B2500", "B2500");
      dtm.ColumnMappings.Add("P1550", "P1550");
      dtm.ColumnMappings.Add("P1750", "P1750");
      dtm.ColumnMappings.Add("DPP1750", "Dpp1750");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand { Connection = Odac.DbConnection };
      adapter.UpdateCommand = new OracleCommand { Connection = Odac.DbConnection };

      //Update Command
      adapter.UpdateCommand.CommandText = "UPDATE LIMS.ML_SAMPLE_SIEMENS SET B100 = :PB100, B800 = :PB800, B2500 = :PB2500, P1550 = :PP1550, P1750 = :PP1750 " +
                                          "WHERE (ID = :Original_ID)";
      adapter.UpdateCommand.CommandType = CommandType.Text;

      var param = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "PB100",
        Precision = 10,
        Scale = 2,
        SourceColumn = "B100",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "PB800",
        Precision = 10,
        Scale = 2,
        SourceColumn = "B800",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "PB2500",
        Precision = 10,
        Scale = 2,
        SourceColumn = "B2500",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "PP1550",
        Precision = 10,
        Scale = 2,
        SourceColumn = "P1550",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "PP1750",
        Precision = 10,
        Scale = 2,
        SourceColumn = "P1750",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "Original_ID",
        SourceColumn = "ID",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Original
      };
      adapter.UpdateCommand.Parameters.Add(param);

      adapter.SelectCommand.ParameterCheck = false;
      adapter.UpdateCommand.ParameterCheck = false;

      foreach (OracleParameter prm in adapter.UpdateCommand.Parameters)
        prm.IsNullable = true;

      foreach (OracleParameter prm in adapter.SelectCommand.Parameters)
        prm.IsNullable = true;

    }

    public int GetListSimensSample(DateTime dateStart, DateTime dateEnd, int shift)
    {
      DateTime dt1 = DateTime.MinValue;
      DateTime dt2 = DateTime.MinValue;

      (this.DataSet as DsMgLab)?.GetShiftDate(dateStart, dateEnd, shift, out dt1, out dt2);

      const string sqlStmt = "SELECT ID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, SAMPLEPOS2, MATLOCALNUMBER, MATMARKINGINFO, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTBEGIN) DTBEGIN, " +
                             "VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTEND) DTEND, STATE, MD, TSTEP, SL_FLG, USR_INS, USR_UPD, B100, B800, B2500, P1550, P1750, DPP1750 " +
                             "FROM LIMS.V_SIEMENSSMP WHERE (DTBEGIN BETWEEN VIZ_PRN.VAR_RPT.TOUTCTIMEZONE(:DT1) AND VIZ_PRN.VAR_RPT.TOUTCTIMEZONE(:DT2)) ORDER BY DTBEGIN";

      adapter.SelectCommand.Parameters.Clear();
      adapter.SelectCommand.CommandText = sqlStmt;
      adapter.SelectCommand.CommandType = CommandType.Text;

      var prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Date,
        ParameterName = "DT1"
      };
      adapter.SelectCommand.Parameters.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Date,
        ParameterName = "DT2"
      };
      adapter.SelectCommand.Parameters.Add(prm);

      var lstPrmValue = new List<Object> { dt1, dt2 };
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

    public int SearchByLocNum(string locNum)
    {

      const string sqlStmt = "SELECT ID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, SAMPLEPOS2, MATLOCALNUMBER, MATMARKINGINFO, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTBEGIN) DTBEGIN, " +
                             "VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTEND) DTEND, STATE, MD, TSTEP, SL_FLG, USR_INS, USR_UPD, B100, B800, B2500, P1550, P1750, DPP1750 " +
                             "FROM LIMS.V_SIEMENSSMP " +
                             "WHERE MATLOCALNUMBER = :PLOCNUM";

      adapter.SelectCommand.Parameters.Clear();
      adapter.SelectCommand.CommandText = sqlStmt;
      adapter.SelectCommand.CommandType = CommandType.Text;

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        OracleDbType = OracleDbType.VarChar,
        Direction = ParameterDirection.Input,
        ParameterName = "PLOCNUM"
      };
      adapter.SelectCommand.Parameters.Add(prm);
      
      var lstPrmValue = new List<Object> { locNum };
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

    public int SaveData()
    {
      return Odac.SaveChangedData(this, adapter);
    }

  }

  public sealed class MlSamplesDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

     public MlSamplesDataTable() : base()
     {
       this.TableName = "MlSamples";
       adapter = new OracleDataAdapter();
              
       DataColumn col = null;

       col = new DataColumn("SampleId", typeof(string), null, MappingType.Element){AllowDBNull = false};
       this.Columns.Add(col);
       
       col = new DataColumn("TestType", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("SteelType", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("ThickNessNominal", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("Line", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("LaserFlag", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("SamplePos", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("MatLocalNumber", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("MatMarkingInfo", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("DtSample", typeof(DateTime), null, MappingType.Element);
       this.Columns.Add(col);
       
       col = new DataColumn("State", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("Md", typeof(String), null, MappingType.Element);
       this.Columns.Add(col);
 
       col = new DataColumn("SampleNum", typeof(String), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Tstep", typeof(String), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("DtSend", typeof(DateTime), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("IsValidate", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("StFlg", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("SlFlg", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col);

      col = new DataColumn("UsrSample", typeof(String), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("IsQm", typeof(Int32), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("MlSamples", new[] { this.Columns["SampleId"] }, true));
       this.Columns["SampleId"].Unique = true;

       adapter.TableMappings.Clear();
       var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlSamples");
       dtm.ColumnMappings.Add("SAMPLEID", "SampleId");
       dtm.ColumnMappings.Add("TESTTYPE", "TestType");
       dtm.ColumnMappings.Add("STEELTYPE", "SteelType");
       dtm.ColumnMappings.Add("THICNESSNOMINAL", "ThickNessNominal");
       dtm.ColumnMappings.Add("LINE", "Line");
       dtm.ColumnMappings.Add("LASERFLAG", "LaserFlag");
       dtm.ColumnMappings.Add("SAMPLEPOS", "SamplePos");
       dtm.ColumnMappings.Add("MATLOCALNUMBER", "MatLocalNumber");
       dtm.ColumnMappings.Add("MATMARKINGINFO", "MatMarkingInfo");
       dtm.ColumnMappings.Add("DTSAMPLE", "DtSample");
       dtm.ColumnMappings.Add("STATE", "State");
       dtm.ColumnMappings.Add("MD", "Md");
       dtm.ColumnMappings.Add("SAMPLENUM", "SampleNum");
       dtm.ColumnMappings.Add("TSTEP", "Tstep");
       dtm.ColumnMappings.Add("DTSEND", "DtSend");
       dtm.ColumnMappings.Add("ISVALIDATE", "IsValidate");
       dtm.ColumnMappings.Add("ST_FLG", "StFlg");
       dtm.ColumnMappings.Add("SL_FLG", "SlFlg");
       dtm.ColumnMappings.Add("USR_SAMPLE", "UsrSample");
       dtm.ColumnMappings.Add("ISQM", "IsQm");

      adapter.TableMappings.Add(dtm);
       adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
     }

     public int GetListSample(DateTime dateStart, DateTime dateEnd, int shift)
     {
       DateTime dt1 = DateTime.MinValue;
       DateTime dt2 = DateTime.MinValue;

       (this.DataSet as DsMgLab)?.GetShiftDate(dateStart, dateEnd, shift, out dt1, out dt2);

       const string sqlStmt = "SELECT SAMPLEID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, MATLOCALNUMBER, MATMARKINGINFO, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTSAMPLE) DTSAMPLE, STATE, MD,  SAMPLENUM, TSTEP, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTSEND) DTSEND, " +
                              "ISVALIDATE, ST_FLG, SL_FLG, USR_SAMPLE, ISQM " +
                              "FROM LIMS.V_SAMPLEMEAS WHERE (DTSAMPLE BETWEEN VIZ_PRN.VAR_RPT.TOUTCTIMEZONE(:DT1) AND VIZ_PRN.VAR_RPT.TOUTCTIMEZONE(:DT2)) ORDER BY DTSAMPLE";
       adapter.SelectCommand.Parameters.Clear();
       adapter.SelectCommand.CommandText = sqlStmt;
       adapter.SelectCommand.CommandType = CommandType.Text;

       var prm = new OracleParameter
                   {
                     DbType = DbType.DateTime,
                     Direction = ParameterDirection.Input,
                     OracleDbType = OracleDbType.Date,
                     ParameterName = "DT1"
                   };
       adapter.SelectCommand.Parameters.Add(prm);

       prm = new OracleParameter
               {
                 DbType = DbType.DateTime,
                 Direction = ParameterDirection.Input,
                 OracleDbType = OracleDbType.Date,
                 ParameterName = "DT2"
               };
       adapter.SelectCommand.Parameters.Add(prm);

       var lstPrmValue = new List<Object> {dt1, dt2};
       return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
     }

     public int SerchBySampleId(String SampleId)
     {
       const string sqlStmt = "SELECT SAMPLEID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, MATLOCALNUMBER, MATMARKINGINFO, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTSAMPLE) DTSAMPLE, STATE, MD,  SAMPLENUM, TSTEP, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTSEND) DTSEND, " +
                              "ISVALIDATE, ST_FLG, SL_FLG, USR_SAMPLE, ISQM " +
                              "FROM LIMS.V_SAMPLEMEAS WHERE MATLOCALNUMBER IN (SELECT MATLOCALNUMBER FROM LIMS.SAMPLEMEAS WHERE SAMPLENUM LIKE :SMPID)";

       adapter.SelectCommand.Parameters.Clear();
       adapter.SelectCommand.CommandText = sqlStmt;
       adapter.SelectCommand.CommandType = CommandType.Text;

       var prm = new OracleParameter
                   {
                     DbType = DbType.String,
                     Direction = ParameterDirection.Input,
                     OracleDbType = OracleDbType.VarChar,
                     Size = 64,
                     ParameterName = "SMPID"
                   };
       adapter.SelectCommand.Parameters.Add(prm);
       //prm.Value = SampleId + "%"; 

       var lstPrmValue = new List<Object> {SampleId + "%"};
       return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
     }

     public int SerchByMatLocalNum(String matLocalNum)
     {
       const string sqlStmt = "SELECT SAMPLEID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, MATLOCALNUMBER, MATMARKINGINFO, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTSAMPLE) DTSAMPLE, STATE, MD,  SAMPLENUM, TSTEP, VIZ_PRN.VAR_RPT.TOLOCALTIMEZONE(DTSEND) DTSEND, " +
                              "ISVALIDATE, ST_FLG, SL_FLG, USR_SAMPLE, ISQM " +
                              "FROM LIMS.V_SAMPLEMEAS WHERE MATLOCALNUMBER LIKE :MLOC";

       adapter.SelectCommand.Parameters.Clear();
       adapter.SelectCommand.CommandText = sqlStmt;
       adapter.SelectCommand.CommandType = CommandType.Text;

       var prm = new OracleParameter
                   {
                     DbType = DbType.String,
                     Direction = ParameterDirection.Input,
                     OracleDbType = OracleDbType.VarChar,
                     Size = 40,
                     ParameterName = "MLOC"
                   };
       adapter.SelectCommand.Parameters.Add(prm);

       var lstPrmValue = new List<Object> {matLocalNum + "%"};
       return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
     }

     public int SerchByMatMarkNum(String matMarkNum)
     {
       var lstPrmValue = new List<Object>();
       DateTime DateStart = DateTime.Now.AddDays(-10000);
       DateTime DateEnd = DateTime.Now.AddDays(10000);
       lstPrmValue.Add(DateStart);
       lstPrmValue.Add(DateEnd);
       lstPrmValue.Add("Z");
       lstPrmValue.Add(matMarkNum);
       lstPrmValue.Add("Z");
       return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
     }

  }
   
   
  public sealed class MlDataDataTable : DataTable
  {
    public int Utype { get; set; }
    public int MesDevice { get; set; }

    private readonly OracleDataAdapter adapter;

    //Коррекционные коэффициенты по Эпштейну
    /*
    private static decimal p1750ApCorr;
    private static decimal p1550ApCorr;
    private static decimal b100ApCorr;
    private static decimal b800ApCorr;
    private static decimal b2500ApCorr;
    */
     public MlDataDataTable() : base()
     {
       this.TableName = "MlData";
       adapter = new OracleDataAdapter();

       /* 
       p1750ApCorr = GetCorrVal("P1750AP");
       p1550ApCorr = GetCorrVal("P1550AP");
       b100ApCorr  = GetCorrVal("B100AP");
       b800ApCorr =  GetCorrVal("B800AP");
       b2500ApCorr = GetCorrVal("B2500AP");
       */
       DataColumn col = null;

       col = new DataColumn("Id", typeof(Int32), null, MappingType.Element){AllowDBNull = false};
       this.Columns.Add(col);

       col = new DataColumn("SampleId", typeof(string), null, MappingType.Element){AllowDBNull = false};
       this.Columns.Add(col);

       col = new DataColumn("Utype", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col); 

       col = new DataColumn("Massa", typeof(int), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B3", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col); 

       col = new DataColumn("B30", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B100", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B800", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col); 

       col = new DataColumn("B2500", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B5000", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1050", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1350", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1550", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1750", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P004500", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Hd004500", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P01500", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Hd01500", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P0041000", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Hd0041000", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P011000", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Hd01100", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Iup1", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Iup2", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Iup3", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Idown1", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Idown2", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Idown3", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Adout", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Adin", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Gib1", typeof(int), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Gib2", typeof(int), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Iup4", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Iup5", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Idown4", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("Idown5", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1550ApProd", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B2500ApProd", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1350ApPoper", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B2500ApPoper", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("P1550ApPoper", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B40ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B60ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B120ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B1000ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B1200ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B2500ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B500ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B5000ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B10000ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B24000ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("B30000ni", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("HcPoperAoNi", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("HcPoperBoNi", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("HcProdAoNi", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("HcProdBoNi", typeof(decimal), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("RavgNi", typeof(Int32), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("SteelMarkNi", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("AdOutR", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("AdInR", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("AdOutL", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("AdInL", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("AdOutSimens", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);

       col = new DataColumn("AdInSimens", typeof(string), null, MappingType.Element);
       this.Columns.Add(col);
      

       this.Constraints.Add(new UniqueConstraint("Pk_MlData", new[] { this.Columns["Id"] }, true));
       this.Columns["Id"].Unique = true;
      
       adapter.TableMappings.Clear();
       var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlData");
       dtm.ColumnMappings.Add("ID", "Id");
       dtm.ColumnMappings.Add("SAMPLEID", "SampleId");
       dtm.ColumnMappings.Add("UTYPE", "Utype");
       dtm.ColumnMappings.Add("MASSA", "Massa");
       dtm.ColumnMappings.Add("B3", "B3");
       dtm.ColumnMappings.Add("B30", "B30");
       dtm.ColumnMappings.Add("B100", "B100");
       dtm.ColumnMappings.Add("B800", "B800");
       dtm.ColumnMappings.Add("B2500", "B2500");
       dtm.ColumnMappings.Add("B5000", "B5000");
       dtm.ColumnMappings.Add("P1050", "P1050");
       dtm.ColumnMappings.Add("P1350", "P1350");
       dtm.ColumnMappings.Add("P1550", "P1550");
       dtm.ColumnMappings.Add("P1750", "P1750");
       dtm.ColumnMappings.Add("P004500", "P004500");
       dtm.ColumnMappings.Add("HD004500", "Hd004500");
       dtm.ColumnMappings.Add("P01500", "P01500");
       dtm.ColumnMappings.Add("HD01500", "Hd01500");
       dtm.ColumnMappings.Add("P0041000", "P0041000");
       dtm.ColumnMappings.Add("HD0041000", "Hd0041000");
       dtm.ColumnMappings.Add("P011000", "P011000");
       dtm.ColumnMappings.Add("HD01100", "Hd01100");
       dtm.ColumnMappings.Add("IUP1", "Iup1");
       dtm.ColumnMappings.Add("IUP2", "Iup2");
       dtm.ColumnMappings.Add("IUP3", "Iup3");
       dtm.ColumnMappings.Add("IDOWN1", "Idown1");
       dtm.ColumnMappings.Add("IDOWN2", "Idown2");
       dtm.ColumnMappings.Add("IDOWN3", "Idown3");
       dtm.ColumnMappings.Add("ADOUT", "Adout");
       dtm.ColumnMappings.Add("ADIN", "Adin");
       dtm.ColumnMappings.Add("GIB1", "Gib1");
       dtm.ColumnMappings.Add("GIB2", "Gib2");
       dtm.ColumnMappings.Add("IUP4", "Iup4");
       dtm.ColumnMappings.Add("IUP5", "Iup5");
       dtm.ColumnMappings.Add("IDOWN4", "Idown4");
       dtm.ColumnMappings.Add("IDOWN5", "Idown5");
       dtm.ColumnMappings.Add("P1550AP_PROD", "P1550ApProd");
       dtm.ColumnMappings.Add("B2500AP_PROD", "B2500ApProd");
       dtm.ColumnMappings.Add("P1350AP_POPER", "P1350ApPoper");
       dtm.ColumnMappings.Add("B2500AP_POPER", "B2500ApPoper");
       dtm.ColumnMappings.Add("P1550AP_POPER", "P1550ApPoper");
       dtm.ColumnMappings.Add("B40NI", "B40ni");
       dtm.ColumnMappings.Add("B60NI", "B60ni");
       dtm.ColumnMappings.Add("B120NI", "B120ni");
       dtm.ColumnMappings.Add("B1000NI", "B1000ni");
       dtm.ColumnMappings.Add("B1200NI", "B1200ni");
       dtm.ColumnMappings.Add("B2500NI", "B2500ni");
       dtm.ColumnMappings.Add("B500NI", "B500ni");
       dtm.ColumnMappings.Add("B5000NI", "B5000ni");
       dtm.ColumnMappings.Add("B10000NI", "B10000ni");
       dtm.ColumnMappings.Add("B24000NI", "B24000ni");
       dtm.ColumnMappings.Add("B30000NI", "B30000ni");
       dtm.ColumnMappings.Add("HCPOPERAONI", "HcPoperAoNi");
       dtm.ColumnMappings.Add("HCPOPERBONI", "HcPoperBoNi");
       dtm.ColumnMappings.Add("HCPRODAONI", "HcProdAoNi");
       dtm.ColumnMappings.Add("HCPRODBONI", "HcProdBoNi");
       dtm.ColumnMappings.Add("RAVGNI", "RavgNi");
       dtm.ColumnMappings.Add("STEELMARK_NI", "SteelMarkNi");
       dtm.ColumnMappings.Add("ADOUTR", "AdOutR");
       dtm.ColumnMappings.Add("ADINR", "AdInR");
       dtm.ColumnMappings.Add("ADOUTL", "AdOutL");
       dtm.ColumnMappings.Add("ADINL", "AdInL");
       dtm.ColumnMappings.Add("ADOUTSIMENS", "AdOutSimens");
       dtm.ColumnMappings.Add("ADINSIMENS", "AdInSimens");
      
       adapter.TableMappings.Add(dtm);

       //--Commands
       adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
       adapter.UpdateCommand = new OracleCommand {Connection = Odac.DbConnection};

       //Select Command
       adapter.SelectCommand.CommandText = "SELECT ID, SAMPLEID, UTYPE, MASSA, B3, B30, B100, B800, B2500, B5000, P1050, P1350, " +
                                        "P1550, P1750, P004500, HD004500, P01500, HD01500, P0041000, HD0041000, " +
                                        "P011000, HD01100, IUP1, IUP2, IUP3, IDOWN1, IDOWN2, IDOWN3,  ADOUT, ADIN, " +
                                        "GIB1, GIB2, IUP4, IUP5, IDOWN4, IDOWN5, P1550AP_PROD, B2500AP_PROD, P1350AP_POPER, B2500AP_POPER, P1550AP_POPER, " +
                                        "B40NI, B60NI, B120NI, B1000NI, B1200NI, B2500NI, B500NI, B5000NI, B10000NI, B24000NI, B30000NI, HCPOPERAONI, HCPOPERBONI, HCPRODAONI, HCPRODBONI, RAVGNI, STEELMARK_NI, " +
                                        "ADOUTR, ADINR, ADOUTL, ADINL, ADOUTSIMENS, ADINSIMENS " +
                                        "FROM LIMS.ML_MDATA " +
                                        "WHERE (SAMPLEID = :SMPLID) AND (UTYPE = :UTP)";
       adapter.SelectCommand.CommandType = CommandType.Text;

       var param = new OracleParameter
                     {
                       DbType = DbType.String,
                       Direction = ParameterDirection.Input,
                       IsNullable = false,
                       ParameterName = "SMPLID",
                       Size = 64,
                       SourceColumn = "SAMPLEID",
                       SourceColumnNullMapping = false,
                       SourceVersion = DataRowVersion.Current
                     };
       adapter.SelectCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Int32,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "UTP",
                   Size = 0,
                   SourceColumn = "UTYPE",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.SelectCommand.Parameters.Add(param);

       //Update Command
       adapter.UpdateCommand.CommandText = "UPDATE LIMS.ML_MDATA SET MASSA = :MASSA, B3 = :B3, B30 = :B30, B100 = :B100, B800 = :B800, B2500 = :B2500, B5000 = :B5000, P1050 = :P1050, P1350 = :P1350, " +
                                           "P1550 = :P1550, P1750 = :P1750, P004500 = :P004500, HD004500 = :HD004500, P01500 = :P01500, HD01500 = :HD01500, P0041000 = :P0041000, " +
                                           "HD0041000 = :HD0041000, P011000 = :P011000, HD01100 = :HD01100, IUP1 = :IUP1, IUP2 = :IUP2, IUP3 = :IUP3, IDOWN1 = :IDOWN1, " +
                                           "IDOWN2 = :IDOWN2, IDOWN3 = :IDOWN3, ADOUT = :ADOUT, ADIN = :ADIN, GIB1 = :GIB1, GIB2 = :GIB2, IUP4 = :IUP4, IUP5 = :IUP5, IDOWN4 = :IDOWN4, IDOWN5 = :IDOWN5, " +
                                           "P1550AP_PROD = :P1550AP_PROD, B2500AP_PROD = :B2500AP_PROD, P1350AP_POPER = :P1350AP_POPER, B2500AP_POPER = :B2500AP_POPER, P1550AP_POPER = :P1550AP_POPER, " +
                                           "B40NI = :B40NI, B60NI = :B60NI, B120NI = :B120NI, B1000NI = :B1000NI, B1200NI = :B1200NI, B2500NI = :B2500NI, B500NI = :B500NI, B5000NI = :B5000NI, B10000NI = :B10000NI, B24000NI = :B24000NI, B30000NI = :B30000NI, HCPOPERAONI = :HCPOPERAONI, HCPOPERBONI = :HCPOPERBONI, HCPRODAONI = :HCPRODAONI, HCPRODBONI = :HCPRODBONI, RAVGNI = :RAVGNI, STEELMARK_NI = :STEELMARK_NI, " +
                                           "ADOUTR = :ADOUTR, ADINR = :ADINR, ADOUTL = :ADOUTL, ADINL = :ADINL, ADOUTSIMENS = :ADOUTSIMENS, ADINSIMENS = :ADINSIMENS " +
                                           "WHERE (ID = :Original_ID)";
       adapter.UpdateCommand.CommandType = CommandType.Text;

       param = new OracleParameter
                 {
                   DbType = DbType.Int32,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "MASSA",
                   Size = 0,
                   SourceColumn = "MASSA",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B3",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B3",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B30",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B30",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B100",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B100",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B800",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B800",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B2500",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B2500",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B5000",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B5000",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P1050",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P1050",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P1350",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P1350",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P1550",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P1550",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P1750",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P1750",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P004500",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P004500",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "HD004500",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "HD004500",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P01500",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P01500",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "HD01500",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "HD01500",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P0041000",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P0041000",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "HD0041000",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "HD0041000",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P011000",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P011000",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "HD01100",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "HD01100",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IUP1",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IUP1",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IUP2",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IUP2",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IUP3",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IUP3",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IDOWN1",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IDOWN1",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IDOWN2",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IDOWN2",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IDOWN3",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IDOWN3",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.String,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "ADOUT",
                   Size = 2,
                   SourceColumn = "ADOUT",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.String,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "ADIN",
                   Size = 2,
                   SourceColumn = "ADIN",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Int32,
                   Direction = ParameterDirection.Input,
                   IsNullable = true,
                   ParameterName = "GIB1",
                   Size = 0,
                   SourceColumn = "GIB1",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Int32,
                   Direction = ParameterDirection.Input,
                   IsNullable = true,
                   ParameterName = "GIB2",
                   Size = 0,
                   SourceColumn = "GIB2",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IUP4",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IUP4",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IUP5",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IUP5",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IDOWN4",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IDOWN4",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "IDOWN5",
                   Size = 0,
                   Precision = 10,
                   Scale = 3,
                   SourceColumn = "IDOWN5",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P1550AP_PROD",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P1550AP_PROD",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B2500AP_PROD",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B2500AP_PROD",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "P1350AP_POPER",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "P1350AP_POPER",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
                 {
                   DbType = DbType.Decimal,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "B2500AP_POPER",
                   Size = 0,
                   Precision = 10,
                   Scale = 2,
                   SourceColumn = "B2500AP_POPER",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Current
                 };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "P1550AP_POPER",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "P1550AP_POPER",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B40NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B40NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B60NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B60NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B120NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B120NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B1000NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B1000NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B1200NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B1200NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B2500NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B2500NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B500NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B500NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B5000NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B5000NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B10000NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B10000NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B24000NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B24000NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "B30000NI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "B30000NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "HCPOPERAONI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "HCPOPERAONI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "HCPOPERBONI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "HCPOPERBONI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "HCPRODAONI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "HCPRODAONI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Decimal,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "HCPRODBONI",
         Size = 0,
         Precision = 10,
         Scale = 2,
         SourceColumn = "HCPRODBONI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.Int32,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "RAVGNI",
         SourceColumn = "RAVGNI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         Size = 20,
         ParameterName = "STEELMARK_NI",
         SourceColumn = "STEELMARK_NI",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "ADOUTR",
         Size = 2,
         SourceColumn = "ADOUTR",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "ADINR",
         Size = 2,
         SourceColumn = "ADINR",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "ADOUTL",
         Size = 2,
         SourceColumn = "ADOUTL",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "ADINL",
         Size = 2,
         SourceColumn = "ADINL",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "ADOUTSIMENS",
         Size = 3,
         SourceColumn = "ADOUTSIMENS",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);

       param = new OracleParameter
       {
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         IsNullable = false,
         ParameterName = "ADINSIMENS",
         Size = 3,
         SourceColumn = "ADINSIMENS",
         SourceColumnNullMapping = false,
         SourceVersion = DataRowVersion.Current
       };
       adapter.UpdateCommand.Parameters.Add(param);
      
       param = new OracleParameter
                 {
                   DbType = DbType.Int32,
                   Direction = ParameterDirection.Input,
                   IsNullable = false,
                   ParameterName = "Original_ID",
                   Size = 0,
                   SourceColumn = "ID",
                   SourceColumnNullMapping = false,
                   SourceVersion = DataRowVersion.Original
                 };
       adapter.UpdateCommand.Parameters.Add(param);
       //--------------------------------------------------------
       adapter.SelectCommand.ParameterCheck = false;
       adapter.UpdateCommand.ParameterCheck = false;
       //this.ColumnChanging += Column_Changing;

       foreach (OracleParameter prm in adapter.UpdateCommand.Parameters)
         prm.IsNullable = true;

       foreach (OracleParameter prm in adapter.SelectCommand.Parameters)
         prm.IsNullable = true;
     }

     private static void Column_Changing(object sender, DataColumnChangeEventArgs e)
     {
      /*
      if (e.ProposedValue == null) e.ProposedValue = DBNull.Value;

       //Это Эпштейн - делаем корректировку
       if (e.Column.Caption != "2") return;

       if ((e.Column.ColumnName == "P1750") && (e.ProposedValue != DBNull.Value))
         e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + p1750ApCorr;
       else if ((e.Column.ColumnName == "P1550") && (e.ProposedValue != DBNull.Value))
         e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + p1550ApCorr;
       else if ((e.Column.ColumnName == "B100") && (e.ProposedValue != DBNull.Value))
         e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + b100ApCorr;
       else if ((e.Column.ColumnName == "B800") && (e.ProposedValue != DBNull.Value))
         e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + b800ApCorr;
       else if ((e.Column.ColumnName == "B2500") && (e.ProposedValue != DBNull.Value))
         e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + b2500ApCorr;
         */
     }

     private decimal GetCorrVal(string NameVal)
     {
       const string sqlStmt = "SELECT CORR FROM LIMS.ML_VALDATA WHERE STEELTYPE = 'АН' AND MEASUREMENTTYPE = :MT";

       var lstPrm = new List<OracleParameter>();
       var prm = new OracleParameter()
       {
         ParameterName = "MT",
         DbType = DbType.String,
         Direction = ParameterDirection.Input,
         OracleDbType = OracleDbType.VarChar,
         Size = NameVal.Length,
         Value = NameVal
       };
       lstPrm.Add(prm);

       return Convert.ToDecimal(Odac.ExecuteScalar(sqlStmt, CommandType.Text, false, lstPrm));
     }


     public int LoadData(String SampleId, int UnitType)
     {
       var lstPrmValue = new List<Object> {SampleId, UnitType};
       return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
     }

     public int SaveData()
     {
       return Odac.SaveChangedData(this, adapter);
     }


   }

  
  public sealed class MlDataProbeDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter; 
   
    public MlDataProbeDataTable() : base()
    {
      //call base method DataTable
      this.TableName = "MlDataProbe";
      adapter = new OracleDataAdapter();       

      DataColumn col = null;
      col = new DataColumn("Id", typeof(string), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("Tstep", typeof(string), null, MappingType.Element) {AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("DenMat", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("KrvBeg", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("KrvEnd", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("VpKrvBeg", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("VpKrvEnd", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("OstNapr", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("CofStar", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("CofZapol", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("AnzInduc", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("TmpResist", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("OtnosUdl", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Tverd", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("TpokrUp", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("TpokrDown", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MgstrD", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MgstrDpp", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("KorForce", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("IsStat", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("CofAnzUmp", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("SteelMarkNiX", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("OsypFlag", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_MlDataProbe", new[] { this.Columns["Id"], this.Columns["Tstep"] }, true));
      
      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlDataProbe");
      dtm.ColumnMappings.Add("ID", "Id");
      dtm.ColumnMappings.Add("TSTEP", "Tstep");
      dtm.ColumnMappings.Add("DENMAT", "DenMat");
      dtm.ColumnMappings.Add("KRVBEG", "KrvBeg");
      dtm.ColumnMappings.Add("KRVEND", "KrvEnd");
      dtm.ColumnMappings.Add("VP_KRVBEG", "VpKrvBeg");
      dtm.ColumnMappings.Add("VP_KRVEND", "VpKrvEnd");
      dtm.ColumnMappings.Add("OSTNAPR", "OstNapr");
      dtm.ColumnMappings.Add("COFSTAR", "CofStar");
      dtm.ColumnMappings.Add("COFZAPOL", "CofZapol");
      dtm.ColumnMappings.Add("ANZINDUC", "AnzInduc");
      dtm.ColumnMappings.Add("TMPRESIST", "TmpResist");
      dtm.ColumnMappings.Add("OTNOSUDL", "OtnosUdl");
      dtm.ColumnMappings.Add("TVERD", "Tverd");
      dtm.ColumnMappings.Add("TPOKRUP", "TpokrUp");
      dtm.ColumnMappings.Add("TPOKRDOWN", "TpokrDown");
      dtm.ColumnMappings.Add("MGSTRD", "MgstrD");
      dtm.ColumnMappings.Add("MGSTRDPP", "MgstrDpp");
      dtm.ColumnMappings.Add("KORFORCE", "KorForce");
      dtm.ColumnMappings.Add("ISSTAT", "IsStat");
      dtm.ColumnMappings.Add("COFANZUMP", "CofAnzUmp");
      dtm.ColumnMappings.Add("STEELMARK_NI_X", "SteelMarkNiX");
      dtm.ColumnMappings.Add("OSYPFLAG", "OsypFlag");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
      adapter.UpdateCommand = new OracleCommand {Connection = Odac.DbConnection};

      //Select Command
      adapter.SelectCommand.CommandText = "SELECT ID, TSTEP, DENMAT, KRVBEG, KRVEND, VP_KRVBEG, VP_KRVEND, OSTNAPR, COFSTAR, COFZAPOL, " +
                                       "ANZINDUC, TMPRESIST, OTNOSUDL, TVERD, TPOKRUP, TPOKRDOWN, MGSTRD, MGSTRDPP, " +
                                       "KORFORCE, ISSTAT, COFANZUMP, STEELMARK_NI_X, OSYPFLAG " +
                                       "FROM LIMS.ML_MDATAP WHERE (ID = :PID) AND (TSTEP = :TS)";
      adapter.SelectCommand.CommandType = CommandType.Text;
      var param = new OracleParameter
                                {
                                  DbType = DbType.String,
                                  Direction = ParameterDirection.Input,
                                  IsNullable = false,
                                  ParameterName = "PID",
                                  Size = 40,
                                  SourceColumn = "ID",
                                  SourceColumnNullMapping = false,
                                  SourceVersion = DataRowVersion.Current
                                };
      adapter.SelectCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "TS",
                  Size = 20,
                  SourceColumn = "TSTEP",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.SelectCommand.Parameters.Add(param);

      //Update Command
      adapter.UpdateCommand.CommandText = "UPDATE LIMS.ML_MDATAP SET DENMAT = :DENMAT, KRVBEG = :KRVBEG, KRVEND = :KRVEND, " +
                                       "VP_KRVBEG = :VP_KRVBEG, VP_KRVEND = :VP_KRVEND, OSTNAPR = :OSTNAPR, " +
                                       "COFSTAR = :COFSTAR, COFZAPOL = :COFZAPOL, ANZINDUC = :ANZINDUC, " +
                                       "TMPRESIST = :TMPRESIST, OTNOSUDL = :OTNOSUDL, TVERD = :TVERD, " +
                                       "TPOKRUP = :TPOKRUP, TPOKRDOWN = :TPOKRDOWN, MGSTRD = :MGSTRD, MGSTRDPP = :MGSTRDPP, " +
                                       "KORFORCE = :KORFORCE, ISSTAT = :ISSTAT, COFANZUMP = :COFANZUMP, STEELMARK_NI_X = :STEELMARK_NI_X, OSYPFLAG = :OSYPFLAG " +
                                       "WHERE (ID = :Original_ID) AND (TSTEP = :Original_TSTEP)";
      adapter.UpdateCommand.CommandType = CommandType.Text;
      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "DENMAT",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "DENMAT",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "KRVBEG",
                  Size = 0,
                  SourceColumn = "KRVBEG",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "KRVEND",
                  Size = 0,
                  SourceColumn = "KRVEND",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "VP_KRVBEG",
                  Size = 1024,
                  SourceColumn = "VP_KRVBEG",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "VP_KRVEND",
                  Size = 1024,
                  SourceColumn = "VP_KRVEND",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = true,
                  ParameterName = "OSTNAPR",
                  Size = 0,
                  SourceColumn = "OSTNAPR",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = true,
                  ParameterName = "COFSTAR",
                  Size = 0,
                  SourceColumn = "COFSTAR",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = true,
                  ParameterName = "COFZAPOL",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "COFZAPOL",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "ANZINDUC",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "ANZINDUC",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "TMPRESIST",
                  Size = 0,
                  SourceColumn = "TMPRESIST",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "OTNOSUDL",
                  Size = 0,
                  SourceColumn = "OTNOSUDL",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "TVERD",
                  Size = 0,
                  SourceColumn = "TVERD",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "TPOKRUP",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "TPOKRUP",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "TPOKRDOWN",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "TPOKRDOWN",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "MGSTRD",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "MGSTRD",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "MGSTRDPP",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "MGSTRDPP",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Decimal,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "KORFORCE",
                  Size = 0,
                  Precision = 10,
                  Scale = 2,
                  SourceColumn = "KORFORCE",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "ISSTAT",
                  Size = 0,
                  SourceColumn = "ISSTAT",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.Int32,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "COFANZUMP",
                  Size = 0,
                  SourceColumn = "COFANZUMP",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Current
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "STEELMARK_NI_X",
        Size = 1024,
        SourceColumn = "STEELMARK_NI_X",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "OSYPFLAG",
        Size = 1,
        SourceColumn = "OSYPFLAG",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "Original_ID",
                  Size = 40,
                  SourceColumn = "ID",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Original
                };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
                {
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  IsNullable = false,
                  ParameterName = "Original_TSTEP",
                  Size = 20,
                  SourceColumn = "TSTEP",
                  SourceColumnNullMapping = false,
                  SourceVersion = DataRowVersion.Original
                };
      adapter.UpdateCommand.Parameters.Add(param);

      //------------------------------------------------------------
      adapter.SelectCommand.ParameterCheck = false;
      adapter.UpdateCommand.ParameterCheck = false;
      this.ColumnChanging += Column_Changing;

      foreach (OracleParameter prm in adapter.UpdateCommand.Parameters)
        prm.IsNullable = true;

      foreach (OracleParameter prm in adapter.SelectCommand.Parameters)
        prm.IsNullable = true;
    }

    private static void Column_Changing(object sender, DataColumnChangeEventArgs e)
    {
      if (e.ProposedValue == null)
        e.ProposedValue = DBNull.Value;
    }

    public int LoadData(String Id, String Tstep)
    {
      var lstPrmValue = new List<Object> {Id, Tstep};
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

    public int SaveData()
    {
      return Odac.SaveChangedData(this, adapter);
    }


  }
  
  
  public sealed class MlUsetDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter; 
   
    public MlUsetDataTable() : base()
    {
      //call base method DataTable
      this.TableName = "MlUset";
      adapter = new OracleDataAdapter();


      var col = new DataColumn("Utype", typeof(int), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("MeasurementType", typeof(string), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);
     
      col = new DataColumn("Ftag", typeof(int), null, MappingType.Element) {AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("IsSample", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_MlUset", new[] {this.Columns["Utype"], this.Columns["Ftag"]}, true));

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlUset");
      dtm.ColumnMappings.Add("UTYPE", "Utype");
      dtm.ColumnMappings.Add("MEASUREMENTTYPE", "MeasurementType");
      dtm.ColumnMappings.Add("FTAG", "Ftag");
      dtm.ColumnMappings.Add("IS_SAMPLE", "IsSample");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand
                                {
                                  Connection = Odac.DbConnection,
                                  CommandText = "SELECT UTYPE, MEASUREMENTTYPE, FTAG, IS_SAMPLE FROM LIMS.V_MLUSET WHERE (SAMPLEID = :SID)",
                                  CommandType = CommandType.Text
                                };

      //Select Command
      var param = new OracleParameter
                                {
                                  DbType = DbType.String,
                                  Direction = ParameterDirection.Input,
                                  IsNullable = false,
                                  ParameterName = "SID",
                                  Size = 64,
                                  SourceColumn = "SAMPLEID",
                                  SourceColumnNullMapping = false,
                                  SourceVersion = DataRowVersion.Current
                                };
      adapter.SelectCommand.Parameters.Add(param);
    }

    public int LoadData(String SampleId)
    {
      var lstPrmValue = new List<Object> {SampleId};
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }


  }
  
  public sealed class MlValDataDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter; 

    public MlValDataDataTable() : base()
    {
      //call base method DataTable
      this.TableName = "MlValData";
      adapter = new OracleDataAdapter();

      DataColumn col = null;
      col = new DataColumn("SteelType", typeof(string), null, MappingType.Element) {AllowDBNull = false};
      this.Columns.Add(col);
      
      col = new DataColumn("Ftag", typeof(int), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("Utype", typeof(int), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("IsValidate", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MinVal", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MaxVal", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_MlValData", new[] {
                           this.Columns["SteelType"],
                           this.Columns["Ftag"],
                           this.Columns["Utype"]}, true));


      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlValData");
      dtm.ColumnMappings.Add("STEELTYPE", "SteelType");
      dtm.ColumnMappings.Add("FTAG", "Ftag");
      dtm.ColumnMappings.Add("IS_VALIDATE", "IsValidate");
      dtm.ColumnMappings.Add("MIN_VALUE", "MinVal");
      dtm.ColumnMappings.Add("MAX_VALUE", "MaxVal");
      dtm.ColumnMappings.Add("UTYPE", "Utype");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand
                                {
                                  Connection = Odac.DbConnection,
                                  CommandText = "SELECT STEELTYPE, UTYPE, FTAG, IS_VALIDATE, MIN_VALUE, MAX_VALUE FROM LIMS.ML_VALDATA",
                                  CommandType = CommandType.Text
                                };

      //Select Command
    }

    public int LoadData()
    {
      return Odac.LoadDataTable(this, adapter, true, null);
    }

  }

  public sealed class MlUtypeInfoDataTable : DataTable
  {

    private readonly OracleDataAdapter adapter;

    public MlUtypeInfoDataTable()  : base()
    {
      //call base method DataTable
      this.TableName = "MlUtypeInfo";
      adapter = new OracleDataAdapter();

      System.Data.DataColumn col = null;
      col = new DataColumn("Utype", typeof(Int32), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("IsEnable", typeof(Int32), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("IsFull", typeof(Int32), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_MlUtypeInfo", new[] {this.Columns["Utype"]}, true));
      this.Columns["Utype"].Unique = true;


      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlUtypeInfo");
      dtm.ColumnMappings.Add("UTYPE", "Utype");
      dtm.ColumnMappings.Add("IS_ENABLE", "IsEnable");
      dtm.ColumnMappings.Add("IS_FULL", "IsFull");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand
                                {
                                  Connection = Odac.DbConnection,
                                  CommandText = "SELECT UTYPE, IS_ENABLE, IS_FULL FROM TABLE(LIMS.MagLab.GetUtypeInfo(:SID))",
                                  CommandType = CommandType.Text
                                };

      //Select Command
      var param = new OracleParameter
                    {
                      DbType = DbType.String,
                      Direction = ParameterDirection.Input,
                      IsNullable = false,
                      ParameterName = "SID",
                      Size = 64,
                      SourceColumn = "SAMPLEID",
                      SourceColumnNullMapping = false,
                      SourceVersion = DataRowVersion.Current
                    };
      adapter.SelectCommand.Parameters.Add(param);
    }

    public int LoadData(String SampleId)
    {
      var lstPrmValue = new List<Object> {SampleId};
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

  }

  public sealed class MlListApInfoDataTable : DataTable
  {

    private readonly OracleDataAdapter adapter;

    public MlListApInfoDataTable()
      : base()
    {
      //call base method DataTable
      this.TableName = "ListApInfo";
      adapter = new OracleDataAdapter();

      DataColumn col = null;
      col = new DataColumn("Name", typeof(String), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("B100", typeof(Decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("B800", typeof(Decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("B2500", typeof(Decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("P1550", typeof(Decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("P1750", typeof(Decimal), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_ListApInfo", new[] {this.Columns["Name"]}, true));
      this.Columns["Name"].Unique = true;


      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlUtypeInfo");
      dtm.ColumnMappings.Add("NAME", "Name");
      dtm.ColumnMappings.Add("B100", "B100");
      dtm.ColumnMappings.Add("B800", "B800");
      dtm.ColumnMappings.Add("B2500", "B2500");
      dtm.ColumnMappings.Add("P1550", "P1550");
      dtm.ColumnMappings.Add("P1750", "P1750");
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand
                                {
                                  Connection = Odac.DbConnection,
                                  CommandText =
                                    "SELECT CASE WHEN UTYPE = 1 THEN 'Лист'  WHEN UTYPE = 2 THEN 'Эпшт' END NAME, " +
                                    "B100, B800, B2500, P1550, P1750 " +
                                    "FROM LIMS.ML_MDATA " +
                                    "WHERE SAMPLEID = :SID " +
                                    "AND (UTYPE IN (1,2)) " +
                                    "ORDER BY UTYPE",
                                  CommandType = CommandType.Text
                                };

      //Select Command

      var param = new OracleParameter
                    {
                      DbType = DbType.String,
                      Direction = ParameterDirection.Input,
                      IsNullable = false,
                      ParameterName = "SID",
                      Size = 64,
                      SourceColumn = "SAMPLEID",
                      SourceColumnNullMapping = false,
                      SourceVersion = DataRowVersion.Current
                    };
      adapter.SelectCommand.Parameters.Add(param);
    }

    public int LoadData(String SampleId)
    {
      var lstPrmValue = new List<Object> {SampleId};
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

  }


  public sealed class MlMatGnlDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlMatGnlDataTable() : base()
    {
      this.TableName = "MlMatGnl";
      adapter = new OracleDataAdapter();

      DataColumn col = null;

      col = new DataColumn("MatId", typeof(string), null, MappingType.Element) {AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("ParentMatId", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_MlMatGnl", new[] { this.Columns["MatId"] }, true));
      this.Columns["MatId"].Unique = true;

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlMatGnl");
      dtm.ColumnMappings.Add("MLOCID", "MatId");
      dtm.ColumnMappings.Add("MLOCIDP", "ParentMatId");
      adapter.TableMappings.Add(dtm);

      adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
    }

    public int LoadData(String MatLocId)
    {
      const string SqlStmt = "SELECT BEZEICHNUNGBASE MLOCID /*Рулон*/, BEZEICHNUNG MLOCIDP /*рулон родитель*/ " +
                             "FROM VIZ.MATGENEALOGY_VD " +
                             "WHERE (SYS_KZ = 'A') " +
                             "START WITH BEZEICHNUNG = :MID " +
                             "CONNECT BY BEZEICHNUNG = PRIOR BEZEICHNUNGBASE";

      adapter.SelectCommand.Parameters.Clear();
      adapter.SelectCommand.CommandText = SqlStmt;
      adapter.SelectCommand.CommandType = CommandType.Text;

      var prm = new OracleParameter
                  {
                    DbType = DbType.String,
                    Direction = ParameterDirection.Input,
                    OracleDbType = OracleDbType.VarChar,
                    ParameterName = "MID",
                    Size = MatLocId.Length
                  };
      adapter.SelectCommand.Parameters.Add(prm);

      var lstPrmValue = new List<Object> {MatLocId};
      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }


  }

  public sealed class MlMk4auDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlMk4auDataTable() : base()
    {
      this.TableName = "MlMk4au";
      adapter = new OracleDataAdapter();

      var col = new DataColumn("Lsimple", typeof(decimal), null, MappingType.Element){AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("Wsimple", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Density", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlMk4au");
      dtm.ColumnMappings.Add("LSIMPLE", "Lsimple");
      dtm.ColumnMappings.Add("WSIMPLE", "Wsimple");
      dtm.ColumnMappings.Add("DENSITY", "Density");
      adapter.TableMappings.Add(dtm);

      adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
    }

    public int LoadData(String SteelType, decimal ThickNom, int Utp)
    {
      const string SqlStmt = "SELECT LSIMPLE, WSIMPLE, DENSITY FROM LIMS.ML_MK4AU " +
                             "WHERE (STEELTYPE = :STTYPE) AND (THICKNESSNOMINAL = :TN) AND (UTYPE = :UTP)";

      adapter.SelectCommand.Parameters.Clear();
      adapter.SelectCommand.CommandText = SqlStmt;
      adapter.SelectCommand.CommandType = CommandType.Text;

      var lstPrmValue = new List<Object>();
      var prm = new OracleParameter
                  {
                    DbType = DbType.String,
                    Direction = ParameterDirection.Input,
                    OracleDbType = OracleDbType.VarChar,
                    ParameterName = "STTYPE",
                    Size = SteelType.Length
                  };
      adapter.SelectCommand.Parameters.Add(prm);
      lstPrmValue.Add(SteelType);


      prm = new OracleParameter
              {
                DbType = DbType.Decimal,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Number,
                ParameterName = "TN",
                Precision = 5,
                Scale = 2
              };
      adapter.SelectCommand.Parameters.Add(prm);
      lstPrmValue.Add(ThickNom);


      prm = new OracleParameter
              {
                DbType = DbType.Int32,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Integer,
                ParameterName = "UTP"
              };
      adapter.SelectCommand.Parameters.Add(prm);
      lstPrmValue.Add(Utp);


      return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
    }

  }


  public sealed class MlMk4apDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlMk4apDataTable() : base()
    {
      this.TableName = "MlMk4ap";
      adapter = new OracleDataAdapter();

      DataColumn col = null;

      col = new DataColumn("Ftag", typeof(int), null, MappingType.Element) {AllowDBNull = false};
      this.Columns.Add(col);

      col = new DataColumn("MeasMl", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MeasP", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MeasName", typeof(string), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("TypIzm", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("ValIzm", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("OutVal", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("IsActive", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("IsValidate", typeof(int), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MinValue", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("MaxValue", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("Corr", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);


      this.Constraints.Add(new UniqueConstraint("Pk_MlMk4ap", new[] { this.Columns["Ftag"] }, true));
      this.Columns["Ftag"].Unique = true;

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("SourceTable1", "MlMk4ap");
      dtm.ColumnMappings.Add("FTAG", "Ftag");
      dtm.ColumnMappings.Add("MEASUREMENTTYPE_ML", "MeasMl");
      dtm.ColumnMappings.Add("MEASUREMENTTYPE_P", "MeasP");
      dtm.ColumnMappings.Add("PNAME", "MeasName");
      dtm.ColumnMappings.Add("TYPIZM", "TypIzm");
      dtm.ColumnMappings.Add("VALIZM", "ValIzm");
      dtm.ColumnMappings.Add("OUTVAL", "OutVal");
      dtm.ColumnMappings.Add("IS_ACTIVE", "IsActive");
      dtm.ColumnMappings.Add("IS_VALIDATE", "IsValidate");
      dtm.ColumnMappings.Add("MIN_VALUE", "MinValue");
      dtm.ColumnMappings.Add("MAX_VALUE", "MaxValue");
      dtm.ColumnMappings.Add("CORR", "Corr");

      adapter.TableMappings.Add(dtm);
      adapter.SelectCommand = new OracleCommand {Connection = Odac.DbConnection};
    }

    public int LoadData(int Utp, int[] Ftag)
    {
      this.Clear();
      const string SqlStmt = "LIMS.MAGLAB.GetParamMk4a";

      adapter.SelectCommand.Parameters.Clear();
      adapter.SelectCommand.CommandText = SqlStmt;
      adapter.SelectCommand.CommandType = CommandType.StoredProcedure;

      var prm = new OracleParameter
                  {
                    DbType = DbType.Int32,
                    Direction = ParameterDirection.Input,
                    OracleDbType = OracleDbType.Integer,
                    ParameterName = "UTP",
                    Value = Utp
                  };
      adapter.SelectCommand.Parameters.Add(prm);

      prm = new OracleParameter
              {
                DbType = DbType.Int32,
                Direction = ParameterDirection.Input,
                OracleDbType = OracleDbType.Integer,
                ParameterName = "FTAG",
                ArrayLength = Ftag.Length,
                Value = Ftag
              };
      adapter.SelectCommand.Parameters.Add(prm);

      prm = new OracleParameter
              {
                Direction = ParameterDirection.ReturnValue, 
                OracleDbType = OracleDbType.Cursor
              };
      adapter.SelectCommand.Parameters.Add(prm);

      if (adapter.SelectCommand.Connection.State == ConnectionState.Closed)
        adapter.SelectCommand.Connection.Open();
  
      adapter.SelectCommand.ExecuteNonQuery();
      var oraCursor = (OracleCursor)prm.Value;
      adapter.Fill(this, oraCursor);
      oraCursor.Dispose();
 
      return 0;
    }

  }


  public sealed class MlStZapDataTable : SmvDataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlStZapDataTable(string tblName)
    {
      TableName = tblName;
      adapter = new OracleDataAdapter();

      var col = new DataColumn("SmpYear", typeof(Int32), null, MappingType.Element){AllowDBNull = false};
      Columns.Add(col);

      col = new DataColumn("Qart", typeof(Int32), null, MappingType.Element) { AllowDBNull = false };
      Columns.Add(col);

      col = new DataColumn("Thicknessnominal", typeof(decimal), null, MappingType.Element) { AllowDBNull = false };
      Columns.Add(col);

      col = new DataColumn("CofStar", typeof(Int32), null, MappingType.Element);
      Columns.Add(col);

      col = new DataColumn("CofZapol", typeof(decimal), null, MappingType.Element);
      Columns.Add(col);

      Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { Columns["SmpYear"], Columns["Qart"], Columns["Thicknessnominal"] }, true));
      //Columns["Id"].Unique = true;

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("V_STZAP", tblName);
      dtm.ColumnMappings.Add("SMP_YEAR", "SmpYear");
      dtm.ColumnMappings.Add("QART", "Qart");
      dtm.ColumnMappings.Add("THICKNESSNOMINAL", "Thicknessnominal");
      dtm.ColumnMappings.Add("COFSTAR", "CofStar");
      dtm.ColumnMappings.Add("COFZAPOL", "CofZapol");
      adapter.TableMappings.Add(dtm);

      //Select Command
      adapter.SelectCommand = new OracleCommand
      {
        Connection = Odac.DbConnection,
        CommandType = CommandType.Text,
        UpdatedRowSource = UpdateRowSource.None
      };

      //Update Command
    }
    public int LoadData()
    {
      adapter.SelectCommand.CommandText =
        "SELECT SMP_YEAR, QART, THICKNESSNOMINAL, COFSTAR, COFZAPOL " +
        "FROM LIMS.V_STZAP ORDER BY 1 DESC, 2 DESC, 3 ASC";

      adapter.SelectCommand.Parameters.Clear();
      return Odac.LoadDataTable(this, adapter, true, null);
    }

  }

  public sealed class MlMesurCofDataTable : DataTable
  {
    private readonly OracleDataAdapter adapter;

    public MlMesurCofDataTable() : base()
    {
      //call base method DataTable
      this.TableName = "MlMesurCof";
      adapter = new OracleDataAdapter();

      DataColumn col = null;
      col = new DataColumn("Md", typeof(string), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("MeasurementTypeMl", typeof(string), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("Utype", typeof(int), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("MesDevice", typeof(int), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("Corr", typeof(decimal), null, MappingType.Element);
      this.Columns.Add(col);

      col = new DataColumn("TypCor", typeof(string), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      col = new DataColumn("MlComment", typeof(string), null, MappingType.Element) { AllowDBNull = false };
      this.Columns.Add(col);

      this.Constraints.Add(new UniqueConstraint("Pk_MlMesurCof", new[] {
                           this.Columns["Md"],
                           this.Columns["MeasurementTypeMl"],
                           this.Columns["Utype"],
                           this.Columns["MesDevice"]}, true));

      adapter.TableMappings.Clear();
      var dtm = new System.Data.Common.DataTableMapping("ML_MESURCF", "MlMesurCof");
      dtm.ColumnMappings.Add("MD", "Md");
      dtm.ColumnMappings.Add("MEASUREMENTTYPE_ML", "MeasurementTypeMl");
      dtm.ColumnMappings.Add("UTYPE", "Utype");
      dtm.ColumnMappings.Add("MES_DEVICE", "MesDevice");
      dtm.ColumnMappings.Add("CORR", "Corr");
      dtm.ColumnMappings.Add("TYP_COR", "TypCor");
      dtm.ColumnMappings.Add("ML_COMMENT", "MlComment");
      
      adapter.TableMappings.Add(dtm);

      //--Commands
      adapter.SelectCommand = new OracleCommand
      {
        Connection = Odac.DbConnection,
        CommandText = "SELECT MD, MEASUREMENTTYPE_ML, UTYPE, MES_DEVICE, CORR, TYP_COR, ML_COMMENT FROM LIMS.ML_MESURCF",
        CommandType = CommandType.Text
      };

      adapter.UpdateCommand = new OracleCommand
      {
        Connection = Odac.DbConnection,
        CommandType = CommandType.Text
      };

      //Select Command
      //Update Command
      adapter.UpdateCommand.CommandText = "UPDATE LIMS.ML_MESURCF SET CORR = :PCORR WHERE (MD = :Original_MD) AND (MEASUREMENTTYPE_ML = :Original_MEASUREMENTTYPE_ML) AND (UTYPE = :Original_UTYPE) AND (MES_DEVICE = :Original_MES_DEVICE)";
      adapter.UpdateCommand.CommandType = CommandType.Text;
      var param = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "PCORR",
        Precision = 10,
        Scale = 2,
        SourceColumn = "CORR",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Current
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "Original_MD",
        SourceColumn = "MD",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Original
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "Original_MEASUREMENTTYPE_ML",
        SourceColumn = "MEASUREMENTTYPE_ML",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Original
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "Original_UTYPE",
        SourceColumn = "UTYPE",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Original
      };
      adapter.UpdateCommand.Parameters.Add(param);

      param = new OracleParameter
      {
        DbType = DbType.Int32,
        Direction = ParameterDirection.Input,
        IsNullable = false,
        ParameterName = "Original_MES_DEVICE",
        SourceColumn = "MES_DEVICE",
        SourceColumnNullMapping = false,
        SourceVersion = DataRowVersion.Original
      };
      adapter.UpdateCommand.Parameters.Add(param);

      adapter.SelectCommand.ParameterCheck = false;
      adapter.UpdateCommand.ParameterCheck = false;

      foreach (OracleParameter prm in adapter.UpdateCommand.Parameters)
        prm.IsNullable = true;

      foreach (OracleParameter prm in adapter.SelectCommand.Parameters)
        prm.IsNullable = true;

    }

    public int LoadData()
    {
      return Odac.LoadDataTable(this, adapter, true, null);
    }
    public int SaveData()
    {
      return Odac.SaveChangedData(this, adapter);
    }


  }





}