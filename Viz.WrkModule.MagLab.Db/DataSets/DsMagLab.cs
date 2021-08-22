using System;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.MagLab.Db.DataSets {


  public partial class DsMagLab 
  {

    public partial class MlSamplesDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit(); 
        this.FetchAll = true;
        this.Connection = Odac.DbConnection;
      }

      public int GetListSimple(DateTime DateStart, DateTime DateEnd)
      {

        string SqlStmt = "SELECT SAMPLEID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, MATLOCALNUMBER, MATMARKINGINFO, DTSAMPLE, STATE, ERR_TEXT,  SAMPLENUM " +
                         "FROM LIMS.V_SAMPLEMEAS WHERE (DTSAMPLE BETWEEN :DT1 AND :DT2)";
        this.SelectCommand.Parameters.Clear();
        this.SelectCommand.CommandText = SqlStmt;
        this.SelectCommand.CommandType = System.Data.CommandType.Text;

        OracleParameter prm = new OracleParameter();
        prm.DbType = System.Data.DbType.DateTime;
        prm.Direction = System.Data.ParameterDirection.Input;
        prm.OracleDbType = OracleDbType.Date;
        prm.ParameterName = "DT1";
        this.SelectCommand.Parameters.Add(prm);

        prm = new OracleParameter();
        prm.DbType = System.Data.DbType.DateTime;
        prm.Direction = System.Data.ParameterDirection.Input;
        prm.OracleDbType = OracleDbType.Date;
        prm.ParameterName = "DT2";
        this.SelectCommand.Parameters.Add(prm);

        List<Object> lstPrmValue = new List<Object>();
        lstPrmValue.Add(DateStart.AddHours(-6).AddHours(8));
        lstPrmValue.Add(DateEnd.AddHours(-6).AddHours(7).AddMinutes(59).AddSeconds(59));
        return Odac.LoadDataTable(this, true, lstPrmValue);
      }

      public int SerchBySampleId(String SampleId)
      {
        string SqlStmt = "SELECT SAMPLEID, TESTTYPE, STEELTYPE, THICKNESSNOMINAL, LINE, LASERFLAG, SAMPLEPOS, MATLOCALNUMBER, MATMARKINGINFO, DTSAMPLE, STATE, ERR_TEXT,  SAMPLENUM " +  
                         "FROM LIMS.V_SAMPLEMEAS WHERE SAMPLENUM LIKE :SMPID ";
        this.SelectCommand.Parameters.Clear();
        this.SelectCommand.CommandText = SqlStmt;
        this.SelectCommand.CommandType = System.Data.CommandType.Text;
                          
        OracleParameter prm = new OracleParameter();
        prm.DbType = System.Data.DbType.String;
        prm.Direction = System.Data.ParameterDirection.Input;
        prm.OracleDbType = OracleDbType.VarChar;
        prm.Size = 20;
        prm.ParameterName = "SMPID";
        this.SelectCommand.Parameters.Add(prm);
        //prm.Value = SampleId + "%"; 
        
        List<Object> lstPrmValue = new List<Object>();
        lstPrmValue.Add(SampleId + "%");
        return Odac.LoadDataTable(this, true, lstPrmValue);
      }

      public int SerchByMatLocalNum(String MatLocalNum)
      {
        List<Object> lstPrmValue = new List<Object>();
        DateTime DateStart = DateTime.Now.AddDays(-10000);
        DateTime DateEnd = DateTime.Now.AddDays(10000);
        lstPrmValue.Add(DateStart);
        lstPrmValue.Add(DateEnd);
        lstPrmValue.Add("Z");
        lstPrmValue.Add("Z");
        lstPrmValue.Add(MatLocalNum);
        return Odac.LoadDataTable(this, true, lstPrmValue);
      }

      public int SerchByMatMarkNum(String MatMarkNum)
      {
        List<Object> lstPrmValue = new List<Object>();
        DateTime DateStart = DateTime.Now.AddDays(-10000);
        DateTime DateEnd = DateTime.Now.AddDays(10000);
        lstPrmValue.Add(DateStart);
        lstPrmValue.Add(DateEnd);
        lstPrmValue.Add("Z");
        lstPrmValue.Add(MatMarkNum);
        lstPrmValue.Add("Z");
        return Odac.LoadDataTable(this, true, lstPrmValue);
      }


    }

    public partial class FindModeDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit();
        System.Data.DataRow row = this.NewRow();
        row[0] = 1;
        row[1] = "По № образца";
        row[2] = 0.1;
        this.Rows.Add(row);

        row = this.NewRow();
        row[0] = 2;
        row[1] = "По лок. № материала";
        row[2] = 0.2;
        this.Rows.Add(row);

        row = this.NewRow();
        row[0] = 3;
        row[1] = "По маркировке";
        row[2] = 0.3;
        this.Rows.Add(row);

        this.AcceptChanges();
      }
    }

    public partial class MlDataDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit();
        this.FetchAll = true;
        this.Connection = Odac.DbConnection;
        this.SelectCommand.ParameterCheck = false;
        this.UpdateCommand.ParameterCheck = false;
        this.ColumnChanging += Column_Changing; 

        foreach(OracleParameter prm in this.UpdateCommand.Parameters)
          prm.IsNullable = true;

        foreach (OracleParameter prm in this.SelectCommand.Parameters)
          prm.IsNullable = true;

        
      }

      private static void Column_Changing(object sender, System.Data.DataColumnChangeEventArgs e)
      {
        if (e.ProposedValue == null) 
          e.ProposedValue = DBNull.Value;  
      }

      public int LoadData(String SampleId, int UnitType)
      {
        List<Object> lstPrmValue = new List<Object>();
        lstPrmValue.Add(SampleId);
        lstPrmValue.Add(UnitType);
        return Odac.LoadDataTable(this, true, lstPrmValue);    
      }

      public int SaveData()
      {
        return Odac.SaveChangedData(this);
      }

    }

    public partial class MlUsetDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit();
        this.FetchAll = true;
        this.Connection = Odac.DbConnection;
      }

      public int LoadData(String SampleId)
      {
        List<Object> lstPrmValue = new List<Object>();
        lstPrmValue.Add(SampleId);
        return Odac.LoadDataTable(this, true, lstPrmValue);      
      }


    }

    public partial class MlValDataDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit();
        this.FetchAll = true;
        this.Connection = Odac.DbConnection;
      }

      public int LoadData()
      {
        return Odac.LoadDataTable(this, true, null);
      }


    }




    public partial class MlDataProbeDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit();
        this.FetchAll = true;
        this.Connection = Odac.DbConnection;
        this.SelectCommand.ParameterCheck = false;
        this.UpdateCommand.ParameterCheck = false;
        this.ColumnChanging += Column_Changing;

        foreach (OracleParameter prm in this.UpdateCommand.Parameters)
          prm.IsNullable = true;

        foreach (OracleParameter prm in this.SelectCommand.Parameters)
          prm.IsNullable = true;


      }

      private static void Column_Changing(object sender, System.Data.DataColumnChangeEventArgs e)
      {
        if (e.ProposedValue == null)
          e.ProposedValue = DBNull.Value;
      }

      public int LoadData(String Id)
      {
        List<Object> lstPrmValue = new List<Object>();
        lstPrmValue.Add(Id);
        return Odac.LoadDataTable(this, true, lstPrmValue);
      }

      public int SaveData()
      {
        return Odac.SaveChangedData(this);
      }


    }






  }

 
}