using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Text;
using System.Windows;
using Devart.Data.Oracle;
using System.IO;
using Smv.Data.Oracle;
using Smv.Utils;


namespace Viz.WrkModule.Thp.Db
{
  public static class Lob
  {
    public static Boolean UploadBlob(string fileName, string fieldNameBlob, string fieldNameTypeDoc, Int64 iD)
    {
      Boolean res = false;
      FileStream fs = null;
      BinaryReader r = null;
     

      try{
        string extFile = Path.GetExtension(fileName).ToUpper().Replace(".","");

        fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        r = new BinaryReader(fs);
        Odac.DbConnection.Open();
        OracleLob myLob = new OracleLob(Odac.DbConnection, OracleDbType.Blob);
        int streamLength = (int)fs.Length;
        myLob.Write(r.ReadBytes(streamLength), 0, streamLength);
        OracleCommand myCommand = new OracleCommand("UPDATE LIMS.THP_DATA SET " + fieldNameBlob + " = :BLB, " + fieldNameTypeDoc + " = :TDOC WHERE ID = :PID", Odac.DbConnection);
        OracleParameter paramBlob = myCommand.Parameters.Add("BLB", OracleDbType.Blob);
        OracleParameter paramTypeDoc = myCommand.Parameters.Add("TDOC", OracleDbType.VarChar);
        OracleParameter paramId = myCommand.Parameters.Add("PID", OracleDbType.Number);
        paramId.Value = iD;
        paramTypeDoc.Value = extFile; 
        paramBlob.OracleValue = myLob;
      
        myCommand.ExecuteNonQuery();
        res = true;
      }
      catch (Exception ex){
        DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Error);
      }
      finally{
        Odac.DbConnection.Close();
        r.Close();
        fs.Close();
      }
      return res;
    }

    public static Boolean DownloadBlob(string fileName, string fieldNameBlob, Int64 iD)
    {
      Boolean res = false;
      OracleCommand myCommand = new OracleCommand("SELECT " + fieldNameBlob + " FROM LIMS.THP_DATA WHERE ID = :PID", Odac.DbConnection);
      OracleParameter paramId = myCommand.Parameters.Add("PID", OracleDbType.Number);
      paramId.Value = iD;

      Odac.DbConnection.Open();
      OracleDataReader myReader = myCommand.ExecuteReader(CommandBehavior.Default);
      try{
        while (myReader.Read()){
          OracleLob myLob = myReader.GetOracleLob(myReader.GetOrdinal(fieldNameBlob));

          if (!myLob.IsNull){
            FileStream fs = new FileStream(fileName, FileMode.Create);
            BinaryWriter w = new BinaryWriter(fs);
            w.Write((byte[])myLob.Value);
            w.Close();
            fs.Close();
          }
        }
        res = true; 
      }
      catch (Exception ex){
        DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Error);
      }
      finally{
        myReader.Close(); 
        //xxx-xxx TFS TEST 
        Odac.DbConnection.Close();
      }
      return res;
    } 







  }

}
