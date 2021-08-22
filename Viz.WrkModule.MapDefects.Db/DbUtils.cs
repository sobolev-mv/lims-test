using System;
using System.Collections.Generic;
using System.Data;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.WrkModule.MapDefects.Db
{
  public static class MapDefectsAction
  {

    public static void DeleteDefectsData(Int64 zDn)
    {
      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "PZDN",
        DbType = DbType.Int64,
        OracleDbType = OracleDbType.Number,
        Direction = ParameterDirection.Input,
        Value = zDn
      };
      lstPrm.Add(prm);
      Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN)", CommandType.Text, false, lstPrm);

    }

    public static void CreateDefectsData(Int64 zDn, string matLocId, Boolean isStrann)
    {
      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        DbType = DbType.Int64,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Value = zDn
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocId.Length,
        Value = matLocId
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = 1,
        Value = '1'
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery(isStrann ? "VIZ_PRN.Razdel_16.CreateDefTable" : "VIZ_PRN.Razdel_16.CreateDefTablePack", CommandType.StoredProcedure, false, lstPrm);     
    }

    public static void CreateDefectsDataNlmk(Int64 zDn, Int64 nlmkCoilId, decimal coilWgt)
    {
      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        DbType = DbType.Int64,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Value = zDn
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.Int64,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Value = nlmkCoilId
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Precision = 10,
        Scale = 3,
        Value = coilWgt
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery("VIZ_PRN.Razdel_16.CreateDefTableNlmkProd", CommandType.StoredProcedure, false, lstPrm);
    }


    public static Boolean IsMatLocked(string matLocId)
    {
      const string stmt = "SELECT COUNT(*) FROM VIZ_PRN.Z_MATLOC WHERE (LOCID = :PLOCID)";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "PLOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocId.Length,
        Value = matLocId
      };
      lstPrm.Add(prm);

      return (Convert.ToInt32(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm)) > 0);
    }



    public static string GetRealLocNumStrann(string matLocId)
    {
      const string stmt = "SELECT MATBEZEICHNUNGOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGINPUT = :LOCID) and MATBEZEICHNUNGOUTPUT NOT LIKE ('%/OTM%') AND (AGTYP = 'STRANN')";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocId.Length,
        Value = matLocId
      };
      lstPrm.Add(prm);
    
      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static decimal GetCoilWgt(string matLocId, string techStep)
    {
      decimal GetWgt(string stmt, string locId, string step)
      {
        var lstPrm = new List<OracleParameter>();

        var prm = new OracleParameter
        {
          ParameterName = "LOCID",
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = locId.Length,
          Value = locId
        };
        lstPrm.Add(prm);

        if (!String.IsNullOrEmpty(step)){
          prm = new OracleParameter
          {
            ParameterName = "TS",
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = step.Length,
            Value = step
          };

          lstPrm.Add(prm);
        }

        return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
      }

      var stmtSql = "SELECT GEWOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = :TS)";
      var rez = GetWgt(stmtSql, matLocId, techStep);

      if (rez == 0){
        stmtSql = "SELECT N.WEIGHTBASE FROM VIZ.MG_NODE N WHERE N.NODEID = (" +
                  "SELECT MIN(NODEID) FROM VIZ.MG_NODE WHERE MESMATNAME = :LOCID AND MESTAG IS NOT NULL)";

        rez = GetWgt(stmtSql, matLocId, null);
      }

      return rez;
    }

    public static decimal GetCoilWgtUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT GEWOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static decimal GetCoilWgtAndLengthNlmkPack(Int64 nlmkCoilId, string fieldName)
    {
      //decimal rez;
      var stmt = "select " + fieldName + " from CCM.V_LIMPS_TOPO@CCM2_NLMK where em_shipped_id = :ID";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "ID",
        DbType = DbType.Int64,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.Number,
        Value = nlmkCoilId
      };
      lstPrm.Add(prm);

      return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }


    public static decimal GetCoilWidth(string matLocId, string techStep)
    {
      decimal GetWidth(string stmt, string locId, string step)
      {
        var lstPrm = new List<OracleParameter>();

        var prm = new OracleParameter
        {
          ParameterName = "LOCID",
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = locId.Length,
          Value = locId
        };
        lstPrm.Add(prm);

        if (!string.IsNullOrEmpty(step)){
          prm = new OracleParameter
          {
            ParameterName = "TS",
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = step.Length,
            Value = step
          };

          lstPrm.Add(prm);
        }

        return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
      }

      var stmtSql = "SELECT BREITEOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = :TS)";
      var rez = GetWidth(stmtSql, matLocId, techStep);

      if (rez == 0){
        stmtSql = "SELECT N.WIDTHBASE FROM VIZ.MG_NODE N WHERE N.NODEID = (" +
                  "SELECT MIN(NODEID) FROM VIZ.MG_NODE WHERE MESMATNAME = :LOCID AND MESTAG IS NOT NULL)";

        rez = GetWidth(stmtSql, matLocId, null);
      }

      return rez;
    }

    public static decimal GetCoilWidthUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT BREITEOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetStrannDateTimeCoil(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT TO_CHAR(VIZ_PRN.VAR_RPT.toLocalTimeZone(AENDDATUM),'DD.MM.YYYY HH24:MI:SS') || '/' || ANLAGE  FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = 'STRANN')";      

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
                  ParameterName = "LOCID",
                  DbType = DbType.String,
                  Direction = ParameterDirection.Input,
                  OracleDbType = OracleDbType.VarChar,
                  Size = MatLocId.Length,
                  Value = MatLocId
                };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetDateTimeCoilUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT TO_CHAR(VIZ_PRN.VAR_RPT.toLocalTimeZone(AENDDATUM),'DD.MM.YYYY HH24:MI:SS') || '/' || ANLAGE  FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetStendCoil(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT ANNEALINGLOT || '/' || ANNEALINGLOTSEQNO FROM VIZ.MAT WHERE (BEZEICHNUNG = :LOCID)";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static decimal GetTolsCoil(string matLocId, string techStep)
    {
      decimal GetTols(string stmt, string locId, string step)
      {
        var lstPrm = new List<OracleParameter>();

        var prm = new OracleParameter
        {
          ParameterName = "LOCID",
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = locId.Length,
          Value = locId
        };
        lstPrm.Add(prm);

        if (!string.IsNullOrEmpty(step))
        {
          prm = new OracleParameter
          {
            ParameterName = "TS",
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = step.Length,
            Value = step
          };

          lstPrm.Add(prm);
        }

        return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
      }

      var stmtSql = "SELECT DICKEOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = :TS)";
      var rez = GetTols(stmtSql, matLocId, techStep);

      if (rez == 0){
        stmtSql = "SELECT N.THICKNESSBASE FROM VIZ.MG_NODE N WHERE N.NODEID = (" +
                  "SELECT MIN(NODEID) FROM VIZ.MG_NODE WHERE MESMATNAME = :LOCID AND MESTAG IS NOT NULL)";

        rez = GetTols(stmtSql, matLocId, null);
      }

      return rez;
    }

    public static decimal GetTolsCoilUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT DICKEOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static decimal GetLenCoil(string matLocId, string techStep)
    {
      decimal GetLen(string stmt, string locId, string step)
      {
        var lstPrm = new List<OracleParameter>();

        var prm = new OracleParameter
        {
          ParameterName = "LOCID",
          DbType = DbType.String,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.VarChar,
          Size = locId.Length,
          Value = locId
        };
        lstPrm.Add(prm);

        if (!string.IsNullOrEmpty(step))
        {
          prm = new OracleParameter
          {
            ParameterName = "TS",
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = step.Length,
            Value = step
          };

          lstPrm.Add(prm);
        }

        return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
      }

      var stmtSql = "SELECT LAENGEOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = :TS)";
      var rez = GetLen(stmtSql, matLocId, techStep);

      if (rez == 0){
        stmtSql = "SELECT N.LENGTHBASE FROM VIZ.MG_NODE N WHERE N.NODEID = (" +
                  "SELECT MIN(NODEID) FROM VIZ.MG_NODE WHERE MESMATNAME = :LOCID AND MESTAG IS NOT NULL)";

        rez = GetLen(stmtSql, matLocId, null);
      }

      return rez;
    }

    public static decimal GetLenCoilUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT LAENGEOUTPUT FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToDecimal(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetPlaceNumUo(string matLocId)
    {
      //decimal rez;
      const string stmt = "SELECT P.SVALUEACT FROM VIZ.PPAR P WHERE P.PARAMETER = 'PLACEMENT_NUMBER' AND P.ME_IDINPUT = (SELECT ME_ID FROM VIZ.MAT WHERE BEZEICHNUNG = :LOCID) AND P.CATEGORY = 'PACK' AND P.SUBCATEGORY = 'PACKING'";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocId.Length,
        Value = matLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetAnLot(string matLocId)
    {
      //decimal rez;
      const string stmt = "SELECT ANNEALINGLOT /*, ANNEALINGLOTSEQNO*/ FROM VIZ.MAT WHERE (BEZEICHNUNG = :LOCID)"; 

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocId.Length,
        Value = matLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }


    public static string GetBrigada(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT VIZ_PRN.UTL_RPT.GetBrigada(DTENDE, ANLAGE) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = 'STRANN')";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetBrigadaUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT VIZ_PRN.UTL_RPT.GetBrigada(DTENDE, ANLAGE) FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetController(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT REPLACE(AENDERER,'_','.') FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID) AND (AGTYP = 'STRANN')";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetControllerUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT REPLACE(AENDERER,'_','.') FROM VIZ.PRODUCTIONJOURNAL WHERE PJID = (SELECT MAX(PJID) FROM VIZ.PRODUCTIONJOURNAL WHERE (MATBEZEICHNUNGOUTPUT = :LOCID))";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }



    public static string GetMapDefInfo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT VIZ_PRN.Razdel_16.GetMapDefInfo(:LOCID) FROM DUAL";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static string GetMapDefInfoUo(string MatLocId)
    {
      //decimal rez;
      const string stmt = "SELECT VIZ_PRN.Razdel_16.GetMapDefInfo(:LOCID) FROM DUAL";
      int idx  = MatLocId.IndexOf('/');

      if (MatLocId.IndexOf('/') != -1)
        MatLocId = MatLocId.Substring(0, MatLocId.IndexOf('/'));

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      return Convert.ToString(Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm));
    }

    public static decimal? GetK2sUo(string MatLocId)
    {
      const string stmt = "SELECT VIZ_PRN.Razdel_16.GetK2s(:LOCID) FROM DUAL";

      var lstPrm = new List<OracleParameter>();
      var prm = new OracleParameter
      {
        ParameterName = "LOCID",
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = MatLocId.Length,
        Value = MatLocId
      };
      lstPrm.Add(prm);

      object res = Odac.ExecuteScalar(stmt, CommandType.Text, false, lstPrm);

      if (res == DBNull.Value)
        return null;
      else
        return Convert.ToDecimal(res);
    }

    public static void CreateCutMatData(string matLocId)
    {
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        OracleDbType = OracleDbType.VarChar,
        Size = matLocId.Length,
        Value = matLocId
      };
      lstPrm.Add(prm);

      Odac.ExecuteNonQuery("VIZ_PRN.CUTMATAVO.CRETECUTAVOMAP", CommandType.StoredProcedure, false, lstPrm);
    }




  }




  
}






