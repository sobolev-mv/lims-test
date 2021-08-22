using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Threading;
using System.Threading;
using System.Windows;
using Viz.DbApp.Psi;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.WrkModule.RptManager.Db
{

  public class RptWithF5Param : Smv.Xls.XlsInstanceParam
  {
    public string PathScriptsDir { get; set; }
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int TypeFilter { get; set; } //0-Без фильтра; 1-Фильтр; 2-Список стендов
    public string ListStendF5  { get; set; }
    public int TypeListValueF5 { get; set; }
    public Boolean IsAroF5 { get; set; }
    public Boolean Is1200F5 { get; set; }
    public Boolean IsApr1F5  { get; set; }
    public Boolean IsAooF5 { get; set; }
    public Boolean IsVtoF5 { get; set; }
    public Boolean IsAvoF5 { get; set; }
    public Boolean IsDateAroF5 { get; set; }
    public Boolean IsDate1200F5 { get; set; }
    public Boolean IsDateApr1F5 { get; set; }
    public Boolean IsDateAooF5 { get; set; }
    public Boolean IsMgOF5 { get; set; }
    public Boolean IsPppF5 { get; set; }
    public Boolean IsWgtCoverF5 { get; set; }
    public Boolean IsStVtoF5 { get; set; }
    public Boolean IsKlpVtoF5 { get; set; }
    public Boolean IsDiskVtoF5 { get; set; }
    public Boolean IsTimeAooVtoF5 { get; set; }
    public Boolean IsBrgAooF5 { get; set; }
    public DateTime DateBeginAroF5 { get; set; }
    public DateTime DateEndAroF5 { get; set; }
    public DateTime DateBegin1200F5 { get; set; }
    public DateTime DateEnd1200F5 { get; set; }
    public DateTime DateBeginApr1F5 { get; set; }
    public DateTime DateEndApr1F5 { get; set; }
    public DateTime DateBeginAooF5 { get; set; }
    public DateTime DateEndAooF5 { get; set; }
    public string AroF5Item { get; set; }
    public string Stan1200F5Item { get; set; }
    public string TolsF5Item { get; set; }
    public string BrgApr1F5Item { get; set; }
    public string BrgVtoF5Item { get; set; }
    public string BrgAooF5Item { get; set; }
    public string BrgAvoF5Item { get; set; }
    public string AooF5Item { get; set; }
    public string AvoF5Item { get; set; }
    public string ShirApr1F5Item { get; set; }
    public string DiskVtoF5Item { get; set; }
    public int AooF5MgOFrom { get; set; }
    public int AooF5MgOTo { get; set; }
    public decimal AooF5PppFrom { get; set; }
    public decimal AooF5PppTo { get; set; }
    public int AooF5WgtCoverFrom { get; set; }
    public int AooF5WgtCoverTo { get; set; }
    public string VtoF5Stend { get; set; }
    public string VtoF5Cap { get; set; }
    public int VtoF5TimeAooVto { get; set; }
    public Boolean IsDateAvoLstF5 { get; set; }
    public DateTime DateBeginAvoLstF5 { get; set; }
    public DateTime DateEndAvoLstF5 { get; set; }


    public RptWithF5Param(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}

    public Boolean IsFilter5Applayed()
    {
      return (IsAroF5 || Is1200F5 || IsApr1F5 || IsAooF5 || IsVtoF5 || IsAvoF5);
    }

    public string GetFilterCriteria()
    {
      string flt = null;

      if (IsAroF5){
        flt += "АРО\r\n" + "Агрегаты: " + AroF5Item + "  ";
        if (IsDateAroF5)
          flt += string.Format("Дата АРО  с {0:dd.MM.yyyy  HH:mm:ss}", DateBeginAroF5) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEndAroF5);

        flt += "\r\n";
      }

      if (Is1200F5){
        flt += "Стан 1200\r\n" + "Агрегаты: " + Stan1200F5Item + "  " + "Толщины: " + TolsF5Item + "  ";
        if (IsDate1200F5)
          flt += string.Format("Дата Стан 1200 с {0:dd.MM.yyyy  HH:mm:ss}", DateBegin1200F5) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEnd1200F5);

        flt += "\r\n";
      }

      if (IsApr1F5){
        flt += "АПР1\r\n" + "Ширины: " + ShirApr1F5Item + "  " + "Бригады: " + BrgApr1F5Item + "  ";
        if (IsDateApr1F5)
          flt += string.Format("Дата АПР1 с {0:dd.MM.yyyy  HH:mm:ss}", DateBeginApr1F5) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEndApr1F5);

        flt += "\r\n";
      }

      if (IsAooF5){
        flt += "АОО\r\n" + "Агрегаты: " + AooF5Item + " ";
        if (IsDateAooF5)
          flt += string.Format("Дата АОО с {0:dd.MM.yyyy HH:mm:ss}", DateBeginAooF5) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEndAooF5) + "  ";

        if (IsMgOF5)
          flt += "MgO c " + AooF5MgOFrom.ToString(CultureInfo.InvariantCulture) + " по " + AooF5MgOTo.ToString(CultureInfo.InvariantCulture) + "  ";

        if (IsPppF5)
          flt += "ППП с " + AooF5PppFrom.ToString(CultureInfo.InvariantCulture) + " по " + AooF5PppTo.ToString(CultureInfo.InvariantCulture) + "  ";

        if (IsWgtCoverF5)
          flt += "Вес покрытия с " + AooF5WgtCoverFrom.ToString(CultureInfo.InvariantCulture) + " по " + AooF5WgtCoverTo.ToString(CultureInfo.InvariantCulture) + "  ";

        if (IsBrgAooF5)
          flt += "Бригады: " + BrgAooF5Item + "  "; 

        flt += "\r\n";
      }

      if (IsVtoF5){
        flt += "ВТО\r\n";
        flt += "Бригады: " + BrgVtoF5Item + "  ";  
        
        if (IsStVtoF5)
          flt += "Стенд ВТО: " + VtoF5Stend + "  ";
        if (IsKlpVtoF5)
          flt += "Колпак: " + VtoF5Cap + "  ";
        if (IsDiskVtoF5)
          flt += "Диск ВТО: " + DiskVtoF5Item + "  ";
        if (IsTimeAooVtoF5)
          flt += "Время АОО-ВТО час: " + VtoF5TimeAooVto + "  ";
        flt += "\r\n";
      }

      if (IsAvoF5){
        flt += "АВО\r\n" + "Агрегаты: " + AvoF5Item + "  " + "Бригады: " + BrgAvoF5Item;
      }

      return flt;
    }

    public void GetFilterLstCriteria(int ExcelSheet)
    {
      dynamic wrkSheet = null;
      //выбираем лист
      ExcelApp.ActiveWorkbook.WorkSheets[ExcelSheet].Select();
      wrkSheet = ExcelApp.ActiveSheet;
      
      /*
      string hdrIncl = null;
      hdrIncl = TypeInclList == 0 ? "Включая значения из списка " : "Исключая значения списка ";
      */

      switch (TypeListValueF5){
        case 0:
          wrkSheet.Cells[2, 1].Value = "Стендовые партии:";
          break;
        case 1:
          wrkSheet.Cells[2, 1].Value = "Стенды ВТО:";
          break;
        default:
          Console.WriteLine("Default case");
          break;
      }

      const int row = 3;
      string[] strArr = ListStendF5.Split(new char[] {','});
      for (int i = 0; i < strArr.Length; i++) 
        wrkSheet.Cells[row + i, 1].Value = strArr[i];
    }

  }

  public class RptWithF5 : Smv.Xls.XlsRpt
  {

    public Boolean PrepareFilterRpt(RptWithF5Param prm, String strThicknessSql)
    {
      List<OracleParameter> lstParam = null;
      string sqlStmt = null;

      //Корректироем даты для учета заводской смены
      prm.DateBeginAroF5 = prm.DateBeginAroF5.AddHours(8);
      prm.DateEndAroF5 = prm.DateEndAroF5.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBegin1200F5 = prm.DateBegin1200F5.AddHours(8);
      prm.DateEnd1200F5 = prm.DateEnd1200F5.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBeginApr1F5 = prm.DateBeginApr1F5.AddHours(8);
      prm.DateEndApr1F5 = prm.DateEndApr1F5.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBeginAooF5 = prm.DateBeginAooF5.AddHours(8);
      prm.DateEndAooF5 = prm.DateEndAooF5.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBeginAvoLstF5 = prm.DateBeginAvoLstF5.AddHours(8);
      prm.DateEndAvoLstF5 = prm.DateEndAvoLstF5.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);

      DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);

      switch (prm.TypeFilter)
      {
        case 0:
          sqlStmt = "begin " +
                    "delete from VIZ_PRN.TMP_FINCUT_FILTR_CORE; " +
                    "insert into VIZ_PRN.TMP_FINCUT_FILTR_CORE " +
                    "select * from VIZ_PRN.V_FINCUT_FILTR_CORE where tols in (SELECT TO_NUMBER(VL_STRING) FROM TABLE(VIZ_PRN.VAR_RPT.GetTabOfStrDelim(:P1, ','))); " +
                    "end;";

          List<OracleParameter> lstPrm = new List<OracleParameter>();
          OracleParameter prmOra = new OracleParameter()
          {
            ParameterName = "P1",
            DbType = DbType.String,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.VarChar,
            Size = strThicknessSql.Length,
            Value = strThicknessSql
          };
          lstPrm.Add(prmOra);

          Odac.ExecuteNonQuery(sqlStmt, CommandType.Text, false, lstPrm, true);

          return true;
        case 1:
          sqlStmt = System.IO.File.ReadAllText(prm.PathScriptsDir + "\\MgrFilterCoreFinCut.sql", Encoding.GetEncoding(1251)).Replace("\r", " ");

          if (prm.IsFilter5Applayed())
          {
            lstParam = new List<OracleParameter>();

            var param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "ARO",
              Value = prm.AroF5Item,
              Size = prm.AroF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FARO",
              Value = prm.IsAroF5 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1ARO",
              Value = prm.DateBeginAroF5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2ARO",
              Value = prm.DateEndAroF5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDTARO",
              Value = (prm.IsDateAroF5 && prm.IsAroF5) ? 1 : 0
            };
            lstParam.Add(param);

            //ST1200
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "S1200",
              Value = prm.Stan1200F5Item,
              Size = prm.Stan1200F5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "F1200",
              Value = prm.Is1200F5 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "TOLS1200",
              Value = prm.TolsF5Item,
              Size = prm.TolsF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT11200",
              Value = prm.DateBegin1200F5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT21200",
              Value = prm.DateEnd1200F5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDT1200",
              Value = (prm.IsDate1200F5 && prm.Is1200F5) ? 1 : 0
            };
            lstParam.Add(param);

            //АПР1 
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "APR1WDTH",
              Value = prm.ShirApr1F5Item,
              Size = prm.ShirApr1F5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAPR1",
              Value = prm.IsApr1F5 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "APR1BRG",
              Value = prm.BrgApr1F5Item,
              Size = prm.BrgApr1F5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1APR1",
              Value = prm.DateBeginApr1F5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2APR1",
              Value = prm.DateEndApr1F5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDTAPR1",
              Value = (prm.IsApr1F5 && prm.IsDateApr1F5) ? 1 : 0
            };
            lstParam.Add(param);

            //АОО
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "AOO",
              Value = prm.AooF5Item,
              Size = prm.AooF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAOO",
              Value = prm.IsAooF5 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1AOO",
              Value = prm.DateBeginAooF5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2AOO",
              Value = prm.DateEndAooF5
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDTAOO",
              Value = (prm.IsAooF5 && prm.IsDateAooF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Decimal,
              OracleDbType = OracleDbType.Number,
              Direction = ParameterDirection.Input,
              Precision = 10,
              Scale = 0,
              ParameterName = "MGO1",
              Value = Convert.ToDecimal(prm.AooF5MgOFrom)
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Decimal,
              OracleDbType = OracleDbType.Number,
              Direction = ParameterDirection.Input,
              Precision = 10,
              Scale = 0,
              ParameterName = "MGO2",
              Value = Convert.ToDecimal(prm.AooF5MgOTo)
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FMGO",
              Value = (prm.IsAooF5 && prm.IsMgOF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Decimal,
              OracleDbType = OracleDbType.Number,
              Direction = ParameterDirection.Input,
              Precision = 10,
              Scale = 1,
              ParameterName = "PPP1",
              Value = prm.AooF5PppFrom
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Decimal,
              OracleDbType = OracleDbType.Number,
              Direction = ParameterDirection.Input,
              Precision = 10,
              Scale = 1,
              ParameterName = "PPP2",
              Value = prm.AooF5PppTo
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FPPP",
              Value = (prm.IsAooF5 && prm.IsPppF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Decimal,
              OracleDbType = OracleDbType.Number,
              Direction = ParameterDirection.Input,
              Precision = 10,
              Scale = 0,
              ParameterName = "VESPOKR1",
              Value = Convert.ToDecimal(prm.AooF5WgtCoverFrom)
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Decimal,
              OracleDbType = OracleDbType.Number,
              Direction = ParameterDirection.Input,
              Precision = 10,
              Scale = 0,
              ParameterName = "VESPOKR2",
              Value = Convert.ToDecimal(prm.AooF5WgtCoverTo)
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FWGTCOVER",
              Value = (prm.IsAooF5 && prm.IsWgtCoverF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "AOOBRG",
              Value = prm.BrgAooF5Item,
              Size = prm.BrgAooF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAOOBRG",
              Value = (prm.IsAooF5 && prm.IsBrgAooF5) ? 1 : 0
            };
            lstParam.Add(param);

            //ВТО
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "STVTO",
              Value = prm.VtoF5Stend,
              //Size = prm.VtoF5Stend.Length
            };
            if (prm.VtoF5Stend != null)
              param.Size = prm.VtoF5Stend.Length;
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FSTVTO",
              Value = (prm.IsVtoF5 && prm.IsStVtoF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "KLPVTO",
              Value = prm.VtoF5Cap,
              //Size = prm.VtoF5Cap.Length
            };
            if (prm.VtoF5Cap != null)
              param.Size = prm.VtoF5Cap.Length;
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FKLPVTO",
              Value = (prm.IsVtoF5 && prm.IsKlpVtoF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "DISKVTO",
              Value = prm.DiskVtoF5Item,
              Size = prm.DiskVtoF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDISKVTO",
              Value = (prm.IsVtoF5 && prm.IsDiskVtoF5) ? 1 : 0
            };
            lstParam.Add(param);

            /* 
            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FVTO",
              Value = prm.IsVtoF5 ? 1 : 0
            };
            lstParam.Add(param);*/

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "BRGVTO",
              Value = prm.BrgVtoF5Item,
              Size = prm.BrgVtoF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FBRGVTO",
              Value = 0//(prm.IsVtoF5 && prm.IsDiskVtoF5) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "AOOVTO",
              Value = prm.VtoF5TimeAooVto
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAOOVTO",
              Value = (prm.IsVtoF5 && prm.IsTimeAooVtoF5) ? 1 : 0
            };
            lstParam.Add(param);

            //АВО
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "AVO",
              Value = prm.AvoF5Item,
              Size = prm.AvoF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAVO",
              Value = prm.IsAvoF5 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "BRGAVO",
              Value = prm.BrgAvoF5Item,
              Size = prm.BrgAvoF5Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter()
            {
              ParameterName = "MTOLS",
              DbType = DbType.String,
              Direction = ParameterDirection.Input,
              OracleDbType = OracleDbType.VarChar,
              Size = strThicknessSql.Length,
              Value = strThicknessSql
            };
            lstParam.Add(param);

            Odac.ExecuteNonQuery(sqlStmt, CommandType.Text, false, lstParam, true);
          }
          break;
        case 2:
          sqlStmt = "begin " +
                    "delete from VIZ_PRN.TMP_FINCUT_FILTR_CORE; " +
                    "insert into VIZ_PRN.TMP_FINCUT_FILTR_CORE " +
                    "select * from  VIZ_PRN.V_FINCUT_FILTR_STEND_CORE;" +
                    "end;"; 

          DbVar.SetStringList(prm.ListStendF5, ",");
          Odac.ExecuteNonQuery(sqlStmt, CommandType.Text, false, null, true);

          break;
      }

      //int rs = Convert.ToInt32(Odac.ExecuteScalar("select count(*) from VIZ_PRN.TMP_FINCUT_FILTR_CORE", CommandType.Text, false, null));
      //MessageBox.Show(rs.ToString());
      

      return true;
    }


  }

}
