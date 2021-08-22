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

namespace Viz.WrkModule.RptOtk.Db
{

  public class RptWithF1Param : Smv.Xls.XlsInstanceParam
  {
    public string PathScriptsDir { get; set; }
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int TypeFilter { get; set; } //0-Без фильтра; 1-Фильтр; 2-Список стендов
    public string ListStendF1  { get; set; }
    public int TypeListValueF1 { get; set; }
    public Boolean IsAroF1 { get; set; }
    public Boolean Is1200F1 { get; set; }
    public Boolean IsApr1F1  { get; set; }
    public Boolean IsAooF1 { get; set; }
    public Boolean IsVtoF1 { get; set; }
    public Boolean IsAvoF1 { get; set; }
    public Boolean IsDateAroF1 { get; set; }
    public Boolean IsDate1200F1 { get; set; }
    public Boolean IsDateApr1F1 { get; set; }
    public Boolean IsDateAooF1 { get; set; }
    public Boolean IsMgOF1 { get; set; }
    public Boolean IsPppF1 { get; set; }
    public Boolean IsWgtCoverF1 { get; set; }
    public Boolean IsStVtoF1 { get; set; }
    public Boolean IsKlpVtoF1 { get; set; }
    public Boolean IsDiskVtoF1 { get; set; }
    public Boolean IsTimeAooVtoF1 { get; set; }
    public Boolean IsBrgAooF1 { get; set; }
    public DateTime DateBeginAroF1 { get; set; }
    public DateTime DateEndAroF1 { get; set; }
    public DateTime DateBegin1200F1 { get; set; }
    public DateTime DateEnd1200F1 { get; set; }
    public DateTime DateBeginApr1F1 { get; set; }
    public DateTime DateEndApr1F1 { get; set; }
    public DateTime DateBeginAooF1 { get; set; }
    public DateTime DateEndAooF1 { get; set; }
    public string AroF1Item { get; set; }
    public string Stan1200F1Item { get; set; }
    public string TolsF1Item { get; set; }
    public string BrgApr1F1Item { get; set; }
    public string BrgVtoF1Item { get; set; }
    public string BrgAooF1Item { get; set; }
    public string BrgAvoF1Item { get; set; }
    public string AooF1Item { get; set; }
    public string AvoF1Item { get; set; }
    public string ShirApr1F1Item { get; set; }
    public string DiskVtoF1Item { get; set; }
    public int AooF1MgOFrom { get; set; }
    public int AooF1MgOTo { get; set; }
    public decimal AooF1PppFrom { get; set; }
    public decimal AooF1PppTo { get; set; }
    public int AooF1WgtCoverFrom { get; set; }
    public int AooF1WgtCoverTo { get; set; }
    public string VtoF1Stend { get; set; }
    public string VtoF1Cap { get; set; }
    public int VtoF1TimeAooVto { get; set; }
    public Boolean IsDateAvoLstF1 { get; set; }
    public DateTime DateBeginAvoLstF1 { get; set; }
    public DateTime DateEndAvoLstF1 { get; set; }


    public RptWithF1Param(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}

    public Boolean IsFilter1Applayed()
    {
      return (IsAroF1 || Is1200F1 || IsApr1F1 || IsAooF1 || IsVtoF1 || IsAvoF1);
    }

    public string GetFilterCriteria()
    {
      string flt = null;

      if (IsAroF1){
        flt += "АРО\r\n" + "Агрегаты: " + AroF1Item + "  ";
        if (IsDateAroF1)
          flt += string.Format("Дата АРО  с {0:dd.MM.yyyy  HH:mm:ss}", DateBeginAroF1) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEndAroF1);

        flt += "\r\n";
      }

      if (Is1200F1){
        flt += "Стан 1200\r\n" + "Агрегаты: " + Stan1200F1Item + "  " + "Толщины: " + TolsF1Item + "  ";
        if (IsDate1200F1)
          flt += string.Format("Дата Стан 1200 с {0:dd.MM.yyyy  HH:mm:ss}", DateBegin1200F1) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEnd1200F1);

        flt += "\r\n";
      }

      if (IsApr1F1){
        flt += "АПР1\r\n" + "Ширины: " + ShirApr1F1Item + "  " + "Бригады: " + BrgApr1F1Item + "  ";
        if (IsDateApr1F1)
          flt += string.Format("Дата АПР1 с {0:dd.MM.yyyy  HH:mm:ss}", DateBeginApr1F1) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEndApr1F1);

        flt += "\r\n";
      }

      if (IsAooF1){
        flt += "АОО\r\n" + "Агрегаты: " + AooF1Item + " ";
        if (IsDateAooF1)
          flt += string.Format("Дата АОО с {0:dd.MM.yyyy HH:mm:ss}", DateBeginAooF1) + " по " + string.Format("{0:dd.MM.yyyy  HH:mm:ss}", DateEndAooF1) + "  ";

        if (IsMgOF1)
          flt += "MgO c " + AooF1MgOFrom.ToString(CultureInfo.InvariantCulture) + " по " + AooF1MgOTo.ToString(CultureInfo.InvariantCulture) + "  ";

        if (IsPppF1)
          flt += "ППП с " + AooF1PppFrom.ToString(CultureInfo.InvariantCulture) + " по " + AooF1PppTo.ToString(CultureInfo.InvariantCulture) + "  ";

        if (IsWgtCoverF1)
          flt += "Вес покрытия с " + AooF1WgtCoverFrom.ToString(CultureInfo.InvariantCulture) + " по " + AooF1WgtCoverTo.ToString(CultureInfo.InvariantCulture) + "  ";

        if (IsBrgAooF1)
          flt += "Бригады: " + BrgAooF1Item + "  "; 

        flt += "\r\n";
      }

      if (IsVtoF1){
        flt += "ВТО\r\n";
        flt += "Бригады: " + BrgVtoF1Item + "  ";  
        
        if (IsStVtoF1)
          flt += "Стенд ВТО: " + VtoF1Stend + "  ";
        if (IsKlpVtoF1)
          flt += "Колпак: " + VtoF1Cap + "  ";
        if (IsDiskVtoF1)
          flt += "Диск ВТО: " + DiskVtoF1Item + "  ";
        if (IsTimeAooVtoF1)
          flt += "Время АОО-ВТО час: " + VtoF1TimeAooVto + "  ";
        flt += "\r\n";
      }

      if (IsAvoF1){
        flt += "АВО\r\n" + "Агрегаты: " + AvoF1Item + "  " + "Бригады: " + BrgAvoF1Item;
      }

      return flt;
    }

    public void GetFilter1LstCriteria(int ExcelSheet)
    {
      dynamic wrkSheet = null;
      //выбираем лист
      ExcelApp.ActiveWorkbook.WorkSheets[ExcelSheet].Select();
      wrkSheet = ExcelApp.ActiveSheet;
      
      /*
      string hdrIncl = null;
      hdrIncl = TypeInclList == 0 ? "Включая значения из списка " : "Исключая значения списка ";
      */

      switch (TypeListValueF1){
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
      string[] strArr = ListStendF1.Split(new char[] {','});
      for (int i = 0; i < strArr.Length; i++) 
        wrkSheet.Cells[row + i, 1].Value = strArr[i];
    }

  }

  public class RptWithF1 : Smv.Xls.XlsRpt
  {
    public Boolean PrepareFilterRpt(RptWithF1Param prm)
    {
      List<OracleParameter> lstParam = null;
      string sqlStmt = null;

      //Корректироем даты для учета заводской смены
      prm.DateBeginAroF1 = prm.DateBeginAroF1.AddHours(8);
      prm.DateEndAroF1 = prm.DateEndAroF1.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBegin1200F1 = prm.DateBegin1200F1.AddHours(8);
      prm.DateEnd1200F1 = prm.DateEnd1200F1.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBeginApr1F1 = prm.DateBeginApr1F1.AddHours(8);
      prm.DateEndApr1F1 = prm.DateEndApr1F1.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBeginAooF1 = prm.DateBeginAooF1.AddHours(8);
      prm.DateEndAooF1 = prm.DateEndAooF1.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);
      prm.DateBeginAvoLstF1 = prm.DateBeginAvoLstF1.AddHours(8);
      prm.DateEndAvoLstF1 = prm.DateEndAvoLstF1.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);

      DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);

      switch (prm.TypeFilter){
        case 0:
          sqlStmt = "BEGIN DELETE FROM VIZ_PRN.TMP_OTK_FILTR_CORE; INSERT INTO VIZ_PRN.TMP_OTK_FILTR_CORE SELECT * FROM VIZ_PRN.OTK_FILTR_CORE; END;";
          break;
        case 1:
          sqlStmt = System.IO.File.ReadAllText(prm.PathScriptsDir + "\\OtkFilterCore.sql", Encoding.GetEncoding(1251)).Replace("\r", " ");

          if (prm.IsFilter1Applayed()){
            lstParam = new List<OracleParameter>();

            var param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "ARO",
              Value = prm.AroF1Item,
              Size = prm.AroF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FARO",
              Value = prm.IsAroF1 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1ARO",
              Value = prm.DateBeginAroF1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2ARO",
              Value = prm.DateEndAroF1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDTARO",
              Value = (prm.IsDateAroF1 && prm.IsAroF1) ? 1 : 0
            };
            lstParam.Add(param);

            //ST1200
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "S1200",
              Value = prm.Stan1200F1Item,
              Size = prm.Stan1200F1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "F1200",
              Value = prm.Is1200F1 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "TOLS",
              Value = prm.TolsF1Item,
              Size = prm.TolsF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT11200",
              Value = prm.DateBegin1200F1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT21200",
              Value = prm.DateEnd1200F1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDT1200",
              Value = (prm.IsDate1200F1 && prm.Is1200F1) ? 1 : 0
            };
            lstParam.Add(param);

            //АПР1 
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "APR1WDTH",
              Value = prm.ShirApr1F1Item,
              Size = prm.ShirApr1F1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAPR1",
              Value = prm.IsApr1F1 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "APR1BRG",
              Value = prm.BrgApr1F1Item,
              Size = prm.BrgApr1F1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1APR1",
              Value = prm.DateBeginApr1F1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2APR1",
              Value = prm.DateEndApr1F1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDTAPR1",
              Value = (prm.IsApr1F1 && prm.IsDateApr1F1) ? 1 : 0
            };
            lstParam.Add(param);

            //АОО
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "AOO",
              Value = prm.AooF1Item,
              Size = prm.AooF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAOO",
              Value = prm.IsAooF1 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1AOO",
              Value = prm.DateBeginAooF1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2AOO",
              Value = prm.DateEndAooF1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDTAOO",
              Value = (prm.IsAooF1 && prm.IsDateAooF1) ? 1 : 0
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
              Value = Convert.ToDecimal(prm.AooF1MgOFrom)
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
              Value = Convert.ToDecimal(prm.AooF1MgOTo)
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FMGO",
              Value = (prm.IsAooF1 && prm.IsMgOF1) ? 1 : 0
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
              Value = prm.AooF1PppFrom
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
              Value = prm.AooF1PppTo
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FPPP",
              Value = (prm.IsAooF1 && prm.IsPppF1) ? 1 : 0
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
              Value = Convert.ToDecimal(prm.AooF1WgtCoverFrom)
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
              Value = Convert.ToDecimal(prm.AooF1WgtCoverTo)
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FWGTCOVER",
              Value = (prm.IsAooF1 && prm.IsWgtCoverF1) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "AOOBRG",
              Value = prm.BrgAooF1Item,
              Size = prm.BrgAooF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAOOBRG",
              Value = (prm.IsAooF1 && prm.IsBrgAooF1) ? 1 : 0
            };
            lstParam.Add(param);

            //ВТО
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "STVTO",
              Value = prm.VtoF1Stend,
              //Size = prm.VtoF1Stend.Length
            };
            if (prm.VtoF1Stend != null)
              param.Size = prm.VtoF1Stend.Length;
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FSTVTO",
              Value = (prm.IsVtoF1 && prm.IsStVtoF1) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "KLPVTO",
              Value = prm.VtoF1Cap,
              //Size = prm.VtoF1Cap.Length
            };
            if (prm.VtoF1Cap != null)
              param.Size = prm.VtoF1Cap.Length;
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FKLPVTO",
              Value = (prm.IsVtoF1 && prm.IsKlpVtoF1) ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "DISKVTO",
              Value = prm.DiskVtoF1Item,
              Size = prm.DiskVtoF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FDISKVTO",
              Value = (prm.IsVtoF1 && prm.IsDiskVtoF1) ? 1 : 0
            };
            lstParam.Add(param);

            /* 
            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FVTO",
              Value = prm.IsVtoF1 ? 1 : 0
            };
            lstParam.Add(param);*/
            
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "BRGVTO",
              Value = prm.BrgVtoF1Item,
              Size = prm.BrgVtoF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FBRGVTO",
              Value = 0//(prm.IsVtoF1 && prm.IsDiskVtoF1) ? 1 : 0
            };
            lstParam.Add(param);
            
            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "AOOVTO",
              Value = prm.VtoF1TimeAooVto
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAOOVTO",
              Value = (prm.IsVtoF1 && prm.IsTimeAooVtoF1) ? 1 : 0
            };
            lstParam.Add(param);

            
            //АВО
            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "AVO",
              Value = prm.AvoF1Item,
              Size = prm.AvoF1Item.Length
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.Int32,
              OracleDbType = OracleDbType.Integer,
              Direction = ParameterDirection.Input,
              ParameterName = "FAVO",
              Value = prm.IsAvoF1 ? 1 : 0
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.String,
              OracleDbType = OracleDbType.VarChar,
              Direction = ParameterDirection.Input,
              ParameterName = "BRGAVO",
              Value = prm.BrgAvoF1Item,
              Size = prm.BrgAvoF1Item.Length
            };

            lstParam.Add(param);
          }
          break;
        case 2:
          if (prm.IsDateAvoLstF1){
            //0-№ Стенда, 1-№ Стенда ВТО 
            if (prm.TypeListValueF1 == 0)
              sqlStmt =
                "BEGIN DELETE FROM VIZ_PRN.TMP_OTK_FILTR_CORE; INSERT INTO VIZ_PRN.TMP_OTK_FILTR_CORE SELECT * FROM VIZ_PRN.OTK_FILTR_STEND_CORE WHERE DTE_AVO BETWEEN :DT1 AND :DT2; END;";
            else
              sqlStmt =
                "BEGIN DELETE FROM VIZ_PRN.TMP_OTK_FILTR_CORE; INSERT INTO VIZ_PRN.TMP_OTK_FILTR_CORE SELECT * FROM VIZ_PRN.OTK_FILTR_CORE WHERE ST_VTO IN (SELECT VL_STRING FROM TABLE(VIZ_PRN.VAR_RPT.GetListStr1)) AND DTE_AVO BETWEEN :DT1 AND :DT2; END;";

            lstParam = new List<OracleParameter>();
            var param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT1",
              Value = prm.DateBeginAvoLstF1
            };
            lstParam.Add(param);

            param = new OracleParameter
            {
              DbType = DbType.DateTime,
              OracleDbType = OracleDbType.Date,
              Direction = ParameterDirection.Input,
              ParameterName = "DT2",
              Value = prm.DateEndAvoLstF1
            };
            lstParam.Add(param);

          }
          else{
            //0-№ Стенда, 1-№ Стенда ВТО 
            if (prm.TypeListValueF1 == 0)
              sqlStmt =
                "BEGIN DELETE FROM VIZ_PRN.TMP_OTK_FILTR_CORE; INSERT INTO VIZ_PRN.TMP_OTK_FILTR_CORE SELECT * FROM VIZ_PRN.OTK_FILTR_STEND_CORE; END;";
            else
              sqlStmt =
                "BEGIN DELETE FROM VIZ_PRN.TMP_OTK_FILTR_CORE; INSERT INTO VIZ_PRN.TMP_OTK_FILTR_CORE SELECT * FROM VIZ_PRN.OTK_FILTR_CORE WHERE ST_VTO IN (SELECT VL_STRING FROM TABLE(VIZ_PRN.VAR_RPT.GetListStr1)); END;";
          }

          DbVar.SetStringList(prm.ListStendF1,",");
          break;
      }

      Odac.ExecuteNonQuery(sqlStmt, CommandType.Text, false, lstParam);

      //int rs = Convert.ToInt32(Odac.ExecuteScalar("select count(*) from VIZ_PRN.TMP_OTK_FILTR_CORE", CommandType.Text, false, null));
      //MessageBox.Show(rs.ToString());


      //MessageBox.Show("Point2");

      return true;
    }



  }

}
