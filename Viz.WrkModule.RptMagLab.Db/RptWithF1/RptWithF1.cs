using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Threading;
using System.Threading;
using Viz.DbApp.Psi;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.WrkModule.RptMagLab.Db
{

  public class RptWithF1Param : Smv.Xls.XlsInstanceParam
  {
    public string PathScriptsDir   { get; set; }
    public DateTime DateBegin      { get; set; }
    public DateTime DateEnd        { get; set; }
    public int TypeFilter          { get; set; } //0-Без фильтра; 1-Фильтр; 2-Список стендов

    public Boolean IsSort          { get; set; }
    public DataRowView SelSortItem { get; set; }
    public DataRowView SelPlskItem { get; set; }

    public Boolean Is1200          { get; set; }
    public DataRowView Sel1200Item { get; set; }
    public Boolean IsDate1200      { get; set; }
    public DateTime DateBegin1200  { get; set; }
    public DateTime DateEnd1200    { get; set; }

    public Boolean IsAoo           { get; set; }
    public DataRowView SelAooItem  { get; set; }
    public Boolean IsDateAoo       { get; set; }
    public DateTime DateBeginAoo   { get; set; }
    public DateTime DateEndAoo     { get; set; }

    public Boolean IsAro           { get; set; }
    public DataRowView SelAroItem  { get; set; }
    public Boolean IsDateAro       { get; set; }
    public DateTime DateBeginAro   { get; set; }
    public DateTime DateEndAro     { get; set; }

    public Boolean IsAvo           { get; set; }
    public DataRowView SelAvoItem  { get; set; }
    public Boolean IsDateAvo       { get; set; }
    public DateTime DateBeginAvo   { get; set; }
    public DateTime DateEndAvo     { get; set; }

    public Boolean IsVto           { get; set; }
    public string StdVto           { get; set; }

    public Boolean IsApr           { get; set; }
    public DataRowView SelAprItem  { get; set; }

    public int TypeListValue       { get; set; } //0-№ Стенда, 1-№ Стенда ВТО   
    public int TypeInclList        { get; set; } //0-из списка, 1-исключая из списка     
    public string ListValue        { get; set; }

    public RptWithF1Param(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}

    public Boolean IsFilter1Applayed()
    {
      return (IsSort || Is1200 || IsAoo || IsAro || IsAvo || IsVto || IsApr);
    }

    public string GetFilter1Criteria()
    {
      string flt = null;

      if (IsSort)
        flt += "Сорт: " + Convert.ToString(SelSortItem.Row["StrDlg"]) + "  Плоскостность: " + Convert.ToString(SelPlskItem.Row["StrDlg"]) + "\r\n";

      if (Is1200){
        flt += "Стан 1200\r\n" + "Агрегаты: " + Convert.ToString(Sel1200Item.Row["StrDlg"]) + "  ";
        if (IsDate1200)
          flt += string.Format("Дата Ст1200 с {0:dd.MM.yyyy HH:mm:ss}", DateBegin1200) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", DateEnd1200) + "  ";

        flt += "\r\n";
      }

      if (IsAoo){
        flt += "АОО\r\n" + "Агрегаты: " + Convert.ToString(SelAooItem.Row["StrDlg"]) + " ";
        if (IsDateAoo)
          flt += string.Format("Дата АОО с {0:dd.MM.yyyy HH:mm:ss}", DateBeginAoo) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", DateEndAoo) + "  ";

        flt += "\r\n";
      }

      if (IsAro){
        flt += "АРО\r\n" + "Агрегаты: " + Convert.ToString(SelAroItem.Row["StrDlg"]) + " ";
        if (IsDateAro)
          flt += string.Format("Дата АРО с {0:dd.MM.yyyy HH:mm:ss}", DateBeginAro) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", DateEndAro) + "  ";

        flt += "\r\n";
      }

      if (IsAvo){
        flt += "АВО\r\n" + "Агрегаты: " + Convert.ToString(SelAvoItem.Row["StrDlg"]) + " ";
        if (IsDateAvo)
          flt += string.Format("Дата АВО с {0:dd.MM.yyyy HH:mm:ss}", DateBeginAvo) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", DateEndAvo) + "  ";

        flt += "\r\n";
      }

      if (IsVto){
        flt += "ВТО\r\n";
        flt += "№ Стенда ВТО: " + StdVto + "  ";
        flt += "\r\n";
      }

      if (IsApr){
        flt += "АПР\r\n" + "Агрегаты: " + Convert.ToString(SelAprItem.Row["StrDlg"]) + "  ";
        flt += "\r\n";
      }
      return flt;
    }

    public void GetFilter1LstCriteria()
    {
      dynamic wrkSheet = null;
      //выбираем лист
      ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
      wrkSheet = ExcelApp.ActiveSheet;
      string hdrIncl = null;

      hdrIncl = TypeInclList == 0 ? "Включая значения из списка " : "Исключая значения списка ";

      switch (TypeListValue)
      {
        case 0:
          wrkSheet.Cells[2, 1].Value = hdrIncl + "Стендовые партии:";
          break;
        case 1:
          wrkSheet.Cells[2, 1].Value = hdrIncl + "Стенды ВТО:";
          break;
        default:
          Console.WriteLine("Default case");
          break;
      }

      const int row = 3;
      string[] strArr = ListValue.Split(new char[] { ',' });
      for (int i = 0; i < strArr.Length; i++) wrkSheet.Cells[row + i, 1].Value = strArr[i];

    }
  }

  public class RptWithF1 : Smv.Xls.XlsRpt
  {
    public Boolean PrepareFilterRpt(RptWithF1Param prm)
    {
      IAsyncResult iar = null;
      List<OracleParameter> lstParam = null;
      string sqlStmt = null;

     //Корректироем даты для учета заводской смены
      prm.DateBegin1200 = prm.DateBegin1200.AddHours(8);
      prm.DateEnd1200 = prm.DateEnd1200.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);

      prm.DateBeginAoo = prm.DateBeginAoo.AddHours(8);
      prm.DateEndAoo = prm.DateEndAoo.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);

      prm.DateBeginAro = prm.DateBeginAro.AddHours(8);
      prm.DateEndAro = prm.DateEndAro.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);

      prm.DateBeginAvo = prm.DateBeginAvo.AddHours(8);
      prm.DateEndAvo = prm.DateEndAvo.AddDays(1).AddHours(7).AddMinutes(59).AddSeconds(59);

      prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));

      switch (prm.TypeFilter){
        case 0:
          sqlStmt = "BEGIN DELETE FROM VIZ_PRN.TMP_CZL_1_FILTR_CORE; INSERT INTO VIZ_PRN.TMP_CZL_1_FILTR_CORE SELECT * FROM VIZ_PRN.CZL_FILTR_CORE; END;";
          break;
        case 1:
          sqlStmt = System.IO.File.ReadAllText(prm.PathScriptsDir + "\\RptMagLabFilter1.sql", Encoding.GetEncoding(1251)).Replace("\r", " ");
          lstParam = new List<OracleParameter>();

          //Sort & Neplosk
          var param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "SORT",
            Value = Convert.ToString(prm.SelSortItem.Row["StrSql"]),
            Size = Convert.ToString(prm.SelSortItem.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "PLSK",
            Value = Convert.ToString(prm.SelPlskItem.Row["StrSql"]),
            Size = Convert.ToString(prm.SelPlskItem.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FSORT",
            Value = prm.IsSort ? 1 : 0
          };
          lstParam.Add(param);

          //ST1200
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "RM1200",
            Value = Convert.ToString(prm.Sel1200Item.Row["StrSql"]),
            Size = Convert.ToString(prm.Sel1200Item.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "F1200",
            Value = prm.Is1200 ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT1RM1200",
            Value = prm.DateBegin1200
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT2RM1200",
            Value = prm.DateEnd1200
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FDT1200",
            Value = (prm.IsDate1200 && prm.Is1200) ? 1 : 0
          };
          lstParam.Add(param);

          //АОО
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "AOO",
            Value = Convert.ToString(prm.SelAooItem.Row["StrSql"]),
            Size = Convert.ToString(prm.SelAooItem.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FAOO",
            Value = prm.IsAoo ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT1AOO",
            Value = prm.DateBeginAoo
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT2AOO",
            Value = prm.DateEndAoo
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FDTAOO",
            Value = (prm.IsAoo && prm.IsDateAoo) ? 1 : 0
          };
          lstParam.Add(param);

          //ARO
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "ARO",
            Value = Convert.ToString(prm.SelAroItem.Row["StrSql"]),
            Size = Convert.ToString(prm.SelAroItem.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FARO",
            Value = prm.IsAro ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT1ARO",
            Value = prm.DateBeginAro
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT2ARO",
            Value = prm.DateEndAro
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FDTARO",
            Value = (prm.IsDateAro && prm.IsAro) ? 1 : 0
          };
          lstParam.Add(param);

          //AVO
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "AVO",
            Value = Convert.ToString(prm.SelAvoItem.Row["StrSql"]),
            Size = Convert.ToString(prm.SelAvoItem.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FAVO",
            Value = prm.IsAvo ? 1 : 0
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT1AVO",
            Value = prm.DateBeginAro
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.DateTime,
            OracleDbType = OracleDbType.Date,
            Direction = ParameterDirection.Input,
            ParameterName = "DT2AVO",
            Value = prm.DateEndAro
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FDTAVO",
            Value = (prm.IsDateAvo && prm.IsAvo) ? 1 : 0
          };
          lstParam.Add(param);

          //ВТО
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "STVTO",
            Value = prm.StdVto,
            //Size = prm.VtoF1Stend.Length
          };
          if (prm.StdVto != null)
            param.Size = prm.StdVto.Length;

          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FSTVTO",
            Value = (prm.IsVto) ? 1 : 0
          };
          lstParam.Add(param);

          //APR
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "APR",
            Value = Convert.ToString(prm.SelAprItem.Row["StrSql"]),
            Size = Convert.ToString(prm.SelAprItem.Row["StrSql"]).Length
          };
          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FAPR",
            Value = prm.IsApr ? 1 : 0
          };
          lstParam.Add(param);
         
          break;
        case 2:
          Boolean isStend = false;
          Boolean isStendVto = false;

          //0-№ Стенда, 1-№ Стенда ВТО 
          if (prm.TypeListValue == 0){
            isStend = true;
            isStendVto = false;
          }else{
            isStend = false;
            isStendVto = true;
          }

          sqlStmt = prm.TypeInclList == 0 ? System.IO.File.ReadAllText(prm.PathScriptsDir + "\\RptMagLabFilter1LstInclude.sql", Encoding.GetEncoding(1251)).Replace("\r", " ") : System.IO.File.ReadAllText(prm.PathScriptsDir + "\\RptMagLabFilter1LstExclude.sql", Encoding.GetEncoding(1251)).Replace("\r", " ");
          lstParam = new List<OracleParameter>();

          //Стенд
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "STEND",
            Value = prm.ListValue
          };
          if (prm.ListValue != null)
            param.Size = prm.ListValue.Length;

          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FSTEND",
            Value = isStend ? 1 : 0
          };
          lstParam.Add(param);

          //Стенд ВТО
          param = new OracleParameter
          {
            DbType = DbType.String,
            OracleDbType = OracleDbType.VarChar,
            Direction = ParameterDirection.Input,
            ParameterName = "STVTO",
            Value = prm.ListValue
          };
          if (prm.ListValue != null)
            param.Size = prm.ListValue.Length;

          lstParam.Add(param);

          param = new OracleParameter
          {
            DbType = DbType.Int32,
            OracleDbType = OracleDbType.Integer,
            Direction = ParameterDirection.Input,
            ParameterName = "FSTVTO",
            Value = isStendVto ? 1 : 0
          };
          lstParam.Add(param);

          break;
      }

      prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(sqlStmt, CommandType.Text, false, true, lstParam); }));
      if (iar != null)
        iar.AsyncWaitHandle.WaitOne();
      else
        return false;
      
      var oracleCommand = iar.AsyncState as OracleCommand;
      if (oracleCommand != null){
        oracleCommand.EndExecuteNonQuery(iar);
        iar = null;
      }

      return true;
    }



  }

}
