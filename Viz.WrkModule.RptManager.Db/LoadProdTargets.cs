using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;
using Microsoft.Win32;
using DevExpress.Spreadsheet;
using Smv.Utils;


namespace Viz.WrkModule.RptManager.Db
{
  public sealed class LoadProdTargetsRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public string CfgFile { get; set; }
    public LoadProdTargetsRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class LoadProdTargets : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as LoadProdTargetsRptParam);
      
      try{
        this.RunRpt(prm);
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка загрузки", ex.Message, MessageBoxImage.Stop)));
      }
      finally{
        GC.Collect();
      }
    }

    private Boolean ClearRptPeriod(DateTime rptPeriod)
    {
      const string stmtSql = "DELETE FROM VIZ_PRN.TV_PLNTARGPROD WHERE RPT_PERIOD = :RPT_PERIOD";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        ParameterName = "RPT_PERIOD",
        OracleDbType = OracleDbType.Date,
        Value = rptPeriod
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.Text, true, lstPrm, true);
    }

    private Boolean InsertEventsToDashBoard(DateTime rptPeriod)
    {
      const string stmtSql = "INSERT INTO DSB.JB_PLNPROD(RPT_PERIOD) VALUES(:RPT_PERIOD)";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        ParameterName = "RPT_PERIOD",
        OracleDbType = OracleDbType.Date,
        Value = rptPeriod
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.Text, true, lstPrm, true);
    }
    
    private Boolean LoadToPlanTable(DateTime rptPeriod, string teStep, string agr, DateTime plnDay, decimal plnValue, string adInstr)
    {
      const string stmtSql = "INSERT INTO VIZ_PRN.TV_PLNTARGPROD VALUES(:RPT_PERIOD, :TESTEP, :AGR, :PLN_DAY, :PLN_VALUE, :AD_INSTR)";
      var lstPrm = new List<OracleParameter>();

      var prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        ParameterName = "RPT_PERIOD",
        OracleDbType = OracleDbType.Date,
        Value = rptPeriod
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        ParameterName = "TESTEP",
        OracleDbType = OracleDbType.VarChar,
        Size = teStep.Length,
        Value = teStep
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        ParameterName = "AGR",
        OracleDbType = OracleDbType.VarChar,
        Size = agr.Length,
        Value = agr
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.DateTime,
        Direction = ParameterDirection.Input,
        ParameterName = "PLN_DAY",
        OracleDbType = OracleDbType.Date,
        Value = plnDay
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.Decimal,
        Direction = ParameterDirection.Input,
        ParameterName = "PLN_VALUE",
        OracleDbType = OracleDbType.Number,
        Precision = 17,
        Scale = 2,
        Value = plnValue
      };
      lstPrm.Add(prm);

      prm = new OracleParameter
      {
        DbType = DbType.String,
        Direction = ParameterDirection.Input,
        ParameterName = "AD_INSTR",
        OracleDbType = OracleDbType.VarChar,
        Size = agr.Length,
        Value = adInstr
      };
      lstPrm.Add(prm);

      return Odac.ExecuteNonQuery(stmtSql, CommandType.Text, false, lstPrm, true);
    }
    
    private Boolean RunRpt(LoadProdTargetsRptParam prm)
    {
      Boolean result;
      DateTime fistDayWrkPeriod = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
      DateTime lastDayWrkPeriod = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month,  DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));

      string strFirstCell;
      string strTeStep;
      string strAgr;

      /*
      const string strNzpCell = "D5,D10,D13,D24,D35,D38,D52,D54,D56,D73,D100";
      const string strNzpTeStep = "1STROLL,1STCUT,DECARB,2NDROLL,2NDCUT,ISOGO,HTANNBF,IN_HTANNBF,STRANN,UO,IN_SGP";
      const string strNzpAgr = "NZPALL,NZPALL,NZPALL,NZPALL,NZPALL,NZPALL,NZPALL,NZPALL,NZPALL,NZPALL,NZPALL";
      */

      /*берем из конфига
      const string strFirstCell = "E4,E5,E8,E10,E11,E13,E15,E18,E21,E24,E26,E29,E32,E35,E36,E38,E40,E43,E46,E49,E52,E53,E54,E55,E56,E58,E61,E64,E67,E70,E73,E75,E78,E81,E84,E87,E90,E93,E96,E99,E106";
      const string strTeStep = "HRC,1STROLL,1STROLL,1STCUT,1STCUT,DECARB,DECARB,DECARB,DECARB,2NDROLL,2NDROLL,2NDROLL,2NDROLL,2NDCUT,2NDCUT,ISOGO,ISOGO,ISOGO,ISOGO,ISOGO,HTANNBF,HTANNBF,HTANNBF,HTANNBF,STRANN,STRANN,STRANN,STRANN,STRANN,STRANN,UO,UO,UO,UO,UO,UO,UO,UO,UO,TOTALACCEPT,TOTALSHIP";
      const string strAgr = "HRC,NZP,RM1300,NZP,APR1,NZP,ARO1,ARO2,AOO1A,NZP,RM12001,RM12002,RRM,NZP,APR8,NZP,AOO3A,AOO3B,AOO4A,AOO4B,NZP,BAF_PACK,BAF_IN,BAF_UNPACK,NZP,AVO3,AVO4,AVO5,AVO6,AVO7,NZP,APR3,APR4,APR5,APR6,APR9,APR10,APR12,APR14,TOTALACCEPT_GO,TOTALSHIP_GO";
      */

      //Читаем строки из конфиг.файла
      try
      {
        strFirstCell = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + prm.CfgFile, "StrFirstCell");
        strTeStep = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + prm.CfgFile, "StrTeStep");
        strAgr = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + prm.CfgFile, "StrAgr");
      }
      catch (Exception){
        DxInfo.ShowDxBoxInfo("Ошибка конфигурации", "Ошибка при чтении конфигурационных параметров!", MessageBoxImage.Error);
        return true;
      }

      var ofd = new OpenFileDialog
      {
        AddExtension = true,
        DefaultExt = ".xslx",
        Filter = "xlsx file (.xlsx)|*.xlsx"
      };

      if (!ofd.ShowDialog().GetValueOrDefault(false))
        return true;
      
      Workbook workbook = new Workbook();
      // Load a workbook from the file. 
      workbook.LoadDocument(ofd.FileName, DocumentFormat.Xlsx);
      //MessageBox.Show(workbook.Worksheets[0].Cells["B8"].Value.TextValue);


      try{
        result = true;
        ClearRptPeriod(fistDayWrkPeriod);
        
        /*
        //Загрузка НЗП
        var strNzpCellList = strNzpCell.Split(',');
        var strNzpTeStepList = strNzpTeStep.Split(',');
        var strNzpAgrList = strNzpAgr.Split(',');

        for (int i = 0; i < strNzpCellList.Length; i++)
          LoadToPlanTable(fistDayWrkPeriod, strNzpTeStepList[i], strNzpAgrList[i], fistDayWrkPeriod.AddDays(-1), Convert.ToDecimal(workbook.Worksheets[0].Cells[strNzpCellList[i]].Value.NumericValue));
        */

        //Загрузка по-агрегатно
        var strFirstCellList = strFirstCell.Split(',');
        var strTeStepList = strTeStep.Split(',');
        var strAgrList = strAgr.Split(',');
        
        for (var i = 0; i < strFirstCellList.Length; i++){

          var j = 0;
          var row = workbook.Worksheets[0].Cells[strFirstCellList[i]].RowIndex;
          var col = workbook.Worksheets[0].Cells[strFirstCellList[i]].ColumnIndex;

          /*
          if (strFirstCellList[i] == "E53"){
            MessageBox.Show(row.ToString() + "," + col.ToString());
            MessageBox.Show(workbook.Worksheets[0].Cells[row, col + j].Value.NumericValue.ToString());
          }
          */

          for (DateTime dt = fistDayWrkPeriod; dt <= lastDayWrkPeriod; dt = dt.AddDays(1), j++){
            var r = LoadToPlanTable(fistDayWrkPeriod, strTeStepList[i], strAgrList[i], dt, Convert.ToDecimal(workbook.Worksheets[0].Cells[row, col + j].Value.NumericValue), "D");
            if (!r)
              return true;

            //здесь загружаем данные из столбца "План Мес"
            if (dt == lastDayWrkPeriod){
              r = LoadToPlanTable(fistDayWrkPeriod, strTeStepList[i], strAgrList[i], fistDayWrkPeriod, Convert.ToDecimal(workbook.Worksheets[0].Cells[row, col + j + 1].Value.NumericValue), "M");
              if (!r)
                return true;
            }
          }
        }

        InsertEventsToDashBoard(fistDayWrkPeriod);
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка загрузки", ex.Message, MessageBoxImage.Stop)));
        result = false;
      }
      finally{
        result = false;
      }
    
      return result;
    }


  }






}

