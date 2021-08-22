using System;
using System.Data;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.Isc.Db
{
  public sealed class VizPruductRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public Boolean IsDateShippingChoose { get; set; }
    public Boolean IsManufacturerChoose { get; set; }
    public Boolean IsContractNoChoose { get; set; }
    public int MnfIdValue { get; set; }
    public string ContractNoValue { get; set; }
    public Boolean IsSpecificationChoose { get; set; }
    public string SpecificationValue { get; set; }
    public VizPruductRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class VizPruduct : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as VizPruductRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm, "Isc Reports");
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally{
        prm.WorkBook.Close();
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        if (prm.ExcelApp != null)
          Marshal.ReleaseComObject(prm.ExcelApp);

        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(VizPruductRptParam prm, dynamic currentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean result = false;
      var oef = new OdacErrorInfo();
 
      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateBegin, 0);
        currentWrkSheet.Cells[2, 1].Value = $"Date report from {prm.DateBegin:dd.MM.yyyy} to {prm.DateEnd:dd.MM.yyyy}";
       
        const string sqlStmtRpt = "SELECT * FROM VIZ_PRN.V_ISC_SPP_RPT_ALL " +
                                  "WHERE ((DATE_SHIPPING BETWEEN :DSH1 AND :DSH2) OR (:ISDSH = 0)) " +
                                  "AND ((CONTRACTNO = :CNTNO) OR (:ISCNTNO = 0)) " +
                                  "AND ((SPECNO = :SPCNO) OR (:ISSPCNO = 0)) " +
                                  "AND ((TYP_MNF = :TYPMNF) OR (:ISTYPMNF = 0))";

        var lstParam = new List<OracleParameter>();

        var param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "DSH1",
          Value = prm.DateBegin
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "DSH2",
          Value = prm.DateEnd
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISDSH",
          Value = prm.IsDateShippingChoose ? 1 : 0
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "CNTNO",
          Value = prm.ContractNoValue
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISCNTNO",
          Value = prm.IsContractNoChoose ? 1 : 0
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "SPCNO",
          Value = prm.SpecificationValue
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISSPCNO",
          Value = prm.IsSpecificationChoose ? 1 : 0
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "TYPMNF",
          SourceColumnNullMapping = false,
          Value = prm.MnfIdValue
        };
        lstParam.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISTYPMNF",
          SourceColumnNullMapping = false,
          Value = prm.IsManufacturerChoose ? 1 : 0
        };
        lstParam.Add(param);

        odr = Odac.GetOracleReader(sqlStmtRpt, CommandType.Text, false, lstParam, null);
        if (odr != null){

          const int firstExcelColumn = 1;
          const int lastExcelColumn = 42;
          var row = 6;
          //var flds = odr.FieldCount;


          while (odr.Read()){
            currentWrkSheet.Range[currentWrkSheet.Cells[row, firstExcelColumn], currentWrkSheet.Cells[row, lastExcelColumn]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[row + 1, firstExcelColumn], currentWrkSheet.Cells[row + 1, lastExcelColumn]]);
            currentWrkSheet.Cells[row, 1].Value = odr.GetValue("ME_ID");
            currentWrkSheet.Cells[row, 2].Value = odr.GetValue("DATE_SHIPPING");
            currentWrkSheet.Cells[row, 3].Value = odr.GetValue("NAME_MNF");

            currentWrkSheet.Cells[row, 4].Value = odr.GetValue("CONTRACTNO");
            currentWrkSheet.Cells[row, 5].Value = odr.GetValue("SPECNO");
            currentWrkSheet.Cells[row, 6].Value = odr.GetValue("NET");
            currentWrkSheet.Cells[row, 7].Value = odr.GetValue("GROSS");
            currentWrkSheet.Cells[row, 8].Value = odr.GetValue("THICKNESS");
            currentWrkSheet.Cells[row, 9].Value = odr.GetValue("WIDTH");
            currentWrkSheet.Cells[row, 10].Value = odr.GetValue("P1550AP");
            currentWrkSheet.Cells[row, 11].Value = odr.GetValue("P1750AP");
            currentWrkSheet.Cells[row, 12].Value = odr.GetValue("P1750LST");
            currentWrkSheet.Cells[row, 13].Value = odr.GetValue("B800LST");
            currentWrkSheet.Cells[row, 14].Value = odr.GetValue("B800AP");
            currentWrkSheet.Cells[row, 15].Value = odr.GetValue("P1550LST");

            currentWrkSheet.Cells[row, 16].Value = odr.GetValue("NUMOFWELDS");
            currentWrkSheet.Cells[row, 17].Value = odr.GetValue("HEATNO");
            currentWrkSheet.Cells[row, 18].Value = odr.GetValue("STOGRADE");
            currentWrkSheet.Cells[row, 19].Value = odr.GetValue("KESIAVG");
            currentWrkSheet.Cells[row, 20].Value = odr.GetValue("GIB");
            currentWrkSheet.Cells[row, 21].Value = odr.GetValue("PLACEMENT_NUM");
            currentWrkSheet.Cells[row, 22].Value = odr.GetValue("ANNEALINGLOT");
            currentWrkSheet.Cells[row, 23].Value = odr.GetValue("GRADE");
            currentWrkSheet.Cells[row, 24].Value = odr.GetValue("STANDART");
            currentWrkSheet.Cells[row, 25].Value = odr.GetValue("UPOS_NRIST");
            currentWrkSheet.Cells[row, 26].Value = odr.GetValue("CERT_NR");
            currentWrkSheet.Cells[row, 27].Value = odr.GetValue("CERT_POS");
            currentWrkSheet.Cells[row, 28].Value = odr.GetValue("C");
            currentWrkSheet.Cells[row, 29].Value = odr.GetValue("SI");
            currentWrkSheet.Cells[row, 30].Value = odr.GetValue("AL");
            currentWrkSheet.Cells[row, 31].Value = odr.GetValue("VYSOTA_VOLN");
            currentWrkSheet.Cells[row, 32].Value = odr.GetValue("DLINA_VOLN");
            currentWrkSheet.Cells[row, 33].Value = odr.GetValue("KOEF_VOLN");
            currentWrkSheet.Cells[row, 34].Value = odr.GetValue("SR_FACT_TOLS");
            currentWrkSheet.Cells[row, 35].Value = odr.GetValue("OTKL_FACT_TOLS_PLUS");
            currentWrkSheet.Cells[row, 36].Value = odr.GetValue("OTKL_FACT_TOLS_MINUS");
            currentWrkSheet.Cells[row, 37].Value = odr.GetValue("ZAUS");
            currentWrkSheet.Cells[row, 38].Value = odr.GetValue("COAT");
            currentWrkSheet.Cells[row, 39].Value = odr.GetValue("KONOSAMENT");
            currentWrkSheet.Cells[row, 40].Value = odr.GetValue("DATE_CERT");
            currentWrkSheet.Cells[row, 41].Value = odr.GetValue("NUM_CONTAINER");
            currentWrkSheet.Cells[row, 42].Value = odr.GetValue("CTG");
            /*
            for (int i = 0; i < flds; i++)
              currentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            */

            row++;
          }
        }

        result = true;
      }
      catch (Exception e){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.Message, MessageBoxImage.Stop)));
        result = false;
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


  }



}
