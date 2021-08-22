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

namespace Viz.WrkModule.RptOpr.Db
{
  public sealed class ReasonSettleMetalRptParam : Smv.Xls.XlsInstanceParam
  {
    public Boolean IsGroupDateRangeAvoF2;
    public Boolean IsGroupDateRangeUoF2;
    public DateTime DateIncomplProd1;
    public DateTime DateIncomplProd2;
    public DateTime DateRangeBeginAvoF2;
    public DateTime DateRangeEndAvoF2;
    public DateTime DateRangeBeginUoF2;
    public DateTime DateRangeEndUoF2;
    public ReasonSettleMetalRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class ReasonSettleMetal : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as ReasonSettleMetalRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);
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

    private Boolean RunRpt(ReasonSettleMetalRptParam prm, dynamic currentWrkSheet)
    {
      var result = false;
      OracleDataReader odr = null;
      DateTime? dtIncomplProd1, dtIncomplProd2, dtAvoBegin, dtAvoEnd, dtUoBegin, dtUoEnd;

      try{
        DbVar.SetRangeDate(prm.DateIncomplProd1, prm.DateIncomplProd1, 1);
        dtIncomplProd1 = DbVar.GetDateBeginEnd(true, true);

        DbVar.SetRangeDate(prm.DateIncomplProd2, prm.DateIncomplProd2, 1);
        dtIncomplProd2 = DbVar.GetDateBeginEnd(true, true);

        DbVar.SetRangeDate(prm.DateRangeBeginAvoF2, prm.DateRangeEndAvoF2, 1);
        dtAvoBegin = DbVar.GetDateBeginEnd(true, true);
        dtAvoEnd = DbVar.GetDateBeginEnd(false, true);

        DbVar.SetRangeDate(prm.DateRangeBeginUoF2, prm.DateRangeEndUoF2, 1);
        dtUoBegin = DbVar.GetDateBeginEnd(true, true);
        dtUoEnd = DbVar.GetDateBeginEnd(false, true);

        var hdr = "";

        if ((!prm.IsGroupDateRangeAvoF2) && (!prm.IsGroupDateRangeUoF2))
          hdr = "Объединенные причины осевшего металла с " + $"{dtIncomplProd1:dd.MM.yyyy HH:mm}" + " по " + $"{dtIncomplProd2:dd.MM.yyyy HH:mm}";
        else if ((prm.IsGroupDateRangeAvoF2) && (!prm.IsGroupDateRangeUoF2))
          hdr = "Объединенные причины осевшего металла с " + $"{dtIncomplProd1:dd.MM.yyyy HH:mm}" + " по " + $"{dtIncomplProd2:dd.MM.yyyy HH:mm}" + ", обработанного на АВО с " + $"{dtAvoBegin:dd.MM.yyyy HH:mm}" + " по " + $"{dtAvoEnd:dd.MM.yyyy HH:mm}";
        else if ((!prm.IsGroupDateRangeAvoF2) && (prm.IsGroupDateRangeUoF2))
          hdr = "Объединенные причины осевшего металла с " + $"{dtIncomplProd1:dd.MM.yyyy HH:mm}" + " по " + $"{dtIncomplProd2:dd.MM.yyyy HH:mm}" + ", обработанного на УО с " + $"{dtUoBegin:dd.MM.yyyy HH:mm}" + " по " + $"{dtUoEnd:dd.MM.yyyy HH:mm}";
        else if ((prm.IsGroupDateRangeAvoF2) && (prm.IsGroupDateRangeUoF2))
          hdr = "Объединенные причины осевшего металла с " + $"{dtIncomplProd1:dd.MM.yyyy HH:mm}" + " по " + $"{dtIncomplProd2:dd.MM.yyyy HH:mm}" + ", обработанного на АВО с " + $"{dtAvoBegin:dd.MM.yyyy HH:mm}" + " по " + $"{dtAvoEnd:dd.MM.yyyy HH:mm}" + " и обработанного на УО с " + $"{dtUoBegin:dd.MM.yyyy HH:mm}" + " по " + $"{dtUoEnd:dd.MM.yyyy HH:mm}";

        currentWrkSheet.Cells[1, 2].Value = hdr;

        const string sqlStmt1 = "VIZ_PRN.OS_MET.PreReasonSettleMetal";
        var lstOraPrm = new List<OracleParameter>()
        {
          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value = prm.DateIncomplProd1
          },

          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value = prm.DateIncomplProd2
          },

          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value = prm.IsGroupDateRangeAvoF2 ? prm.DateRangeBeginAvoF2 : (null as DateTime?)
          },

          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value =  prm.IsGroupDateRangeAvoF2 ? prm.DateRangeEndAvoF2 : (null as DateTime?)
          },

          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value = prm.IsGroupDateRangeUoF2 ? prm.DateRangeBeginUoF2 : (null as DateTime?)
          },

          new OracleParameter()
          {
            DbType = DbType.DateTime,
            Direction = ParameterDirection.Input,
            OracleDbType = OracleDbType.Date,
            Value =  prm.IsGroupDateRangeUoF2 ? prm.DateRangeEndUoF2 : (null as DateTime?)
          }

        };

        if (!Odac.ExecuteNonQuery(sqlStmt1, CommandType.StoredProcedure, false, lstOraPrm)) 
          throw new NotImplementedException("StoredProcedure: VIZ_PRN.OS_MET.PreReasonSettleMetal");


        string[] sqlStmtArr = new[] { "SELECT * FROM VIZ_PRN.V_NZP WHERE PR_NZP = '1' ORDER BY TOLS", "SELECT * FROM VIZ_PRN.V_NZP WHERE PR_NZP = '2' ORDER BY TOLS" };
        int[] rowArr = new[] {7, 20};

        for (int i = 0; i < 2; i++){

          odr = Odac.GetOracleReader(sqlStmtArr[i], CommandType.Text, false, null, null);
          if (odr != null){

            while (odr.Read()){
              currentWrkSheet.Cells[rowArr[i], 2].Value = odr.GetValue("S12_UP");
              currentWrkSheet.Cells[rowArr[i], 3].Value = odr.GetValue("S12_ICT");
              currentWrkSheet.Cells[rowArr[i], 4].Value = odr.GetValue("S12_IUK");
              currentWrkSheet.Cells[rowArr[i], 5].Value = odr.GetValue("S12_IRW");

              currentWrkSheet.Cells[rowArr[i], 7].Value = odr.GetValue("S23_UP");
              currentWrkSheet.Cells[rowArr[i], 8].Value = odr.GetValue("S23_ICT");
              currentWrkSheet.Cells[rowArr[i], 9].Value = odr.GetValue("S23_IUK");
              currentWrkSheet.Cells[rowArr[i], 10].Value = odr.GetValue("S23_IRW");

              currentWrkSheet.Cells[rowArr[i], 12].Value = odr.GetValue("S24_UP");
              currentWrkSheet.Cells[rowArr[i], 13].Value = odr.GetValue("S24_ICT");
              currentWrkSheet.Cells[rowArr[i], 14].Value = odr.GetValue("S24_IUK");
              currentWrkSheet.Cells[rowArr[i], 15].Value = odr.GetValue("S24_IRW");

              currentWrkSheet.Cells[rowArr[i], 17].Value = odr.GetValue("S30_OS");
              currentWrkSheet.Cells[rowArr[i], 18].Value = odr.GetValue("S30_DEF");
              currentWrkSheet.Cells[rowArr[i], 19].Value = odr.GetValue("S30_MRK");
              currentWrkSheet.Cells[rowArr[i], 20].Value = odr.GetValue("S30_ADG");

              currentWrkSheet.Cells[rowArr[i], 22].Value = odr.GetValue("VES");
              rowArr[i]++;
            }
            odr.Close();
            odr.Dispose();
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[3].Select();
        currentWrkSheet = prm.ExcelApp.ActiveSheet;

        sqlStmtArr = new[] { "SELECT * FROM VIZ_PRN.V_NZP_PRICH WHERE PR_NZP = '1' ORDER BY TOLS", "SELECT * FROM VIZ_PRN.V_NZP_PRICH WHERE PR_NZP = '2' ORDER BY TOLS" };
        rowArr = new[] { 7, 20 };

        for (int i = 0; i < 2; i++){

          odr = Odac.GetOracleReader(sqlStmtArr[i], CommandType.Text, false, null, null);
          if (odr != null){

            while (odr.Read()){
              currentWrkSheet.Cells[rowArr[i], 2].Value = odr.GetValue("S12_UP");
              currentWrkSheet.Cells[rowArr[i], 3].Value = odr.GetValue("S12_ICT");
              currentWrkSheet.Cells[rowArr[i], 4].Value = odr.GetValue("S12_RW");
              currentWrkSheet.Cells[rowArr[i], 5].Value = odr.GetValue("S12_DB");
              currentWrkSheet.Cells[rowArr[i], 6].Value = odr.GetValue("S12_ILST");
              currentWrkSheet.Cells[rowArr[i], 7].Value = odr.GetValue("S12_IUK");
              currentWrkSheet.Cells[rowArr[i], 8].Value = odr.GetValue("S12_IRW");
              currentWrkSheet.Cells[rowArr[i], 9].Value = odr.GetValue("S12_NMR");
              currentWrkSheet.Cells[rowArr[i], 10].Value = odr.GetValue("S12_PDB");

              currentWrkSheet.Cells[rowArr[i], 12].Value = odr.GetValue("S23_UP");
              currentWrkSheet.Cells[rowArr[i], 13].Value = odr.GetValue("S23_ICT");
              currentWrkSheet.Cells[rowArr[i], 14].Value = odr.GetValue("S23_RW");
              currentWrkSheet.Cells[rowArr[i], 15].Value = odr.GetValue("S23_DB");
              currentWrkSheet.Cells[rowArr[i], 16].Value = odr.GetValue("S23_ILST");
              currentWrkSheet.Cells[rowArr[i], 17].Value = odr.GetValue("S23_IUK");
              currentWrkSheet.Cells[rowArr[i], 18].Value = odr.GetValue("S23_IRW");
              currentWrkSheet.Cells[rowArr[i], 19].Value = odr.GetValue("S23_NMR");
              currentWrkSheet.Cells[rowArr[i], 20].Value = odr.GetValue("S23_PDB");
              
              currentWrkSheet.Cells[rowArr[i], 22].Value = odr.GetValue("S24_UP");
              currentWrkSheet.Cells[rowArr[i], 23].Value = odr.GetValue("S24_ICT");
              currentWrkSheet.Cells[rowArr[i], 24].Value = odr.GetValue("S24_RW");
              currentWrkSheet.Cells[rowArr[i], 25].Value = odr.GetValue("S24_DB");
              currentWrkSheet.Cells[rowArr[i], 26].Value = odr.GetValue("S24_ILST");
              currentWrkSheet.Cells[rowArr[i], 27].Value = odr.GetValue("S24_IUK");
              currentWrkSheet.Cells[rowArr[i], 28].Value = odr.GetValue("S24_IRW");
              currentWrkSheet.Cells[rowArr[i], 29].Value = odr.GetValue("S24_NMR");
              currentWrkSheet.Cells[rowArr[i], 30].Value = odr.GetValue("S24_PDB");

              currentWrkSheet.Cells[rowArr[i], 32].Value = odr.GetValue("S30_OS");
              currentWrkSheet.Cells[rowArr[i], 33].Value = odr.GetValue("S30_DEF");
              currentWrkSheet.Cells[rowArr[i], 34].Value = odr.GetValue("S30_MRK");
              currentWrkSheet.Cells[rowArr[i], 35].Value = odr.GetValue("S30_ADG");

              currentWrkSheet.Cells[rowArr[i], 37].Value = odr.GetValue("VES");
              rowArr[i]++;
            }
            odr.Close();
            odr.Dispose();
          }
        }

        sqlStmtArr = new[] { "SELECT * FROM VIZ_PRN.V_NZP_RUL WHERE PR_NZP = '1'", "SELECT * FROM VIZ_PRN.V_NZP_RUL WHERE PR_NZP = '2'" };
        int[] sheetIdxArr = new[] { 4, 5 };


        for (int i = 0; i < 2; i++){

          odr = Odac.GetOracleReader(sqlStmtArr[i], CommandType.Text, false, null, null);
          if (odr != null){

            const int firstExcelColumn = 1;
            const int lastExcelColumn = 9;
            var row = 2;
            var flds = odr.FieldCount;

            prm.ExcelApp.ActiveWorkbook.WorkSheets[sheetIdxArr[i]].Select();
            currentWrkSheet = prm.ExcelApp.ActiveSheet;

            while (odr.Read()){
              currentWrkSheet.Range[currentWrkSheet.Cells[row, firstExcelColumn], currentWrkSheet.Cells[row, lastExcelColumn]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[row + 1, firstExcelColumn], currentWrkSheet.Cells[row + 1, lastExcelColumn]]);

              for (int j = 1; j < lastExcelColumn + 1; j++)
                currentWrkSheet.Cells[row, j].Value = odr.GetValue(j);

              row++;
            }
            odr.Close();
            odr.Dispose();
          }
        }
        
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        result = true;
      }
      catch (Exception e){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.Message, MessageBoxImage.Stop)));
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
