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

  public sealed class ShiftRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public string Unit { get; set; }
    public string Team { get; set; }
    public string ShiftForeman { get; set; }
    public string SeniorWorker { get; set; }
    public string Shift { get; set; }
    public int LngId { get; set; }
    public ShiftRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class ShiftRptSl : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as ShiftRptParam);
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

    private Boolean RunRpt(ShiftRptParam prm, dynamic currentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean result = false;
      var oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateBegin, 0);
        DbVar.SetString(prm.Unit, prm.Shift);
        dtBegin = DbVar.GetDateBeginEnd(true, false);
        dtEnd = DbVar.GetDateBeginEnd(false, false);

        currentWrkSheet.Cells[2, 11].Value = $"{prm.DateBegin:dd.MM.yyyy}";
        currentWrkSheet.Cells[2, 14].Value = prm.Shift;
        currentWrkSheet.Cells[2, 15].Value = prm.Unit;
        currentWrkSheet.Cells[2, 17].Value = prm.Team;

        if (prm.LngId == 1){
          //Shift foreman
          currentWrkSheet.Cells[4, 6].Value = "Quality Engineer: " + prm.ShiftForeman;
          //Senior worker
          currentWrkSheet.Cells[4, 13].Value = "Production Engineer: " + prm.SeniorWorker;
        }
        else{
          currentWrkSheet.Cells[4, 6].Value = "Инженер по качеству: " + prm.ShiftForeman;
          currentWrkSheet.Cells[4, 13].Value = "Инженер-технолог: " + prm.SeniorWorker;
        }
        
        int qntInsert = 0;

        const string sqlStmtApr121 = "SELECT * FROM VIZ_PRN.V_ISC_SMRP_SL_OBR";
        odr = Odac.GetOracleReader(sqlStmtApr121, CommandType.Text, false, null, null);
        if (odr != null){

          int inRow1 = 12;
          int inRowInsert1 = 14;
            
          while (odr.Read()){
            if (inRow1 == inRowInsert1){
              currentWrkSheet.Rows[inRow1].Insert();
              currentWrkSheet.Range[currentWrkSheet.Cells[inRow1 + 1, 1], currentWrkSheet.Cells[inRow1 + 1, 19]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow1, 1], currentWrkSheet.Cells[inRow1, 19]]);
              inRowInsert1++;
              qntInsert++;
            }

            currentWrkSheet.Cells[inRow1, 1].Value = odr.GetValue("LOT_NO");
            currentWrkSheet.Cells[inRow1, 2].Value = odr.GetValue("COIL_NO");
            currentWrkSheet.Cells[inRow1, 3].Value = odr.GetValue("SORT");
            currentWrkSheet.Cells[inRow1, 4].Value = odr.GetValue("THICKNESS");
            currentWrkSheet.Cells[inRow1, 5].Value = odr.GetValue("WIDTH");
            currentWrkSheet.Cells[inRow1, 6].Value = odr.GetValue("WEIGHT");
            currentWrkSheet.Cells[inRow1, 8].Value = odr.GetValue("RASKROY");
            currentWrkSheet.Cells[inRow1, 13].Value = odr.GetValue("YELD_WEIGHT");
            currentWrkSheet.Cells[inRow1, 14].Value = odr.GetValue("WIDTH_OUT");
            currentWrkSheet.Cells[inRow1, 15].Value = odr.GetValue("EDGE_CROP");
            currentWrkSheet.Cells[inRow1, 16].Value = odr.GetValue("CROSS_CUT");
            currentWrkSheet.Cells[inRow1, 17].Value = odr.GetValue("RESIDUES");
            currentWrkSheet.Cells[inRow1, 18].Value = odr.GetValue("TXTCOMMENT");
            inRow1++;
              
         }

          odr.Close();
          odr.Dispose();
        }

        const string sqlStmtApr122 = "SELECT * FROM VIZ_PRN.V_ISC_SMRP_SL_ISH";
        odr = Odac.GetOracleReader(sqlStmtApr122, CommandType.Text, false, null, null);
        if (odr != null){

          int inRow2 = 27 + qntInsert;
          int inRowInsert2 = 29 + qntInsert;

          while (odr.Read()){
            if (inRow2 == inRowInsert2){
              currentWrkSheet.Rows[inRow2].Insert();
              currentWrkSheet.Range[currentWrkSheet.Cells[inRow2 + 1, 1], currentWrkSheet.Cells[inRow2 + 1, 8]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRow2, 1], currentWrkSheet.Cells[inRow2, 8]]);
              inRowInsert2++;
              qntInsert++;
            }

            currentWrkSheet.Cells[inRow2, 1].Value = odr.GetValue(0);
            currentWrkSheet.Cells[inRow2, 2].Value = odr.GetValue(1);
            currentWrkSheet.Cells[inRow2, 3].Value = odr.GetValue(2);
            currentWrkSheet.Cells[inRow2, 4].Value = odr.GetValue(3);
            currentWrkSheet.Cells[inRow2, 5].Value = odr.GetValue(4);
            currentWrkSheet.Cells[inRow2, 7].Value = odr.GetValue(5);
            inRow2++;
          }

          odr.Close();
          odr.Dispose();
        }


        const string sqlStmtApr12Futer2 = "SELECT * FROM VIZ_PRN.V_ISC_SMRP_DT";
        odr = Odac.GetOracleReader(sqlStmtApr12Futer2, CommandType.Text, false, null, null);
        if (odr != null){

          int inRowFuter2 = 36 + qntInsert;
          int inRowInsertFuter2 = 38 + qntInsert;

          while (odr.Read()){
            if (inRowFuter2 == inRowInsertFuter2){
              currentWrkSheet.Rows[inRowFuter2].Insert();
              currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter2 + 1, 1], currentWrkSheet.Cells[inRowFuter2 + 1, 17]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[inRowFuter2, 1], currentWrkSheet.Cells[inRowFuter2, 17]]);
              inRowInsertFuter2++;
              //qntInsert++;
            }

            currentWrkSheet.Cells[inRowFuter2, 1].Value = odr.GetValue(0);
            currentWrkSheet.Cells[inRowFuter2, 3].Value = odr.GetValue(1);
            currentWrkSheet.Cells[inRowFuter2, 4].Value = odr.GetValue(2);
            currentWrkSheet.Cells[inRowFuter2, 8].Value = odr.GetValue(3);
            inRowFuter2++;
          }
          odr.Close();
          odr.Dispose();
        }

        const string sqlStmtApr12Futer3 = "SELECT * FROM VIZ_PRN.V_ISC_SMRP_FIO";
        odr = Odac.GetOracleReader(sqlStmtApr12Futer3, CommandType.Text, false, null, null);
        if (odr != null){
          int inRowFuter3 = 36 + qntInsert;

           while (odr.Read()){
             currentWrkSheet.Cells[inRowFuter3, 13].Value = odr.GetValue(0);
             currentWrkSheet.Cells[inRowFuter3, 14].Value = odr.GetValue(1);
             currentWrkSheet.Cells[inRowFuter3, 16].Value = odr.GetValue(2);
             inRowFuter3++;
           }

           odr.Close();
           odr.Dispose();
        }

        //currentWrkSheet.PageSetup.PrintArea = "$A$1:$U$" + (50 + qntInsertAll).ToString();
        //currentWrkSheet.Cells[1, 5].Select();
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
