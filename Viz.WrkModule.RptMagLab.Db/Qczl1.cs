using System;
using System.Collections.Generic;
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

namespace Viz.WrkModule.RptMagLab.Db
{
  public sealed class QCzl1RptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string   Rm1200 { get; set; }
    public string   Aro { get; set; }
    public string   Aoo { get; set; }
    public string   Avo { get; set; }
    public string   Apr { get; set; }
    public Boolean  IsRm1200 { get; set; }
    public Boolean  IsAro { get; set; }
    public Boolean  IsAoo { get; set; }
    public Boolean  IsAvo { get; set; }
    public Boolean  IsApr { get; set; }



    public QCzl1RptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd, 
                         string Rm1200, string Aro, string Aoo, string Avo, string Apr,
                         Boolean IsRm1200, Boolean IsAro, Boolean IsAoo, Boolean IsAvo, Boolean IsApr)
           : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.Rm1200 = Rm1200;
      this.Aro = Aro;
      this.Aoo = Aoo;
      this.Avo = Avo;
      this.Apr = Apr;
      this.IsRm1200 = IsRm1200;
      this.IsAro = IsAro;
      this.IsAoo = IsAoo;
      this.IsAvo = IsAvo;
      this.IsApr = IsApr;
    }
  }

  public sealed class QCzl1 : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as QCzl1RptParam);
      dynamic wrkSheet = null;

      try{

        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally{
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(QCzl1RptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      int row = 0;

      try{
        const string SqlStmt1 = "SELECT * FROM VIZ_PRN.CZL_MARKA_FINCUT_PROC ORDER BY 1";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString("FINCUT", "PACK", prm.Rm1200, prm.Aro, prm.Aoo, prm.Avo, prm.Apr)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[1, 2].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
        string strFlt = "";
        if (prm.IsRm1200)
          strFlt += ":Ст1200 =" + prm.Rm1200;
        if (prm.IsAro)
          strFlt += ":АРО=" + prm.Aro;
        if (prm.IsAoo)
          strFlt += ":АОО=" + prm.Aoo;
        if (prm.IsAvo)
          strFlt += ":АВО=" + prm.Avo;
        if (prm.IsApr)
          strFlt += ":АПР=" + prm.Apr;
        CurrentWrkSheet.Cells[3, 3].Value = strFlt;

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt1, System.Data.CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          row = 12;

          while (odr.Read()){
            if (Convert.ToInt32(odr.GetValue("TOLS")) != 99){
              CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
              CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
              CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("B100AP_X");
              CurrentWrkSheet.Cells[row + 1, 5].Value = odr.GetValue("STD100");
              CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("B800AP_X");
              CurrentWrkSheet.Cells[row + 1, 6].Value = odr.GetValue("STD800");
              CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B2500AP_X");
              CurrentWrkSheet.Cells[row + 1, 7].Value = odr.GetValue("STD2500");
              CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1550AP_X");
              CurrentWrkSheet.Cells[row + 1, 8].Value = odr.GetValue("STD1550");
              CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("P1750AP_X");
              CurrentWrkSheet.Cells[row + 1, 9].Value = odr.GetValue("STD1750");

              CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("PROC_8_9");
              CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("PROC9");
              CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("PROC_8");
              CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("PROC_7");
              CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("PROC_6");
              CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("PROC_5");
              CurrentWrkSheet.Cells[row, 16].Value = odr.GetValue("PROC_NM");
            }
            else{
              CurrentWrkSheet.Cells[20, 3].Value = odr.GetValue("VES");
              CurrentWrkSheet.Cells[20, 10].Value = odr.GetValue("PROC_8_9");
              CurrentWrkSheet.Cells[20, 11].Value = odr.GetValue("PROC9");
              CurrentWrkSheet.Cells[20, 12].Value = odr.GetValue("PROC_8");
              CurrentWrkSheet.Cells[20, 13].Value = odr.GetValue("PROC_7");
              CurrentWrkSheet.Cells[20, 14].Value = odr.GetValue("PROC_6");
              CurrentWrkSheet.Cells[20, 15].Value = odr.GetValue("PROC_5");
              CurrentWrkSheet.Cells[20, 16].Value = odr.GetValue("PROC_NM");
            }

            row++;
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        const string SqlStmt2 = "SELECT * FROM VIZ_PRN.CZL_MARKA_PACK_PROC ORDER BY 1";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt2, System.Data.CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          row = 29;

          while (odr.Read()){
            if (Convert.ToInt32(odr.GetValue("TOLS")) != 99){
              CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
              CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
              CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("B100AP_X");
              CurrentWrkSheet.Cells[row + 1, 5].Value = odr.GetValue("STD100");
              CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("B800AP_X");
              CurrentWrkSheet.Cells[row + 1, 6].Value = odr.GetValue("STD800");
              CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("B2500AP_X");
              CurrentWrkSheet.Cells[row + 1, 7].Value = odr.GetValue("STD2500");
              CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("P1550AP_X");
              CurrentWrkSheet.Cells[row + 1, 8].Value = odr.GetValue("STD1550");
              CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("P1750AP_X");
              CurrentWrkSheet.Cells[row + 1, 9].Value = odr.GetValue("STD1750");

              CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("PROC_8_9");
              CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("PROC9");
              CurrentWrkSheet.Cells[row, 12].Value = odr.GetValue("PROC_8");
              CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("PROC_7");
              CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("PROC_6");
              CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("PROC_5");
              CurrentWrkSheet.Cells[row, 16].Value = odr.GetValue("PROC_NM");
            }
            else{
              CurrentWrkSheet.Cells[37, 3].Value = odr.GetValue("VES");
              CurrentWrkSheet.Cells[37, 10].Value = odr.GetValue("PROC_8_9");
              CurrentWrkSheet.Cells[37, 11].Value = odr.GetValue("PROC9");
              CurrentWrkSheet.Cells[37, 12].Value = odr.GetValue("PROC_8");
              CurrentWrkSheet.Cells[37, 13].Value = odr.GetValue("PROC_7");
              CurrentWrkSheet.Cells[37, 14].Value = odr.GetValue("PROC_6");
              CurrentWrkSheet.Cells[37, 15].Value = odr.GetValue("PROC_5");
              CurrentWrkSheet.Cells[37, 16].Value = odr.GetValue("PROC_NM");
            }

            row++;
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //Возвращаемся на первую страницу
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      finally{
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return Result;
    }


  }






}



