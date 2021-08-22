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
  public sealed class CzlPackRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string TechStepInspLot { get; set; }
    //public string TechStepPrjJornal { get; set; }

    public CzlPackRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd, string RptTechStepIl/*, string RptTechStepPj*/)
           : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.TechStepInspLot = RptTechStepIl;
      //this.TechStepPrjJornal = RptTechStepPj;
    }
  }

  public sealed class CzlPack : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      CzlPackRptParam prm = (e.Argument as CzlPackRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        //Здесь формирование самого отчета
        //wrkSheet.Range("A1").Value = prm.ExcelApp.Version;
        //wrkSheet.Range("A2").Value = "asdadsdgsfgsfsg";

        //Здесь визуализация Экселя
        //prm.ExcelApp.ScreenUpdating = true;
        //prm.ExcelApp.Visible = true;
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
        
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(CzlPackRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      string SqlStmt = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        SqlStmt = "SELECT * FROM VIZ_PRN.CZL_PACK2 ORDER BY MLOCID, TSDATE";
        DbVar.SetString(prm.TechStepInspLot);
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);

        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef);
        if (odr == null) return false;

        CurrentWrkSheet.Cells[2, 3].Value2 = $"{dtBegin:dd.MM.yyyy HH:mm:ss}" + " - " + $"{dtEnd:dd.MM.yyyy HH:mm:ss}"; 

        int flds = odr.FieldCount;
        int row = 7;
        
        string prevLocId = null; 
        string curLocId = null;
        int ColorRow = 49407;

        while (odr.Read()){
          curLocId = Convert.ToString(odr.GetValue("MLOCID"));
          CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 187]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 187]]);
          
          if (curLocId == prevLocId){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 187]].Interior.Pattern = 1;//xlSolid
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 187]].Interior.Color = ColorRow;
            //===================
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row - 1, 1], CurrentWrkSheet.Cells[row - 1, 187]].Interior.Pattern = 1;//xlSolid
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row - 1, 1], CurrentWrkSheet.Cells[row - 1, 187]].Interior.Color = ColorRow;
            ColorRow -= 100;
          }  

          prevLocId = curLocId;

          for (int i = 0; i < flds; i++)
            CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);

          row++;
        }

        CurrentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return Result;
    }


  }






}
