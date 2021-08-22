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
  public sealed class CzlEfLsr9tRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public decimal P1750023 { get; set; }
    public decimal P1750027 { get; set; }
    public decimal P1750030 { get; set; }
    public decimal B800 { get; set; }
    public decimal KesiAvg { get; set; }
    public decimal CoefVoln { get; set; }
    public decimal QntShov { get; set; }
    public string Sort { get; set; }
    public string AdgIn { get; set; }
    public string AdgOut { get; set; }
    public int TypeRpt { get; set; }
    public string ListVal { get; set; }
    public string AdgInFlt { get; set; }
    public string AdgOutFlt { get; set; }


    public CzlEfLsr9tRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class CzlEfLsr9t : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as CzlEfLsr9tRptParam);
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

    private void ListFilterInfoToExcel(CzlEfLsrRptParam prm)
    {
      dynamic wrkSheet = null;
      //выбираем лист
      prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
      wrkSheet = prm.ExcelApp.ActiveSheet;
      const int row = 3;
      string[] strArr = prm.ListVal.Split(new char[] {','});
      for (int i = 0; i < strArr.Length; i++) wrkSheet.Cells[row + i, 1].Value = strArr[i];
    }


    private Boolean RunRpt(CzlEfLsr9tRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      IAsyncResult iar = null;


      try{
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        CurrentWrkSheet.Cells[4, 2].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        const string SqlStmt = "SELECT * FROM VIZ_PRN.CZL_ELK9T";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int row = 8;
          int flds = odr.FieldCount;

          while(odr.Read()){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 8]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 8]]);

            for (int i = 0; i < flds; i++){
              if ((i == 0) && (odr.GetInt32(0) == 99))
                CurrentWrkSheet.Cells[row, i + 1].Value = "Всего";
              else  
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            }
            row++;
          }
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





