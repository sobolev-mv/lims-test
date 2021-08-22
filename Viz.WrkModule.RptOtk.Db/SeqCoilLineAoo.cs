using System;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class SeqCoilLineAooRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public SeqCoilLineAooRptParam(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class SeqCoilLineAoo : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as SeqCoilLineAooRptParam);
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
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        /*
        if (prm.WorkBook != null)
          Marshal.ReleaseComObject(prm.WorkBook);
        */

        if (prm.ExcelApp != null)
          Marshal.ReleaseComObject(prm.ExcelApp);

        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(SeqCoilLineAooRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_LINE_AOO";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[2, 1].Value = string.Format("за период с {0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          string agr = null;

          int rowAoo3a = 6;
          int rowAoo3b = 6;
          int rowAoo4a = 6;
          int rowAoo4b = 6;

        
          var flds = odr.FieldCount;

          while (odr.Read()){
            agr = odr.GetString(0);

            if (agr == "AOO3A"){
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo3a, 1], CurrentWrkSheet.Cells[rowAoo3a, 4]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo3a + 1, 1], CurrentWrkSheet.Cells[rowAoo3a + 1, 4]]);
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[rowAoo3a, i].Value = odr.GetValue(i);

              rowAoo3a++;
            } else if (agr == "AOO3B"){
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo3b, 5], CurrentWrkSheet.Cells[rowAoo3b, 8]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo3b + 1, 5], CurrentWrkSheet.Cells[rowAoo3b + 1, 8]]);
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[rowAoo3b, i + 4].Value = odr.GetValue(i);

              rowAoo3b++;
            } else if (agr == "AOO4A"){
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo4a, 9], CurrentWrkSheet.Cells[rowAoo4a, 12]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo4a + 1, 9], CurrentWrkSheet.Cells[rowAoo4a + 1, 12]]);
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[rowAoo4a, i + 8].Value = odr.GetValue(i);

              rowAoo4a++;
            } else if (agr == "AOO4B"){
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo4b, 13], CurrentWrkSheet.Cells[rowAoo4b, 16]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowAoo4b + 1, 13], CurrentWrkSheet.Cells[rowAoo4b + 1, 16]]);
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[rowAoo4b, i + 12].Value = odr.GetValue(i);

              rowAoo4b++;
            }
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception e){
        MessageBox.Show(e.Message);
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

