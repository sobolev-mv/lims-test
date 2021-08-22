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
  public sealed class OtkFreqDistrDefectAvoRptParam : RptWithF1Param
  {
    public string Defect { get; set; }
    public OtkFreqDistrDefectAvoRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class OtkFreqDistrDefectAvo : RptWithF1
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkFreqDistrDefectAvoRptParam);
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

        //Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(OtkFreqDistrDefectAvoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        PrepareFilterRpt(prm);
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        const string SqlStmt1 = "SELECT * FROM VIZ_PRN.CZL_RASPRED";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.Defect)));

        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        CurrentWrkSheet.Cells[1, 6].Value = prm.Defect;
        CurrentWrkSheet.Cells[2, 1].Value = string.Format("за период с {0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[4, 2].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt1, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          while (odr.Read()){
            CurrentWrkSheet.Cells[6, 4].Value = odr.GetDecimal("ves_vsego");
            CurrentWrkSheet.Cells[7, 4].Value = odr.GetDecimal("ves_rul");
          }
          odr.Close();
          odr.Dispose();
        }

        const string SqlStmt2 = "SELECT * FROM VIZ_PRN.CZL_RASPRED_TBL";

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt2, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 12;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetDecimal("kolvo");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetDecimal("ves_def");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetDecimal("ves_uch"); 
            row++;
          }
        }

        
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





