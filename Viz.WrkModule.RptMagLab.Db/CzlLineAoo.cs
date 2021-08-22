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

namespace Viz.WrkModule.RptMagLab.Db
{
  public sealed class CzlLineAooRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string TypeLineAoo { get; set; }
    public int TypeMat { get; set; }

    public CzlLineAooRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class CzlLineAoo : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      CzlLineAooRptParam prm = (e.Argument as CzlLineAooRptParam);
      dynamic wrkSheet = null;

      try
      {
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
      catch (Exception ex)
      {
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart) (() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally
      {
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

    private Boolean RunRpt(CzlLineAooRptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      string SqlStmt = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      int lenTab;      

      try
      {
        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart) (() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart) (() => DbVar.SetString(prm.TypeLineAoo)));

        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart) (() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart) (() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        SqlStmt = "VIZ_PRN.CZL_CMPLINE_AOO.PrepareData";
        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart)
            (() => { iar = Odac.ExecuteNonQueryAsync(SqlStmt, CommandType.StoredProcedure, false, false, null); }));

        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
        {
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }

        if (prm.TypeMat == 1){
          CurrentWrkSheet.Cells[2, 9].Value = prm.TypeLineAoo;
          CurrentWrkSheet.Cells[3, 3].Value = " c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
          SqlStmt = "SELECT * FROM VIZ_PRN.CZL_CMPLINE_AOO_STEND";
          lenTab = 5;
        }
        else{
          CurrentWrkSheet.Cells[2, 5].Value = prm.TypeLineAoo;
          CurrentWrkSheet.Cells[3, 2].Value = " c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
          SqlStmt = "SELECT * FROM VIZ_PRN.CZL_CMPLINE_AOO_RULON";
          lenTab = 7;
        }

        prm.Disp.Invoke(DispatcherPriority.Normal,
          (ThreadStart) (() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null)
        {
          var row = 6;
          var flds = odr.FieldCount;

          while (odr.Read())
          {
            if (row >= 7)
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, lenTab]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, lenTab]]);

            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }
        }

        CurrentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart) (() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
        Result = false;
      }
      finally
      {
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


