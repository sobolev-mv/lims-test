using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class MonitorDefRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public MonitorDefRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class MonitorDef : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as MonitorDefRptParam);
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

    private void FillMonthHeader(dynamic currentWrkSheet)
    {
      const string sqlStmt = "SELECT DATA FROM VIZ_PRN.TMP_DAY ORDER BY NPP";
      OracleDataReader odr = null;
      odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

      if (odr != null){
        int col = 5;
        const int row = 5;

        while (odr.Read())
        {
          //MessageBox.Show(odr.GetValue(0).ToString());
          currentWrkSheet.Cells[row, col].Value = odr.GetValue(0);
          col++;
        }

        odr.Close();
        odr.Dispose();
      }
    }

    private Boolean RunRpt(MonitorDefRptParam prm, dynamic CurrentWrkSheet)
    {
      //IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;

      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      string[] strThickness = {"0.23,0.27,0.30", "0.23", "0.27", "0.30"};

      try{
        Odac.ExecuteNonQuery("VIZ_PRN.QUARTILE_UO1.QRT_PERIOD", CommandType.StoredProcedure, false, null);
        const string sqlStmt = "SELECT * FROM VIZ_PRN.V_QUARTILE_UO1";

        for (int sheetIdx = 0; sheetIdx < 4; sheetIdx++){

          prm.ExcelApp.ActiveWorkbook.WorkSheets[sheetIdx + 1].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

          DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 0);
          DbVar.SetString(strThickness[sheetIdx]);
          dtBegin = DbVar.GetDateBeginEnd(true, false);
          dtEnd = DbVar.GetDateBeginEnd(false, false);

          CurrentWrkSheet.Cells[3, 3].Value = $"с {dtBegin:dd.MM.yyyy} по {dtEnd:dd.MM.yyyy}";

          odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);
          if (odr != null){

            int flds = odr.FieldCount;
            int row = 6;

            while (odr.Read()){

              for (int i = 0; i < flds; i++)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
              row++;
            }

            odr.Close();
            odr.Dispose();
          }

          FillMonthHeader(CurrentWrkSheet);
          CurrentWrkSheet.Cells[5, 10].Value = DateTime.Today.AddDays(-1);
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
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

