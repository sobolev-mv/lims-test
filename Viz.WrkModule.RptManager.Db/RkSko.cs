using System;
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

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class RkSkoRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }


    public  RkSkoRptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd) : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
    }
  }

  public sealed class RkSko : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      RkSkoRptParam prm = (e.Argument as RkSkoRptParam);
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
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
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

    private Boolean RunRpt(RkSkoRptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_PK_CKO_PRN";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        CurrentWrkSheet.Cells[2, 7].Value = string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int[] exelRow = new[] { 6, 12 };

          while (odr.Read()){

            for (int i = 0, j = 4; i < 20; i++, j++)
              CurrentWrkSheet.Cells[exelRow[0], j].Value = odr.GetValue(i);

            for (int i = 20, j = 4; i < 40; i++, j++)
              CurrentWrkSheet.Cells[exelRow[1], j].Value = odr.GetValue(i);

            for (int i = 0; i < 2; i++)
              exelRow[i]++;
          }

          odr.Close();
          odr.Dispose();
        }

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_PK_CKO_PRN_UO";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int[] exelRow = new[] { 19, 27, 35 };

          while (odr.Read()){

            for (int i = 0, j = 4; i < 20; i++, j++)
              CurrentWrkSheet.Cells[exelRow[0], j].Value = odr.GetValue(i);

            for (int i = 20, j = 4; i < 40; i++, j++)
              CurrentWrkSheet.Cells[exelRow[1], j].Value = odr.GetValue(i);
            
            for (int i = 40, j = 4; i < 50; i++, j++)
              CurrentWrkSheet.Cells[exelRow[2], j].Value = odr.GetValue(i);

            for (int i = 0; i < 3; i++)
              exelRow[i]++;
          }
        }

        CurrentWrkSheet.Cells[1, 1].Select();
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

