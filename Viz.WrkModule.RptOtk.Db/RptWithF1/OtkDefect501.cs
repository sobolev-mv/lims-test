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
  public sealed class OtkDefect501RptParam : RptWithF1Param
  {
    public decimal Glubina { get; set; }

    public OtkDefect501RptParam(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class OtkDefect501 : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkDefect501RptParam);
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

    private Boolean RunRpt(OtkDefect501RptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      //IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
     

      try{
        PrepareFilterRpt(prm);
        
        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF501_516";
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        DbVar.SetNum(prm.Glubina);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);


        Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null);

        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        CurrentWrkSheet.Cells[2, 4].Value = prm.Glubina;

        if (prm.TypeFilter >= 1){
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;
        }

        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //var oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_501M");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_DEF_501B");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 516
        prm.ExcelApp.ActiveWorkbook.WorkSheets[9].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        
        if (prm.TypeFilter >= 1){
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;
        }

        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_516M");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_DEF_516B");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 437
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF501_VES";
        prm.ExcelApp.ActiveWorkbook.WorkSheets[7].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_437");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 441
        prm.ExcelApp.ActiveWorkbook.WorkSheets[8].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);        
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_441");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 202
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF501_AGR";
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("AGR");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_202");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_DEF_202ALL");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 602
        prm.ExcelApp.ActiveWorkbook.WorkSheets[3].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
       
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("AGR");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_602");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 603
        prm.ExcelApp.ActiveWorkbook.WorkSheets[4].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
       
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("AGR");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_603");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 604
        prm.ExcelApp.ActiveWorkbook.WorkSheets[5].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("AGR");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_604");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_DEF_604ALL");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 607
        prm.ExcelApp.ActiveWorkbook.WorkSheets[6].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 7].Value = string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);

        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 4].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        //oracleCommand = iar.AsyncState as OracleCommand;
        //if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 8;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("AGR");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_DEF_607");
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("VES_DEF_607ALL");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception){
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



