using System;
using System.Collections.Generic;
using System.Data;
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
using Smv.Collections.Matrix;

namespace Viz.WrkModule.RptOpr.Db
{
  public sealed class ReasonOfStripBreakageRmAreaRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public ReasonOfStripBreakageRmAreaRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class ReasonOfStripBreakageRmArea : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as ReasonOfStripBreakageRmAreaRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        //Здесь визуализация Экселя
        //prm.ExcelApp.ScreenUpdating = true;
        //prm.ExcelApp.Visible = true;
        this.SaveResult(prm);
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
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

    private Boolean RunRpt4(ReasonOfStripBreakageRmAreaRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);

        //CurrentWrkSheet.Range["H1", "L1"].ClearContents();
        //CurrentWrkSheet.Range["H1:L1"].ClearContents();
        CurrentWrkSheet.Cells[2, 1].Value = $"за период с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");
        Odac.ExecuteNonQuery("VIZ_PRN.OTK_248_247.preCountDef", CommandType.StoredProcedure, false, null);
        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");


        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.V_OTK_PRICH248_247";
        odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);
        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");


        if (odr != null){
          const int rowRoot = 5;
          var row = rowRoot;
          const int firstExcelColumn = 1;
          const int lastExcelColumn = 30;

          int flds = odr.FieldCount;

          while (odr.Read()){
            //CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, firstExcelColumn], CurrentWrkSheet.Cells[row + 1, lastExcelColumn]]);
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }
          
          if (row > rowRoot)
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowRoot, firstExcelColumn], CurrentWrkSheet.Cells[row - 1, lastExcelColumn]].Cells.Borders.LineStyle = 1;//XlLineStyle.xlContinuous
        }

        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");
        

        CurrentWrkSheet.Cells[1, 1].Select();
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

    private Boolean RunRpt(ReasonOfStripBreakageRmAreaRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      
      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);

        //CurrentWrkSheet.Range["H1", "L1"].ClearContents();
        //CurrentWrkSheet.Range["H1:L1"].ClearContents();
        CurrentWrkSheet.Cells[2, 1].Value = $"за период с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");
        Odac.ExecuteNonQuery("VIZ_PRN.OTK_248_247.preCountDef", CommandType.StoredProcedure, false, null);
        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");


        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.V_OTK_PRICH248_247";
        odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);
        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");


        if (odr != null){
          const int rowRoot = 5;
          var row = rowRoot;
          const int firstExcelColumn = 1;
          const int lastExcelColumn = 30;
          
          int flds = odr.FieldCount;
          int j = 0;
          
          var data = new Object[1, flds];

          while (odr.Read()){
            data = (object[,])ArrayUtl.ResizeArray(data, new[] { j + 1, flds });

            for (int i = 0; i < flds; i++)
              data[j, i] = odr.GetValue(i);
            
            j++;
            row++;
          }

          if (row > rowRoot){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowRoot, firstExcelColumn], CurrentWrkSheet.Cells[row - 1, lastExcelColumn]].Value = data;
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowRoot, firstExcelColumn], CurrentWrkSheet.Cells[row - 1, lastExcelColumn]].Cells.Borders.LineStyle = 1; //XlLineStyle.xlContinuous
          }
        }

        //MessageBox.Show($"за период с {DateTime.Now:dd.MM.yyyy HH:mm}");


        CurrentWrkSheet.Cells[1, 1].Select();
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


