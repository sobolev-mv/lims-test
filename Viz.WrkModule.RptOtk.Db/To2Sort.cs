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

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class To2SortParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string Thickness { get; set; }
    public string FilterThickness { get; set; }

    public To2SortParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class To2Sort : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as To2SortParam);
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

        Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null; 
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(To2SortParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        const string sqlStmt = "SELECT * FROM VIZ_PRN.OTK_2_SORT_KAT1 ORDER BY VES";
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        DbVar.SetString(prm.Thickness, "4");        

        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[1, 6].Value = "c " + $"{dtBegin:dd.MM.yyyy HH:mm:ss}" + " по " + $"{dtEnd:dd.MM.yyyy HH:mm:ss}";
        CurrentWrkSheet.Cells[2, 3].Value = prm.FilterThickness;

        odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, null);
        if (odr != null){
          int row = 6;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("KOD_DEF");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        DbVar.SetString(prm.Thickness, "Б/К");
        odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, null);

        if (odr != null){
          int row = 6;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("KOD_DEF");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        DbVar.SetString(prm.Thickness, "4,Б/К");
        odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, null);

        if (odr != null){
          int row = 6;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 6].Value = odr.GetValue("KOD_DEF");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("VES");
            row++;
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 6].Value = "c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
        CurrentWrkSheet.Cells[2, 3].Value = prm.FilterThickness;

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

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


