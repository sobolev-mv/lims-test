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
  public sealed class OtkDefYearAvoRptParam : RptWithF1Param
  {
    public OtkDefYearAvoRptParam(string sourceXlsFile, string destXlsFile, DateTime dateBegin, DateTime dateEnd)
      : base(sourceXlsFile, destXlsFile)
    {
      DateBegin = new DateTime(dateBegin.Year, 01, 01);
      DateEnd = new DateTime(dateBegin.Year, 12, 31);      
    }
  }

  public sealed class OtkDefYearAvo : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkDefYearAvoRptParam);
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

    private Boolean RunRpt(OtkDefYearAvoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      Int64 zdn = 0;


      try{
        //генерим отрицательный номер задания
        var rm = new Random();
        zdn = rm.Next(10000000, 99999999) * -1;
        PrepareFilterRpt(prm);

        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEFECT_YEAR";
        //Здесь выставляем дату за весь год
        prm.DateBegin = new DateTime(prm.DateBegin.Year, 01, 01);
        prm.DateEnd = new DateTime(prm.DateBegin.Year, 12, 31);

        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        DbVar.SetNum(zdn);

        //MessageBox.Show("Point1");

        Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null);

        //MessageBox.Show("Point2");

        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true); 
        CurrentWrkSheet.Cells[1, 13].Value = $"{dtBegin:yyyy}";
        if (prm.TypeFilter >= 1){
          CurrentWrkSheet.Cells[3, 2].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;
        }

        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);
        int row = 7; 

        if (odr != null){

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("FEHLERTYP");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES_DEF_01");
            //CurrentWrkSheet.Cells[row, 3].Value = 100;

            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_DEF_02");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("VES_DEF_03");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("VES_DEF_04");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_DEF_05");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_DEF_06");
            CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("VES_DEF_07");
            CurrentWrkSheet.Cells[row, 17].Value = odr.GetValue("VES_DEF_08");
            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("VES_DEF_09");
            CurrentWrkSheet.Cells[row, 21].Value = odr.GetValue("VES_DEF_10");
            CurrentWrkSheet.Cells[row, 23].Value = odr.GetValue("VES_DEF_11");
            CurrentWrkSheet.Cells[row, 25].Value = odr.GetValue("VES_DEF_12");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }


        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEFECT_YEAR_VSEGO";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          row++;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES_DEF_01");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_DEF_02");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("VES_DEF_03");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("VES_DEF_04");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_DEF_05");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_DEF_06");
            CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("VES_DEF_07");
            CurrentWrkSheet.Cells[row, 17].Value = odr.GetValue("VES_DEF_08");
            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("VES_DEF_09");
            CurrentWrkSheet.Cells[row, 21].Value = odr.GetValue("VES_DEF_10");
            CurrentWrkSheet.Cells[row, 23].Value = odr.GetValue("VES_DEF_11");
            CurrentWrkSheet.Cells[row, 25].Value = odr.GetValue("VES_DEF_12");
            row++;
          }
        }

        if (prm.TypeFilter >= 1){
          for (int i = 2; i < 26; i++){
            prm.ExcelApp.ActiveWorkbook.WorkSheets[i].Select();
            CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
            CurrentWrkSheet.Cells[2, 2].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;
          }
          prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        }

        //Здесь вызываем код очистки.
        //Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null);

        Result = true;
      }
      catch (Exception ex){
        Result = false;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Error)));
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
