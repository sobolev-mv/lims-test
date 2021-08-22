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
  public sealed class OtkQntDefMonthRptParam : RptWithF1Param
  {

    public OtkQntDefMonthRptParam(string sourceXlsFile, string destXlsFile, DateTime dateBegin, DateTime dateEnd)
      : base(sourceXlsFile, destXlsFile)
    {
      DateBegin = new DateTime(dateBegin.Year, dateBegin.Month, 1);   
      DateEnd = new DateTime(dateBegin.Year, dateBegin.Month, DateTime.DaysInMonth(dateBegin.Year, dateBegin.Month));
    }



  }

  public sealed class OtkQntDefMonth : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkQntDefMonthRptParam);
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
        //GC.WaitForPendingFinalizers();
        //GC.Collect();
      }
    }

    //За кажды сутки в течении месяца
    private Boolean RunRpt(OtkQntDefMonthRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      Int64 zdn = 0;

      //генерим отрицательный номер задания
      var rm = new Random();
      zdn = rm.Next(10000000, 99999999) * -1;

      var dt1 = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
      var dt2 = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));

      try{
        PrepareFilterRpt(prm);

        DbVar.SetRangeDate(dt1, dt2, 1);
        DbVar.SetNum(zdn);

        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);

        Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null);

        const string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEFECT_GRAF";

        CurrentWrkSheet.Cells[1, 15].Value = string.Format(" с {0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[3, 2].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;


        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);
        if (odr != null){
          var row = 6;
          var flds = odr.FieldCount;

          while (odr.Read()){

            for (int i = 0; i < flds; i++){
              if (i > 1)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            }
            row++;
          }
        }

        if (prm.TypeFilter >= 1){
          for (int i = 2; i < 26; i++){
            prm.ExcelApp.ActiveWorkbook.WorkSheets[i].Select();
            CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
            CurrentWrkSheet.Cells[3, 3].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;
          }
          prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        }

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

