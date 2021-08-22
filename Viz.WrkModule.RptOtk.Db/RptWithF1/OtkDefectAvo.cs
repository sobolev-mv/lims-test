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
  public sealed class OtkDefectAvoRptParam : RptWithF1Param
  {
    public OtkDefectAvoRptParam(string sourceXlsFile, string destXlsFile): base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class OtkDefectAvo : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkDefectAvoRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;

        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);
      }
      catch (Exception ex)
      {
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
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

    private Boolean RunRpt(OtkDefectAvoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      //IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      


      try{
        PrepareFilterRpt(prm);

        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        DbVar.SetNum(1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);

        //MessageBox.Show("Point3");

        Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null, false);
        //Odac.ExecuteNonQuery("BEGIN VIZ_PRN.VAR_RPT.SetNum1(1); VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT; END;", CommandType.Text, false, null, false);

        //MessageBox.Show("Point4");

        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEFECT_PERIOD_RUL";



        CurrentWrkSheet.Cells[2, 1].Value = string.Format("за период с " + "{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[4, 2].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;

        //MessageBox.Show("Point5");

        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        int row = 0;

        if (odr != null){
          row = 9;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 1].Value = odr.GetString("DISPLAYTEXT");
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetString("FEHLERTYP");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetDouble("VES_01");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetDouble("VES_02");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetDouble("VES_03");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetDouble("VES_04");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetDouble("VES_05");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetDouble("VES_06");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //MessageBox.Show("Point6");

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEFECT_PERIOD_VSEGO_RUL";
        odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, null);
        
        if (odr != null){
          row++;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES_01");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_02");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("VES_03");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("VES_04");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_05");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_06");
            row++;
          }
        }

        //MessageBox.Show("Point7");

        //Здесь вызываем код очистки.
        //Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null, false);

        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка выполнения", ex.Message, MessageBoxImage.Stop)));
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


