using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Smv.Utils;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class OtkQualityAvoRptParam : RptWithF1Param
  {
    public Boolean IsRpt2 { get; set; }

    public OtkQualityAvoRptParam(string sourceXlsFile, string destXlsFile)
           : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class OtkQualityAvo : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkQualityAvoRptParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;

        this.RunRpt(prm, wrkSheet);
        SaveResult(prm);
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
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

    private Boolean RunRpt(OtkQualityAvoRptParam prm, dynamic CurrentWrkSheet)
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
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true); 
        DbVar.SetNum(zdn);   
         
        
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        Odac.ExecuteNonQuery("VIZ_PRN.OTK_AVO.preOTK_CAT_AVO", CommandType.StoredProcedure, false, null);

        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[2, 15].Value = $"{dtBegin:dd.MM.yyyy}" + " по " + $"{dtEnd:dd.MM.yyyy}";

        if (prm.TypeFilter >= 1){
          CurrentWrkSheet.Cells[4, 5].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;
        }

        string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_CAT_DEF_AVO";
        DbVar.SetNum(zdn,860);  
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 9;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("VES_VH");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_CAT1");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("VES_CAT2");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("VES_CAT3");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_CAT4");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_OTM");
            CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("VES_OBR");

            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("VES_860");
            CurrentWrkSheet.Cells[row, 21].Value = odr.GetValue("VES_960");
            CurrentWrkSheet.Cells[row, 23].Value = odr.GetValue("VES_1000");
            
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        /*
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF_VES_SHIR";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetNum(zdn)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 9;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("VES_860");
            CurrentWrkSheet.Cells[row, 21].Value = odr.GetValue("VES_960");
            CurrentWrkSheet.Cells[row, 23].Value = odr.GetValue("VES_1000");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        */ 
        /*
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetNum(zdn, 1000)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 9;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 23].Value = odr.GetValue("VES_SORT1");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }
      */

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_T2";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 19;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("D500");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("D405");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("D505");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("D406");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("D506");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("D508_608");
            CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("D035");
            CurrentWrkSheet.Cells[row, 17].Value = odr.GetValue("D606");
            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("D437");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        SqlStmt = prm.IsRpt2 ? "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_NT3" : "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_T3";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 30;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("D501");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("D202");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("D603");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("D604");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("D607");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("D009");
            CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("D010");
            CurrentWrkSheet.Cells[row, 17].Value = odr.GetValue("D414_614");
            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("D522");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //Звполняем страницу 2
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_T2_2L";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 10;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("D306");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("D411");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("D17");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("D421");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("D633");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("D635");
            CurrentWrkSheet.Cells[row, 15].Value = odr.GetValue("D542");
            CurrentWrkSheet.Cells[row, 17].Value = odr.GetValue("D441");
            CurrentWrkSheet.Cells[row, 19].Value = odr.GetValue("DALL");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        SqlStmt = prm.IsRpt2 ? "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_T3_N2L" : "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_T3_2L";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 21;
          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("D513");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("D29");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("D31");
            CurrentWrkSheet.Cells[row, 9].Value = odr.GetValue("D507");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("D209");
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        
        //Cтенд-Дефект
        prm.ExcelApp.ActiveWorkbook.WorkSheets[4].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_DEF";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){

          int row = 4;
          int rowInsert = 6;
          int flds = odr.FieldCount;

          while (odr.Read()){

            if (row == rowInsert){
              CurrentWrkSheet.Rows[row].Insert();
              CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 12]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 12]]);
              rowInsert++;
            }

            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //Здесь вызываем код очистки.
        Odac.ExecuteNonQuery("VIZ_PRN.OTK_AVO.postOTK_CAT_AVO", CommandType.StoredProcedure, false, null);

        //Список стендов
        prm.ExcelApp.ActiveWorkbook.WorkSheets[3].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        if (prm.IsDateAvoLstF1)
          CurrentWrkSheet.Cells[1, 10].Value = $"c {prm.DateBeginAvoLstF1:dd.MM.yyyy HH:mm:ss}" + " по " + $"{prm.DateEndAvoLstF1:dd.MM.yyyy HH:mm:ss}";

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_DEF_AVO_STEND";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          var row = 4;
          var flds = odr.FieldCount;
          var nr = 1; 

          while (odr.Read()){
            if (row == 49){
              row = 4;
              nr += 4;
            } 

            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + nr].Value = odr.GetValue(i);

            row++;
          }
        }
 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception ex){
        Result = false;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
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


