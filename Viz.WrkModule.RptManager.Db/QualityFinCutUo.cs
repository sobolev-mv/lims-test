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
  public sealed class QualityFinCutUoRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public QualityFinCutUoRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class QualityFinCutUo : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as QualityFinCutUoRptParam);
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

    private Boolean RunRpt(QualityFinCutUoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      var oef = new OdacErrorInfo();
      DateTime dtBegin = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
      DateTime dtEnd = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month)); 

      try{
        DbVar.SetRangeDate(dtBegin, dtEnd, 1);
        CurrentWrkSheet.Cells[2, 1].Value = "за период с " + $"{dtBegin:dd.MM.yyyy}" + " по " + $"{dtEnd:dd.MM.yyyy}";
        
        //Готовим данные
        const string sqlStmt1 = "BEGIN DELETE FROM VIZ_PRN.TMP_QLT_FINCUT; INSERT INTO VIZ_PRN.TMP_QLT_FINCUT SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_CORE; END;";
        Odac.ExecuteNonQuery(sqlStmt1, CommandType.Text, false, null);
        
        //лист "Таблица" строки 4,5,6 колонки C- AH 
        const string sqlStmt2 = "SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT";
        odr = Odac.GetOracleReader(sqlStmt2, CommandType.Text, false, null, null);
        
        if (odr != null){
          int flds = odr.FieldCount;
          int row = 4;

          while (odr.Read()){
            //CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 91]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 91]]);
            for (int i = 1; i < flds; i++)

              //MessageBox.Show(Convert.ToString(odr.GetValue(k)));
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
              //MessageBox.Show(odr.GetValue(i).ToString());
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        const string sqlStmt3 = "SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_DEF";
        odr = Odac.GetOracleReader(sqlStmt3, CommandType.Text, false, null, null);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 11;

          while (odr.Read()){
            
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);

            row += 2;
          }

          odr.Close();
          odr.Dispose();
        }

        //Переходим на лист "Список рулонов"   со строки 4 и вниз, копируя сетку 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[9].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

        const string sqlStmt4 = "SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_RUL";
        odr = Odac.GetOracleReader(sqlStmt4, CommandType.Text, false, null, null);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 4;

          while (odr.Read()){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 6]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 6]]);

            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }

          odr.Close();
          odr.Dispose();
        }

        //Переходим на лист "% 1 сорта СГП"   
        prm.ExcelApp.ActiveWorkbook.WorkSheets[10].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(dtBegin, dtEnd, 1)));
        CurrentWrkSheet.Cells[2, 1].Value = "за период с " + string.Format("{0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);

        const string sqlStmt5 = "SELECT * FROM VIZ_PRN.OTK_DINAMIKA_SGP_1SORT";
        odr = Odac.GetOracleReader(sqlStmt5, CommandType.Text, false, null, null);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 7;

          while (odr.Read()){
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }
  
        }

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

