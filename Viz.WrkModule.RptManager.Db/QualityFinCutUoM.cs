using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class QualityFinCutUoMRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public QualityFinCutUoMRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class QualityFinCutUoM : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as QualityFinCutUoMRptParam);
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

    private Boolean RunRpt(QualityFinCutUoMRptParam prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;
      var oef = new OdacErrorInfo();
      DateTime dtBegin = new DateTime(prm.DateBegin.Year, 1, 1);
      DateTime dtEnd = new DateTime(prm.DateBegin.Year, 12, 31); 

      try{
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(dtBegin, dtEnd, 1)));
        CurrentWrkSheet.Cells[2, 1].Value = "за " + prm.DateBegin.Year.ToString(CultureInfo.InvariantCulture) + " год";
        
        //Готовим данные
        const string sqlStmt1 = "BEGIN DELETE FROM VIZ_PRN.TMP_QLT_FINCUT; INSERT INTO VIZ_PRN.TMP_QLT_FINCUT SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_CORE; END;";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(sqlStmt1, CommandType.Text, false, false, null); }));
        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        //лист "Таблица" строки 4,5,6 колонки C- O 
        const string sqlStmt2 = "SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_M";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt2, CommandType.Text, false, null, oef); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 4;

          while (odr.Read()){
            //CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 91]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 91]]);
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);

            row++;
          }
        }

        odr.Close();
        odr.Dispose();

        const string sqlStmt3 = "SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_DEF_M";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt3, CommandType.Text, false, null, oef); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 11;

          while (odr.Read()){
            
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);

            row += 2;
          }
        }

        odr.Close();
        odr.Dispose();

        
        //Переходим на лист "Список рулонов"   со строки 4 и вниз, копируя сетку 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[8].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;

        const string sqlStmt4 = "SELECT * FROM VIZ_PRN.OTK_QLT_FINCUT_RUL";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(sqlStmt4, CommandType.Text, false, null, oef); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 4;

          while (odr.Read()){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 1], CurrentWrkSheet.Cells[row, 6]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 1], CurrentWrkSheet.Cells[row + 1, 6]]);

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

