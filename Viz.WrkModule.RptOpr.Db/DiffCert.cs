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

namespace Viz.WrkModule.RptOpr.Db
{
  public sealed class DiffCertRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public DiffCertRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class DiffCert : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DiffCertRptParam);
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

    private Boolean RunRpt(DiffCertRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);

        //CurrentWrkSheet.Range["H1", "L1"].ClearContents();
        //CurrentWrkSheet.Range["H1:L1"].ClearContents();
        CurrentWrkSheet.Cells[1, 4].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.V_DIFFCERT";
        odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);
       
        if (odr != null){
          int flds = odr.FieldCount;
          int row = 6;

          int rowMergStart = 0;
          int rowMergEnd = 0;

          const int firstExcelColumn = 1;
          const int lastExcelColumn = 14;

          string locNoPred = "";
          string locNoCurr = "";

          while (odr.Read()) {

            locNoCurr = odr.GetString(2);
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, firstExcelColumn], CurrentWrkSheet.Cells[row, lastExcelColumn]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, firstExcelColumn], CurrentWrkSheet.Cells[row + 1, lastExcelColumn]]);

            for (int i = 8; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);


            if (locNoCurr != locNoPred) {

              for (int i = 0; i < 8; i++)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

              if (rowMergStart > 0) {

                for (int j = 1; j < 9; j++)
                  CurrentWrkSheet.Range[CurrentWrkSheet.Cells[rowMergStart, j], CurrentWrkSheet.Cells[rowMergEnd + 1, j]].Merge();

                rowMergStart = rowMergEnd = 0;
              }
            }
            else if (rowMergStart == 0)
              rowMergStart = rowMergEnd = row - 1;
            else
              rowMergEnd++;


            locNoPred = locNoCurr;

            row++;
          }
        }

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


