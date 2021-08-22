using System;
using System.Data;
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

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class QntWeldOn2ndRollRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public QntWeldOn2ndRollRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class QntWeldOn2ndRoll : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as QntWeldOn2ndRollRptParam);
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
        prm.WorkBook.Close();
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        if (prm.ExcelApp != null)
          Marshal.ReleaseComObject(prm.ExcelApp);

        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(QntWeldOn2ndRollRptParam prm, dynamic currentWrkSheet)
    {
      
      var result = false;
      OracleDataReader odr = null;

      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);
        currentWrkSheet.Cells[1, 4].Value = "с " + $"{dtBegin:dd.MM.yyyy HH:mm:ss}" + " по " + $"{dtEnd:dd.MM.yyyy HH:mm:ss}";

        const string stmtSql5 = "SELECT * FROM VIZ_PRN.V_QNTWELDON2NDROLL";
        odr = Odac.GetOracleReader(stmtSql5, CommandType.Text, false, null, null);
       
        if (odr != null){
          const int firstExcelColumn = 1;
          const int lastExcelColumn = 74;
          var row = 5;
          while (odr.Read()){

            currentWrkSheet.Range[currentWrkSheet.Cells[row, firstExcelColumn], currentWrkSheet.Cells[row, lastExcelColumn]].Copy(currentWrkSheet.Range[currentWrkSheet.Cells[row + 1, firstExcelColumn], currentWrkSheet.Cells[row + 1, lastExcelColumn]]);

            for (int i = 0; i < odr.FieldCount; i++)
              currentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }

        }

        result = true;
      }
      catch (Exception e){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.Message, MessageBoxImage.Stop)));
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


  }



}
