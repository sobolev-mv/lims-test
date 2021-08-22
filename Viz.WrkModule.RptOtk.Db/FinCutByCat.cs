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
  public sealed class FinCutByCatRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public FinCutByCatRptParam(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class FinCutByCat : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as FinCutByCatRptParam);
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

    private Boolean RunRpt(FinCutByCatRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;
      //DateTime? dtBegin = null;
      //DateTime? dtEnd = null;
      int curYear = prm.DateBegin.Year;

      try
      {

        const string sqlStmt = "SELECT * FROM VIZ_PRN.V_FINCUTBYCAT_M";


        /*За год по месяцам*/
        var prmSql1 =  new[] {"0.23, 0.27, 0.30", "0.23", "0.27", "0.30"};
        var prmSql2 =  new[] {"Порезка 0,23-0,30", "Порезка 0,23", "Порезка 0,27", "Порезка 0,30"};
        var hdrSheet = new[] {" ", " в толщине 0,23 ", " в толщине 0,27 ", " в толщине 0,30 "};

        for (var iSheet = 0; iSheet < 4; iSheet++){

          prm.ExcelApp.ActiveWorkbook.WorkSheets[iSheet + 1].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          DbVar.SetString(prmSql1[iSheet], prmSql2[iSheet]);

          CurrentWrkSheet.Cells[1, 1].Value = $"Категории металла по первичной порезке" + hdrSheet[iSheet] + "за 2019, 2020 " + ", " + curYear + " год";

          //Заполняем шапку по месяцам
          for (var iMonth = 1; iMonth < 13; iMonth++)
            CurrentWrkSheet.Cells[3, iMonth + 25].Value = new DateTime(curYear, iMonth, 1);
          


          prm.DateBegin = new DateTime(curYear, 1, 1);
          prm.DateEnd = new DateTime(curYear, 12, 31);
          DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
          //dtBegin = DbVar.GetDateBeginEnd(true, true);
          //dtEnd = DbVar.GetDateBeginEnd(false, true);

          odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

          if (odr == null)
            continue;

          var row = 5;
          var flds = odr.FieldCount;

          while (odr.Read()){
            CurrentWrkSheet.Cells[row, 1].Value = odr.GetValue(0);
            
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 25].Value = odr.GetValue(i);

            row++;
          }

          odr.Close();
          odr.Dispose();
          
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
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




