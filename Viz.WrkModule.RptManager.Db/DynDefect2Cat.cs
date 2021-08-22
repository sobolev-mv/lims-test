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

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class DynDefect2CatRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public DynDefect2CatRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class DynDefect2Cat : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DynDefect2CatRptParam);
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

    private Boolean RunRpt(DynDefect2CatRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;

      DateTime dtTmpBegin;
      DateTime dtTmpEnd; 

      try{
        dtTmpBegin = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
        dtTmpEnd = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));

        DbVar.SetRangeDate(dtTmpBegin, dtTmpEnd, 1);
        CurrentWrkSheet.Cells[2, 1].Value = $"за период с {dtTmpBegin:dd.MM.yyyy} по {dtTmpEnd:dd.MM.yyyy}";

        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.OTK_DINAMIKA_DEF2KAT";
        odr = Odac.GetOracleReader(sqlStmt1, CommandType.Text, false, null, null);
       
        if (odr != null){
          int flds = odr.FieldCount;
          int row = 6;

          while (odr.Read()){
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
            row++;
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[2, 1].Value = $"за период с {dtTmpBegin:dd.MM.yyyy} по {dtTmpEnd:dd.MM.yyyy}";

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


