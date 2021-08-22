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
  public sealed class QualityIndsUo1RptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public Boolean IsFill1StSheet { get; set; }
    public Boolean IsTypeProdF3 { get; set; }
    public string TypeProdSqlStrF3 { get; set; }
    public Boolean IsThicknessF3 { get; set; }
    public string ThicknessSqlStrF3 { get; set; }
    public Boolean IsSortF3 { get; set; }
    public string SortSqlStrF3 { get; set; }

    public QualityIndsUo1RptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class QualityIndsUo1 : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as QualityIndsUo1RptParam);
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

    private Boolean RunRpt(QualityIndsUo1RptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean Result = false;

      try{
        prm.DateBegin = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
        prm.DateEnd = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));


        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        var dtBegin = DbVar.GetDateBeginEnd(true, true);
        var dtEnd = DbVar.GetDateBeginEnd(false, true);

        //Stopwatch stopWatch = new Stopwatch();
        //TimeSpan ts = stopWatch.Elapsed;
        //string elapsedTime = null;


        //stopWatch.Start();
        var arrSqlPrm = new[]{ "0.23, 0.27, 0.30, 0.35",  "0.23", "0.27", "0.30", "0.35"};
        var arrStartRow = new[] {6, 5, 6, 6, 6 };
        var arrRowHdr = new[] { 2, 1, 2, 2, 2 };

        const string sqlStmt = "SELECT * FROM VIZ_PRN.V_FINCUT_QM ORDER BY 1";

        for (int j = 0; j < arrRowHdr.Length; j++){

          prm.ExcelApp.ActiveWorkbook.WorkSheets[j + 1].Select();
          CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
          CurrentWrkSheet.Cells[arrRowHdr[j], 3].Value = prm.DateBegin;

          if (j == 0)
            CurrentWrkSheet.Cells[2, 15].Value = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}";
          

          DbVar.SetString(arrSqlPrm[j]);
          odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

          if (odr != null){

            int flds = odr.FieldCount;
            
            while (odr.Read()){
              for (int i = 0; i < flds; i++)
                CurrentWrkSheet.Cells[arrStartRow[j], i + 2].Value = odr.GetValue(i);

              arrStartRow[j]++;
            }

            odr.Close();
            odr.Dispose();
          }

        }

        //stopWatch.Stop();
        //ts = stopWatch.Elapsed;
        //elapsedTime = $"SGP_SRV_SHIR_L2ITOG... -  {ts.Hours:00}:{ts.Minutes:00}:{ts.Seconds:00}.{ts.Milliseconds / 10:00}";
        //MessageBox.Show(elapsedTime);


        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
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


