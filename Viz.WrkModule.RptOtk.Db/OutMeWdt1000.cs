using System;
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

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class OutMeWdt1000Param : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin {get; set;}
    public DateTime DateEnd {get; set;}

    public OutMeWdt1000Param(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class OutMeWdt1000 : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OutMeWdt1000Param);
      dynamic wrkSheet = null;

      try
      {
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
      finally{
        prm.ExcelApp.Quit();

        //Здесь код очистки      
        if (wrkSheet != null)
          Marshal.ReleaseComObject(wrkSheet);

        Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(OutMeWdt1000Param prm, dynamic CurrentWrkSheet)
    {
      IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try
      {
        const string SqlStmt = "SELECT * FROM VIZ_PRN.OTK_SHIR";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));
        CurrentWrkSheet.Cells[1, 9].Value = "c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          int aprPrev = 0;
          int aprCurrent = 0;
          int row = 0;

          while (odr.Read()){
            aprCurrent = odr.GetInt32("APR"); 
            
            if (aprCurrent != aprPrev){
              switch (aprCurrent){
                case 3:
                  row = 6;  
                  break;
                case 4:
                  row = 11;  
                  break;
                case 5:
                  row = 16;  
                  break;
                case 6:
                  row = 21;  
                  break;
                case 9:
                  row = 26;  
                  break;
                case 12:
                  row = 31;  
                  break;
                default:
                  break;
              }
            }

            CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue("TOLS");
            CurrentWrkSheet.Cells[row, 4].Value = odr.GetValue("VES_KL_1");
            CurrentWrkSheet.Cells[row, 5].Value = odr.GetValue("VES_SORT_1");
            CurrentWrkSheet.Cells[row, 7].Value = odr.GetValue("VES_KL_2");
            CurrentWrkSheet.Cells[row, 8].Value = odr.GetValue("VES_SORT_2");
            CurrentWrkSheet.Cells[row, 10].Value = odr.GetValue("VES_KL_3");
            CurrentWrkSheet.Cells[row, 11].Value = odr.GetValue("VES_SORT_3");
            CurrentWrkSheet.Cells[row, 13].Value = odr.GetValue("VES_KL_4");
            CurrentWrkSheet.Cells[row, 14].Value = odr.GetValue("VES_SORT_4");

            aprPrev = aprCurrent;
            row++;

          }
        }

        CurrentWrkSheet.Cells[1, 1].Select();
        Result = true;
      }
      catch (Exception)
      {
        Result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return Result;
    }


  }






}


