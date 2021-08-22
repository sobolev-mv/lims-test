using System;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptOtk.Db
{
  public sealed class DistribDefectsOnSurfaceRptParam : RptWithF1Param
  {
    public string Defect { get; set; }
    public DistribDefectsOnSurfaceRptParam(string sourceXlsFile, string destXlsFile): base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class DistribDefectsOnSurface : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DistribDefectsOnSurfaceRptParam);
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

    private Boolean RunRpt(DistribDefectsOnSurfaceRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      try{
        PrepareFilterRpt(prm);
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[1, 2].Value = "Распределение  дефекта " + prm.Defect + " по поверхности рулона с шагом 5%" +
                                             string.Format(" за период с " + "{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        switch (prm.TypeFilter){
          case 1:
            CurrentWrkSheet.Cells[3, 2].Value = prm.GetFilterCriteria();
            break;
          case 2:
            prm.GetFilter1LstCriteria(9);
            //Возвращаемся на первую страницу
            prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
            CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
            break;
        }

        //1.сбор информации по всем рулонам
        string SqlStmt = "begin VIZ_PRN.Raspred_Def_POV.insRaspr('0', 0, '" + prm.Defect + "'); end;";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(SqlStmt, CommandType.Text, false, false, null); }));

        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null){
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_RASPR_POV_PRN";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 6;
          var flds = odr.FieldCount;

          while (odr.Read()){
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //2.сбор информации по каждому рулону отдельно
        for (int k = 0; k < 6; k++){

          SqlStmt = "begin VIZ_PRN.Raspred_Def_POV.insRaspr('" + (k + 1).ToString(CultureInfo.InvariantCulture) + "', 0, '" + prm.Defect + "'); end;";
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(SqlStmt, CommandType.Text, false, false, null); }));

          if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
          else
            return false;

          oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
          }

          SqlStmt = "SELECT * FROM VIZ_PRN.OTK_RASPR_POV_PRN";
          prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt, CommandType.Text, false, null, null); }));
          oracleCommand = iar.AsyncState as OracleCommand;
          if (oracleCommand != null)
            odr = oracleCommand.EndExecuteReader(iar);

          if (odr != null){
            var row = 32 + k * 24;
            var flds = odr.FieldCount;

            while (odr.Read()){
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
              row++;
            }
          }
          odr.Close();
          odr.Dispose();
        }

        Result = true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
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

