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
  public sealed class DistribDefectsOnLengthRptParam : RptWithF1Param
  {
    public string Defect { get; set; }
    public DistribDefectsOnLengthRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class DistribDefectsOnLength : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DistribDefectsOnLengthRptParam);
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

        Marshal.ReleaseComObject(prm.WorkBook);
        Marshal.ReleaseComObject(prm.ExcelApp);
        wrkSheet = null;
        prm.WorkBook = null;
        prm.ExcelApp = null;
        GC.Collect();
      }
    }

    private Boolean RunRpt(DistribDefectsOnLengthRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      //IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;



      try{
        PrepareFilterRpt(prm);
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);

        CurrentWrkSheet.Cells[1, 2].Value = "Распределение  дефекта " + prm.Defect + " по длине рулона с шагом 5%";
        CurrentWrkSheet.Cells[2, 6].Value = string.Format("за период с " + "{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);

        switch (prm.TypeFilter){
          case 1:
            CurrentWrkSheet.Cells[4, 2].Value = prm.GetFilterCriteria();
            break;
          case 2:
            prm.GetFilter1LstCriteria(10);
            //Возвращаемся на первую страницу
            prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
            CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
            break;
        }

        //1.сбор информации по всем рулонам
        string SqlStmt = "begin VIZ_PRN.Raspred_Def.insRaspr('0', 0, '" + prm.Defect + "'); end;";
        Odac.ExecuteNonQuery(SqlStmt, CommandType.Text, false, null);   //(SqlStmt, CommandType.Text, false, false, null);

        SqlStmt = "SELECT * FROM VIZ_PRN.OTK_RASPR_PRN";
        odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);       
        

        if (odr != null){
          var row = 7;
          var flds = odr.FieldCount;

          while (odr.Read()){
            for (int i = 1; i < flds; i++)
              CurrentWrkSheet.Cells[row, i+1].Value = odr.GetValue(i);
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //2.сбор информации по каждому рулону отдельно
        for (int k = 0; k < 6; k++){

          SqlStmt = "begin VIZ_PRN.Raspred_Def.insRaspr('" + (k + 1).ToString(CultureInfo.InvariantCulture) + "', 0, '" + prm.Defect + "'); end;";
          Odac.ExecuteNonQuery(SqlStmt, CommandType.Text, false, null);


          SqlStmt = "SELECT * FROM VIZ_PRN.OTK_RASPR_PRN";
          odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

          if (odr != null){
            var row = 13 + k * 6;
            var flds = odr.FieldCount;

            while (odr.Read()){
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
              row++;
            }
            odr.Close();
            odr.Dispose();
          }

        }

        //3.сбор информации по 2-м, 3-м рулонам
        for (int k = 0; k < 5; k++){

          SqlStmt = "begin VIZ_PRN.Raspred_Def.insRaspr('0', " + (k + 1).ToString(CultureInfo.InvariantCulture) + ", '" + prm.Defect + "'); end;";
          Odac.ExecuteNonQuery(SqlStmt, CommandType.Text, false, null);

          SqlStmt = "SELECT * FROM VIZ_PRN.OTK_RASPR_PRN";
          odr = Odac.GetOracleReader(SqlStmt, CommandType.Text, false, null, null);

          if (odr != null){
            var row = 49 + k * 6;
            var flds = odr.FieldCount;

            while (odr.Read()){
              for (int i = 1; i < flds; i++)
                CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);
              row++;
            }
            odr.Close();
            odr.Dispose();
          }

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

