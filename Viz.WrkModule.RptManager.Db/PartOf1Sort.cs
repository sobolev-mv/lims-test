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
  public sealed class PartOf1SortRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int TypeFilterF3 { get; set; }
    public Boolean IsThicknessF3 { get; set; }
    public decimal ThicknessF3 { get; set; }
    public string   ListLocNumF3 { get; set; }
    public PartOf1SortRptParam(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class PartOf1Sort : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as PartOf1SortRptParam);
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
      finally
      {
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

    private Boolean RunRpt(PartOf1SortRptParam prm, dynamic CurrentWrkSheet)
    {
      
      OracleDataReader odr = null;
      Boolean Result = false;
      Int64 zdn = 0;

      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      //генерим отрицательный номер задания
      var rm = new Random();
      zdn = rm.Next(10000000, 99999999) * -1;
      DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
      DbVar.SetNum(zdn);

      //MessageBox.Show(zdn.ToString());
      Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.OTK_DEF WHERE ZDN = " + zdn.ToString(), CommandType.Text, false, null);

      dtBegin = DbVar.GetDateBeginEnd(true, true);
      dtEnd = DbVar.GetDateBeginEnd(false, true);
      CurrentWrkSheet.Cells[2, 9].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

      try{

        if ((prm.TypeFilterF3 == 0) && (!prm.IsThicknessF3)){
       
          var sqlStmt1 = "BEGIN " +
                         "DELETE FROM VIZ_PRN.TMP_FINCUT; " +
                        "insert into VIZ_PRN.TMP_FINCUT(ME_ID, LOCNO, DICKEOUTPUT, shir, GEWOUTPUT) " +
                        "select ME_ID, LOCNO, DICKEOUTPUT, SHIR, GEWOUTPUT from VIZ_PRN.V_FINCUT_CORE; " +
                        "VIZ_PRN.OTK_FINCUT.preOTK_FINCUT; " +
                        "END;";

          Odac.ExecuteNonQuery(sqlStmt1, CommandType.Text, false, null);

          var sqlStmt2 = "SELECT * FROM VIZ_PRN.V_FINCUT_SORT";
          odr = Odac.GetOracleReader(sqlStmt2, CommandType.Text, false, null, null);

          if (odr != null){
            
            int row = 7;

            while (odr.Read())
            {
              CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue(1);
              CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue(2);

              row++;
            }
          }

          odr.Close();
          odr.Dispose();

          CurrentWrkSheet.Cells[4, 3].Value = "Без фильтрации";
        }
        else if ((prm.TypeFilterF3 == 0) && (prm.IsThicknessF3)){
          // c фильтром по толщине 

          var sqlStmt3 = "BEGIN " +
                         "DELETE FROM VIZ_PRN.TMP_FINCUT; " +
                         "insert into VIZ_PRN.TMP_FINCUT(ME_ID, LOCNO, DICKEOUTPUT, shir, GEWOUTPUT) " +
                         "select ME_ID, LOCNO, DICKEOUTPUT, SHIR, GEWOUTPUT from VIZ_PRN.V_FINCUT_CORE WHERE DICKEOUTPUT = :PTOLS; " +
                         "VIZ_PRN.OTK_FINCUT.preOTK_FINCUT; " +
                         "END;";

          List<OracleParameter> lstPrm = new List<OracleParameter>();

          OracleParameter oprm = new OracleParameter
          {
            ParameterName = "PTOLS",
            DbType = DbType.Decimal,
            OracleDbType = OracleDbType.Number,
            Direction = ParameterDirection.Input,
            Value = prm.ThicknessF3
          };
          lstPrm.Add(oprm);

          Odac.ExecuteNonQuery(sqlStmt3, CommandType.Text, true, lstPrm, true);

          var sqlStmt4 = "SELECT * FROM VIZ_PRN.V_FINCUT_SORT";
          odr = Odac.GetOracleReader(sqlStmt4, CommandType.Text, false, null, null);

          if (odr != null){

            int row = 7;

            while (odr.Read()){
              CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue(1);
              CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue(2);

              row++;
            }
          }

          odr.Close();
          odr.Dispose();

          CurrentWrkSheet.Cells[4, 3].Value = "Толщина: " + prm.ThicknessF3.ToString();

        } else if (prm.TypeFilterF3 == 1){
           DbVar.SetString(prm.ListLocNumF3);

           var sqlStmt5 = "BEGIN " +
                          "DELETE FROM VIZ_PRN.TMP_FINCUT; " +
                          "insert into VIZ_PRN.TMP_FINCUT(ME_ID, LOCNO, DICKEOUTPUT, shir, GEWOUTPUT) " +
                          "select ME_ID, LOCNO, DICKEOUTPUT, SHIR, GEWOUTPUT from VIZ_PRN.V_FINCUT_LOCAL_CORE; " +
                          "VIZ_PRN.OTK_FINCUT.preOTK_FINCUT; " +
                          "END;";

           Odac.ExecuteNonQuery(sqlStmt5, CommandType.Text, false, null);

           var sqlStmt6 = "SELECT * FROM VIZ_PRN.V_FINCUT_SORT";
           odr = Odac.GetOracleReader(sqlStmt6, CommandType.Text, false, null, null);

           if (odr != null){

             int row = 7;

             while (odr.Read()){
               CurrentWrkSheet.Cells[row, 2].Value = odr.GetValue(1);
               CurrentWrkSheet.Cells[row, 3].Value = odr.GetValue(2);

               row++;
             }
           }

           odr.Close();
           odr.Dispose();

           CurrentWrkSheet.Cells[4, 3].Value = "Лок №: " + prm.ListLocNumF3;
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select(); //выбираем лист
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[2, 8].Value = $"с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

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
        Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.OTK_DEF WHERE ZDN = " + zdn.ToString(), CommandType.Text, false, null);
      }

      return Result;
    }


  }






}

