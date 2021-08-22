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
  public sealed class BalanceWrkTimeRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public BalanceWrkTimeRptParam(string sourceXlsFile, string destXlsFile)
      : base(sourceXlsFile, destXlsFile)
    { }
  }

  public sealed class BalanceWrkTime : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as BalanceWrkTimeRptParam);
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

    private Boolean RunRpt(BalanceWrkTimeRptParam prm, dynamic CurrentWrkSheet)
    {
      var nameMonth = new string[] {"ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ", "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ", "НОЯБРЬ", "ДЕКАБРЬ" };

      OracleDataReader odr = null;
      Boolean Result = false;
      Int64 zdn = 0;

      DateTime? dtBegin = null;
      DateTime? dtEnd = null;

      //генерим отрицательный номер задания
      var rm = new Random();
      zdn = rm.Next(10000000, 99999999) * -1;

      prm.DateBegin = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, 1);
      prm.DateEnd = new DateTime(prm.DateBegin.Year, prm.DateBegin.Month, DateTime.DaysInMonth(prm.DateBegin.Year, prm.DateBegin.Month));

      DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
      DbVar.SetNum(zdn);

      //MessageBox.Show(zdn.ToString());
      //Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.OTK_DEF WHERE ZDN = " + zdn.ToString(), CommandType.Text, false, null);

      dtBegin = DbVar.GetDateBeginEnd(true, true);
      dtEnd = DbVar.GetDateBeginEnd(false, true);
      CurrentWrkSheet.Cells[2, 9].Value = $"ЗА {nameMonth[prm.DateBegin.Month - 1]} МЕСЯЦ {prm.DateBegin.Year} Г.";

      try{
        var sqlStmt1 = "VIZ_PRN.ReportDowntimes.PREDOWNTIMESREPORT";

        List<OracleParameter> lstPrm = new List<OracleParameter>();

        var prmProc = new OracleParameter()
        {
          DbType = DbType.DateTime,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.Date,
          Value = prm.DateBegin
        };

        lstPrm.Add(prmProc);

        prmProc = new OracleParameter()
        {
          DbType = DbType.DateTime,
          Direction = ParameterDirection.Input,
          OracleDbType = OracleDbType.Date,
          Value = prm.DateEnd
        };

        lstPrm.Add(prmProc);

        Odac.ExecuteNonQuery(sqlStmt1, CommandType.StoredProcedure, false, lstPrm);

        var sqlStmt2 = "SELECT * FROM VIZ_PRN.V_BALANS_TIME";
        odr = Odac.GetOracleReader(sqlStmt2, CommandType.Text, false, null, null);

        if (odr != null){

          int flds = odr.FieldCount;
          int row = 9;

          while (odr.Read()){

            for (int i = 2; i < flds; i++)
              CurrentWrkSheet.Cells[row, i].Value = odr.GetValue(i);

            row++;
          }

          odr.Close();
          odr.Dispose();
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
        //Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.OTK_DEF WHERE ZDN = " + zdn.ToString(), CommandType.Text, false, null);
      }

      return Result;
    }


  }






}

