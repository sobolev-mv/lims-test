﻿using System;
using System.Collections.Generic;
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

namespace Viz.WrkModule.RptMagLab.Db
{
  public sealed class  CzlPlosk2RptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string Rm1200 { get; set; }
    public string Aro { get; set; }
    public string Aoo { get; set; }
    public string Avo { get; set; }
    public string Apr { get; set; }
    public Boolean IsRm1200 { get; set; }
    public Boolean IsAro { get; set; }
    public Boolean IsAoo { get; set; }
    public Boolean IsAvo { get; set; }
    public Boolean IsApr { get; set; }
    public Boolean IsInList { get; set; }
    public int TypeList { get; set; }
    public string ListVal { get; set; }

    public  CzlPlosk2RptParam(string sourceXlsFile, string destXlsFile, DateTime RptDateBegin, DateTime RptDateEnd,
                              string Rm1200, string Aro, string Aoo, string Avo, string Apr,
                              Boolean IsRm1200, Boolean IsAro, Boolean IsAoo, Boolean IsAvo, Boolean IsApr,
                              Boolean IsInList, int TypeList, string ListVal)
      : base(sourceXlsFile, destXlsFile)
    {
      this.DateBegin = RptDateBegin;
      this.DateEnd = RptDateEnd;
      this.Rm1200 = Rm1200;
      this.Aro = Aro;
      this.Aoo = Aoo;
      this.Avo = Avo;
      this.Apr = Apr;
      this.IsRm1200 = IsRm1200;
      this.IsAro = IsAro;
      this.IsAoo = IsAoo;
      this.IsAvo = IsAvo;
      this.IsApr = IsApr;
      //-----------------
      this.IsInList = IsInList;
      this.TypeList = TypeList;
      this.ListVal = ListVal;
    }
  }

  public sealed class CzlPlosk2 : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as CzlPlosk2RptParam);
      dynamic wrkSheet = null;

      try{

        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;
        this.RunRpt(prm, wrkSheet);
        this.SaveResult(prm);
      }
      catch (Exception ex){
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

    private void ListFilterInfoToExcel(CzlPlosk2RptParam prm)
    {
      dynamic wrkSheet = null;
      //выбираем лист
      prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
      wrkSheet = prm.ExcelApp.ActiveSheet;

      switch (prm.TypeList){
        case 0:
          wrkSheet.Cells[3, 1].Value = "Стендовые партии:";
          break;
        case 1:
          wrkSheet.Cells[2, 1].Value = "Стенды ВТО:";
          break;
        default:
          Console.WriteLine("Default case");
          break;
      }

      const int row = 4;
      string[] strArr = prm.ListVal.Split(new char[] { ',' });
      for (int i = 0; i < strArr.Length; i++) wrkSheet.Cells[row + i, 1].Value = strArr[i];

    }

    private Boolean RunRpt(CzlPlosk2RptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;
      int row = 0;

      try{
        string SqlStmt1 = null;
        SqlStmt1 = prm.IsInList ? "SELECT * FROM VIZ_PRN.CZL_NEPL_SPIS ORDER BY 1" : "SELECT * FROM VIZ_PRN.CZL_NEPL_NSPIS ORDER BY 1";

        switch (prm.TypeList){
          case 0:
            prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.Rm1200, prm.Aro, prm.Aoo, prm.Avo, prm.Apr, prm.ListVal, string.Empty)));
            break;
          case 1:
            prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.Rm1200, prm.Aro, prm.Aoo, prm.Avo, prm.Apr, string.Empty, prm.ListVal)));
            break;
          default:
            Console.WriteLine("Default case");
            break;
        }

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        CurrentWrkSheet.Cells[4, 2].Value = "за период c " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy HH:mm:ss}", dtEnd);
        string strFlt = "";
        if (prm.IsRm1200)
          strFlt += ":Ст1200 =" + prm.Rm1200;
        if (prm.IsAro)
          strFlt += ":АРО=" + prm.Aro;
        if (prm.IsAoo)
          strFlt += ":АОО=" + prm.Aoo;
        if (prm.IsAvo)
          strFlt += ":АВО=" + prm.Avo;
        if (prm.IsApr)
          strFlt += ":АПР=" + prm.Apr;
        CurrentWrkSheet.Cells[3, 3].Value = strFlt;

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt1, System.Data.CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
          odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          row = 11;
          int flds = odr.FieldCount;

          while (odr.Read()){
            for (int i = 0; i < flds; i++) CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
            row++;
          }
        }

        ListFilterInfoToExcel(prm);

        //Возвращаемся на первую страницу
        CurrentWrkSheet = prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        CurrentWrkSheet.Cells[1, 1].Select();
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



