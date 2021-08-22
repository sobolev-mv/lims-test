﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;

namespace Viz.WrkModule.RptManager.Db
{
  public sealed class LasScrAfterFinCutRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }

    public LasScrAfterFinCutRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class LasScrAfterFinCut : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as LasScrAfterFinCutRptParam);
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

    private Boolean RunRpt(LasScrAfterFinCutRptParam prm, dynamic CurrentWrkSheet)
    {
      //IAsyncResult iar = null;
      OracleDataReader odr = null;
      Boolean Result = false;

      DateTime? dtBegin = null;
      DateTime? dtEnd = null;


      try{
        DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1);
        dtBegin = DbVar.GetDateBeginEnd(true, true);
        dtEnd = DbVar.GetDateBeginEnd(false, true);
        CurrentWrkSheet.Cells[2, 1].Value = $"за период с {dtBegin:dd.MM.yyyy HH:mm:ss} по {dtEnd:dd.MM.yyyy HH:mm:ss}";

        const string sqlStmt1 = "VIZ_PRN.LSR_MET.preLSR_MET";
        Odac.ExecuteNonQuery(sqlStmt1, CommandType.StoredProcedure, false, null);

        const string sqlStmt = "SELECT * FROM VIZ_PRN.V_FINCUT_LSR ORDER BY 1";
        odr = Odac.GetOracleReader(sqlStmt, CommandType.Text, false, null, null);

        if (odr != null){
          int flds = odr.FieldCount;
          int row = 7;

          while (odr.Read()){
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 1].Value = odr.GetValue(i);

            row++;
          }
          
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
      }

      return Result;
    }


  }






}

