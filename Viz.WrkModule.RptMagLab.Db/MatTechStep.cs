using System;
using System.Diagnostics;
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
  public sealed class RptMagLabParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public string TechStepInspLot { get; set; }
    public string TechStepPrjJornal { get; set; }

    public RptMagLabParam(string sourceXlsFile, string destXlsFile, DateTime rptDateBegin, DateTime rptDateEnd, string RptTechStepIl, string RptTechStepPj)
           : base(sourceXlsFile, destXlsFile)
    {
      DateBegin = rptDateBegin;
      DateEnd = rptDateEnd;
      TechStepInspLot = RptTechStepIl;
      TechStepPrjJornal = RptTechStepPj; 
    }
  }

  public sealed class MaterialStepRpt : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as RptMagLabParam);
      dynamic wrkSheet = null;

      try{
        //Выбираем нужный лист 
        Debug.Assert(prm != null, "prm != null");
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
        Debug.Assert(prm != null, "prm != null");
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

    private int GetExcelVersion(String verExcel)
    {
      string[] astr = verExcel.Split('.');  

      if (astr != null)
        return Convert.ToInt32(astr[0]); 
      else
        return 0;
    }

    private Boolean RunRpt(RptMagLabParam prm, dynamic currentWrkSheet)
    {
      OracleDataReader odr = null;
      Boolean result = false;
      var oef = new OdacErrorInfo();
      
      int row = 3;

      try{
        //SqlStmt = System.IO.File.ReadAllText(Smv.Utils.Etc.StartPath + "\\Sql\\Viz.WrkModule.RptMagLab-MatTechStep.sql", System.Text.Encoding.GetEncoding(1251)).Replace("\r", " ");
        const string sqlStmt = "SELECT * FROM VIZ_PRN.RPTMAGLAB_MATTECHSTEP";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.TechStepInspLot, prm.TechStepPrjJornal)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;

        int flds = odr.FieldCount;

        for (int i = 16; i < flds; i++){
          currentWrkSheet.Cells[row, i + 1].Value2 = odr.GetName(i);
        }
      
        row = 5;

        while (odr.Read()){

          currentWrkSheet.Rows[row].Select();
          currentWrkSheet.Rows[row].Copy();
          currentWrkSheet.Rows[row + 1].Select();
          currentWrkSheet.Paste();

          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);

          row++;
        }

        currentWrkSheet.Cells[1, 1].Select();

        result = true;
      }
      catch (Exception){
         result = false;
      }
      finally{
        if (odr != null){
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


  }


















}

