using System;
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
  public sealed class AlstIsolRptParam : Smv.Xls.XlsInstanceParam
  {
    public string ClientOrder { get; set; }
    public string ClientOrderPos { get; set; }

    public AlstIsolRptParam(string sourceXlsFile, string destXlsFile, string ClientOrder, string ClientOrderPos)
      : base(sourceXlsFile, destXlsFile)
    {
      this.ClientOrder = ClientOrder;
      this.ClientOrderPos = ClientOrderPos;
    }
  }

  public sealed class AlstomIsol : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      AlstIsolRptParam prm = (e.Argument as AlstIsolRptParam);
      dynamic wrkSheet = null;

      try
      {
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
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally
      {
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

    private int GetExcelVersion(String VerExcel)
    {
      string[] astr = VerExcel.Split('.');

      if (astr != null)
        return Convert.ToInt32(astr[0]);
      else
        return 0;
    }

    private Boolean RunRpt(AlstIsolRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      string SqlStmt = null;
      Boolean Result = false;
      OdacErrorInfo oef = new OdacErrorInfo();

      

      try
      {
        //SqlStmt = System.IO.File.ReadAllText(Smv.Utils.Etc.StartPath + "\\Sql\\Viz.WrkModule.RptMagLab-MatTechStep.sql", System.Text.Encoding.GetEncoding(1251)).Replace("\r", " ");
        SqlStmt = "SELECT * FROM VIZ_PRN.ALSTOM1";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { DbVar.SetString(prm.ClientOrder, prm.ClientOrderPos); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(SqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null) return false;
        CurrentWrkSheet.Cells[2, 19].Value2 = prm.ClientOrder + "/" + prm.ClientOrderPos;      
      
        int flds = odr.FieldCount;
        int row = 7;

        while (odr.Read()){

          CurrentWrkSheet.Rows[row].Select();
          CurrentWrkSheet.Rows[row].Copy();
          CurrentWrkSheet.Rows[row + 1].Select();
          CurrentWrkSheet.Paste();

          for (int i = 0; i < flds; i++){
            CurrentWrkSheet.Cells[row, i + 1].Value2 = odr.GetValue(i);
          }

          row++;
        }

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

