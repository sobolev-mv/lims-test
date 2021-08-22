using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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
  public sealed class OtkNadavVtoRptParam : RptWithF1Param
  {
    public decimal Glubina { get; set; }
    public string Defect { get; set; }
    public OtkNadavVtoRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class OtkNadavVto : RptWithF1
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as OtkNadavVtoRptParam);
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

    private Boolean RunRpt(OtkNadavVtoRptParam prm, dynamic CurrentWrkSheet)
    {
      OracleDataReader odr = null;
      IAsyncResult iar = null;
      Boolean Result = false;
      DateTime? dtBegin = null;
      DateTime? dtEnd = null;


      try{
       PrepareFilterRpt(prm);

       //int rs = Convert.ToInt32(Odac.ExecuteScalar("select count(*) from VIZ_PRN.TMP_OTK_FILTR_CORE", CommandType.Text, false, null));
       //MessageBox.Show(rs.ToString());

       prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtBegin = DbVar.GetDateBeginEnd(true, true); }));
       prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { dtEnd = DbVar.GetDateBeginEnd(false, true); }));

        //Для всеx дефектов кроме 501 ("Надав ВТО") глубина залегания не должна влиять на результаты запроса
      //if (prm.Defect != "501")
        //  prm.Glubina = -1000000;

        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetNum(prm.Glubina)));
        //prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString(prm.Defect)));

        CurrentWrkSheet.Cells[1, 1].Value = "Cписок рулонов, % " + prm.Defect;
        CurrentWrkSheet.Cells[1, 6].Value = string.Format("за период с {0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        CurrentWrkSheet.Cells[2, 5].Value = prm.Glubina;
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[2, 8].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;

        var rm = new Random();
        Double zdn = rm.Next(10000000, 99999999);
        DbVar.SetNum(Convert.ToDecimal(zdn));

        List<OracleParameter> lstPrm = new List<OracleParameter>();

        OracleParameter prmProc = new OracleParameter();
        prmProc.DbType = DbType.Double;
        prmProc.Direction = ParameterDirection.Input;
        prmProc.OracleDbType = OracleDbType.Number;
        //prmProc.Size = 64;
        prmProc.Value = zdn;
        lstPrm.Add(prmProc);

        prmProc = new OracleParameter();
        prmProc.DbType = DbType.String;
        prmProc.Direction = ParameterDirection.Input;
        prmProc.OracleDbType = OracleDbType.VarChar;
        prmProc.Size = prm.Defect.Length;
        prmProc.Value = prm.Defect;
        lstPrm.Add(prmProc);

        prmProc = new OracleParameter();
        prmProc.DbType = DbType.Decimal;
        prmProc.Direction = ParameterDirection.Input;
        prmProc.OracleDbType = OracleDbType.Number;
        prmProc.Value = prm.Glubina;
        lstPrm.Add(prmProc);

        Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.NADAVVTO", CommandType.StoredProcedure, false, lstPrm);

        //Odac.ExecuteNonQuery("VIZ_PRN.OTK_DEFECT.PREOTK_DEFECT", CommandType.StoredProcedure, false, null);

        DbVar.SetNum(Convert.ToDecimal(zdn));

        const string SqlStmt1 = "SELECT * FROM VIZ_PRN.OTK_NADAV_VTO_501R";

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt1, CommandType.Text, false, null, null); }));
        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 5;
          var flds = odr.FieldCount;

          while (odr.Read()){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 2], CurrentWrkSheet.Cells[row, 16]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 2], CurrentWrkSheet.Cells[row + 1, 16]]);
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
            row++;
          }
          odr.Close();
          odr.Dispose();
        }

        //выбираем лист 2
        prm.ExcelApp.ActiveWorkbook.WorkSheets[2].Select();
        CurrentWrkSheet = prm.ExcelApp.ActiveSheet;
        CurrentWrkSheet.Cells[1, 1].Value = "Cписок стендовых партий, % " + prm.Defect;
        CurrentWrkSheet.Cells[1, 6].Value = string.Format("за период с {0:dd.MM.yyyy}", dtBegin) + " по " + string.Format("{0:dd.MM.yyyy}", dtEnd);
        CurrentWrkSheet.Cells[2, 5].Value = prm.Glubina;
        if (prm.TypeFilter >= 1)
          CurrentWrkSheet.Cells[2, 8].Value = prm.TypeFilter == 1 ? prm.GetFilterCriteria() : "Список стендов: " + prm.ListStendF1;

        const string SqlStmt2 = "SELECT * FROM VIZ_PRN.OTK_NADAV_VTO_501S";

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.GetOracleReaderAsync(SqlStmt2, CommandType.Text, false, null, null); }));
        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null) odr = oracleCommand.EndExecuteReader(iar);

        if (odr != null){
          var row = 5;
          var flds = odr.FieldCount;

          while (odr.Read()){
            CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row, 2], CurrentWrkSheet.Cells[row, 11]].Copy(CurrentWrkSheet.Range[CurrentWrkSheet.Cells[row + 1, 2], CurrentWrkSheet.Cells[row + 1, 11]]);
            for (int i = 0; i < flds; i++)
              CurrentWrkSheet.Cells[row, i + 2].Value = odr.GetValue(i);
            row++;
          }
        }

        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select();
        Result = true;
      }
      catch (Exception e){
        MessageBox.Show(e.Message);
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





