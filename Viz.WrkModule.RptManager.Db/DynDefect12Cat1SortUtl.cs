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
  public sealed class DynDefect12Cat1SortUtlRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime DateBegin { get; set; }
    public DateTime DateEnd { get; set; }
    public int TypeAction { get; set; }

    public DynDefect12Cat1SortUtlRptParam(string sourceXlsFile, string destXlsFile) : base(sourceXlsFile, destXlsFile)
    {}
  }

  public sealed class DynDefect12Cat1SortUtl : Smv.Xls.XlsRpt
  {

    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as DynDefect12Cat1SortUtlRptParam);
 
      try{
        if (prm.TypeAction == 1)
          this.Calc(prm);
        else
          this.Undo(prm);
          
      }
      catch (Exception ex){
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", ex.Message, MessageBoxImage.Stop)));
      }
    }

    private Boolean Calc(DynDefect12Cat1SortUtlRptParam prm)
    {
      IAsyncResult iar = null;

      try{
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1)));

        var rm = new Random();
        Int64 zdn = rm.Next(10000000, 99999999) * -1;

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetNum(zdn)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("VIZ_PRN.OTK_DINAMIKA.OTK_RASCHET", CommandType.StoredProcedure, false, false, null); }));

        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null){
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }
        
        //Не убирать второй вызов установки даты здесь нужен!
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.DateBegin, prm.DateEnd, 1))); 
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("VIZ_PRN.OTK_DINAMIKA.OTK_RASCHET_SGP", CommandType.StoredProcedure, false, false, null); }));

        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null)
        {
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }


        //Очистка!!!;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync("VIZ_PRN.OTK_AVO.postOTK_CAT_AVO", CommandType.StoredProcedure, false, false, null); }));

        if (iar != null)
            iar.AsyncWaitHandle.WaitOne();
        else
            return false;

        oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null){
            oracleCommand.EndExecuteNonQuery(iar);
            iar = null;
        }
        //Очистка!!!;



        return true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }

      return false;
    }

    private Boolean Undo(DynDefect12Cat1SortUtlRptParam prm)
    {
      IAsyncResult iar = null;

      try{
        const string stmt = "BEGIN " +
                            "DELETE FROM VIZ_PRN.OTK_DINAMIKA_KAT_SORT WHERE DATA = (SELECT MAX(DATA) FROM VIZ_PRN.OTK_DINAMIKA_KAT_SORT); " +
                            "DELETE FROM VIZ_PRN.OTK_DINAMIKA_SGP WHERE DATA = (SELECT MAX(DATA) FROM VIZ_PRN.OTK_DINAMIKA_SGP); " +
                            "DELETE FROM VIZ_PRN.OTK_DINAMIKA_TOLS WHERE DATA = (SELECT MAX(DATA) FROM VIZ_PRN.OTK_DINAMIKA_TOLS); " +
                            "END;";

        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { iar = Odac.ExecuteNonQueryAsync(stmt, CommandType.Text, false, false, null); }));

        if (iar != null)
          iar.AsyncWaitHandle.WaitOne();
        else
          return false;

        var oracleCommand = iar.AsyncState as OracleCommand;
        if (oracleCommand != null){
          oracleCommand.EndExecuteNonQuery(iar);
          iar = null;
        }

        return true;
      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка выполнения", ex.Message, MessageBoxImage.Stop)));
      }

      return false;
    }



  }






}


