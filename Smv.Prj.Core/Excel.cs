using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

namespace Smv.Xls
{
  public delegate void ProcConnectToTargetDb(int? idReport, string dbAlias = null);

  public delegate string ProcGetCurrentDbAlias();

  public class XlsInstanceParam
  {
    public string  SourceXlsFile{get; private set;}
    public string  DestXlsFile{get; set;}
    public int ExcelVersion{get; set;}
    public dynamic ExcelApp{get; set;}
    public dynamic WorkBook{get; set;}
    public System.Windows.Threading.Dispatcher Disp{ get; set; } 

    public XlsInstanceParam(string sourceXlsFile, string destXlsFile)
    {
      SourceXlsFile = sourceXlsFile;
      DestXlsFile = destXlsFile;
      Disp = System.Windows.Threading.Dispatcher.CurrentDispatcher;
    }
  }

  public sealed class XlsInstanceBackgroundReport
  {
    private DoWorkEventHandler doWorkHandler = null;
    private RunWorkerCompletedEventHandler completeWorkHandler = null;    
    private BackgroundWorker bw = new BackgroundWorker();

    private int GetExcelVersion(String VerExcel)
    {
      string[] astr = VerExcel.Split('.');

      if (astr != null)
        return Convert.ToInt32(astr[0]);
      else
        throw new NotImplementedException("Ver:GetExcelVersion");
    }

    public Boolean RunBackgroundXlsReport(DoWorkEventHandler doWork, RunWorkerCompletedEventHandler completeWork, XlsInstanceParam arg, string local = "ru")
    {
      string verMso;
      string hdrMsg;

      if (local == "ru"){
        hdrMsg = "Ошибка";
        verMso = "Версия пакета MS Offce ниже чем MS Offce 2007!";
      }
      else{
        hdrMsg = "Error";
        verMso = "The version of a MS Offce is lower than MS Offce 2007!";
      }
      
      if (doWorkHandler != null)
        this.bw.DoWork -= doWorkHandler;

      if (completeWorkHandler != null)
        this.bw.RunWorkerCompleted -= completeWorkHandler;      

      this.bw.DoWork += doWork; 
      this.bw.RunWorkerCompleted += completeWork;

      this.doWorkHandler = doWork;
      this.completeWorkHandler = completeWork;

      try{
        /*
        Process proc = Process.GetProcesses().FirstOrDefault(p => p.ProcessName.StartsWith("EXCEL", StringComparison.InvariantCultureIgnoreCase));
        if (proc != null)
          throw new Exception(excelApp);  
        */ 

        Type xlAppType = Type.GetTypeFromProgID("Excel.Application");
        arg.ExcelApp = Activator.CreateInstance(xlAppType);
        arg.ExcelVersion = GetExcelVersion(arg.ExcelApp.Version);

        //Проверяем что оффисе MS Office не ниже 2007
        if (arg.ExcelVersion <= 12)
          throw new Exception(verMso);

        arg.ExcelApp.ScreenUpdating = false;
        arg.ExcelApp.DisplayAlerts = false;

        arg.WorkBook = arg.ExcelApp.Workbooks.Open(arg.SourceXlsFile);
        this.bw.RunWorkerAsync(arg);

        arg.ExcelApp.DisplayAlerts = true;
        return true;
      }
      catch (Exception ex){
        Utils.DxInfo.ShowDxBoxInfo(hdrMsg, ex.Message, MessageBoxImage.Stop);

        if (arg.ExcelApp != null){
          Marshal.ReleaseComObject(arg.ExcelApp);
          arg.ExcelApp = null;
          GC.Collect();
        }
        return false;
      }
    }

    public Boolean RunBackground(DoWorkEventHandler doWork, RunWorkerCompletedEventHandler completeWork, XlsInstanceParam arg)
    {
      if (doWorkHandler != null)
        this.bw.DoWork -= doWorkHandler;

      if (completeWorkHandler != null)
        this.bw.RunWorkerCompleted -= completeWorkHandler;

      this.bw.DoWork += doWork;
      this.bw.RunWorkerCompleted += completeWork;

      this.doWorkHandler = doWork;
      this.completeWorkHandler = completeWork;

      try{
        this.bw.RunWorkerAsync(arg);
        return true;
      }
      catch (Exception ex){
        Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop);
        return false;
      }
    } 
  }

  public abstract class XlsRpt
  {
    public int IdReport { get; set; } = -1;
    public string OldDbAlias { get; set; } = null;
    public ProcConnectToTargetDb ConnectToTargetDb { get; set; } = null;
    public ProcGetCurrentDbAlias GetCurrentDbAlias { get; set; } = null;

    //Вызывать в наследнике в конце, в случае переключения целевой БД
    protected virtual void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      if (this.IdReport > 0){
        ConnectToTargetDb(null, this.OldDbAlias);

        //отцепляем делегаты
        GetCurrentDbAlias = null;
        ConnectToTargetDb = null;
      }
    }

    protected virtual void SaveResult(XlsInstanceParam prm, string rptNameFolder = "Отчеты Lims")
    {
      //Проверяем существует ли папка
      if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder))
        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder);

      if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile)))
        File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile));

      prm.WorkBook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile));
      prm.DestXlsFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile);
      //System.Diagnostics.Process.Start("explorer.exe", @"/root," + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @",/select," + System.IO.Path.GetFileName(prm.DestXlsFile));
      Process.Start("explorer.exe", @"/e,/select," + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile));
    }

    protected virtual void SaveResultToXlsFormat(XlsInstanceParam prm)
    {
      const string rptNameFolder = "Отчеты Lims";

      //Проверяем существует ли папка
      if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder))
        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder);

      if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile)))
        File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile));

      if (prm.ExcelVersion <= 12)
        prm.WorkBook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile));
      else
        prm.WorkBook.SaveAs(
          Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" +
          Path.GetFileName(prm.DestXlsFile), 18);

      //System.Diagnostics.Process.Start("explorer.exe", @"/root," + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @",/select," + System.IO.Path.GetFileName(prm.DestXlsFile));
      Process.Start("explorer.exe", @"/e,/select," + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + rptNameFolder + "\\" + Path.GetFileName(prm.DestXlsFile));
    }

    public virtual Boolean RunXls(XlsInstanceBackgroundReport rpt, RunWorkerCompletedEventHandler completeWork, XlsInstanceParam prm, string local = "ru")
    {
      //Здесь будем переключать базу отчета при необходимости
      if (IdReport > 0){
        this.OldDbAlias = GetCurrentDbAlias();
        ConnectToTargetDb(IdReport);
      }
      
      return rpt.RunBackgroundXlsReport(DoWorkXls, completeWork, prm, local);
    }

    public virtual Boolean Run(XlsInstanceBackgroundReport rpt, RunWorkerCompletedEventHandler completeWork, XlsInstanceParam prm)
    {
      return rpt.RunBackground(DoWorkXls, completeWork, prm);
    }

  }


}
