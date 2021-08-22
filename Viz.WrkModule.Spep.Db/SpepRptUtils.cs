using System;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Collections.ObjectModel;
using Devart.Data.Oracle;
using Smv.Data.Oracle;
using Viz.DbApp.Psi;
using System.Net;
using Smv.Utils;


namespace Viz.WrkModule.Spep.Db
{

  public sealed class SpepStageResult
  {
    public string Message { get; set; }
    public ImageSource Glyph { get; set; }

    public SpepStageResult(string message, ImageSource glyph)
    {
      this.Message = message;
      this.Glyph = glyph;
    }
  }


  public sealed class SpepRptParam : Smv.Xls.XlsInstanceParam
  {
    public DateTime SpepDateTime { get; set; }
    public Boolean IsSendTo { get; set; }
    public Boolean IsAutomat { get; set; }
    public string ConStrMagLab { get; set; }
    public string RarCmd { get; set; }
    public string RarCmdParam { get; set; }
    public string FtpServerIP { get; set; }
    public string FtpUserID { get; set; }
    public string FtpPassword { get; set; }
    public string WputCmd { get; set; }
    public string WputCmdParam { get; set; }
    public ObservableCollection<SpepStageResult> CollectResult { get; set; }

    public SpepRptParam(string sourceXlsFile, string destXlsFile, DateTime spepDate, Boolean isSendTo, ObservableCollection<SpepStageResult> col)
      : base(sourceXlsFile, destXlsFile)
    {
      SpepDateTime = spepDate;
      this.IsSendTo = isSendTo;
      this.CollectResult = col;
    }
  }

  public sealed class SpepRpt : Smv.Xls.XlsRpt
  {
    protected override void DoWorkXls(object sender, DoWorkEventArgs e)
    {
      var prm = (e.Argument as SpepRptParam);
      dynamic wrkSheet = null;
      int row = 1;
      string errorMsg = " ";

      DateTime? spDate1 = null;
      DateTime? spDate2 = null;

      try{
        Debug.Assert(prm != null, "prm != null");
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => prm.CollectResult.Clear()));
        int ver = GetExcelVersion(prm.ExcelApp.Version);

        if (ver == 0)
          throw new ArgumentNullException("Ver:GetExcelVersion");

        //Здесь коррекция даты для внешнего автомата
        if (prm.IsAutomat)
          if ((DateTime.Now.Hour >= 0) && (DateTime.Now.Hour < 8))
            prm.SpepDateTime = prm.SpepDateTime.AddDays(-1);


        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetRangeDate(prm.SpepDateTime, prm.SpepDateTime, 1)));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { spDate1 = DbVar.GetDateBeginEnd(true, true); }));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { spDate2 = DbVar.GetDateBeginEnd(false, true); }));

        DateTime fdt1 = Convert.ToDateTime(spDate1);
        DateTime fdt2 = Convert.ToDateTime(spDate2);

        AddResultToProtocol(prm, true, "Период", "с " + fdt1.ToString("dd.MM.yyyy HH:mm:ss") + " по " + fdt2.ToString("dd.MM.yyyy HH:mm:ss"), "");

        string outFileName = GetFullNameTagetSpepFile(prm);
        if (outFileName == null)
          throw new ArgumentNullException("OutFileName:GetFullNameTagetSpepFile");

        //Выбираем нужный лист 
        prm.ExcelApp.ActiveWorkbook.WorkSheets[1].Select(); //выбираем лист
        wrkSheet = prm.ExcelApp.ActiveSheet;

        //Здесь сбор данных по соответсвующим объектам контроля согласно спецификации
        //Boolean resSpObj5 = SpepObj5(prm, ref row, wrkSheet, ref errorMsg);
        //AddResultToProtocol(prm, resSpObj5, "Объект Контроля №5", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);

        //Boolean resSpObj6 = SpepObj6(prm, ref row, wrkSheet, ref errorMsg);
        //AddResultToProtocol(prm, resSpObj6, "Объект Контроля №6", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);

        //Boolean resSpObj7 = SpepObj7(prm, ref row, wrkSheet, ref errorMsg);
        //AddResultToProtocol(prm, resSpObj7, "Объект Контроля №7", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);

        //Boolean resSpObj12 = SpepObj12(prm, ref row, wrkSheet, ref errorMsg);
        //AddResultToProtocol(prm, resSpObj12, "Объект Контроля №12", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);

        Boolean resSpObj13 = SpepObj13(prm, ref row, wrkSheet, ref errorMsg, prm.SpepDateTime);
        AddResultToProtocol(prm, resSpObj13, "Объект Контроля №13", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);

        //Boolean resSpObj14 = SpepObj14(prm, ref row, wrkSheet, ref errorMsg);
        //AddResultToProtocol(prm, resSpObj14, "Объект Контроля №14", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);

        //Boolean resSpObj8 = SpepObj8(prm, ref row, wrkSheet, ref errorMsg);
        //AddResultToProtocol(prm, resSpObj8, "Объект Контроля №8", "Успешный сбор данных.", "Ошибка! - Сбор данных не выполнен: " + errorMsg);


        //Здесь формирование самого отчета
        //wrkSheet.Range("A1").Value = prm.ExcelApp.Version;
        //wrkSheet.Range("A2").Value = "asdadsdgsfgsfsg";

        //Здесь визуализация Экселя
        //prm.ExcelApp.ScreenUpdating = true;
        //prm.ExcelApp.Visible = true;

        if (ver > 12)
          prm.WorkBook.SaveAs(outFileName + ".xls", 18);
        else
          prm.WorkBook.SaveAs(outFileName + ".xls");

        prm.WorkBook.Close(true);
        prm.ExcelApp.Quit();

        //Запуск архивации
        /*
        var proc =
          new Process
          {
            StartInfo =
            {
              UseShellExecute = false,
              FileName = prm.RarCmd,
              Arguments = string.Format(prm.RarCmdParam, outFileName + ".rar", outFileName + ".xls"),
              CreateNoWindow = true
            }
          };
        proc.Start();
        proc.WaitForExit();
        proc.Dispose();
        */

        if ((File.Exists(outFileName + ".xls")) && (prm.IsSendTo)){
          //Boolean resFtp = this.UploadFtp(outFileName + ".rar", prm.FtpServerIP, prm.FtpUserID, prm.FtpPassword, ref errorMsg);
          File.Copy(outFileName + ".xls", outFileName + ".rrr", true);
          Boolean resFtp = this.UploadFtpWput(outFileName + ".xls", prm.WputCmd, prm.WputCmdParam, ref errorMsg);
          AddResultToProtocol(prm, resFtp, "Отправка файла на FTP-сервер НЛМК", "Файл успешно отправлен.", "Ошибка! - Отправка файла не выполнена: " + errorMsg);

          if (resFtp){
            File.Copy(outFileName + ".rrr", outFileName + ".xls", true);
            File.Delete(outFileName + ".rrr");
          }
        }


      }
      catch (Exception ex){
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка Excel", ex.Message, MessageBoxImage.Stop)));
      }
      finally
      {
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

    private string GetFullNameTagetSpepFile(SpepRptParam prm)
    {
      try
      {
        string destDir = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "SpepOutCatalog");
        prm.RarCmd = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "RarCmd");
        prm.RarCmdParam = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "RarCmdParam");
        prm.FtpServerIP = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "FtpServer");
        prm.FtpUserID = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "FtpUser");
        prm.FtpPassword = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "FtpPsw");
        prm.ConStrMagLab = Smv.App.Config.ConfigParam.ReadConnectionStringParamValue(prm.DestXlsFile, "ConStrMagLab");
        prm.WputCmd = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "WputCmd");
        prm.WputCmdParam = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(prm.DestXlsFile, "WputCmdParam");

        var di = new DirectoryInfo(destDir);
        if (!di.Exists)
          throw new Exception("Каталог: " + destDir + " не существует!");

        string strDate = prm.SpepDateTime.ToString("ddMMyyyy");

        int num = (from f in di.EnumerateFiles("viz" + strDate + "*.xls")
                   select Convert.ToInt32(f.Name.Replace(".xls", null).Substring(11))).Count<int>();

        if (num > 0)
          num = (from f in di.EnumerateFiles("viz" + strDate + "*.xls")
                 select Convert.ToInt32(f.Name.Replace(".xls", null).Substring(11))).Max<int>();

        return destDir + "\\" + "viz" + strDate + (num + 1).ToString();
      }
      catch (Exception)
      {
        return null;
      }
    }

    private int GetExcelVersion(String verExcel)
    {
      var astr = verExcel.Split('.');
      return Convert.ToInt32(astr[0]);
    }


    private void AddResultToProtocol(SpepRptParam prm, Boolean res, string spepObj, string message, string errorMessage)
    {
      if (!res)
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => prm.CollectResult.Add(new SpepStageResult(spepObj + ": " + errorMessage, new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.Spep.Db;Component/Images/StageError-16x16.png"))))));
      else
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => prm.CollectResult.Add(new SpepStageResult(spepObj + ": " + message, new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.Spep.Db;Component/Images/StageOk-16x16.png"))))));
    }

    private Boolean SpepObj5(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg)
    {
      OracleDataReader odr = null;
      Boolean result = false;
      var oef = new OdacErrorInfo();

      try
      {
        string stmt = "VIZ_PRN.SPEP_MGO"; ;
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => Odac.ExecuteNonQuery(stmt, System.Data.CommandType.StoredProcedure, false, null)));

        stmt = "SELECT * FROM VIZ_PRN.REP_SPEP_MGO";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(stmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null)
        {
          errorMsg = oef.ErrorMsg;
          return false;
        }

        int flds = odr.FieldCount;
        while (odr.Read())
        {

          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value = typeof(DateTime) == odr.GetFieldType(i) ? odr.GetDateTime(i) : odr.GetValue(i);
          row++;
        }
        result = true;
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }

    private Boolean SpepObj6(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg)
    {
      OracleDataReader odr = null;
      bool result;
      var oef = new OdacErrorInfo();

      try
      {
        const string sqlStmt = "SELECT * FROM VIZ_PRN.REP_SPEP_S2";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null)
        {
          errorMsg = oef.ErrorMsg;
          return false;
        }

        int flds = odr.FieldCount;
        while (odr.Read())
        {

          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value = typeof(DateTime) == odr.GetFieldType(i) ? odr.GetDateTime(i) : odr.GetValue(i);
          row++;
        }
        result = true;
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }

    private Boolean SpepObj7(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg)
    {
      OracleDataReader odr = null;
      Boolean result;
      var oef = new OdacErrorInfo();

      try
      {
        const string sqlStmt = "SELECT * FROM VIZ_PRN.REP_SPEP_C_OO";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null)
        {
          errorMsg = oef.ErrorMsg;
          return false;
        }

        int flds = odr.FieldCount;
        while (odr.Read())
        {

          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value = typeof(DateTime) == odr.GetFieldType(i) ? odr.GetDateTime(i) : odr.GetValue(i);
          row++;
        }
        result = true;
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


    private Boolean SpepObj12(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg)
    {
      OracleDataReader odr = null;
      bool result;
      var oef = new OdacErrorInfo();

      try
      {
        const string sqlStmt = "SELECT * FROM VIZ_PRN.REP_SPEP_PPP";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null)
        {
          errorMsg = oef.ErrorMsg;
          return false;
        }

        int flds = odr.FieldCount;
        while (odr.Read())
        {
          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value = typeof(DateTime) == odr.GetFieldType(i) ? odr.GetDateTime(i) : odr.GetValue(i);
          row++;
        }
        result = true;
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


    private Boolean SpepObj13(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg, DateTime dtMagLab)
    {
      Boolean result;
      var con = new OleDbConnection();
      OleDbDataReader rd = null;

      try
      {
        con.ConnectionString = prm.ConStrMagLab;
        con.Open();
        var cmd = new OleDbCommand("dbo.SpepI9", con)
        {
          CommandType = System.Data.CommandType.StoredProcedure,
          CommandTimeout = 4 * 60
        };

        cmd.Parameters.Add(new OleDbParameter());
        cmd.Parameters[0].DbType = System.Data.DbType.DateTime;
        cmd.Parameters[0].Direction = System.Data.ParameterDirection.Input;
        cmd.Parameters[0].Value = dtMagLab;

        rd = cmd.ExecuteReader();

        Debug.Assert(rd != null, "rd != null");
        int flds = rd.FieldCount;
        while (rd.Read()){

          //Только для АВО-7
          if (/*(rd.GetInt32(5) != 12) &&*/ (rd.GetInt32(5) != 14))
            continue;

          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value = typeof(DateTime) == rd.GetFieldType(i) ? rd.GetDateTime(i) : rd.GetValue(i);

          row++;
        }

        result = true;
        con.Dispose();
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (rd != null)
        {
          rd.Close();
          rd.Dispose();
          con.Dispose();
        }
      }

      return result;
    }


    private Boolean SpepObj14(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg)
    {
      OracleDataReader odr = null;
      bool result;
      var oef = new OdacErrorInfo();

      try
      {
        const string sqlStmt = "SELECT * FROM VIZ_PRN.REP_SPEP_14";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString("FINISHED_GOODS")));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null)
        {
          errorMsg = oef.ErrorMsg;
          return false;
        }

        int flds = odr.FieldCount;
        while (odr.Read())
        {

          for (int i = 0; i < flds; i++)
            currentWrkSheet.Cells[row, i + 1].Value = typeof(DateTime) == odr.GetFieldType(i) ? odr.GetDateTime(i) : odr.GetValue(i);

          row++;
        }
        result = true;
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }


    private Boolean SpepObj8(SpepRptParam prm, ref int row, dynamic currentWrkSheet, ref string errorMsg)
    {
      OracleDataReader odr = null;
      Boolean result;
      var oef = new OdacErrorInfo();

      try
      {
        const string sqlStmt1 = "SELECT * FROM VIZ_PRN.REP_SPEP_8";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DbVar.SetString("FINISHED_GOODS")));
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt1, System.Data.CommandType.Text, false, null, oef); }));

        if (odr == null)
        {
          errorMsg = oef.ErrorMsg;
          return false;
        }

        int flds = odr.FieldCount;

        if (flds == 0)
        {
          odr.Close();
          odr.Dispose();
          return (true);
        }

        currentWrkSheet.Cells[row, 1].Value = 8;
        currentWrkSheet.Cells[row, 2].Value = prm.SpepDateTime.ToString("ddMMyyyy");
        currentWrkSheet.Cells[row, 3].Value = "I";
        currentWrkSheet.Cells[row, 5].Value = prm.SpepDateTime;

        while (odr.Read())
        {
          int fTols = Convert.ToInt32(Convert.ToDecimal(odr.GetValue("TOLS")) * 100);

          if (fTols == 27)
          {
            currentWrkSheet.Cells[row, 6].Value = odr.GetValue("PAP");
            currentWrkSheet.Cells[row, 8].Value = odr.GetValue("DP1750");
          }
          else
          {
            currentWrkSheet.Cells[row, 7].Value = odr.GetValue("PAP");
            currentWrkSheet.Cells[row, 9].Value = odr.GetValue("DP1750");
          }
        }

        //--------------
        odr.Close();
        odr.Dispose();
        const string sqlStmt2 = "select  viz_prn.P1750AP_P1750L('HEAD',0.27) IB_27, viz_prn.P1750AP_P1750L('HEAD',0.3) IB_30, P1750AP_P1750L('TAIL',0.27) IE_27, P1750AP_P1750L('TAIL',0.3) IE_30 from dual";
        prm.Disp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => { odr = Odac.GetOracleReader(sqlStmt2, System.Data.CommandType.Text, false, null, oef); }));

        while (odr.Read())
        {
          currentWrkSheet.Cells[row, 10].Value = odr.GetValue("IB_27");
          currentWrkSheet.Cells[row, 11].Value = odr.GetValue("IB_30");
          currentWrkSheet.Cells[row, 12].Value = odr.GetValue("IE_27");
          currentWrkSheet.Cells[row, 13].Value = odr.GetValue("IE_30");
        }

        result = true;
      }
      catch (Exception ex)
      {
        errorMsg = ex.Message;
        result = false;
      }
      finally
      {
        if (odr != null)
        {
          odr.Close();
          odr.Dispose();
        }
      }

      return result;
    }

    private Boolean UploadFtp(string filename, string ftpServerIP, string ftpUserID, string ftpPassword, ref string errorMsg)
    {

      Boolean res = true;

      FileInfo fileInf = new FileInfo(filename);
      //string uri = "ftp://" + ftpServerIP + "/" + fileInf.Name;

      FtpWebRequest reqFTP;

      // Create FtpWebRequest object from the Uri provided

      reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("ftp://" + ftpServerIP + "/" + fileInf.Name));
      reqFTP.Proxy = null;
      // Provide the WebPermission Credintials

      reqFTP.Credentials = new NetworkCredential(ftpUserID, ftpPassword);

      // By default KeepAlive is true, where the control connection is not closed
      // after a command is executed.
      reqFTP.KeepAlive = false;


      // Specify the command to be executed.
      reqFTP.Method = WebRequestMethods.Ftp.UploadFile;

      // Specify the data transfer type.
      reqFTP.UseBinary = true;
      reqFTP.UsePassive = true;
      reqFTP.Timeout = 20000;

      // Notify the server about the size of the uploaded file

      reqFTP.ContentLength = fileInf.Length;
      // The buffer size is set to 2kb

      int buffLength = 2048;

      byte[] buff = new byte[buffLength];

      int contentLen;

      // Opens a file stream (System.IO.FileStream) to read the file to be uploaded

      FileStream fs = fileInf.OpenRead();

      try
      {

        // Stream to which the file to be upload is written

        Stream strm = reqFTP.GetRequestStream();

        // Read from the file stream 2kb at a time

        contentLen = fs.Read(buff, 0, buffLength);

        // Till Stream content ends

        while (contentLen != 0)
        {

          // Write Content from the file stream to the FTP Upload Stream

          strm.Write(buff, 0, contentLen);

          contentLen = fs.Read(buff, 0, buffLength);

        }

        // Close the file stream and the Request Stream

        strm.Close();
        fs.Close();
      }

      catch (Exception ex)
      {
        errorMsg = ex.Message;
        res = false;
        //MessageBox.Show(ex.Message, "Upload Error");

      }
      return res;
    }

    private Boolean UploadFtpWput(string filename, string wputCmd, string wputCmdParam, ref string errorMsg)
    {
      Boolean res = true;

      FileInfo fileInf = new FileInfo(filename);
      //string uri = "ftp://" + ftpServerIP + "/" + fileInf.Name;

      //Запуск архивации
      var proc = new Process
      {
        StartInfo =
        {
          UseShellExecute = false,
          FileName = Etc.StartPath + "\\" + wputCmd,
          Arguments = string.Format(wputCmdParam, filename, fileInf.Name),
          CreateNoWindow = true
        }
      };
      proc.Start();
      proc.WaitForExit();
      proc.Dispose();

      if (File.Exists(filename))
      {
        errorMsg = "Wput.exe ftp - ошибка отправки";
        res = false;
      }

      return res;
    }



  }
}
