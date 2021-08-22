using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Windows;
using Viz.Lims.Sh;
using System.IO.Ports;
using System.Threading;
using System.Globalization;


namespace Viz.MagLab.MeasureUnits
{
  
  internal sealed class BrokhausIsolUnit  : MeasureIsolUnit
  {
   
    private SerialPort spPressure = null;
    private SerialPort spCurrent = null;
    private object threadLock = new object(); 
    
    private System.Timers.Timer measureTimer;
    private int  timeMeasure = 0;
    private byte pressureValue = 0;
    private Boolean OldStatePressure = false;
    private Boolean CurrentStatePressure = false;

    private byte[] StringToASCIIByteArray(string text)
    {
      ASCIIEncoding encoding = new ASCIIEncoding();
      return encoding.GetBytes(text);
    }

    private string ASCIIByteArrayToString(byte[] characters)
    {
      ASCIIEncoding encoding = new ASCIIEncoding();
      return encoding.GetString(characters);
    }
    

    public BrokhausIsolUnit(string ConfigFile, uint MeasureCount)
    {
      string curUnit = "Измеритель давления:\n";
      this.IsError = false;
      this.mCount = MeasureCount;

      try{

        spPressure = new SerialPort();
        spCurrent = new SerialPort();
               
        ComPortSectionHandler sh = this.ReadParamValue(ConfigFile, "PressureComPort");
        spPressure.PortName = sh.Port;
        spPressure.BaudRate = sh.BaudRate;
        spPressure.Parity = (Parity)Enum.Parse(typeof(Parity), sh.Parity, true); 
        spPressure.DataBits = 8;
        spPressure.StopBits = (StopBits)Enum.Parse(typeof(StopBits), sh.StopBits, true);
        spPressure.Handshake = (Handshake)Enum.Parse(typeof(Handshake), sh.HandShake, true);
        spPressure.RtsEnable = false;
        spPressure.DtrEnable = true;
        spPressure.ReadTimeout = sh.SyncReadTimeOut;
        spPressure.WriteTimeout = sh.SyncWriteTimeOut;
        spPressure.Open();

        if (!InitPressure())
          throw new ApplicationException("Ошибка инициализации. Работа не возможна!");
                
        curUnit = "Измеритель тока:\n";
        //===========================================================================
        sh = this.ReadParamValue(ConfigFile, "CurrentComPort");
        spCurrent.PortName = sh.Port;
        spCurrent.BaudRate = sh.BaudRate;
        spCurrent.Parity = (Parity)Enum.Parse(typeof(Parity), sh.Parity, true);
        spCurrent.DataBits = 8;
        spCurrent.StopBits = (StopBits)Enum.Parse(typeof(StopBits), sh.StopBits, true);
        spCurrent.Handshake = (Handshake)Enum.Parse(typeof(Handshake), sh.HandShake, true);
        spCurrent.ReadTimeout = sh.SyncReadTimeOut;
        spCurrent.WriteTimeout = sh.SyncWriteTimeOut;
        spCurrent.Open();
        //============================================================================
         
        string str = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "TimeMeasure");
        this.timeMeasure = Convert.ToInt32(str);
        this.measureTimer = new System.Timers.Timer(timeMeasure);
        this.measureTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);

        str = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "PressureValue");
        this.pressureValue = Convert.ToByte(str);

        this.soundFile = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "SoundFile");        

        this.measureTimer.Start();
        
      }
      catch(Exception ex){
        this.IsError = true;
        DevExpress.Xpf.Core.DXMessageBox.Show(curUnit + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
      }

    }

    private Boolean InitPressure()
    {
      
      for (int i = 0; i < 7; i++){

        byte[] data1 = {0x55};

        spPressure.Write(data1, 0, 1);
        spPressure.Close();
        Thread.Sleep(50);
        spPressure.Parity = Parity.Odd;
        spPressure.Open();

        byte[] data2 = {0xAA, 0xAA};
        spPressure.Write(data2,0,2);
        Thread.Sleep(200);
        int rez = spPressure.Read(data2, 0, 2);
      
        if ((data2[0] == 0x55) && (data2[1] == 0x55)){
          this.IsError = false;
          return true;
        }
        else{
          this.IsError = true;
          spPressure.Close();
          Thread.Sleep(200);
          spPressure.Parity = Parity.Even;
          spPressure.Open();
        }
      }
      
      return false; 
    }

    private ComPortSectionHandler ReadParamValue(string ExeConfigFile, string Section)
    {
      ExeConfigurationFileMap filemap = new ExeConfigurationFileMap();
      filemap.ExeConfigFilename = ExeConfigFile;
      Configuration config = ConfigurationManager.OpenMappedExeConfiguration(filemap, ConfigurationUserLevel.None);

      /*create a PortSectionHandler object and use oConfiguration.GetSection method to get the targeted section*/
      try{
        ComPortSectionHandler oSection = config.GetSection(Section) as Viz.Lims.Sh.ComPortSectionHandler; 
        return oSection;
      } 
      //If there is not a corresponding section, then the exception is raised
      catch (NullReferenceException caught) {MessageBox.Show(caught.Message);}
      return null;  
    }

    private void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
    {
      
      lock (threadLock){
        
        if (!this.BeforeRead()){
          this.measureTimer.Stop();
          MessageBox.Show("Reinitial Pressure unit");
          this.spPressure.Close();
          this.spPressure.Parity = Parity.Even;
          this.spPressure.Open();
          this.InitPressure();
          this.measureTimer.Start();
          return;
        }
            
        byte[] data2 = {0x00, 0x00};
        Thread.Sleep(100);
        int rez = spPressure.Read(data2,0,2);

        this.CurrentStatePressure = (data2[0] >= this.pressureValue);

        if ((!this.OldStatePressure) && (this.CurrentStatePressure)){
          this.measureTimer.Stop();

          Thread.Sleep(200);
        
          byte[] barr = this.StringToASCIIByteArray("TA*");
          this.spCurrent.Write(barr, 0, barr.Length);
          Thread.Sleep(300);

          byte[] data64 = new byte[64];
          rez = spCurrent.Read(data64, 0, 64);
          string ValStr = this.ASCIIByteArrayToString(data64);
          ValStr = ValStr.Replace('.', CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]);
          decimal IsolVal = Convert.ToDecimal(ValStr);
          IsolVal = decimal.Round(IsolVal, 3);
          if (IsolVal <= 0)
            IsolVal = 0.001M; 

          Smv.Utils.Sound.PlaySoundFile(this.soundFile);
          this.OnMeasuredValue(IsolVal);

          this.OldStatePressure = this.CurrentStatePressure;   

          this.measureTimer.Start();
        }

        if ((this.OldStatePressure) && (!this.CurrentStatePressure))
          this.OldStatePressure = this.CurrentStatePressure;     
      }
    }
    
    private Boolean BeforeRead()
    {
      
      byte[] data1 = {0x00};
      
      spPressure.Write(data1,0,1);
      Thread.Sleep(100);
      int rez = spPressure.Read(data1, 0,1);

      if (data1[0] != 0xFF)
        return false;

      data1[0] = 0x00;
      spPressure.Write(data1,0,1);
      Thread.Sleep(100);
      rez = spPressure.Read(data1, 0,1);
      if (data1[0] != 0x00)
        return false;

      data1[0] = 0x02;
      spPressure.Write(data1,0,1);
      Thread.Sleep(100);
      rez = spPressure.Read(data1, 0, 1);
      if (data1[0] != 0x02)
        return false;

      data1[0] = 0x00;
      spPressure.Write(data1, 0, 1);
      Thread.Sleep(100);
      rez = spPressure.Read(data1, 0,1);
      if (data1[0] != 0x00)
        return false;

      data1[0] = 0x09;
      spPressure.Write(data1,0,1);
      Thread.Sleep(100);
      rez = spPressure.Read(data1, 0,1);
      if (data1[0] != 0x09)
        return false;
      
      return true;
    }

    public override void StartMeasure()
    {
      measureTimer.Start();
    }

    public override void StopMeasure()
    {
      lock (threadLock){

        if (measureTimer != null)  
          measureTimer.Stop();

        OldStatePressure = false;
      }
    }

    public override void Close()
    {

      this.StopMeasure();      

      if (spPressure != null){
        spPressure.Close();
        spPressure.Dispose();
        //MessageBox.Show("Порт давления освобожден!");
      }
   
      if (spCurrent != null){
        spCurrent.Close();
        spCurrent.Dispose();
        //MessageBox.Show("Порт измерителя тока освобожден!");
      }
    }


  }
}
