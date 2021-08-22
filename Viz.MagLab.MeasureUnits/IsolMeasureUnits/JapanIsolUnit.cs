using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Text;
using System.Timers;
using System.Windows;
using System.Windows.Threading;
using System.Threading;

namespace Viz.MagLab.MeasureUnits
{
  internal sealed class JapanIsolUnit : MeasureIsolUnit
  {
    private int timeMeasure = 0;
    private int timeBaseCycle = 0;
    private int channelA = 0;
    private int channelD = 0;
    private string connStr = null;
    private System.Timers.Timer measureTimer;
    private OleDbConnection dbCon = null;
    private OleDbCommand cmd1 = null;
    private OleDbCommand cmd2 = null;
    private object threadLock = new object(); 
    private Dispatcher dsp =  Dispatcher.CurrentDispatcher; 
    
    internal JapanIsolUnit(string ConfigFile, uint MeasureCount)
    {
      this.IsError = false;
      this.mCount = MeasureCount; 
      
      string str = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "TimeMeasure");
      this.timeMeasure = Convert.ToInt32(str);

      str = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "ChannelA");
      this.channelA = Convert.ToInt32(str);

      str = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "ChannelD");
      this.channelD = Convert.ToInt32(str);

      str = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "TimeBaseCycle");
      this.timeBaseCycle = Convert.ToInt32(str);

      this.soundFile = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(ConfigFile, "SoundFile");

      this.connStr = Smv.App.Config.ConfigParam.ReadConnectionStringParamValue(ConfigFile, "ConStrMeasureBase");
      this.dbCon = new OleDbConnection(connStr);
      this.cmd1 = dbCon.CreateCommand();
      this.cmd1.Connection = this.dbCon;
      this.cmd2 = dbCon.CreateCommand();
      this.cmd2.Connection = this.dbCon;

      this.cmd1.CommandType = CommandType.Text;
      this.cmd1.CommandText = "SELECT Value FROM dbo.Currents WHERE ID_Channel = ?";
      OleDbParameter prm = new OleDbParameter();
      prm.DbType = DbType.Int32;
      prm.Direction = ParameterDirection.Input;
      this.cmd1.Parameters.Add(prm);

      this.cmd2.CommandType = CommandType.Text;
      this.cmd2.CommandText = "SELECT (CAST(ROUND(ABS(Cur.Value),3) AS NUMERIC(5,3)))  " +
                              "FROM dbo.Currents Cur, dbo.Channels Chan " +
                              "WHERE (Cur.ID_Channel = Chan.ID_Channel) " +
                                   "AND (Cur.ID_Channel = ?) " +
                                   //"AND (Chan.ID_Channel = ?) " +
                                   "AND (Chan.ID_USPD = 1) " +
                                   "AND (Cur.State & Chan.ProtocolMask = 0)";

      prm = new OleDbParameter();
      prm.DbType = DbType.Int32;
      prm.Direction = ParameterDirection.Input;
      this.cmd2.Parameters.Add(prm);

      this.measureTimer = new System.Timers.Timer(timeMeasure);
      this.measureTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);

      try{
        this.dbCon.Open();
        measureTimer.Start();
      }
      catch (Exception ex){
        this.IsError = true; 
        DevExpress.Xpf.Core.DXMessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
      }
      finally{
        this.dbCon.Close();
      }
      
    }
    
    private void OnTimedEvent(object source, ElapsedEventArgs e)
    {
      lock (threadLock){
        (source as System.Timers.Timer).Stop();
        decimal IsolVal = TimerMeasure();

       if (IsolVal != -1){
        Smv.Utils.Sound.PlaySoundFile(this.soundFile);
        this.OnMeasuredValue(IsolVal);
       } 

        if (!this.IsError) 
          (source as System.Timers.Timer).Start();
      } 
    }

    private decimal TimerMeasure()
    {
      decimal aVal = -1;
      decimal MaxValue = Convert.ToDecimal(0.001);
      DateTime dt;

      if (this.Measure1() == 1)
      {
        dt = DateTime.Now.AddSeconds(this.timeBaseCycle);
        
        while (DateTime.Now <= dt){
          aVal = Measure2();

          if (aVal > MaxValue)
            MaxValue = aVal;

          if (this.Measure1() == 0)
            break;
        }
        aVal = MaxValue;
        //SET @DChannel = @ID_Channel_D
        //break;
      }
      return aVal;    
    }


    private int Measure1()
    {
      int rez = -1;

      try{
        if (this.cmd1.Connection.State != ConnectionState.Open)
          this.cmd1.Connection.Open();
        cmd1.Parameters[0].Value = channelD;
        rez = Convert.ToInt32(cmd1.ExecuteScalar());
      }
      catch (Exception ex){
        this.IsError = true;
        this.dsp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => {DevExpress.Xpf.Core.DXMessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);}));
        rez = -1;    
      }
      finally{
        this.dbCon.Close(); 
      }
      return rez;
    }

    private decimal Measure2()
    {
      decimal rez = -1;

      try{
        if (this.cmd2.Connection.State != ConnectionState.Open)
          this.cmd2.Connection.Open();

        cmd2.Parameters[0].Value = channelA;
        rez = Convert.ToDecimal(cmd2.ExecuteScalar());
      }
      catch (Exception ex){
        this.IsError = true;
        this.dsp.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => {DevExpress.Xpf.Core.DXMessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }));
        rez = -1;
      }
      finally{
        this.dbCon.Close(); 
      }
      return rez;
    }


    public override void StartMeasure()
    {
      this.measureTimer.Start();
    }

    public override void StopMeasure()
    {
      lock (threadLock){ 
        this.measureTimer.Stop();
      }
    }

    public override void Close()
    {
      lock (threadLock){
        measureTimer.Stop();
        measureTimer.Dispose();
        this.dbCon.Dispose();
      }
    }

  }
}
