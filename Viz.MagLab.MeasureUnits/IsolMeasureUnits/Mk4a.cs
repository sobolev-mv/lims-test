using System;
using System.Data;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Sockets;
using System.Windows;
using Smv.Network.Tcp;
using DevExpress.Xpf.Editors;

namespace Viz.MagLab.MeasureUnits
{
  public sealed class Mk4a : AsyncTcpSockClient
  {

    #region Error Device
    //Ошибки при работе.
    //0x001 - передача калибровочных параметров в устройство
    //0x002 - Амплитуда Test_ОС
    //0x003 - Установка флагов
    //0x004 - Запуск  Test_ОС
    //0x005 - вышел TIMEOUT Test_ОС
    //0x006 - получения результата
    //0x007 - Нет регистрации тока (слабый сигнал, обрыв провода)
    //0x008 - Нет регистрации индукции
    //0x009 - внутренняя ошибка контроллера
    //0x011 - Передачи задания на измерение
    //0x012 - запуска измерения
    //0x013 - вышел TIMEOUT результата
    //0x014 - Считывания результата
    //0x015 - Проверка корректности результата1
    //0x016 - Проверка корректности результата2
    //$1016 -  Признак срабатывания аппаратной защиты.
    //$2016 -  За отведенное количество итераций не найдены необходимые пределы измерения.
    //$4016 -  Прекращение измерения по внешнему флагу
    //$0z16 -  z- количество найденных ошибок во входном задании
    //0x017 - Вычисления
    #endregion

    #region Constants
    //статусы сервера
    private const uint CS_Zero        = 0x00; // - не определено
    private const uint CS_NSet        = 0x01; // - нет устройства
    private const uint CS_YSet        = 0x12; //- есть устройство
    private const uint CS_GotovNdat   = 0x13; // - готов к работе, нет измеренных данных
    private const uint CS_AnalizIn    = 0x04; // - анализ входных данных
    private const uint CS_ErrIn       = 0x15; // - входные задания не правильны, готов к новому заданию
    private const uint CS_Work        = 0x06; // - выполнение работы
    private const uint CS_YRezGotov   = 0x17; // - есть результат, готов к новому заданию
    private const uint CS_ErrWork     = 0x18; // - Ошибка выполнения, готов к новому заданию
    #endregion

    #region Structure
    [StructLayout(LayoutKind.Explicit, Pack = 1, Size = 8)]
    public struct Izm1
    {
      [FieldOffset(0)]
      public Single Zn;//измерено, отклонение по серии измерений 1/(N*(N-1))
      [FieldOffset(4)]
      public Single Si;//измерено, отклонение по серии измерений 1/(N*(N-1))
    }

    [StructLayout(LayoutKind.Explicit, Pack = 1, Size = 8, CharSet = CharSet.Ansi)]
    public struct NtZakaz1
    {
      [FieldOffset(0)]
      public UInt32 TypIzm;
      [FieldOffset(4)]
      public Single Zzdn;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, Size = 188, CharSet = CharSet.Ansi)]
    public struct NtZakazArr
    {
      public UInt32 TypApp;
      public Single Obr_shirina;
      public Single Obr_Dlina;
      public Single Obr_massa;
      public Single Obr_plotn;
      public UInt32 C;
      public UInt32 NZ;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 160)]
      public NtZakaz1[] Z;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, Size = 192)]
    public struct C2SBuf
    {
      public UInt32 T;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 188)]
      public Byte[] Data;
      //case
      //1: (ZA: T_NtZakazArr) формируем заказ
      //2: (qw: LongWord )    Опрос состояния
      //3: (GR: LongWord)     Вернуть результат.
      //4: (SA: LongWord )    Вернуть измерители
      //5  (Co: LongWord )    Команда
    }

    //Для результатов измерений
    [StructLayout(LayoutKind.Explicit, Pack = 1, Size = 36)]
    public struct NtRez1
    {
      [FieldOffset(0)]
      public UInt32 TypIzm;  //Тип измерения : (поле или индукция)
      [FieldOffset(4)]
      public Izm1 B;
      [FieldOffset(12)]
      public Izm1 H;
      [FieldOffset(20)]
      public Izm1 S;
      [FieldOffset(28)]
      public Izm1 P;  //Измеренные значения
    }


    [StructLayout(LayoutKind.Explicit, Pack = 1, Size = 8, CharSet = CharSet.Ansi)]
    public struct NtStErr
    {
      [FieldOffset(0)]
      public UInt32 Stat;
      [FieldOffset(4)]
      public UInt32 Err;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, Size=728)]
    public struct NtRezArr
    {
      public UInt32 C;   //
      public UInt32 NZ;  //
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 720)]
      public Byte[] RA; //NtRez1[]
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, Size = 304)]
    public struct NtAppArr
    {
      public UInt32 C;   //
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 300)]
      public Byte[] SA;
    }


    [StructLayout(LayoutKind.Sequential, Pack = 1, Size = 732)]
    public struct S2CBuf
    {
      public UInt32 T;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = 728)]
      public Byte[] Data;
      //case
      //2: (An : NtStErr );  //Возвращаем состояние
      //3: (RT : NtRezArr ); //Возвращаем результат
      //4: (OI : NtAppArr) ; //Описание измерителей
    }
    #endregion 

    #region Fields
    private System.Timers.Timer mTimer;
    private ProgressBarEdit pgb;
    private string host;
    private int port;
    private DataTable data;
    private DataView dvFlt;
    private uint nDivice;
    #endregion

    #region Public Property
    public event EventHandler<EventArgs> MeasureComplete;
    #endregion    

    #region Constructors
    public Mk4a(int RxBuffer, int uType, System.Windows.Threading.Dispatcher dsp, ProgressBarEdit pgb, string Host, int Port, DataTable dtData) : base(RxBuffer, dsp)
    {
      this.pgb = pgb;
      this.host = Host;
      this.port = Port; 
      this.data = dtData;
      this.Connected += ConnectedEventHandler;
      this.SendedData += SendedDataEventHandler;
      this.ReceivedData += ReceivedDataEventHandler;
      this.Error += ErrorEventHandler;
      this.mTimer = new System.Timers.Timer(2000);
      this.mTimer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);

      //Создаем фильтр только для выбранных измерений
      this.dvFlt = new DataView(this.data) { RowFilter = "IsActive <> 0" }; 
      this.nDivice = Convert.ToUInt32((uType == 1) ? 2 : 3); 
    }
    #endregion

    #region Methods

    private void OnMeasureComplete()
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<EventArgs> temp = this.MeasureComplete;

      if (temp != null){
         temp(this, new EventArgs());
      }
    }

    private void ConnectedEventHandler(Object sender, EventArgs e)
    {
      
    }

    private void ErrorEventHandler(Object sender, EventArgs e)
    {
      this.pgb.StyleSettings = new ProgressBarStyleSettings();
    }


    private void SendedDataEventHandler(Object sender, SendedDataEventArgs e)
    {
      
    }

    private void ReceivedDataEventHandler(Object sender, ReceivedDataEventArgs e)
    {
      var st = new S2CBuf(); 
      var ntStErr =  new NtStErr();
      var ntRezArr = new NtRezArr();
      var ntAppArr = new NtAppArr();

      if (e.ReceivedBytes != Marshal.SizeOf(st)){
        this.mTimer.Stop();
        Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "Количество принятых данных не верно!", MessageBoxImage.Stop);
        return;
      }

      st = (S2CBuf)ReadStruct(e.Data, typeof(S2CBuf));
      switch(st.T){
        case 2:
          ntStErr = (NtStErr)ReadStruct(st.Data, typeof(NtStErr)); 
          
          switch(ntStErr.Stat & 0x00ff){
            case CS_Zero:
              this.mTimer.Stop();
              this.pgb.StyleSettings = new ProgressBarStyleSettings();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "устройство не определено!", MessageBoxImage.Stop);
              
              break;
            case CS_NSet:
              this.mTimer.Stop();
              this.pgb.StyleSettings = new ProgressBarStyleSettings();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "нет устройства!", MessageBoxImage.Stop);
              
              break;
            case CS_YSet:
              this.mTimer.Stop();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Инфо", "есть устройства!", MessageBoxImage.Stop);
              
              break;
            case CS_GotovNdat:
              this.mTimer.Stop();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Инфо", "готов к работе, нет измеренных данных!", MessageBoxImage.Stop);
              break;
            case CS_AnalizIn:
              //Smv.Utils.DbErrorInfo.ShowDbErorInfo("Инфо", "анализ входных данных!");
              break;
            case CS_ErrIn:
              this.mTimer.Stop();
              this.pgb.StyleSettings = new ProgressBarStyleSettings();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "входные задания не правильны, готов к новому заданию!", MessageBoxImage.Stop);
              break;
            case CS_Work:
              //Smv.Utils.DbErrorInfo.ShowDbErorInfo("Инфо", "выполнение работы!");
              break;
            case CS_YRezGotov:
              this.mTimer.Stop();
              this.SendQweMeasureData();
              //Smv.Utils.DbErrorInfo.ShowDbErorInfo("Инфо", "есть результат, готов к новому заданию!");
              break;
            case CS_ErrWork:
              this.mTimer.Stop();
              this.pgb.StyleSettings = new ProgressBarStyleSettings();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "Ошибка выполнения, готов к новому заданию!", MessageBoxImage.Stop);
              break;
            default:
              this.mTimer.Stop();
              this.pgb.StyleSettings = new ProgressBarStyleSettings();
              Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "Не нормативный ответ!", MessageBoxImage.Stop);
              break;
          }

          if (ntStErr.Err != 0) {
            this.pgb.StyleSettings = new ProgressBarStyleSettings();
            Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", "Ошибка выполнения: " + ntStErr.Err.ToString(), MessageBoxImage.Stop); 
          }
          break;
        case 3:
          this.mTimer.Stop();
          this.pgb.StyleSettings = new ProgressBarStyleSettings(); 
          ntRezArr = (NtRezArr)ReadStruct(st.Data, typeof(NtRezArr));
          this.SetMeasureResult(ntRezArr); 
          OnMeasureComplete();
          //Smv.Utils.DbErrorInfo.ShowDbErorInfo("Инфо", this.FormatMeasure(ntRezArr));    
          break;
        case 4:
          this.mTimer.Stop();
          ntAppArr = (NtAppArr)ReadStruct(st.Data, typeof(NtAppArr));
          this.FormatMeasureDevices(ntAppArr);
          /*
          for i:=1 to Bin.OI.C do
             Memo1.Lines.Add( Bin.OI.SA[i]);
          end;
          */ 
          break;
		    default:
          this.mTimer.Stop();
          break;
	    } 

    }

    private void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
    {   
      IAsyncResult ias;
      ias = SendQweStatus();
      ias.AsyncWaitHandle.WaitOne();
    }


    private string FormatMeasure(NtRezArr ntRezArr)
    {
      String rez = null;
      NtRez1 ntRez1 = new NtRez1();
      Byte[] tbuffer = new Byte[Marshal.SizeOf(typeof(NtRez1))];

      for (int i = 0; i < ntRezArr.C; i++){
        Array.Copy(ntRezArr.RA, i * Marshal.SizeOf(ntRez1), tbuffer, 0, Marshal.SizeOf(ntRez1));
        ntRez1 = (NtRez1)ReadStruct(tbuffer, typeof(NtRez1));

        if ((ntRez1.TypIzm & 0x0f) == 1)
          rez += "Задание Индукции" + "\n";

        if ((ntRez1.TypIzm & 0x0f) == 2)
          rez += "Задание Поля" + "\n";
      
        rez += "B = " + ntRez1.B.Zn.ToString("n") + " ±" + ntRez1.B.Si.ToString("n") + " Тл\n";
        rez += "H = " + ntRez1.H.Zn.ToString("n") + " ±" + ntRez1.H.Si.ToString("n") + " А/м\n";
        rez += "S = " + ntRez1.S.Zn.ToString("n") + " ±" + ntRez1.S.Si.ToString("n") + " Дж\n";
        rez += "P = " + ntRez1.P.Zn.ToString("n") + " ±" + ntRez1.P.Si.ToString("n") + " Вт/кг\n";
      }

      return rez;  
    }

    private void SetMeasureResult(NtRezArr ntRezArr)
    {
      NtRez1 ntRez1 = new NtRez1();
      Byte[] tbuffer = new Byte[Marshal.SizeOf(typeof(NtRez1))];

      for (int i = 0; i < ntRezArr.C; i++){

        Array.Copy(ntRezArr.RA, i * Marshal.SizeOf(ntRez1), tbuffer, 0, Marshal.SizeOf(ntRez1));
        ntRez1 = (NtRez1)ReadStruct(tbuffer, typeof(NtRez1));

        DataRow row = this.dvFlt[i].Row; //this.data.Rows[i];
        row.BeginEdit();
 
        if ((ntRez1.TypIzm & 0x0f) == 1)
          //rez += "Задание Индукции" + "\n";
          row["OutVal"] = /*Convert.ToDouble(row["Corr"]) + */  Math.Round(ntRez1.P.Zn, 2);

        if ((ntRez1.TypIzm & 0x0f) == 2)
         //rez += "Задание Поля" + "\n";
          row["OutVal"] =  /*Convert.ToDouble(row["Corr"]) + */ Math.Round(ntRez1.B.Zn, 2);         

        row.EndEdit();  
        
        /*
        rez += "B = " + ntRez1.B.Zn.ToString("n") + " ±" + ntRez1.B.Si.ToString("n") + " Тл\n";
        rez += "H = " + ntRez1.H.Zn.ToString("n") + " ±" + ntRez1.H.Si.ToString("n") + " А/м\n";
        rez += "S = " + ntRez1.S.Zn.ToString("n") + " ±" + ntRez1.S.Si.ToString("n") + " Дж\n";
        rez += "P = " + ntRez1.P.Zn.ToString("n") + " ±" + ntRez1.P.Si.ToString("n") + " Вт/кг\n";
        */ 
      }
     
    }


    private string FormatMeasureDevices(NtAppArr ntAppArr)
    {
      //Byte[] tbuffer = new Byte[30];

      for (int i = 0; i < ntAppArr.C; i++){
        Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", Encoding.ASCII.GetString(ntAppArr.SA, i * 30, 30), MessageBoxImage.Stop);
      } 

      return null;
    }

    public IAsyncResult SendQweMeasure(UInt32 MeasDevice, decimal? Mass, decimal? sLen, decimal? sWid, decimal? sDens)
    {
      C2SBuf c2sBuf = new C2SBuf();
      c2sBuf.T = 1;
      NtZakazArr ntZakazArr = new NtZakazArr();
      
      /*
      ntZakazArr.TypApp = 2;
      ntZakazArr.Obr_shirina = 0.30f;
      ntZakazArr.Obr_Dlina = 0.9f;
      ntZakazArr.Obr_massa = 0.257f;
      ntZakazArr.Obr_plotn = 7650.0f;
      ntZakazArr.C = 3;
      ntZakazArr.NZ = 1002;

      ntZakazArr.Z = new NtZakaz1[160];
      ntZakazArr.Z[0].TypIzm = 2;  // H
      ntZakazArr.Z[0].Zzdn = 80.0f;
      ntZakazArr.Z[1].TypIzm = 1;  // B
      ntZakazArr.Z[1].Zzdn = 0.75f;
      ntZakazArr.Z[2].TypIzm = 1;  // B
      ntZakazArr.Z[2].Zzdn = 1.0f;
      */

      int countIzm = dvFlt.Count;

      ntZakazArr.TypApp = MeasDevice;
      ntZakazArr.Obr_shirina = Convert.ToSingle(sWid / 1000);
      ntZakazArr.Obr_Dlina = Convert.ToSingle(sLen / 1000 );
      ntZakazArr.Obr_massa = Convert.ToSingle(Mass / 1000);
      ntZakazArr.Obr_plotn = Convert.ToSingle(sDens);
      ntZakazArr.C = Convert.ToUInt32(countIzm);
      ntZakazArr.NZ = 1002;
      ntZakazArr.Z = new NtZakaz1[160];

      for (int i = 0; i < countIzm; i++){
        ntZakazArr.Z[i].TypIzm = Convert.ToUInt32(dvFlt[i].Row["TypIzm"]);
        ntZakazArr.Z[i].Zzdn = Convert.ToSingle(dvFlt[i].Row["ValIzm"]);     
      }

      byte[] byteDatantZakaz = RawSerialize(ntZakazArr);
      c2sBuf.Data = byteDatantZakaz;
      byte[] byteData = RawSerialize(c2sBuf);
      return this.SendData(byteData);
    }

    public IAsyncResult SendQweStatus()
    {
      this.BeginReceiveData(1000);
      //System.Threading.Thread.Sleep(200);
      C2SBuf c2sBuf = new Mk4a.C2SBuf();
      c2sBuf.T = 2;
      //MessageBox.Show(Marshal.SizeOf(c2sBuf).ToString());
      byte[] byteData = Mk4a.RawSerialize(c2sBuf);
      return this.SendData(byteData);
    }

    public void SendQweMesuareDevice()
    {
      this.BeginReceiveData(1000);
      C2SBuf c2sBuf = new Mk4a.C2SBuf();
      c2sBuf.T = 4;
      byte[] byteData = Mk4a.RawSerialize(c2sBuf);
      this.SendData(byteData);
    }

    public void SendQweMeasureData()
    {
      this.BeginReceiveData(1000);
      //System.Threading.Thread.Sleep(200);
      C2SBuf c2sBuf = new Mk4a.C2SBuf();
      c2sBuf.T = 3;
      //MessageBox.Show(Marshal.SizeOf(c2sBuf).ToString());
      byte[] byteData = Mk4a.RawSerialize(c2sBuf);
      this.SendData(byteData);
    }

    public void StartMeasure(decimal? Mass, decimal? sLen, decimal? sWid, decimal? sDens)
    {
      IAsyncResult iasCon;
      this.pgb.StyleSettings = new ProgressBarMarqueeStyleSettings();
      this.Close();
      iasCon = this.Connect(this.host, this.port);
      iasCon.AsyncWaitHandle.WaitOne();

      if (!this.IsConnected) 
        return;

      iasCon = this.SendQweMeasure(this.nDivice, Mass, sLen, sWid, sDens);
      iasCon.AsyncWaitHandle.WaitOne();
      this.mTimer.Start();
    }
    #endregion


  }
}
