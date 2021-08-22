/*Реализация работы с ПО измерительного уст-ва "Brockhaus Messtechnik" MPG200D */

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Smv.Network.Tcp;

namespace Viz.MagLab.MeasureUnits
{
  
  public sealed class BrockhausMpg200D : TcpClientAsync
  {
    #region Consts
    private const byte ProtocolTypeReq   = 1;
    private const byte ProtocolTypeRcv   = 2;
    private const int RequestMeasureSize = 95;
    private const int ReceiveMeasureSize = 153;
    private const int StringParamLength  = 32;
    private const int NumOfMeasurements  = 5;
    #endregion
    
    #region Public Struct & Enum
    public enum SampleTypeMeasure
    {
      Sheet = 1,
      RingCore = 2,
      Epstein = 4
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, Size = RequestMeasureSize, CharSet = CharSet.Ansi)]
    public struct RequestMeasure
    {
      public Byte ProtocolType;
      public Int32 DataLength;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = StringParamLength)]
      public byte[] SampleName;
      public Int16 SampleType;
      public float Weight;
      public float Density;
      public float LengthOrDA;
      public float WidthOrDI;
      public float NominalThickness;
      public float Quantity;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = StringParamLength)]
      public byte[] CoilSystemName;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, Size = ReceiveMeasureSize, CharSet = CharSet.Ansi)]
    public struct ReceiveMeasure
    {
      public Byte  ProtocolType;
      public Int32 DataLength;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = StringParamLength)]
      public byte[] SampleName;
      public float Frequency;
      public float Jn;
      public float Jmax;
      public float Jr;
      public float Jeff;
      public float Hn;
      public float Hmax;
      public float Hc;
      public float Heff;
      public float Ps;
      public float Ss;
      public float µr;
      public float FF;
      public float DC;
      public float Korr;
      public float Ph;
      public float Pw;
      public float Imax;
      public float Ieff;
      public float Umax;
      public float Ueff;
      [MarshalAs(UnmanagedType.ByValArray, SizeConst = StringParamLength)]
      public byte[] State;
    }
    public struct MeasurementResult
    {
      public decimal B100;
      public decimal B800;
      public decimal B2500;
      public decimal P1550;
      public decimal P1750;
      public string State1;
      public string State2;
      public string State3;
      public string State4;
      public string State5;
      public int CodeError;
    }
    #endregion
    
    #region Private Method
    private MeasurementResult ProcessMeasRes(List<ReceiveMeasure> rcvList, List<byte[]> rcvBinDataList)
    {
      var me = new MeasurementResult();

      for (int i = 0; i < NumOfMeasurements; i++)
      {
        rcvList[i] = (ReceiveMeasure) ReadStruct(rcvBinDataList[i], typeof(ReceiveMeasure));

        switch (rcvList[i].Jn)
        {
          case 1.5f:
            me.P1550 = Convert.ToDecimal(rcvList[i].Ps);
            break;
          case 1.7f:
            me.P1750 = Convert.ToDecimal(rcvList[i].Ps);
            break;
        }

        switch (rcvList[i].Hn)
        {
          case 100f:
            me.B100 = Convert.ToDecimal(rcvList[i].Jmax);
            break;
          case 800f:
            me.B800 = Convert.ToDecimal(rcvList[i].Jmax);
            break;
          case 2500f:
            me.B2500 = Convert.ToDecimal(rcvList[i].Jmax);
            break;
        }

        switch (i)
        {
          case 0:
            me.State1 = Encoding.ASCII.GetString(rcvList[i].State, 0, StringParamLength).Trim();
            break;
          case 1:
            me.State2 = Encoding.ASCII.GetString(rcvList[i].State, 0, StringParamLength).Trim();
            break;
          case 2:
            me.State3 = Encoding.ASCII.GetString(rcvList[i].State, 0, StringParamLength).Trim(); 
            break;
          case 3:
            me.State4 = Encoding.ASCII.GetString(rcvList[i].State, 0, StringParamLength).Trim(); 
            break;
          case 4:
            me.State5 = Encoding.ASCII.GetString(rcvList[i].State, 0, StringParamLength).Trim(); 
            break;
        }


      }

      return me;
    }
    #endregion

    #region Public Method
    public async Task<MeasurementResult> RunMeasurement(string sampleName, SampleTypeMeasure sampleType, decimal weight, decimal density, int length, int width, decimal thickness, int quantity, string coilSystemName)
    {
      ClearError();

      if (!IsConnected()){
        msgInfo?.ShowDlgErrorInfo("Нет соединения", "Соединение не установлено. Отправка и прием данных не возможен!");
        return new MeasurementResult
                   {
                     CodeError = -2,
                     State1 = "Нет соединения. Отправка и прием данных не возможен!",
                     State2 = "Нет соединения. Отправка и прием данных не возможен!",
                     State3 = "Нет соединения. Отправка и прием данных не возможен!",
                     State4 = "Нет соединения. Отправка и прием данных не возможен!",
                     State5 = "Нет соединения. Отправка и прием данных не возможен!"
                   };
      }

      var req = new RequestMeasure
      {
        ProtocolType = ProtocolTypeReq,
        SampleName = new byte[StringParamLength],
        SampleType = (Int16)sampleType,
        Weight = Convert.ToSingle(weight),
        Density = Convert.ToSingle(density),
        LengthOrDA = Convert.ToSingle(length),
        WidthOrDI = Convert.ToSingle(width),
        NominalThickness = Convert.ToSingle(thickness),
        Quantity = Convert.ToSingle(quantity),
        CoilSystemName = new byte[StringParamLength]
      };

      req.DataLength = Marshal.SizeOf(req); //должно быть 95
      Array.Copy(Encoding.ASCII.GetBytes(sampleName), req.SampleName,  Encoding.ASCII.GetBytes(sampleName).Length);
      Array.Copy(Encoding.ASCII.GetBytes(coilSystemName), req.CoilSystemName, Encoding.ASCII.GetBytes(coilSystemName).Length);
      byte[] byteDataSndReq = RawSerialize(req);

      var rcvList = new List<ReceiveMeasure>();
      var rcvBinDataList = new List<byte[]>();
      var rcvListTaskRes = new List<Task<int>>();

      for (int i = 0; i < NumOfMeasurements; i++)
      {
        rcvList.Add(new ReceiveMeasure
                    {
                      SampleName = new byte[StringParamLength],
                      State = new byte[StringParamLength]
                    }
                   );

        rcvBinDataList.Add(new byte[Marshal.SizeOf(rcvList[0])]); //должно быть 153
      }
      
      var taskSnd = SendDataAsync(byteDataSndReq);
      await taskSnd;

      for (int i = 0; i < NumOfMeasurements; i++)
      {
        var taskRcv = ReceiveDataAsync(rcvBinDataList[i]);
        rcvListTaskRes.Add(taskRcv);
        await taskRcv;
      }

      /********************Обработка принятых данных***************************/
      for (int i = 0; i < NumOfMeasurements; i++)
        if (rcvListTaskRes[i].Result != Marshal.SizeOf(rcvList[0]))
          return new MeasurementResult
                 {
                  CodeError = -3,
                  State1 = LastError,
                  State2 = "Размер принятого пакета не верен!"
                 };


      return ProcessMeasRes(rcvList, rcvBinDataList);
    }
    #endregion

    #region Constructor
    public BrockhausMpg200D(string host, int port, int readTimeout = 0, Smv.Utils.MessageInfo msgInfo = null, int connectTimeout = 0) 
           : base(host, port, readTimeout, msgInfo, connectTimeout)
    {
    }
    #endregion

  }
}
