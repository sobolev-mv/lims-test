using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Smv.Utils;

namespace Smv.Network.Tcp
{
  public class TcpClientAsync
  {
    #region Fields

    private TcpClient tcpClient = null;
    private string lastError = String.Empty;
    protected IMessageInfo msgInfo = null;
    #endregion

    #region Public Property

    public string Host { get; }
    public int Port { get; }
    public int ReadTimeout { get; }
    public int ConnectTimeout { get; }
    public string LastError => lastError;

    #endregion

    #region Private Method

    #endregion

    #region Public Method
    public static byte[] RawSerialize(object anything)
    {
      var rawSize = Marshal.SizeOf(anything);
      var rawData = new byte[rawSize];
      GCHandle handle = GCHandle.Alloc(rawData, GCHandleType.Pinned);
      Marshal.StructureToPtr(anything, handle.AddrOfPinnedObject(), false);
      handle.Free();
      return rawData;
    }

    public static object ReadStruct(byte[] data, Type t)
    {
      byte[] buffer = new byte[Marshal.SizeOf(t)];
      Array.Copy(data, 0, buffer, 0, Marshal.SizeOf(t));
      GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
      Object temp = Marshal.PtrToStructure(handle.AddrOfPinnedObject(), t);
      handle.Free();
      return temp;
    }

    public Boolean IsConnected()
    {
      return tcpClient.Connected;
    }

    public void ClearError()
    {
      lastError = String.Empty;
    }


    public Boolean Connect()
    {
      try
      {

        if (tcpClient == null)
          tcpClient = new TcpClient();

        if (tcpClient.Connected)
        {
          ClearError();
          return true;
        }

        tcpClient.Connect(Host, Port);
        tcpClient.ReceiveTimeout = this.ReadTimeout;
        //tcpClient.GetStream().ReadTimeout = this.ReadTimeout;
        ClearError();
        return true;
      }
      catch (Exception ex)
      {
        msgInfo?.ShowDlgErrorInfo("Ошибка соединения", ex.Message);
        lastError = ex.Message;
        return false;
      }
    }

    /*    
        public Boolean Connect()
        {
          if (tcpClient == null)
            tcpClient = new TcpClient();

          if (tcpClient.Connected){
            ClearError();
            return true;
          }

          CancellationToken ct = new CancellationToken(); // Required for "*.Task()" method
          if (tcpClient.ConnectAsync(Host, Port).Wait(ConnectTimeout, ct)) // Connect with timeout as 1 second
          {
            ct.ThrowIfCancellationRequested();
            ClearError();
            return true;
          }
          else
          {
            msgInfo?.ShowDlgErrorInfo("Ошибка соединения", "Превышено время ожидания!");
            lastError = "Превышено время ожидания!";
            return false;
          }
        }
    */
    public void Close()
    {
      if (tcpClient == null)
        return;

      if (!tcpClient.Connected)
        return;

      tcpClient.GetStream().Close();
      tcpClient.Close();
      tcpClient.Dispose();
      tcpClient = null;
      ClearError();
    }

    public async Task SendDataAsync(byte[] data)
    {
      await tcpClient.GetStream().WriteAsync(data, 0, data.Length);
    }

    public async Task<int> ReceiveDataAsync(byte[] data)
    {
      int bytesReceived = 0;

      await Task.Run(() =>
        {
          try{
            bytesReceived = tcpClient.GetStream().Read(data, 0, data.Length);
          }
          catch (Exception ex){
            msgInfo?.ShowDlgErrorInfo("Ошибка приема данных", ex.HResult.ToString(CultureInfo.InvariantCulture) + " " + ex.Message);
            lastError = ex.Message;
            bytesReceived = -1;
          }
        }
      );

      return bytesReceived;

    }


    #endregion

    #region Constructor

    public TcpClientAsync(string host, int port, int readTimeout = 0, IMessageInfo msgInfo = null, int connectTimeout = 0)
    {
      Host = host;
      Port = port;
      ReadTimeout = readTimeout;
      ConnectTimeout = connectTimeout;
      this.msgInfo = msgInfo;
    }
    #endregion


  }
}