using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows;

namespace Smv.Network.Tcp
{
  
  public class SendedDataEventArgs : EventArgs
  {
    private int sendBytes;

    public SendedDataEventArgs(int sndBytes)
    {
      this.sendBytes = sndBytes;
    }

    public int SendBytes
    {
      get { return sendBytes; }
      set { sendBytes = value; }
    }

  }

  public class ReceivedDataEventArgs : EventArgs
  {
    private int rcvBytes;
    private byte[] data;

    public ReceivedDataEventArgs(int rcvBytes, byte[] data)
    {
      this.rcvBytes = rcvBytes;
      this.data = data;
    }

    public int ReceivedBytes
    {
      get { return rcvBytes; }
      set { rcvBytes = value; }
    }

    public byte[] Data
    {
      get { return this.data; }
      set { this.data = value; }
    }

  }

  internal sealed class StateObject
  {
    //Receive buffer.
    public byte[] Buffer = null;
  }

  public class AsyncTcpSockClient
  {
    protected System.Windows.Threading.Dispatcher dsp;
    private Socket sok = null;
    protected byte[] rxBuffer = null;
    protected int rxPosition = 0;
    public  int RxBufferLength {get{return rxBuffer.Length;}}
    public Boolean IsConnected { get { return (this.sok != null) && (this.sok.Connected); } }     

    public event EventHandler<EventArgs> Connected;
    public event EventHandler<EventArgs> Error;
    public event EventHandler<SendedDataEventArgs> SendedData;
    public event EventHandler<ReceivedDataEventArgs> ReceivedData;

    protected static bool SetKeepAlive(Socket sock, ulong time, ulong interval)
    {
      const int bytesperlong = 4; // 32 / 8
      const int bitsperbyte = 8;

      try{
        // resulting structure
        byte[] SIO_KEEPALIVE_VALS = new byte[3 * bytesperlong];

        // array to hold input values
        ulong[] input = new ulong[3];

        // put input arguments in input array
        if (time == 0 || interval == 0) // enable disable keep-alive
          input[0] = (0UL); // off
        else
          input[0] = (1UL); // on

        input[1] = (time); // time millis
        input[2] = (interval); // interval millis

        // pack input into byte struct
        for (int i = 0; i < input.Length; i++){
          SIO_KEEPALIVE_VALS[i * bytesperlong + 3] = (byte)(input[i] >> ((bytesperlong - 1) * bitsperbyte) & 0xff);
          SIO_KEEPALIVE_VALS[i * bytesperlong + 2] = (byte)(input[i] >> ((bytesperlong - 2) * bitsperbyte) & 0xff);
          SIO_KEEPALIVE_VALS[i * bytesperlong + 1] = (byte)(input[i] >> ((bytesperlong - 3) * bitsperbyte) & 0xff);
          SIO_KEEPALIVE_VALS[i * bytesperlong + 0] = (byte)(input[i] >> ((bytesperlong - 4) * bitsperbyte) & 0xff);
        }
        // create bytestruct for result (bytes pending on server socket)
        byte[] result = BitConverter.GetBytes(0);
        // write SIO_VALS to Socket IOControl
        sock.IOControl(IOControlCode.KeepAliveValues, SIO_KEEPALIVE_VALS, result);
      }
      catch (Exception){
        return false;
      }
      return true;
    }

    public static byte[] RawSerialize(object anything)
    { 
      int rawsize = Marshal.SizeOf(anything); 
      byte[] rawdata = new byte[rawsize]; 
      GCHandle handle = GCHandle.Alloc(rawdata, GCHandleType.Pinned); 
      Marshal.StructureToPtr(anything, handle.AddrOfPinnedObject(), false); 
      handle.Free(); 
      return rawdata; 
    }

    public static object ReadStruct(Byte[] Data, Type t) 
    { 
      byte[] buffer = new byte[Marshal.SizeOf(t)];
      Array.Copy(Data, 0, buffer, 0, Marshal.SizeOf(t));
      GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned); 
      Object temp = Marshal.PtrToStructure(handle.AddrOfPinnedObject(), t); 
      handle.Free(); 
      return temp; 
    }

    public AsyncTcpSockClient(int RxBufferLength, System.Windows.Threading.Dispatcher dsp)
    {
      this.rxBuffer = new byte[RxBufferLength];
      this.dsp = dsp; 
    }

    public IAsyncResult Connect(string HostName, int Port)
    {
      //Create a TCP/IP socket.
      sok = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
      //Create EndPoint  
      EndPoint remoteEP = new DnsEndPoint(HostName, Port); 
      //Connect to the remote endpoint.
      return sok.BeginConnect(remoteEP, new AsyncCallback(ConnectCallback), sok);
    }
    
    public void Disconnect()
    {
      if (sok != null)
        sok.Disconnect(true);
    }

    public void Close()
    {
      if (sok != null){
        if (sok.Connected)
          sok.Shutdown(SocketShutdown.Both);
        sok.Close(); 
      }
    }

    private void ConnectCallback(IAsyncResult ar)
    {
      try{
        //Retrieve the socket from the state object.
        Socket s = (Socket)ar.AsyncState;

        //Complete the connection.
        s.EndConnect(ar);
        this.dsp.Invoke(System.Windows.Threading.DispatcherPriority.Normal, (ThreadStart)(() => { OnConnected(); }));
      }
      catch (Exception e){
        this.dsp.Invoke(System.Windows.Threading.DispatcherPriority.Normal, (ThreadStart)(() => { OnError(); }));
        this.dsp.Invoke(System.Windows.Threading.DispatcherPriority.Normal, (ThreadStart)(() => Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка соединения", e.ToString(), MessageBoxImage.Stop)));
      }
    }

    private void OnConnected()
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<EventArgs> temp = this.Connected;

      if (temp != null)
        temp(this, new EventArgs());
    }

    private void OnError()
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<EventArgs> temp = this.Error;

      if (temp != null)
        temp(this, new EventArgs());
    }


    public IAsyncResult SendData(byte[] Data)
    {
      if (sok.Connected) 
        //Begin sending the data to the remote device.
        return sok.BeginSend(Data, 0, Data.Length, SocketFlags.None, new AsyncCallback(SendCallback), sok);
      else
        return null;
    }

    private void SendCallback(IAsyncResult ar)
    {
      try{
        //Retrieve the socket from the state object.
        Socket s = (Socket)ar.AsyncState;

        // Complete sending the data to the remote device.
        int bytesSent = s.EndSend(ar);
        OnSendedData(bytesSent);
      }
      catch (Exception e){
        Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.ToString(), MessageBoxImage.Stop);
      }
    }

    private void OnSendedData(int sndBytes)
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<SendedDataEventArgs> temp = this.SendedData;

      if (temp != null)
        temp(this, new SendedDataEventArgs(sndBytes));
    }

    public void BeginReceiveData(int bufferSize)
    {
      if (!sok.Connected) return;

      this.rxPosition = 0;
      Array.Clear(this.rxBuffer, 0, this.rxBuffer.Length);      

      try{
        StateObject state = new StateObject();
        state.Buffer = new byte[bufferSize];
                 
        //Begin receiving the data from the remote device.
        sok.BeginReceive(state.Buffer, 0, bufferSize, SocketFlags.None, new AsyncCallback(ReceiveCallback), state);
      }
      catch (Exception e){
        Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.ToString(), MessageBoxImage.Stop);
      }
    }

    private void ReceiveCallback(IAsyncResult ar)
    {
      try{
        //Retrieve the state object and the client socket 
        //from the asynchronous state object.
        byte[] buff = (ar.AsyncState as StateObject).Buffer;
                        
        //Read data from the remote device.
        int bytesRead = sok.EndReceive(ar);

        if (bytesRead > 0){

          //There might be more data, so store the data received so far.
          Array.Copy(buff, 0, rxBuffer, this.rxPosition, bytesRead);
          this.rxPosition += bytesRead;
          Array.Clear(buff, 0, buff.Length);

          //Get the rest of the data.
          if (bytesRead == buff.Length)
            sok.BeginReceive(buff, 0, buff.Length, 0, new AsyncCallback(ReceiveCallback), ar.AsyncState);
          else
            this.dsp.Invoke(System.Windows.Threading.DispatcherPriority.Normal, (ThreadStart)(() => {OnReceivedData(this.rxPosition, this.rxBuffer);}));
        }
        else{
          //All the data has arrived; put it in response.
          OnReceivedData(this.rxPosition, this.rxBuffer);
        }
      }
      catch (Exception e){
        Smv.Utils.DxInfo.ShowDxBoxInfo("Ошибка", e.ToString(), MessageBoxImage.Stop);
      }
    }

    private void OnReceivedData(int rcvBytes, byte[] data)
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<ReceivedDataEventArgs> temp = this.ReceivedData;
      if (temp != null)
        temp(this, new ReceivedDataEventArgs(rcvBytes, data));
    }

    

  }
}
