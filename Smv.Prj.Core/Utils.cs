using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;
using DevExpress.Xpf.Core;
using Microsoft.Win32;

namespace Smv.Utils
{
  public interface IMessageInfo
  {
    void ShowDlgErrorInfo(string errorsTitle, string errorsMsg);
  }


  public static class WindowsOption
  {
    public static Window ActveWindow
    {get; set;}

    public static String CurrentTheme
    { get; set; }

    public static void AdjustWindowTextOption(Window wnd)
    {
      wnd.SetValue(TextOptions.TextFormattingModeProperty, Environment.OSVersion.Version.Major >= 6 ? TextFormattingMode.Ideal : TextFormattingMode.Display);
    }
  }

  //Singleton
  public sealed class Etc
  {
    private static Etc etc = null;
    public static string StartPath { get; private set; }
   
    private Etc(){}

    public static void Create(Assembly asm)
    {
      if (etc != null) return;
      etc = new Etc();
      StartPath = GetAssemblyPath(asm);
    }

    public static string GetAssemblyVersion(Assembly CurrentAsm)
    {
      return CurrentAsm.GetName().Version.ToString();
    }

    public static string GetAssemblyVersion(Type type)
    {
      string rez = null;
      var asm = Assembly.GetAssembly(type);
      if (asm != null){
        var asmName = asm.GetName();
        /*
        Console.WriteLine("Type {0} from assembly {1}", type.Name, asmName.Name);
        Console.WriteLine("Version={0} ", asmName.Version);
        Console.WriteLine("VersionCompatibility={0} ", asmName.VersionCompatibility);
        Console.WriteLine("ProcessorArchitecture={0} ", asmName.ProcessorArchitecture);
        Console.WriteLine("FullName={0} ", asmName.FullName);
        Console.WriteLine("---------------------------------------");
        Console.WriteLine();
        */
        rez = asmName.Version.ToString();
      }
      return rez;
    }

    public static String GetAssemblyPath(Assembly CurrentAsm)
    {
      //System.Reflection.Assembly.GetExecutingAssembly();
      String mp = new Uri(CurrentAsm.CodeBase).LocalPath;
      return Path.GetDirectoryName(mp);
    }

    public static void ExecFileAssociationApp(string fileName)
    {
      var proc = new System.Diagnostics.Process();
      proc.StartInfo.FileName = fileName;
      proc.StartInfo.UseShellExecute = true;
      proc.Start();      
    }

    public static void WriteToEndTxtFile(string fileName, string writeStr, Encoding encoding)
    {
      //Encoding.GetEncoding("windows-1251")
      var fileStream = new FileStream(fileName, FileMode.Append);
      var streamWriter = new StreamWriter(fileStream, encoding);
      streamWriter.Write(writeStr + "\r\n");
      streamWriter.Close();
      fileStream.Close();
    }

    public static string GetStringWithDelimFromTxtFile(Encoding encoding, string delimiter)
    {
      var ofd = new OpenFileDialog { DefaultExt = ".txt", Filter = "text format (.txt)|*.txt" };
      bool? result = ofd.ShowDialog();
      if (result != true)
        return string.Empty;
      return File.ReadAllText(ofd.FileName, encoding).Replace(" ", "").Replace("\r\n", " ").Trim().Replace(" ", delimiter);
    }
  }

  public static class Sound
  {

    [System.Flags]
    private enum PlaySoundFlags : int
    {
      SND_SYNC = 0x0000,
      SND_ASYNC = 0x0001,
      SND_NODEFAULT = 0x0002,
      SND_LOOP = 0x0008,
      SND_NOSTOP = 0x0010,
      SND_NOWAIT = 0x00002000,
      SND_FILENAME = 0x00020000,
      SND_RESOURCE = 0x00040004
    }
    
    [DllImport("winmm.DLL", EntryPoint = "PlaySound", SetLastError = true, CharSet = CharSet.Unicode, ThrowOnUnmappableChar = true)]
    private static extern bool PlaySound(string szSound, System.IntPtr hMod, PlaySoundFlags flags);

    [DllImport("kernel32.dll", EntryPoint = "Beep", SetLastError = true, CharSet = CharSet.Unicode, ThrowOnUnmappableChar = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool Beep(int dwFreq, int dwDuration);

    public static bool PlaySoundFile(String FileName)
    {
      return PlaySound(FileName, new IntPtr(), PlaySoundFlags.SND_SYNC);
    }

  }

  public static class DxInfo
  {
    public static void ShowDxBoxInfo(string strTitle, string strMsg, MessageBoxImage typeImage)
    {
      if (!Application.Current.CheckAccess())
        if (WindowsOption.ActveWindow != null)
          Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal,(ThreadStart)(() => DXMessageBox.Show(WindowsOption.ActveWindow, strMsg, strTitle, MessageBoxButton.OK, typeImage)));
        else
          Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal,(ThreadStart) (() => DXMessageBox.Show(strMsg, strTitle, MessageBoxButton.OK, typeImage)));
      else
        if (WindowsOption.ActveWindow != null)
          DXMessageBox.Show(WindowsOption.ActveWindow, strMsg, strTitle, MessageBoxButton.OK, typeImage);
        else
          DXMessageBox.Show(strMsg, strTitle, MessageBoxButton.OK, typeImage);
   }

    public static void ShowDxBoxInfo(string strTitle, string strMsg, MessageBoxImage typeImage, Window parentWindow)
    {
      if (!Application.Current.CheckAccess())
        if (parentWindow != null)
          Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DXMessageBox.Show(parentWindow, strMsg, strTitle, MessageBoxButton.OK, typeImage)));
        else
          Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)(() => DXMessageBox.Show(strMsg, strTitle, MessageBoxButton.OK, typeImage)));
      else
      if (parentWindow != null)
        DXMessageBox.Show(parentWindow, strMsg, strTitle, MessageBoxButton.OK, typeImage);
      else
        DXMessageBox.Show(strMsg, strTitle, MessageBoxButton.OK, typeImage);
    }

    public static Boolean ShowDxBoxQuestionYn(FrameworkElement owner, string strTitle, string strMsg, MessageBoxImage typeImage)
    {
      if (owner == null)
        owner = Application.Current.MainWindow;

      MessageBoxResult mbRes = DXMessageBox.Show(owner, strMsg, strTitle, MessageBoxButton.YesNo, typeImage);
      return (mbRes == MessageBoxResult.Yes);
    }
  }

  public class MessageInfo : IMessageInfo
  {
    private readonly Window parentWin;

    public void ShowDlgErrorInfo(string errorsTitle, string errorsMsg)
    {
      Smv.Utils.DxInfo.ShowDxBoxInfo(errorsTitle, errorsMsg, MessageBoxImage.Error, parentWin);
    }

    public MessageInfo(Window parentView)
    {
      parentWin = parentView;
    }

  }

  public static class VizDateTime
  {
    public static void GetDateTimeVizWorkPeriod(DateTime wrkDate, ref DateTime startDateTime, ref DateTime endDateTime)  
    {
      startDateTime = endDateTime = wrkDate;
      startDateTime = startDateTime.AddHours(8);

      endDateTime = endDateTime.AddDays(1);
      endDateTime = endDateTime.AddHours(7);
      endDateTime = endDateTime.AddMinutes(59);
      endDateTime = endDateTime.AddSeconds(59);
    }
  }


  public static class ExecDlg
  {
    public static bool InputQuery(string Title, string Prompt, ref string StrValue, Boolean IsPassword)
    {
      var vq = new Dialogs.ViewInputQuery(IsPassword) {Title = Title, tbPrompt = {Text = Prompt}};
      if (!String.IsNullOrEmpty(StrValue)) vq.teStringValue.Text = StrValue;
      if (vq.ShowDialog() == false) return false;
      StrValue = vq.teStringValue.Text;
      return true;
    }

  }


}
