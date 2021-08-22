using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;

namespace Smv.Mef.Contracts
{
  
  public interface IWorkModuleContract
  {
    string Id {get;}
    string Version {get;}
    string FriendlyName { get; }
    string CaptionControl { get; }
    string HintControl { get; }
    string NameControl { get; }
    ImageSource LargeGlyph { get; }
    Window MainWindow { get; set; }
    ICommand RunModuleCommand { get; }
    Object CmParam {get; set;}

    event EventHandler<Smv.RibbonUserUI.RibbonUIEventArgs> RunEvent;
  }

  public interface IDbModuleContract
  {
    //string ModuleName{get;}
    Boolean Connect();
    void Disconnect(Boolean IsDispose);
    string GetStatusInfo1(string inf);
    string GetStatusInfo2(string inf);
    string GetStatusInfo3(string inf);
    string GetStatusInfo4(string inf);
    string GetActualModuleVersion(string ModuleId);
    string GetModuleNameDescr(string ModuleId);
  }
  
}
