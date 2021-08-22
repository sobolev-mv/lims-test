using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel.Composition;
using Smv.Mef.Contracts;
using DevExpress.Xpf.Core;

namespace Viz.WrkModule.RptMagLab
{

  [Export(typeof(Smv.Mef.Contracts.IWorkModuleContract))]
  public sealed class RptMagLabContract : Smv.Mef.Contracts.IWorkModuleContract
  {
    private ImageSource largeGlyph;
    private Smv.MVVM.Commands.DelegateCommand runModuleCommand;
    
    public event EventHandler<Smv.RibbonUserUI.RibbonUIEventArgs> RunEvent;
    public string FriendlyName { get; set; }
    public string Version
    {
      get { return Smv.Utils.Etc.GetAssemblyVersion(System.Reflection.Assembly.GetExecutingAssembly()); }
    }

    public string Id
    {
      get { return ModuleConst.ModuleId; }
    }

    public UserControl CreateContent(Window owner)
    {
      return null;
    }

    public ImageSource LargeGlyph
    {
      get { return largeGlyph; }
    }

    public ICommand RunModuleCommand
    {
      get
      {
        if (runModuleCommand == null)
        {
          runModuleCommand = new Smv.MVVM.Commands.DelegateCommand(ExecRunModuleCommand);
        }
        return runModuleCommand;
      }
    }

    private void ExecRunModuleCommand()
    {
      EventHandler<Smv.RibbonUserUI.RibbonUIEventArgs> temp = RunEvent;
      if (temp != null)
        temp(this, new Smv.RibbonUserUI.RibbonUIEventArgs(new ViewRptMagLab(CmParam)));
    }
    
    public string CaptionControl
    {
      get { return "Отч. Маг-Лаб"; }
    }

    public string HintControl
    {
      get { return "Отчетность по измерениям магнитной лаборатории"; }
    }

    public string NameControl
    {
      get { return "BtnRptMagLab"; }
    }

    public Window MainWindow
    {
      get;
      set;
    }

    public Object CmParam { get; set; }  

    public RptMagLabContract()
    {
      largeGlyph = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.RptMagLab;Component/Images/ModuleGlyph-32x32.ico"));
    }

  }

}