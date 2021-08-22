﻿using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.ComponentModel.Composition;

namespace Viz.WrkModule.RptOtk
{

  [Export(typeof(Smv.Mef.Contracts.IWorkModuleContract))]
  public sealed class RptMagLabContract : Smv.Mef.Contracts.IWorkModuleContract
  {
    private readonly ImageSource largeGlyph;
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
      get {return runModuleCommand ?? (runModuleCommand = new Smv.MVVM.Commands.DelegateCommand(ExecRunModuleCommand));}
    }

    private void ExecRunModuleCommand()
    {
      EventHandler<Smv.RibbonUserUI.RibbonUIEventArgs> temp = RunEvent;
      if (temp != null)
        temp(this, new Smv.RibbonUserUI.RibbonUIEventArgs(new ViewRptOtk(CmParam)));
    }

    public string CaptionControl
    {
      get { return "Отч. ОТК"; }
    }

    public string HintControl
    {
      get { return "Отчетность Отк"; }
    }

    public string NameControl
    {
      get { return "BtnRptOtk"; }
    }

    public Window MainWindow
    {
      get;
      set;
    }

    public Object CmParam { get; set; }

    public RptMagLabContract()
    {
      largeGlyph = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.RptOtk;Component/Images/ModuleGlyph-32x32.png"));
    }

  }

}