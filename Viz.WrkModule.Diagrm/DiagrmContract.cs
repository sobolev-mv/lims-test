﻿using System;
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

namespace Viz.WrkModule.Diagrm
{

  [Export(typeof(Smv.Mef.Contracts.IWorkModuleContract))]
  public sealed class DiagrmContract : Smv.Mef.Contracts.IWorkModuleContract
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
      get { return Viz.WrkModule.Diagrm.ModuleConst.ModuleId; }
    }

    public UserControl CreateContent(System.Windows.Window owner)
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
        temp(this, new Smv.RibbonUserUI.RibbonUIEventArgs(new ViewDiagrm()));
    }

    public string CaptionControl
    {
      get { return "Диагр"; }
    }

    public string HintControl
    {
      get { return "Работа с диаграммами х-тик производства ЭС"; }
    }

    public string NameControl
    {
      get { return "BtnDiagrm"; }
    }

    public Window MainWindow
    {
      get;
      set;
    }

    public Object CmParam { get; set; }

    public DiagrmContract()
    {
      largeGlyph = new BitmapImage(new Uri("pack://application:,,,/Viz.WrkModule.Diagrm;Component/Images/ModuleGlyph-32x32.ico"));
    }

  }

}