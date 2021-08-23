using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Ribbon;
using Smv.App.Config;
using Smv.Mef.Contracts;
using Smv.RibbonUserUI;
using Smv.Utils;
using ShutdownMode = System.Windows.ShutdownMode;


namespace Smv.Modules.MgrExt
{

  /// <summary>
  /// Interaction logic for Window1.xaml
  /// </summary>
  public partial class MainWindow : DXRibbonWindow
  {
    [ImportMany(typeof(IWorkModuleContract))]
    public IEnumerable<IWorkModuleContract> ModuleContracts { get; set; }

    [ImportMany(typeof(IDbModuleContract))]
    public IEnumerable<IDbModuleContract> DbContracts { get; set; }

    private IDbModuleContract DbContract;

    private CompositionContainer moduleContainer;
    private CompositionContainer dbContainer;

    //private DataSets.DsApp dsApp;
    private readonly TabbedSdiManager sdiManager;
    private int ModuleCount;
    private List<BarItem> lstStartBarItem;
    private List<BarItemLink> lstStartBarItemLink;

    [STAThread]
    public static void Main(string[] args)
    {
      Application app = new Viz.MagLab.Main.App();
      app.DispatcherUnhandledException += AppDispatcherUnhandledException;
      app.ShutdownMode = ShutdownMode.OnMainWindowClose;
      app.Run(new MainWindow());
    }

    private void Window_Closed(object sender, EventArgs e)
    {
      if (DbContract != null) DbContract.Disconnect(true);
    }

    private bool ComposeWorkModule()
    {
      var catalogModule = new DirectoryCatalog("WrkModule");
      moduleContainer = new CompositionContainer(catalogModule);
      
      try{
        moduleContainer.ComposeParts(this);
      }
      catch (CompositionException compositionException){
        DXMessageBox.Show(Application.Current.Windows[0], compositionException.ToString(), "Внимание", MessageBoxButton.OK, MessageBoxImage.Stop);
        return false;
      }
      return true;
    }

    private bool ComposeDbModule()
    {
      var catalogDb = new DirectoryCatalog("DbModule");
      dbContainer = new CompositionContainer(catalogDb);

      try{
        dbContainer.ComposeParts(this);
      }
      catch (CompositionException compositionException){
        DXMessageBox.Show(Application.Current.Windows[0], compositionException.ToString(), "Внимание", MessageBoxButton.OK, MessageBoxImage.Stop);
        return false;
      }
      return true;
    }


    public MainWindow()
    {
      InitializeComponent();
      

      string strTheme = Convert.ToString(ConfigParam.ReadPrivateConfigParam("Themes", "CurrentTheme"));
      cbThemes623510.EditValue = !String.IsNullOrEmpty(strTheme) ? strTheme : "DeepBlue";
      WindowsOption.ActveWindow = this;
      WindowsOption.AdjustWindowTextOption(this);
      Etc.Create(Assembly.GetExecutingAssembly());

      //Utils.Etc.StartPath = Utils.Etc.GetAssemblyPath(System.Reflection.Assembly.GetExecutingAssembly());
      sdiManager = new TabbedSdiManager(this, rcMain, ccMain);
      lstStartBarItem = new List<BarItem>();
      lstStartBarItemLink = new List<BarItemLink>();
    }

    private void cbThemes_EditValueChanged(object sender, RoutedEventArgs e)
    {
      WindowsOption.CurrentTheme = Convert.ToString((sender as BarEditItem).EditValue);
      //this.SetValue(DevExpress.Xpf.Core.ThemeManager.ThemeNameProperty, (sender as BarEditItem).EditValue);                     
      //ThemeManager.ApplicationThemeName = Convert.ToString((sender as BarEditItem).EditValue);
      ApplicationThemeHelper.ApplicationThemeName = Convert.ToString((sender as BarEditItem).EditValue);
    }

    private void btnQuit_ItemClick(object sender, ItemClickEventArgs e)
    {
      ConfigParam.WritePrivateConfigParam("Themes", "CurrentTheme",GetValue(ThemeManager.ThemeNameProperty));
      Close();
    }

    private void btnUpdate_ItemClick(object sender, ItemClickEventArgs e)
    {
      var startInfo = new ProcessStartInfo(Etc.StartPath + "\\Smv.DispUpdate.exe", "UPDATE " + Etc.StartPath + " \\\\vs-sp-fs02.ao.nlmk\\PSI\\Root\\LIMS");
      Process.Start(startInfo);
    }
    
    private void btnConnect_ItemClick(object sender, ItemClickEventArgs e)
    {
      stiStatus2623510.Content = DbContract.GetStatusInfo2("...");
      stiStatus3623510.Content = DbContract.GetStatusInfo3("...");
      stiStatus4623510.Content = DbContract.GetStatusInfo4("...");

      foreach (var item in lstStartBarItem)
        bmMain.Items.Remove(item);

      foreach (var item in lstStartBarItemLink)
        rpgModules.ItemLinks.Remove(item);

      lstStartBarItem.Clear();
      dbContainer?.Dispose();
      moduleContainer?.Dispose();
      DXRibbonWindow_ContentRendered(null,null);
      
    }

    private static void AppDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
      MessageBox.Show(e.Exception.Message + "\r\n" + e.Exception.Source + "r\n" + e.Exception.StackTrace);
      e.Handled = true;
    }


    private void DXRibbonWindow_ContentRendered(object sender, EventArgs e)
    {
      stiModuleVersion623510.Content = "...";

      //MessageBox.Show("Before ComposeDbModule() -> Step1");
             
      //построение контейнера
      if (!ComposeDbModule()) return;

      //MessageBox.Show("After ComposeDbModule() -> ModuleCount: " + DbContracts.Count().ToString());


      DbContract = DbContracts.ToArray()[0];
            
      if (!DbContract.Connect()) return;//!!!!!!!!!!!!!!!!!!!!!!!!!!!!

      stiStatus1623510.Content = DbContract.GetStatusInfo1("Copyright © NLMK-IT Ltd 2008-" + DateTime.Today.Year.ToString(CultureInfo.InvariantCulture));
      stiStatus2623510.Content = DbContract.GetStatusInfo2(String.Empty);
      stiStatus3623510.Content = DbContract.GetStatusInfo3(String.Empty);
      stiStatus4623510.Content = DbContract.GetStatusInfo4(String.Empty).ToUpper();
      stiModuleVersion623510.Content = "DX-XPF Версия: " + Etc.GetAssemblyVersion(stiStatus1623510.GetType());
                  
      //построение контейнера
      if (!ComposeWorkModule()) return;
      ModuleCount = ModuleContracts.Count();

      //Если не найдено не одного внешнего модуля
      if (ModuleCount == 0){
        DXMessageBox.Show(Application.Current.Windows[0], "Модулей расширения не найдено!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Stop);
        return;
      }
     
      foreach (var ModuleContract in ModuleContracts){
        
        if (DbContract.GetActualModuleVersion(ModuleContract.Id) == null)
          continue; 

        Boolean verBool = (DbContract.GetActualModuleVersion(ModuleContract.Id) == ModuleContract.Version);//!!!!!!!!!!!!!!!!!!!!!!!!

        if (!verBool && ccMain.Content == null)
          ccMain.Content = new Label
          {
            Content = "Для модуля: " + '"' + DbContract.GetModuleNameDescr(ModuleContract.Id) + '"' + "  доступно обновление. Рекомендуем запустить операцию обновления.",
            Foreground = Brushes.Blue,
            FontFamily = new FontFamily("Arial"),
            FontSize = 20,
            FontWeight = FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
          };

          /*
          DXMessageBox.Show(Application.Current.Windows[0], "Модуль: " + DbContract.GetModuleNameDescr(ModuleContract.Id) + "  требует обновления!\r\n Процесс обновления будет запущен.", 
                          "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
          btnUpdate_ItemClick(null,null);
          */
          
        var btnItem = new BarButtonItem
                        {
                          Content = ModuleContract.CaptionControl,
                          Name = ModuleContract.NameControl,
                          Hint = ModuleContract.HintControl + "\nВерсия: " + ModuleContract.Version,
                          Command = ModuleContract.RunModuleCommand,
                          LargeGlyph = ModuleContract.LargeGlyph,
                          Tag = 1
                        };
        bmMain.Items.Add(btnItem);
        ModuleContract.RunEvent += sdiManager.Exe;
        ModuleContract.CmParam = pgbMain; //EditSettings as DevExpress.Xpf.Editors.Settings.ProgressBarEditSettings;  
        var btnItemLnk = new BarButtonItemLink
                           {
                             BarItemName = ModuleContract.NameControl,
                             RibbonStyle = RibbonItemStyles.All
                           };
        rpgModules.ItemLinks.Add(btnItemLnk);
        ModuleContract.MainWindow = this;

        lstStartBarItem.Add(btnItem);
        lstStartBarItemLink.Add(btnItemLnk);

      }
      
    }
  }

}
