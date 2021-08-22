using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using DevExpress.Xpf.Ribbon;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Bars;


namespace Smv.RibbonUserUI
{
  public sealed class TabbedSdiManager
  {
    int countWnd = 0;
    RibbonControl rc = null;
    RibbonUserControl ucCurrent = null;
    ContentControl cc = null;
    System.Windows.Window mainWnd = null;
    String oldTitle = null;

    public TabbedSdiManager(System.Windows.Window MainWindow, RibbonControl Rc, ContentControl Cc)
    {
      rc = Rc;
      cc = Cc;
      mainWnd = MainWindow;
      oldTitle = mainWnd.Title;
    }

    private void QuitItemClick(object sender, ItemClickEventArgs e)
    {
      CloseTabbedDoc();
    }

    public void CloseTabbedDoc()
    {
      if (countWnd == 0) return;

      foreach (RibbonPage rp in ucCurrent.UserPages)
        (rc.ActualCategories[0] as RibbonDefaultPageCategory).Pages.Remove(rp);
      

      foreach (BarItem bi in ucCurrent.BarManagerItems){
        //DXMessageBox.Show(bi.Content.ToString());
        rc.Manager.UnregisterName(bi.Name);
        if (Convert.ToString(bi.Tag).CompareTo("NotDispose") != 0)
          rc.Manager.Items.Remove(bi);

        if (Convert.ToString(bi.Tag).CompareTo("CloseUserControl") == 0)
          bi.ItemClick -= QuitItemClick;
      }

      cc.UnregisterName(ucCurrent.RegName);
      cc.Content = null;
      ucCurrent = null;
      
      rc.DataContext = null;
      countWnd--;
      mainWnd.Title = oldTitle;
      (rc.ActualCategories[0] as RibbonDefaultPageCategory).Pages[0].IsVisible = true;
      (rc.ActualCategories[0] as RibbonDefaultPageCategory).Pages[0].IsSelected = true;
      GC.Collect();
    }


    public void NewTabbedDoc(RibbonUserControl UsrControl)
    {
      if (UsrControl == null) return;
      if (countWnd != 0) return;
      if ((countWnd > 0) && (UsrControl.GetType() == ucCurrent.GetType())) return;

      
      ucCurrent = UsrControl;
      cc.RegisterName(ucCurrent.RegName, ucCurrent);
      cc.Content = UsrControl;
      mainWnd.Title = ucCurrent.Caption;

      foreach (BarItem bi in ucCurrent.BarManagerItems){
        rc.Manager.RegisterName(bi.Name, bi);
        rc.Manager.Items.Add(bi);
        if (Convert.ToString(bi.Tag).CompareTo("CloseUserControl") == 0) 
          bi.ItemClick += QuitItemClick;
      }
      
      foreach (RibbonPage rp in ucCurrent.UserPages)
        (rc.ActualCategories[0] as RibbonDefaultPageCategory).Pages.Add(rp);
      

      foreach (BarItem bi in rc.Manager.Items)
        bi.DataContext = ucCurrent.DataContext;

      (rc.ActualCategories[0] as RibbonDefaultPageCategory).Pages[0].IsVisible = false;
      countWnd++;
    }

    public void Exe(object sender, Smv.RibbonUserUI.RibbonUIEventArgs e)
    {
      NewTabbedDoc(e.RibbonUsrControl);
    }


  }
}
