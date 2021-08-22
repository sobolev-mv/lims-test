using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Ribbon;
using System.Windows.Markup;
using System.ComponentModel;

namespace Smv.RibbonUserUI
{
    
  public sealed class RibbonUIEventArgs : EventArgs
  {
    private RibbonUserControl ribbonUsrControl;

    public RibbonUIEventArgs(RibbonUserControl UsrControl)
    {
      ribbonUsrControl = UsrControl;
    }

    public RibbonUserControl RibbonUsrControl 
    {
      get { return ribbonUsrControl; }
      set { ribbonUsrControl = value; }
    }
  }
  
   
  public class RibbonUserControl : UserControl, INotifyPropertyChanged
  {
      public string Caption
      {get; set;}

      public string RegName
      { get; set; }

      public Object ObjectParam
      { get; set; }

      public RibbonUserControl() : base()
      {
        SetValue(BarManagerItemsProperty, new List<BarItem>());
        SetValue(UserPagesProperty, new List<RibbonPage>());
      }
      
      #region Bar Manager Items

      public List<BarItem> BarManagerItems
      {
        get { return (List<BarItem>)GetValue(BarManagerItemsProperty); }
        set { SetValue(BarManagerItemsProperty, value); }
      }

      // Using a DependencyProperty as the backing store for BarManagerItems.  This enables animation, styling, binding, etc...
      public static readonly DependencyProperty BarManagerItemsProperty = DependencyProperty.Register("BarManagerItems", typeof(List<BarItem>), typeof(RibbonUserControl), new UIPropertyMetadata(null));

      #endregion
 

      #region UserPages

        public List<RibbonPage> UserPages
        {
            get { return (List<RibbonPage>)GetValue(UserPagesProperty); }
            set { SetValue(UserPagesProperty, value); }
        }

        // Using a DependencyProperty as the backing store for RibbonItems.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty UserPagesProperty = DependencyProperty.Register("UserPages", typeof(List<RibbonPage>), typeof(RibbonUserControl), new UIPropertyMetadata(null));

        #endregion

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
          PropertyChangedEventHandler handler = PropertyChanged;

          if (handler != null){
            handler(this, new PropertyChangedEventArgs(propertyName));
          }
        }

    }
}
