using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using System.Windows.Media;
using DevExpress.Xpf.LayoutControl;
using DevExpress.Xpf.NavBar;


namespace Viz.WrkModule.MagLab.ViewModel
{
  internal sealed class ViewModelSampleProp : Smv.MVVM.ViewModels.ViewModelBase
  {
    #region Fields
    private DataTable     dtSample;
    private String        matLocalId;
    private Window        win;
    private DataSet       dsSample;
    private LayoutControl lcRoot;
    private NavBarControl nvbcMain;
    private Boolean       isFinishedGoods  = false;
    #endregion

    #region Public Property

    public Boolean IsFinishedGoods
    {
      get { return isFinishedGoods; }
      set
      {
        if (value == isFinishedGoods) return;
        isFinishedGoods = value;
        base.OnPropertyChanged("IsFinishedGoods");
      }
    }




    #endregion

    #region Private Method
    private void ShowSampleProp()
    {
      DataView dv = new DataView(this.dtSample);
      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "'";
      dv.Sort = "DtSample";
      foreach(DataRowView drView in dv){
        string sId = Convert.ToString(drView.Row["SampleId"]);  
        DataTable dt = CreateDataTableForSample();

        LayoutGroup lg = this.lcRoot.CreateGroup(); 
        lg.View = LayoutGroupView.GroupBox;
        lg.Header = "Обр. №: " + Convert.ToString(drView.Row["SampleNum"]) + " " + 
                    Convert.ToString(drView.Row["TestType"]) + "-" + Convert.ToString(drView.Row["SamplePos"]) +
                    "-" + Convert.ToString(drView.Row["Line"]);
        lg.HorizontalAlignment = HorizontalAlignment.Left;
        lg.VerticalAlignment = VerticalAlignment.Stretch;
        lg.Width = 270;
        lg.ItemSpace = 0;
        lg.Orientation = System.Windows.Controls.Orientation.Horizontal;
        this.lcRoot.Children.Add(lg);

                 
        DevExpress.Xpf.Editors.ListBoxEdit lbe = new DevExpress.Xpf.Editors.ListBoxEdit();
        //System.Windows.Controls.ListBox lbe = new System.Windows.Controls.ListBox();
        lbe.ItemsSource = dt.DefaultView;
        //lbe.ItemTemplate = this.win.Resources["SampleInfoItemTemplate"] as DataTemplate; 
        lbe.Style = this.win.Resources["SampleInfoItemTemplate"] as Style;
        lbe.Background = Brushes.WhiteSmoke;
        lbe.BorderBrush = Brushes.WhiteSmoke; 
        lbe.Focusable = true; 
        lbe.IsHitTestVisible = true; 
        lbe.OverridesDefaultStyle = false; 
        lbe.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        lg.Children.Add(lbe);
        Viz.WrkModule.MagLab.Db.LabAction.FillSampleInfo(sId,dt); 
        
      }

      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "' AND Line LIKE '*AVO*'";
      dv.Sort = "DtSample";
        
      if (dv.Count != 0){
        LayoutGroup lgx = this.lcRoot.CreateGroup();
        lgx.View = LayoutGroupView.GroupBox;
        lgx.Header = "АВО";
        lgx.HorizontalAlignment = HorizontalAlignment.Left;
        lgx.VerticalAlignment = VerticalAlignment.Stretch;
        lgx.Width = 270;
        lgx.ItemSpace = 0;
        lgx.Orientation = System.Windows.Controls.Orientation.Horizontal;
        this.lcRoot.Children.Add(lgx);

        DataTable dtx = CreateDataTableForSample();

        DevExpress.Xpf.Editors.ListBoxEdit lbex = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbex.ItemsSource = dtx.DefaultView;
        lbex.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbex.Background = Brushes.WhiteSmoke;
        lbex.BorderBrush = Brushes.WhiteSmoke;
        lbex.Focusable = true;
        lbex.IsHitTestVisible = true;
        lbex.OverridesDefaultStyle = false;
        lbex.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        lgx.Children.Add(lbex);
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%STRANN", "",dtx); 
      }

      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "' AND Line LIKE '*APR*'";
      dv.Sort = "DtSample";

      if (dv.Count != 0)
      {
        LayoutGroup lgx = this.lcRoot.CreateGroup();
        lgx.View = LayoutGroupView.GroupBox;
        lgx.Header = "АПР";
        lgx.HorizontalAlignment = HorizontalAlignment.Left;
        lgx.VerticalAlignment = VerticalAlignment.Stretch;
        lgx.Width = 270;
        lgx.ItemSpace = 0;
        lgx.Orientation = System.Windows.Controls.Orientation.Horizontal;
        this.lcRoot.Children.Add(lgx);

        DataTable dtx = CreateDataTableForSample();

        DevExpress.Xpf.Editors.ListBoxEdit lbex = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbex.ItemsSource = dtx.DefaultView;
        lbex.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbex.Background = Brushes.WhiteSmoke;
        lbex.BorderBrush = Brushes.WhiteSmoke;
        lbex.Focusable = true;
        lbex.IsHitTestVisible = true;
        lbex.OverridesDefaultStyle = false;
        lbex.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        lgx.Children.Add(lbex);
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%FINCUT", "",dtx);

        lgx = this.lcRoot.CreateGroup();
        lgx.View = LayoutGroupView.GroupBox;
        lgx.Header = "ЛАЗЕР";
        lgx.HorizontalAlignment = HorizontalAlignment.Left;
        lgx.VerticalAlignment = VerticalAlignment.Stretch;
        lgx.Width = 270;
        lgx.ItemSpace = 0;
        lgx.Orientation = System.Windows.Controls.Orientation.Horizontal;
        this.lcRoot.Children.Add(lgx);

        dtx = CreateDataTableForSample();

        lbex = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbex.ItemsSource = dtx.DefaultView;
        lbex.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbex.Background = Brushes.WhiteSmoke;
        lbex.BorderBrush = Brushes.WhiteSmoke;
        lbex.Focusable = true;
        lbex.IsHitTestVisible = true;
        lbex.OverridesDefaultStyle = false;
        lbex.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        lgx.Children.Add(lbex);
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%LASSCR", "",dtx);


      }

    }
    
    private void ShowSampleProp2(string InspLotStatus)
    {
      
      //item1.Content = "Home";

      DataView dv = new DataView(this.dtSample);
      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "'";
      dv.Sort = "DtSample";
        

      foreach (DataRowView drView in dv)
      {
        NavBarGroup group1 = new NavBarGroup();
        group1.DisplaySource = DisplaySource.Content;
       
                
        string sId = Convert.ToString(drView.Row["SampleId"]);
        DataTable dt = CreateDataTableForSample();
        
        group1.Header = "Образец № " + Convert.ToString(drView.Row["SampleNum"]) + " " +
                    Convert.ToString(drView.Row["TestType"]) + "-" + Convert.ToString(drView.Row["SamplePos"]) +
                    "-" + Convert.ToString(drView.Row["Line"] + "  Материал № " + this.matLocalId);

        DevExpress.Xpf.Editors.ListBoxEdit lbe = new DevExpress.Xpf.Editors.ListBoxEdit();
        
        //System.Windows.Controls.ListBox lbe = new System.Windows.Controls.ListBox();
        lbe.ItemsSource = dt.DefaultView;
        //lbe.ItemTemplate = this.win.Resources["SampleInfoItemTemplate"] as DataTemplate; 
        lbe.Style = this.win.Resources["SampleInfoItemTemplate"] as Style;
        lbe.Background = Brushes.WhiteSmoke;
        lbe.BorderBrush = Brushes.WhiteSmoke;
        lbe.Focusable = true;
        lbe.IsHitTestVisible = true;
        lbe.OverridesDefaultStyle = false;
        lbe.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        Viz.WrkModule.MagLab.Db.LabAction.FillSampleInfo(sId, dt);
      
        group1.Content = lbe;
        this.nvbcMain.Groups.Add(group1);
      }

      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "' AND Line LIKE '*AVO*'";
      dv.Sort = "DtSample";

      if (dv.Count != 0)
      {

        NavBarGroup group1 = new NavBarGroup();
        group1.DisplaySource = DisplaySource.Content;
        group1.Header = "Материал после АВО № " + this.matLocalId;
        DataTable dtx = CreateDataTableForSample();

        DevExpress.Xpf.Editors.ListBoxEdit lbex = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbex.ItemsSource = dtx.DefaultView;
        lbex.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbex.Background = Brushes.WhiteSmoke;
        lbex.BorderBrush = Brushes.WhiteSmoke;
        lbex.Focusable = true;
        lbex.IsHitTestVisible = true;
        lbex.OverridesDefaultStyle = false;
        lbex.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%STRANN", InspLotStatus, dtx);
        group1.Content = lbex;
        this.nvbcMain.Groups.Add(group1);
      }

      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "' AND Line LIKE '*APR*' " + "AND LaserFlag = 0";
      dv.Sort = "DtSample";
      

      if (dv.Count != 0){
        string tstep = Convert.ToString(dv[0]["Tstep"]); 

        NavBarGroup group1 = new NavBarGroup();
        group1.DisplaySource = DisplaySource.Content;
        group1.Header = "Материал после АПР № " + this.matLocalId;
        DataTable dtx = CreateDataTableForSample();

        DevExpress.Xpf.Editors.ListBoxEdit lbex = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbex.ItemsSource = dtx.DefaultView;
        lbex.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbex.Background = Brushes.WhiteSmoke;
        lbex.BorderBrush = Brushes.WhiteSmoke;
        lbex.Focusable = true;
        lbex.IsHitTestVisible = true;
        lbex.OverridesDefaultStyle = false;
        lbex.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%" + tstep, InspLotStatus, dtx);
        group1.Content = lbex; 
        this.nvbcMain.Groups.Add(group1);

      }

      dv.RowFilter = "MatLocalNumber = " + "'" + this.matLocalId + "' AND Line LIKE '*APR*' " + "AND LaserFlag = 1";
      dv.Sort = "DtSample";

      if (dv.Count != 0)
      {
        NavBarGroup group1 = new NavBarGroup();
        group1.DisplaySource = DisplaySource.Content;
        group1.Header = "Материал после обработки лазером № " + this.matLocalId;
        DataTable dtx = CreateDataTableForSample();

        DevExpress.Xpf.Editors.ListBoxEdit lbex = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbex.ItemsSource = dtx.DefaultView;
        lbex.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbex.Background = Brushes.WhiteSmoke;
        lbex.BorderBrush = Brushes.WhiteSmoke;
        lbex.Focusable = true;
        lbex.IsHitTestVisible = true;
        lbex.OverridesDefaultStyle = false;
        lbex.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%LASSCR", InspLotStatus, dtx);
        group1.Content = lbex;
        this.nvbcMain.Groups.Add(group1);
      }

      if (IsFinishedGoods){ 
        NavBarGroup grp1 = new NavBarGroup();
        grp1.DisplaySource = DisplaySource.Content;
        grp1.Header = "Материал после окончательной приемки № " + this.matLocalId;
        DataTable dtxx = CreateDataTableForSample();

        DevExpress.Xpf.Editors.ListBoxEdit lbexx = new DevExpress.Xpf.Editors.ListBoxEdit();
        lbexx.ItemsSource = dtxx.DefaultView;
        lbexx.Style = this.win.Resources["ProbeInfoItemTemplate"] as Style;
        lbexx.Background = Brushes.WhiteSmoke;
        lbexx.BorderBrush = Brushes.WhiteSmoke;
        lbexx.Focusable = true;
        lbexx.IsHitTestVisible = true;
        lbexx.OverridesDefaultStyle = false;
        lbexx.SelectionMode = System.Windows.Controls.SelectionMode.Extended;
        Viz.WrkModule.MagLab.Db.LabAction.FillProbeInfo(this.matLocalId, "%FINISHED_GOODS", InspLotStatus, dtxx);
        grp1.Content = lbexx;
        this.nvbcMain.Groups.Add(grp1);
      }


    }


    private DataTable CreateDataTableForSample()
    {
      DataTable dt = new DataTable();
      this.dsSample.Tables.Add(dt);

      System.Data.DataColumn col = null;

      col = new System.Data.DataColumn("Utype", typeof(int), null, System.Data.MappingType.Element);
      col.AllowDBNull = false;
      dt.Columns.Add(col);

      col = new System.Data.DataColumn("Ftag", typeof(int), null, System.Data.MappingType.Element);
      dt.Columns.Add(col);

      col = new System.Data.DataColumn("CharName", typeof(string), null, System.Data.MappingType.Element);
      dt.Columns.Add(col);

      col = new System.Data.DataColumn("MeasValue", typeof(string), null, System.Data.MappingType.Element);
      dt.Columns.Add(col);
      dt.Constraints.Add(new System.Data.UniqueConstraint("Pk", new System.Data.DataColumn[] {dt.Columns["Utype"], dt.Columns["Ftag"]}, true));
      return dt; 
    }





    #endregion

    #region Constructor
    internal ViewModelSampleProp(String MatLocalId, System.Data.DataTable dt, Window win)
    {
      this.dtSample = dt;
      this.matLocalId = MatLocalId; 
      this.win = win;
      this.lcRoot = LogicalTreeHelper.FindLogicalNode(this.win, "LayoutRoot") as DevExpress.Xpf.LayoutControl.LayoutControl;
      this.nvbcMain = LogicalTreeHelper.FindLogicalNode(this.win, "nvbcMain") as NavBarControl;
      this.dsSample = new DataSet("dsSample");
      ShowSampleProp2("sent to ERP");
    }
    #endregion

    #region Commands
    private DelegateCommand<Object> showDataCommand;

    public ICommand ShowDataCommand
    {
      get{
        if (showDataCommand == null)
          showDataCommand = new DelegateCommand<Object>(ExecuteShowData, CanExecuteShowData);
        return showDataCommand;
      }
    }

    private void ExecuteShowData(Object parameter)
    {
     string ilStatus = null;
     int prm = Convert.ToInt32(parameter);
     
     if (prm == 2)
       ilStatus = "sent to ERP";

     if (prm == 1)
       ilStatus = "in work";
 
     this.nvbcMain.Groups.Clear();
     ShowSampleProp2(ilStatus);
    }

    private bool CanExecuteShowData(Object parameter)
    {
      return true;
    }

    #endregion

  }
}
