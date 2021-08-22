using System;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DevExpress.Xpf.Bars;
using DevExpress.Xpf.Charts;
using DevExpress.Xpf.Core;
using Microsoft.Win32;
using Smv.MVVM.Commands;
using Smv.MVVM.ViewModels;
using Viz.WrkModule.Diagrm.Db.DataSets;

namespace Viz.WrkModule.Diagrm
{
  internal sealed class ViewModelDiagrm : ViewModelBase
  {

    #region Fields
    private DsDiagrm dsDiagrm = new DsDiagrm(); 
    private ChartControl chr;
    private BarEditItem beiGroup;
    private string  matLocId;
    private string  typeDiagrm;
    private int groupId;
    private decimal ?minAxisY;
    private decimal ?maxAxisY;
    private string titleUp;
    #endregion Fields

    #region Public Property
    public DataTable DgTypeDiagramm
    {
      get { return dsDiagrm.DgTypeDiagramm; }
    }

    public DataTable DgMeasData
    {
      get { return dsDiagrm.DgMeasData; }
    }

    public DataTable DgGroup
    {
      get { return dsDiagrm.DgGroup; }
    }

    public String MatLocId
    {
      get { return matLocId; }
      set
      {
        if (value == matLocId) return;
        matLocId = value;
        OnPropertyChanged("MatLocId");
      }
    }

    public Int32 GroupId
    {
      get { return groupId; }
      set
      {
        if (value == groupId) return;
        groupId = value;
        OnPropertyChanged("GroupId");
      }
    }


    public String TypeDiagrm
    {
      get { return typeDiagrm; }
      set
      {
        if (value == typeDiagrm) return;
        typeDiagrm = value;
        OnPropertyChanged("TypeDiagrm");
      }
    }

    public Decimal ?MinAxisY
    {
      get { return minAxisY; }
      set
      {
        if (value == minAxisY) return;
        minAxisY = value;
        OnPropertyChanged("MinAxisY");
      }
    }

    public Decimal ?MaxAxisY
    {
      get { return maxAxisY; }
      set
      {
        if (value == maxAxisY) return;
        maxAxisY = value;
        OnPropertyChanged("MaxAxisY");
      }
    }

    public string TitleUp
    {
      get => titleUp;
      set
      {
        if (value == titleUp) return;
        matLocId = value;
        OnPropertyChanged("TitleUp");
      }
    }


    #endregion Public Property

    #region Private Method
    private void GroupChanged(object sender, RoutedEventArgs e)
    {
      TypeDiagrm = null;
      dsDiagrm.DgTypeDiagramm.LoadData(GroupId);
    }
    #endregion Private Method

    #region Constructor
    internal ViewModelDiagrm(UserControl control, object LstGroup)
    {
      chr = LogicalTreeHelper.FindLogicalNode(control, "ChrDiag") as ChartControl;
      //this.beiGroup = (LstGroup as DevExpress.Xpf.Bars.BarEditItem).EditSettings as DevExpress.Xpf.Editors.Settings.ComboBoxEditSettings; 
      beiGroup = (LstGroup as BarEditItem);
      beiGroup.EditValueChanged += GroupChanged;
      dsDiagrm.DgGroup.LoadData();
    }
    #endregion Constructor

    #region Commands
    private DelegateCommand<Object> showDiagrmCommand;
    private DelegateCommand<Object> saveDiagrmCommand;
    private DelegateCommand<Object> applyMinMaxValDiagrmCommand;

    public ICommand ShowDiagrmCommand
    {
      get{return showDiagrmCommand ?? (showDiagrmCommand = new DelegateCommand<Object>(ExecuteShowDiagrm, CanExecuteShowDiagrm));}
    }

    private void ExecuteShowDiagrm(Object parameter)
    {
      chr.Diagram = null;
      chr.Titles.Clear();
      chr.AnimationMode = ChartAnimationMode.OnDataChanged;
      dsDiagrm.DgMeasData.LoadData(MatLocId, TypeDiagrm, dsDiagrm.DgTypeDiagramm.GetXType(TypeDiagrm));

      if (dsDiagrm.DgMeasData.Rows.Count == 0){
        DXMessageBox.Show(Application.Current.Windows[0], "Данные по материалу отсутствуют.", "Нет данных", MessageBoxButton.OK, MessageBoxImage.Warning);
        return;
      }

      chr.Titles.Add(new Title()
                     {
                      Content = "Лок. №: " + MatLocId,
                      HorizontalAlignment = HorizontalAlignment.Center
                     }
                    );

      chr.Diagram = new XYDiagram2D();
      chr.Diagram.Series.Add(new LineSeries2D());
      chr.Diagram.Series[0].Label = new SeriesLabel();
      chr.Diagram.Series[0].LabelsVisibility = false;
      //chr.Diagram.Series[0].PointOptions = new PointOptions();
      ((LineSeries2D) chr.Diagram.Series[0]).MarkerVisible = false;
      ((LineSeries2D) chr.Diagram.Series[0])?.SetPointAnimation(new Marker2DSlideFromTopCenterAnimation());
      ((LineSeries2D) chr.Diagram.Series[0])?.SetSeriesAnimation(new Line2DUnwrapHorizontallyAnimation());
      ((LineSeries2D) chr.Diagram.Series[0]).ValueScaleType = ScaleType.Numerical;      
      chr.Diagram.Series[0].ValueDataMember = "Value";
      
      //(chr.Diagram as XYDiagram2D).SetAxisXZoomRatio(0.02);
      ((XYDiagram2D) chr.Diagram).EnableAxisXNavigation = true;
      ((XYDiagram2D) chr.Diagram).DefaultPane = new Pane
      {
        AxisXScrollBarOptions = new ScrollBarOptions
        {
          Visible = true,
          Alignment = ScrollBarAlignment.Far
        } 
      }; 
      
      ((XYDiagram2D) chr.Diagram).AxisX = new AxisX2D
      {
        GridLinesVisible = true,
        GridLinesMinorVisible = true,
        VisualRange = new Range()
      };

      ((XYDiagram2D)chr.Diagram).AxisY = new AxisY2D
      {
        GridLinesVisible = true,
        GridLinesMinorVisible = true,
        VisualRange = new Range()
      };

      ((XYDiagram2D) chr.Diagram).AxisX.VisualRange.MinValue = dsDiagrm.DgMeasData.GetAxisXExtremesVal(MatLocId, TypeDiagrm, dsDiagrm.DgTypeDiagramm.GetXType(TypeDiagrm), AxisExtremes.Min);
      ((XYDiagram2D) chr.Diagram).AxisX.VisualRange.MaxValue = dsDiagrm.DgMeasData.GetAxisXExtremesVal(MatLocId, TypeDiagrm, dsDiagrm.DgTypeDiagramm.GetXType(TypeDiagrm), AxisExtremes.Max);
      ((XYDiagram2D) chr.Diagram).ActualAxisY.VisualRange.MinValue = dsDiagrm.DgMeasData.GetAxisYExtremesVal(MatLocId, TypeDiagrm, AxisExtremes.Min);
      ((XYDiagram2D)chr.Diagram).ActualAxisY.VisualRange.MaxValue = dsDiagrm.DgMeasData.GetAxisYExtremesVal(MatLocId, TypeDiagrm, AxisExtremes.Max);

      if (string.Equals(dsDiagrm.DgTypeDiagramm.GetXType(TypeDiagrm), "NUM", StringComparison.Ordinal)){
        chr.Diagram.Series[0].ArgumentDataMember = "Len";
        ((LineSeries2D) chr.Diagram.Series[0]).ArgumentScaleType = ScaleType.Numerical;
        ((XYDiagram2D) chr.Diagram).AxisX.NumericScaleOptions = new ContinuousNumericScaleOptions()
        {
          AutoGrid = true
        };
      }
      else{
        chr.Diagram.Series[0].ArgumentDataMember = "DgDateTime";
        ((LineSeries2D) chr.Diagram.Series[0]).ArgumentScaleType = ScaleType.DateTime;
        ((XYDiagram2D) chr.Diagram).AxisX.DateTimeScaleOptions = new ContinuousDateTimeScaleOptions()
        {
          AutoGrid = true,
          GridAlignment = DateTimeGridAlignment.Hour
        };

        ((XYDiagram2D) chr.Diagram).AxisX.Label = new AxisLabel()
        {
          TextPattern = "{A:dd.MM.yy HH:mm}",
          Staggered = true
        };
        ((LineSeries2D)chr.Diagram.Series[0]).CrosshairLabelPattern = "{A:dd.MM.yy HH:mm} - {V}";
        ((LineSeries2D)chr.Diagram.Series[0]).ToolTipSeriesPattern = "{A:dd.MM.yy HH:mm} - {V}";
        //Axis2D.SetResolveOverlappingOptions((chr.Diagram as XYDiagram2D).ActualAxisX.Label, new AxisLabelResolveOverlappingOptions { AllowRotate = true, AllowHide = true, AllowStagger = false, MinIndent = 4 });
      }


      chr.Diagram.Series[0].DataSource = dsDiagrm.DgMeasData;
      
    }

    private bool CanExecuteShowDiagrm(Object parameter)
    {
      return (!string.IsNullOrEmpty(TypeDiagrm)) && (!string.IsNullOrEmpty(MatLocId));
    }


    public ICommand SaveDiagrmCommand => saveDiagrmCommand ?? (saveDiagrmCommand = new DelegateCommand<Object>(ExecuteSaveDiagrm, CanExecuteSaveDiagrm));

    private void ExecuteSaveDiagrm(Object parameter)
    {
      SaveFileDialog sfd = new SaveFileDialog();
      sfd.DefaultExt = ".jpg";
      sfd.Filter = "jpeg image (.jpg)|*.jpg";
      bool? result = sfd.ShowDialog();

      if (result != true) return;
      
      double x = chr.ActualWidth;
      double y = chr.ActualHeight;

      chr.Measure(new Size(x, y));
      chr.Arrange(new Rect(chr.DesiredSize));
      chr.Measure(new Size(x + 1, y + 1));
      chr.Arrange(new Rect(chr.DesiredSize));
      chr.Measure(new Size(x, y));
      chr.Arrange(new Rect(chr.DesiredSize));

      VisualBrush brush = new VisualBrush(chr);
      DrawingVisual visual = new DrawingVisual();
      DrawingContext context = visual.RenderOpen();
      context.DrawRectangle(brush, null, new Rect(0, 0, chr.ActualWidth, chr.ActualHeight));
      context.Close();
      RenderTargetBitmap bmp = new RenderTargetBitmap((int)chr.ActualWidth, (int)chr.ActualHeight, 96, 96, PixelFormats.Pbgra32);
      bmp.Render(visual);
      JpegBitmapEncoder encoder = new JpegBitmapEncoder();
      encoder.Frames.Add(BitmapFrame.Create(bmp));

      FileStream file = new FileStream(sfd.FileName, FileMode.Create);
      encoder.Save(file);
      file.Close();
    }

    private bool CanExecuteSaveDiagrm(Object parameter)
    {
      return (chr.Diagram != null);
    }

    public ICommand ApplyMinMaxValDiagrmCommand
    {
      get{return applyMinMaxValDiagrmCommand ?? (applyMinMaxValDiagrmCommand = new DelegateCommand<Object>(ExecuteApplyMinMaxVal, CanExecuteApplyMinMaxVal));}
    }

    private void ExecuteApplyMinMaxVal(Object parameter)
    {
      if ((chr.Diagram as XYDiagram2D).AxisY == null){
        (chr.Diagram as XYDiagram2D).AxisY = new AxisY2D();
        (chr.Diagram as XYDiagram2D).AxisY.VisualRange = new Range();
      }

      (chr.Diagram as XYDiagram2D).AxisY.VisualRange.SetMinMaxValues(minAxisY, maxAxisY); 
    }

    private bool CanExecuteApplyMinMaxVal(Object parameter)
    {
      return (chr.Diagram != null) && (MinAxisY != null) && (MaxAxisY != null);
    }


    #endregion Commands

  }
}
