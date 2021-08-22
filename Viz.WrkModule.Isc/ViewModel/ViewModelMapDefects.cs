using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using System.Windows.Shapes;
using System.Windows.Media;
using DevExpress.Mvvm;
using DevExpress.Mvvm.DataAnnotations;
using DevExpress.Mvvm.POCO;
using DevExpress.Xpf.Core;
using Viz.WrkModule.Isc.Db.DataSets;
using Viz.DbApp.Psi;
using DevExpress.Xpf.Grid;
using DevExpress.Xpf.Editors;
using Viz.WrkModule.MapDefects;
using Viz.WrkModule.MapDefects.Db;
using Viz.WrkModule.MapDefects.Db.DataSets;

namespace Viz.WrkModule.Isc
{
  
  public class ViewModelMapDefects
  {
    #region Fields
    private readonly DataRow dtRowMat;
    private Window view;
    private readonly ProgressBarEdit pgbWait;
    private readonly DXTabControl tcMain;
    private Canvas cnv1;
    private Canvas cnv2;
    #endregion

    #region Public Property
    public virtual Boolean IsControlEnabled { get; set; }
    public virtual decimal ScaleX { get; set; }
    public virtual decimal ScaleY { get; set; }
    #endregion

    #region Private Method

    private void WinLoaded(object sender, RoutedEventArgs e)
    {
      BeforeTask();
      Task.Factory.StartNew(BuildDefectsMap, (TypeTypeMnf)Convert.ToInt32(dtRowMat["TypeMnf"])).ContinueWith(AfterTaskEnd);
    }

    private void BeforeTask()
    {
      IsControlEnabled = false;
      this.pgbWait.StyleSettings = new ProgressBarMarqueeStyleSettings();
      (this.pgbWait.StyleSettings as ProgressBarMarqueeStyleSettings).AccelerateRatio = 10;
    }

    private void AfterTaskEnd(Task obj)
    {
      this.view.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)(() =>
      {
        this.pgbWait.StyleSettings = new ProgressBarStyleSettings();
        IsControlEnabled = true;
        CommandManager.InvalidateRequerySuggested();
      }));
    }

    private void SetScaleXY(Canvas cnv, decimal scaleX, decimal scaleY)
    {
      cnv.LayoutTransform = new ScaleTransform(Convert.ToDouble(scaleX / 100), Convert.ToDouble(scaleY / 100));
    }

    private string GetLabelDefect(DataView dvDef)
    {
      string rez = "(";
      //Делаем скобки с категорией
      /*
      if (Convert.ToString(dvDef[0].Row["Cat"]) == "б/к")
        rez = "(" + Convert.ToString(dvDef[0].Row["Cat"]) + "/";
      else
        rez = "(" + Convert.ToString(dvDef[0].Row["Cat"]) + "к/";
      */

      rez += Convert.ToString(dvDef[0].Row["FehlerTyp"]) + "/" + ((Convert.ToDecimal(dvDef[0].Row["ZoneTo"]) - Convert.ToDecimal(dvDef[0].Row["ZoneFrom"])) / 1000).ToString("n3") + ")";

      for (int i = 1; i <= dvDef.Count - 1; i++)
        rez += "(" + Convert.ToString(dvDef[i].Row["FehlerTyp"]) + "/" + Math.Round(Convert.ToDouble(dvDef[i].Row["YPOSVON"]), 0).ToString("n0") + "-" +
               Math.Round(Convert.ToDouble(dvDef[i].Row["YPOSBIS"]), 0).ToString("n0") + ")";
      /*
      foreach (DataRowView drv in dvDef){
        //MessageBox.Show(Convert.ToString(drv.Row["Cat"]));
      }
      */
      return rez;
    }

    //Работает для Канвы1
    private void PaintCoilRuleForward(double Kx, int wgtUnit, double xMin, double xMax, double yMin, double wgtCoil, int nRnd, Canvas cnv)
    {
      Label lbl = null;
      double xRuleUnit = Math.Round(wgtUnit * Kx, nRnd);
      int rulePartQnt = Convert.ToInt32(wgtCoil / wgtUnit);

      view.Dispatcher.BeginInvoke(new Action(() =>
      { 
        for (int i = 1; i < rulePartQnt; i++){

          cnv.Children.Add(new Line
          {
            X1 = xMax - xRuleUnit * i,
            Y1 = yMin - 7,
            X2 = xMax - xRuleUnit * i,
            Y2 = yMin,
            Stroke = Brushes.Black,
            StrokeThickness = 1
          }
                          );

          lbl = new Label
          {
            Content = ((i * wgtUnit * 0.001)).ToString("n1"),
            Foreground = Brushes.Black,
            FontFamily = new FontFamily("Arial"),
            FontSize = 11,
          };
          Canvas.SetLeft(lbl, (xMax - xRuleUnit * i) + 2);
          Canvas.SetTop(lbl, yMin - 17);
          cnv.Children.Add(lbl);
        }

        lbl = new Label
        {
          Content = (wgtCoil / 1000).ToString("n3"),
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 11,
        };
        Canvas.SetLeft(lbl, xMin);
        Canvas.SetTop(lbl, yMin - 17);
        cnv.Children.Add(lbl);

      }), DispatcherPriority.Render);
    }

    //Работает для Канвы2
    private void PaintCoilRuleBackward(double Kx, int wgtUnit, double xMin, double xMax, double yMin, double wgtCoil, int nRnd, Canvas cnv)
    {
      Label lbl = null;
      double xRuleUnit = Math.Round(wgtUnit * Kx, nRnd);
      int rulePartQnt = Convert.ToInt32(wgtCoil / wgtUnit);

      view.Dispatcher.BeginInvoke(new Action(() =>
      {

        for (int i = 1; i < rulePartQnt; i++){

          cnv.Children.Add(new Line
          {
            X1 = xMin + xRuleUnit * i,
            Y1 = yMin - 7,
            X2 = xMin + xRuleUnit * i,
            Y2 = yMin,
            Stroke = Brushes.Black,
            StrokeThickness = 1
          }
                          );

          lbl = new Label
          {
            Content = (i * wgtUnit * 0.001).ToString("n1"),
            Foreground = Brushes.Black,
            FontFamily = new FontFamily("Arial"),
            FontSize = 11,
          };
          Canvas.SetLeft(lbl, (xMin + xRuleUnit * i) - 24);
          Canvas.SetTop(lbl, yMin - 18);
          cnv.Children.Add(lbl);
        }

        lbl = new Label
        {
          Content = (wgtCoil / 1000).ToString("n3"),
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 11,
        };
        Canvas.SetLeft(lbl, xMax - 26);
        Canvas.SetTop(lbl, yMin - 18);
        cnv.Children.Add(lbl);

      }), DispatcherPriority.Render);
    }

    private Brush GetBrush(int id)
    {
      switch (id)
      {
        case 1:
          return Brushes.Blue;
        case 2:
          return Brushes.Red;
        case 3:
          return Brushes.Green;
        case 4:
          return Brushes.Gold;
        case 5:
          return Brushes.Fuchsia;
        case 6:
          return Brushes.Firebrick;
        case 7:
          return Brushes.DarkOrange;
        default:
          return Brushes.Black;
      }
    }

    private VisualBrush GetHatchBrush(int Id, double Height, double Width)
    {
      double mVal = Math.Min(Height, Width);
      VisualBrush vb = null;
      Canvas cnvVb = null;

      view.Dispatcher.Invoke(new Action(() =>
      {
        cnvVb = new Canvas();

        vb = new VisualBrush
        {
          TileMode = TileMode.Tile,
          Viewport = new Rect(0, 0, mVal, mVal),
          ViewportUnits = BrushMappingMode.Absolute,
          Viewbox = new Rect(0, 0, mVal, mVal),
          ViewboxUnits = BrushMappingMode.Absolute,
        };

      }), DispatcherPriority.Render);


      if ((Id % 2) == 0)

        view.Dispatcher.Invoke(() =>
        {
          cnvVb.Children.Add(new Path()
            {
              Stroke = GetBrush(Id),//Brushes.Black,
              Data = new LineGeometry(new Point(0, 0), new Point(mVal, mVal))
            }
          );
        }, DispatcherPriority.Render);

      else
        view.Dispatcher.Invoke(() =>
        {
          cnvVb.Children.Add(new Path()
          {
            Stroke = GetBrush(Id),//Brushes.Black,
            Data = new LineGeometry(new Point(0, mVal), new Point(mVal, 0))
          }
                            );
        }, DispatcherPriority.Render);

      view.Dispatcher.Invoke(() =>
      {
        vb.Visual = cnvVb;
      }, DispatcherPriority.Render);

      return vb;
    }

    private double GetActualWidthCanvas()
    {
      double res = 0;

      view.Dispatcher.Invoke(new Action(() =>
      { 
        res = this.cnv1.ActualWidth;
      }), DispatcherPriority.Render);

      return res;
    }

    private void DrawDefectsOnCoileSide(DsMapDef dsMapDef, double xMax, double yMax, double yMin, double kx, double ky, double yForward, int nrnd, List<double> lstSf, string beginLabel, ref int zIdx)
    {
      double oldX = xMax;

      foreach (DataRow rowZone in dsMapDef.LstDefZones.Rows){

        double zoneFrom = Convert.ToDouble(rowZone["ZoneFrom"]);
        double zoneTo = Convert.ToDouble(rowZone["ZoneTo"]);
        dsMapDef.MapDef.DefaultView.RowFilter = "ZoneFrom=" + zoneFrom.ToString(CultureInfo.InvariantCulture) + " AND ZoneTo=" + zoneTo.ToString(CultureInfo.InvariantCulture);

        lstSf.Add(xMax - Math.Round(zoneTo * kx, nrnd));
        string lblDefect = zIdx.ToString(CultureInfo.InvariantCulture) + GetLabelDefect(dsMapDef.MapDef.DefaultView);
        double lineY2 = yMax + yForward * zIdx;
        double lblDefectTop = yMax + yForward * zIdx - 13;

        view.Dispatcher.BeginInvoke(new Action(() =>
        {
          var line = new Line
          {
            X1 = xMax - Math.Round(zoneTo * kx, nrnd),
            Y1 = yMin,
            X2 = xMax - Math.Round(zoneTo * kx, nrnd),
            Y2 = lineY2,
            Stroke = Brushes.Black,
            StrokeThickness = 1
          };
          cnv1.Children.Add(line);

          var hlbl = new Label
          {
            Content = lblDefect,
            Foreground = Brushes.Black,
            FontFamily = new FontFamily("Arial"),
            FontSize = 10,
            FontWeight = FontWeights.Bold
          };
          Canvas.SetLeft(hlbl, xMax - Math.Round(zoneTo * kx, nrnd) + 1);
          Canvas.SetTop(hlbl, lblDefectTop);
          cnv1.Children.Add(hlbl);
        }), DispatcherPriority.Render);

        int idBrush = 1;
        foreach (DataRowView drv in dsMapDef.MapDef.DefaultView){

          string strCat = Convert.ToString(drv.Row["Cat"]);

          if ((strCat == "3") || (strCat == "б/к")){

            //Здесь рисуем поперечный основной дефект
            double yPos1 = Convert.ToDouble(drv.Row["YposvOn"]);
            double yPos2 = Convert.ToDouble(drv.Row["YposbIs"]);

            var rectHeight = Math.Round((yPos2 - yPos1) * ky, nrnd);
            var rectWidth = oldX - (xMax - Math.Round(zoneTo * kx, nrnd));
            var rectFill = GetHatchBrush(idBrush, Math.Round((yPos2 - yPos1) * ky, nrnd), oldX - (xMax - Math.Round(zoneTo * kx, nrnd)));
            var rectStroke = this.GetBrush(idBrush);

            view.Dispatcher.BeginInvoke(new Action(() =>
            {
              var rect = new Rectangle()
              {
                Height = rectHeight,
                Width = rectWidth,
                Fill = rectFill,
                Stroke = rectStroke,
                StrokeThickness = 1
              };
              Canvas.SetLeft(rect, xMax - Math.Round(zoneTo * kx, nrnd));
              Canvas.SetTop(rect, yMax - Math.Round(yPos1 * ky, nrnd) - Math.Round((yPos2 - yPos1) * ky, nrnd));
              cnv1.Children.Add(rect);
            }), DispatcherPriority.Render);
          }

          idBrush++;
        }
        oldX = xMax - Math.Round(zoneTo * kx, nrnd);
        zIdx++;
      }

      view.Dispatcher.BeginInvoke(new Action(() =>
      {
        //Делаем подпись "начало"
        var hlbl = new Label
        {
          Content = beginLabel,
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 10,
          FontWeight = FontWeights.Bold,
          RenderTransform = new RotateTransform(90),
        };
        Canvas.SetLeft(hlbl, xMax + 14);
        Canvas.SetTop(hlbl, yMin);
        cnv1.Children.Add(hlbl);

      }), DispatcherPriority.Render);

    }

    private void BuildDefectsMap(object state)
    {
      Label hlbl = null;
      Polyline plCoil = null;
      double lenCoil = 0;
      double coilWgt = 0;

      //Списки для запоминания координаты X дефекта каждой из поверхностей
      var lstSf1 = new List<double> { };
      var lstSf2 = new List<double> { };

      //DataSet
      var dsMapDef = new DsMapDef();

      var rm = new Random();
      Int64 zdn = rm.Next(10000000, 99999999);

      //Получаем длину рулона в метрах
      lenCoil = Convert.ToDouble((TypeTypeMnf)state == TypeTypeMnf.Viz ? MapDefectsAction.GetLenCoilUo(Convert.ToString(this.dtRowMat["LocNo"])) / 1000 : MapDefectsAction.GetCoilWgtAndLengthNlmkPack(Convert.ToInt64(dtRowMat["MeId"]), "EM_LENGTH"));

      //В случае если нет данных для отрисовки
      if (Math.Abs(lenCoil) < 0.001){

        view.Dispatcher.BeginInvoke(new Action(() =>
        {
          var hlblExit = new Label
          {
            Content = "NO DATA FOUND!",
            Foreground = Brushes.Crimson,
            FontFamily = new FontFamily("Arial"),
            FontSize = 50,
            FontWeight = FontWeights.Bold
          };
          Canvas.SetLeft(hlblExit, 100);
          Canvas.SetTop(hlblExit, cnv1.ActualHeight / 2);
          cnv1.Children.Add(hlblExit);

          hlblExit = new Label
          {
            Content = "NO DATA FOUND!",
            Foreground = Brushes.Crimson,
            FontFamily = new FontFamily("Arial"),
            FontSize = 50,
            FontWeight = FontWeights.Bold
          };
          Canvas.SetLeft(hlblExit, 100);
          Canvas.SetTop(hlblExit, cnv1.ActualHeight / 2);
          cnv2.Children.Add(hlblExit);
        }), DispatcherPriority.Render);

        return;
      }

      //получаем полную массу рулона в кг. 
      if ((TypeTypeMnf) state == TypeTypeMnf.Viz){
        coilWgt = Convert.ToDouble(MapDefectsAction.GetCoilWgtUo(Convert.ToString(this.dtRowMat["LocNo"])));
        MapDefectsAction.CreateDefectsData(-zdn, Convert.ToString(this.dtRowMat["LocNo"]), false);
      }
      else{
        coilWgt = lenCoil * 1000 * Convert.ToDouble(dtRowMat["Width"]) * Convert.ToDouble(dtRowMat["Thickness"]) * 7.65 / 1000000;
        MapDefectsAction.CreateDefectsDataNlmk(-zdn, Convert.ToInt64(dtRowMat["MeId"]), Convert.ToDecimal(coilWgt));
      }


      dsMapDef.MapDef.LoadData(-zdn, 1, 3);
      dsMapDef.LstDefZones.LoadData(-zdn, 1, 3);

      //получаем полную ширину рулона в мм. 
      double coilWidth = Convert.ToDouble(dtRowMat["Width"]);

      //Здесь проверяем широкий ли это монитор
      double canvasWidth = GetActualWidthCanvas();
      if (canvasWidth > 1280)
        canvasWidth = 1280;

      view.Dispatcher.BeginInvoke(new Action(() =>
      {
        cnv1.Children.Clear();
        cnv2.Children.Clear();
        cnv1.LayoutTransform = null;
      }), DispatcherPriority.Render);

      const int nrnd = 6;  //кол-во знаков после зяпятой при округлении
      const double xMin = 20;
      const double yMin = 90;
      double xMax = Math.Round(canvasWidth - canvasWidth / 4, nrnd);
      const double yMax = 190;
      const double yForward = 15; //высота на которую увеличивется растояние по y для описания дефектов
      double kx = Math.Round((xMax - xMin) / coilWgt, nrnd);//масштабирование
      double ky = Math.Round((yMax - yMin) / coilWidth, nrnd);//масштабирование

      view.Dispatcher.BeginInvoke(new Action(() =>
      {
        //Здесь рисуем заголовок с необходимыми параметрами рулона на странице 1
        hlbl = new Label
        {
          Content = "HeatNo/PlacementNum: " + Convert.ToString(this.dtRowMat["HeatNo"]) + "/" + Convert.ToString(this.dtRowMat["PlacementNum"]),
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 13,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, xMin);
        Canvas.SetTop(hlbl, 25);
        cnv1.Children.Add(hlbl);

        hlbl = new Label
        {
          Content = "COIL SURFACE",
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 16,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, 380);
        Canvas.SetTop(hlbl, 21);
        cnv1.Children.Add(hlbl);

        hlbl = new Label
        {
          Content = "AnnealingLot: " + Convert.ToString(this.dtRowMat["AnnealingLot"]) + "  " +
                    "Thickness: " + Convert.ToDecimal(this.dtRowMat["Thickness"]).ToString("n2") + "мм    " + "Width: " + coilWidth.ToString("n0") + "мм  " +
                    "Net: " + (coilWgt / 1000).ToString("n3") + " tn" + "  " +
                    "Length: " + lenCoil.ToString("n3") + "м  " + "Manufacturer: " + Convert.ToString(this.dtRowMat["NameMnf"]),

          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 13,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, xMin);
        Canvas.SetTop(hlbl, 43);
        cnv1.Children.Add(hlbl);

        //Здесь рисуем заголовок с необходимыми параметрами рулона на странице 2
        hlbl = new Label
        {
          Content = "HeatNo/PlacementNum: " + Convert.ToString(this.dtRowMat["HeatNo"]) + "/" + Convert.ToString(this.dtRowMat["PlacementNum"]),
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 13,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, xMin);
        Canvas.SetTop(hlbl, 25);
        cnv2.Children.Add(hlbl);

        hlbl = new Label
        {
          Content = "COIL SURFACE",
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 16,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, 380);
        Canvas.SetTop(hlbl, 21);
        cnv2.Children.Add(hlbl);

        hlbl = new Label
        {
          Content = "AnnealingLot: " + Convert.ToString(this.dtRowMat["AnnealingLot"]) + "  " +
                    "Thickness: " + Convert.ToDecimal(this.dtRowMat["Thickness"]).ToString("n2") + "мм    " + "Width: " + coilWidth.ToString("n0") + "мм  " +
                    "Net: " + (coilWgt / 1000).ToString("n3") + " tn" + "  " +
                    "Length: " + lenCoil.ToString("n3") + "м  " + "Manufacturer: " + Convert.ToString(this.dtRowMat["NameMnf"]),
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 13,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, xMin);
        Canvas.SetTop(hlbl, 43);
        cnv2.Children.Add(hlbl);

        //рисуем первую сторону рулона
        plCoil = new Polyline();
        plCoil.Points.Add(new Point(xMin, yMin));
        plCoil.Points.Add(new Point(xMax, yMin));
        plCoil.Points.Add(new Point(xMax, yMax));
        plCoil.Points.Add(new Point(xMin, yMax));
        plCoil.Points.Add(new Point(xMin, yMin));
        plCoil.Stroke = Brushes.Black;
        plCoil.StrokeThickness = 2;
        cnv1.Children.Add(plCoil);

      }), DispatcherPriority.Render);

      //рисуем весовую линейку первой стороны 
      this.PaintCoilRuleForward(kx, 500, xMin, xMax, yMin, coilWgt, nrnd, cnv1);
      
      //здесь начинается сама отрисовка дефектов первой стороны  
      int zIdx = 1;
      DrawDefectsOnCoileSide(dsMapDef, xMax, yMax, yMin, kx, ky, yForward, nrnd, lstSf1, "Begin", ref zIdx);

      //Здесь начинаем рисовать вторую сторону рулона
      dsMapDef.MapDef.LoadData(-zdn, 2, 3);
      dsMapDef.LstDefZones.LoadData(-zdn, 2, 3);
      MapDefectsAction.DeleteDefectsData(-zdn);

      //Определяем yMin для второй стороны рулона
      double yMin2 = yMax + yForward * zIdx + 10;
      double yMax2 = yMin2 + (yMax - yMin);

      view.Dispatcher.BeginInvoke(new Action(() =>
      {
        plCoil = new Polyline();
        plCoil.Points.Add(new Point(xMin, yMin2));
        plCoil.Points.Add(new Point(xMax, yMin2));
        plCoil.Points.Add(new Point(xMax, yMax2));
        plCoil.Points.Add(new Point(xMin, yMax2));
        plCoil.Points.Add(new Point(xMin, yMin2));
        plCoil.Stroke = Brushes.Black;
        plCoil.StrokeThickness = 2;
        cnv1.Children.Add(plCoil);
      }), DispatcherPriority.Render);

      //рисуем весовую линейку второй стороны 
      this.PaintCoilRuleForward(kx, 500, xMin, xMax, yMin2, coilWgt, nrnd, cnv1);

      //double oldX = xMax;
      zIdx = 1; //сбрасываем

      DrawDefectsOnCoileSide(dsMapDef, xMax, yMax2, yMin2, kx, ky, yForward, nrnd, lstSf2, "Begin", ref zIdx);

      double yMin3 = 140;
      double yMax3 = yMin3 + (yMax - yMin);

      //Далее рисуем заполняемый заголовок н странице 2
      //Определяем yMin для раскроечного рулона
      view.Dispatcher.BeginInvoke(new Action(() =>
      {
        hlbl = new Label
        {
          Content = "Date__________________________AGR №_________Team №________Superviser__________________________________" +
                    "      Width______________________мм  Net_________________tn",
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 13,
          FontWeight = FontWeights.Bold
        };
        Canvas.SetLeft(hlbl, 20);
        Canvas.SetTop(hlbl, 100);
        cnv2.Children.Add(hlbl);

        //yMax3 = yMin3 + (yMax - yMin);

        plCoil = new Polyline();
        plCoil.Points.Add(new Point(xMin, yMin3));
        plCoil.Points.Add(new Point(xMax, yMin3));
        plCoil.Points.Add(new Point(xMax, yMax3));
        plCoil.Points.Add(new Point(xMin, yMax3));
        plCoil.Points.Add(new Point(xMin, yMin3));
        plCoil.Stroke = Brushes.Black;
        plCoil.StrokeThickness = 2;
        cnv2.Children.Add(plCoil);

      }), DispatcherPriority.Render);


      //рисуем весовую линейку для раскроечного рулона 
      this.PaintCoilRuleBackward(kx, 500, xMin, xMax, yMin3, coilWgt, nrnd, cnv2);

      view.Dispatcher.BeginInvoke(new Action(() =>
      {
        //рисуем пунктиром дефектные зоны для раскроечного рулона
        foreach (double t in lstSf1)
        {
          var line = new Line
          {
            X1 = t,
            Y1 = yMin3,
            X2 = t,
            Y2 = yMax3,
            Stroke = Brushes.Black,
            StrokeThickness = 1,
            StrokeDashArray = DoubleCollection.Parse("5, 3")
          };
          cnv2.Children.Add(line);
        }

        foreach (double t in lstSf2)
        {
          var line = new Line
          {
            X1 = t,
            Y1 = yMin3,
            X2 = t,
            Y2 = yMax3,
            Stroke = Brushes.Black,
            StrokeThickness = 1,
            StrokeDashArray = DoubleCollection.Parse("5, 3")
          };
          cnv2.Children.Add(line);
        }

        //Делаем подпись конец
        hlbl = new Label
        {
          Content = "End",
          Foreground = Brushes.Black,
          FontFamily = new FontFamily("Arial"),
          FontSize = 10,
          FontWeight = FontWeights.Bold,
          RenderTransform = new RotateTransform(90),
        };
        Canvas.SetLeft(hlbl, xMax + 14);
        Canvas.SetTop(hlbl, yMin3);
        cnv2.Children.Add(hlbl);

      }), DispatcherPriority.Render);
    }

    #endregion

    #region Constructor
    public ViewModelMapDefects(Window view, DataRow dtRowMat)
    {
      this.view = view;
      this.dtRowMat = dtRowMat;
      tcMain = LogicalTreeHelper.FindLogicalNode(this.view, "tcMain") as DXTabControl;
      pgbWait = LogicalTreeHelper.FindLogicalNode(this.view, "PgbMeasure") as ProgressBarEdit;
      this.cnv1 = LogicalTreeHelper.FindLogicalNode(this.view, "Cnv1") as Canvas;
      this.cnv2 = LogicalTreeHelper.FindLogicalNode(this.view, "Cnv2") as Canvas;
      view.Loaded += WinLoaded;
      ScaleY = 100;
      ScaleX = 100;

      IsControlEnabled = true;
    }

    #endregion

    #region Command
    public void ShowDefectMap()
    {
      //BuildMapDefPackViz();
    }

    public bool CanShowDefectMap()
    {
      return true;
    }

    public void SetScaleY()
    {
      this.SetScaleXY(cnv1, ScaleX, ScaleY);
    }

    public bool CanSetScaleY()
    {
      return cnv1.Children.Count > 2;
    }

    public void PrintMapDefects()
    {
      var printDialog = new PrintDialog();

      if (printDialog.ShowDialog().GetValueOrDefault() == true){
        printDialog.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape;
        //printDialog.PrintQueue.GetPrintCapabilities().
        printDialog.PrintVisual(this.cnv1, "Print Defects Map1");
      }

      printDialog = new PrintDialog();
      if (printDialog.ShowDialog().GetValueOrDefault() != true) return;
      printDialog.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape;
      printDialog.PrintVisual(this.cnv2, "Print Defects Map2");
    }

    public bool CanPrintMapDefects()
    {
      return cnv1.Children.Count > 2; 
    }
    #endregion
  }
}
