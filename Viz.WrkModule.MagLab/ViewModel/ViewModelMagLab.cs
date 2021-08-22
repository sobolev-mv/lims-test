using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using DevExpress.XtraEditors.DXErrorProvider;
using Smv.MVVM.Commands;
using System.Windows.Input;
using System.Windows;
using System.Windows.Media;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;
using Smv.MVVM.ViewModels;
using Smv.Utils;
using Viz.DbApp.Psi;
using Viz.MagLab.MeasureUnits;
using Viz.WrkModule.MagLab.Db;
using Viz.WrkModule.MagLab.Db.DataSets;
using Viz.WrkModule.MagLab.View;


namespace Viz.WrkModule.MagLab
{
  internal sealed class ViewModelMagLab : ViewModelBase
  {

    #region Fields
    private Boolean applyCorrectCoeff = true;
    private uint qntIsolPointsPerSide = 3;

    private readonly UserControl usrControl;
    private readonly DsMgLab dsMagLab = new DsMgLab();
    private int currentMeasureUnit = 1;
    private int selectedMeasureDevice;

    private Int32 isQm;
    private string currentSampleId = null;
    private string currentTstep = null;
    private string currentProbeId = null;
    private string currentSteelType = null;
    private string currentMd = null;
    private DataRow currentSampleDataRow = null; 
    private DateTime dateSamplesBegin;
    private DateTime dateSamplesEnd;
    private Int32 findSampleMode = 1;
    private Int32 selectedShift = 3;
    private String findSampleText = null;
    private Int32 sampleState = -1;
    private GridControl gcSampleData;
    private GridControl gcProbeData;
    private GridControl gcSamples;

    //private DevExpress.Xpf.Bars.BarEditItem beiFind;

    private DataTable findModeTable;
    private GridColumn[] samplDataGridColArr;
    private GridColumn[] probeDataGridColArr;

    public ObservableCollection<bool> unitEnable = new ObservableCollection<bool>{false, false, false, false,false};
    private ObservableCollection<Brush> unitColor = new ObservableCollection<Brush> {Brushes.Black, Brushes.Black, Brushes.Black, Brushes.Black, Brushes.Black};      
    private ObservableCollection<int?> kesiValue = new ObservableCollection<int?> {null,null,null,null,null,null};

    private bool accCmdViewSamples;
    private bool accCmdFindSamples;
    private bool accCmdSaveChangedData;
    private bool accCmdMeasureSample;
    private bool accCmdSampleToMes;
    private bool accCmdSampleToMeasure;
    private bool accCmdCopyProp;
    private bool accCmdViewPropProbe;
    private bool accCmdEditSample;
    private bool accCmdEditProbe;
    private bool accCmdDeleteSample;
    private bool accCmdValidateSample;
    private bool accCmdChangeStatFlag;
    private bool accCmdCheckS2L;
    private bool accCmdUnCheckS2L;
    private bool accCmdCopyPropS2L;
    private bool accCmdFr20Import;
    private bool accCmdStZap;
    private bool accCmdMesurCof;
    private bool accCmdCopyApstProp;
    private bool accCmdSiemensSample;
    private bool accCmdAdgRpt;
    private bool accCmdCopyApstCmpCoil;

    #endregion Fields

    #region Public Property

    public ObservableCollection<Brush> UnitColor
    {
      get{ return unitColor; }
      set{
        if (value == unitColor) return;
        unitColor = value;
        base.OnPropertyChanged("UnitColor");
      }
    }

    public ObservableCollection<bool> UnitEnable
    {
      get{ return unitEnable; }
      set{
        if (value == unitEnable) return;
        unitEnable = value;
        base.OnPropertyChanged("UnitEnable");
      }
    }

    public ObservableCollection<int?> KesiValue
    {
      get{ return kesiValue; }
      set{
        if (value == kesiValue) return;
        kesiValue = value;
        base.OnPropertyChanged("KesiValue");
      }
    }

    public DataTable DeviceLst
    {
      get { return dsMagLab.MlDeviceLst; }
    }

    public DataTable Samples
    {
      get{ return dsMagLab.MlSamples;}
    }

    public DataTable FindMode
    {
      get{ return findModeTable; }
    }

    public DataTable Shift
    {
      get{ return dsMagLab.MlShift; }
    }

    public DataTable SamplesData
    {
      get{ return dsMagLab.MlData; }
    }

    public DataTable ProbeData
    {
      get{ return dsMagLab.MlDataProbe; }
    }

    public DataTable ListApInfoData
    {
      get{ return dsMagLab.MlListApInfo; }
    }


    public DateTime DateSamplesBegin
    {
      get{ return dateSamplesBegin; }
      set{
        if (value == dateSamplesBegin) return;
        dateSamplesBegin = value;
        base.OnPropertyChanged("DateSamplesBegin");
      }
    }

    public DateTime DateSamplesEnd
    {
      get{ return dateSamplesEnd; }
      set{
        if (value == dateSamplesEnd) return;
        dateSamplesEnd = value;
        base.OnPropertyChanged("DateSamplesEnd");
      }
    }

    public Int32 FindSampleMode
    {
      get{ return findSampleMode; }
      set{
        if (value == findSampleMode) return;
        findSampleMode = value;
        base.OnPropertyChanged("FindSampleMode");
      }
    }

    public Int32 SelectedShift
    {
      get{ return selectedShift; }
      set{
        if (value == selectedShift) return;
        selectedShift = value;
        base.OnPropertyChanged("SelectedShift");
      }
    }


    public String FindSampleText
    {
      get{ return findSampleText; }
      set{
        if (value == findSampleText) return;
        findSampleText = value;
        base.OnPropertyChanged("FindSampleText");
      }
    }
    
    public Int32 SelectedMeasureDevice
    {
      get { return selectedMeasureDevice; }
      set
      {
        if (value == selectedMeasureDevice) return;
        selectedMeasureDevice = value;
        base.OnPropertyChanged("SelectedMeasureDevice");
      }
    }

    public GridControl GcSampleData
    {
      get{ return gcSampleData; }
      set{
        if (value == gcSampleData) return;
        gcSampleData = value;
        base.OnPropertyChanged("GcSampleData");
      }
    }

    #endregion Public Property

    #region Private Method

    private void Column_MlDataChanging(object sender, DataColumnChangeEventArgs e)
    {
      if (!applyCorrectCoeff)
        return;

      if (e.ProposedValue == null)
        e.ProposedValue = DBNull.Value;

      dsMagLab.MlMesurCof.DefaultView.ApplyDefaultSort = true;
      int i = dsMagLab.MlMesurCof.DefaultView.Find(new Object[] { this.currentMd, e.Column.ColumnName, (sender as MlDataDataTable).Utype, (sender as MlDataDataTable).MesDevice });

      if ((i != -1) && (e.ProposedValue != DBNull.Value) && (Convert.ToChar(dsMagLab.MlMesurCof.DefaultView[i]["TypCor"]) == (char)MlTypeCorrect.General))
        //MessageBox.Show("Найдено!");
        //Делаем корректировку если параметр Тип корректировки = 'G' (Общая корректирвка, делается в здесь)
        e.ProposedValue = Convert.ToDecimal(e.ProposedValue) + Convert.ToDecimal(dsMagLab.MlMesurCof.DefaultView[i]["Corr"]);
    }

    private void CreateFindModeTable()
    {
      findModeTable = new DataTable();
      findModeTable.Columns.Add("Id", Type.GetType("System.Int32"));
      findModeTable.Columns.Add("NameMode", Type.GetType("System.String"));

      DataRow row = findModeTable.NewRow();
      row[0] = 1;
      row[1] = "По № образца";
      findModeTable.Rows.Add(row);

      row = findModeTable.NewRow();
      row[0] = 2;
      row[1] = "По лок. № материала";
      findModeTable.Rows.Add(row);

      row = findModeTable.NewRow();
      row[0] = 3;
      row[1] = "По маркировке";
      findModeTable.Rows.Add(row);

      row = findModeTable.NewRow();
      row[0] = 4;
      row[1] = "По лок. № материала + дети";
      findModeTable.Rows.Add(row);

      findModeTable.AcceptChanges();
    }

    private void ValidateValueCell(object sender, GridCellValidationEventArgs e)
    {
      int IsValidate = 0;
      decimal MinVal = 0;
      decimal MaxVal = 0;
      int tag = Convert.ToInt32((sender as GridColumn).Tag);
      Boolean IsSample = (Convert.ToInt32((e.Column.View as GridViewBase).Grid.Tag) == 1);

      if (IsSample)
        dsMagLab.MlValData.DefaultView.RowFilter = "SteelType = '" + currentSteelType + "'" + " AND Utype = " + currentMeasureUnit.ToString();
      else
        dsMagLab.MlValData.DefaultView.RowFilter = "SteelType = '" + currentSteelType + "'" + " AND Utype = 0";
      
      dsMagLab.MlValData.DefaultView.Sort = "Ftag";
      int i = dsMagLab.MlValData.DefaultView.Find(tag);

      if (i == -1){
        e.IsValid = false;
        e.ErrorType = ErrorType.Critical;
        e.ErrorContent = "Ошибки в настройке валидации";
        return;        
      }  
      else{
        IsValidate = Convert.ToInt32(dsMagLab.MlValData.DefaultView[i]["IsValidate"]);
        MinVal = Convert.ToDecimal(dsMagLab.MlValData.DefaultView[i]["MinVal"]);
        MaxVal = Convert.ToDecimal(dsMagLab.MlValData.DefaultView[i]["MaxVal"]);
      }

      if ((IsValidate == 0) || (e.Value == null)) {
        e.IsValid = true;
        return;
      }

      decimal Val = Convert.ToDecimal(e.Value);
      Boolean IsVal = ((Val >= MinVal) && (Val <= MaxVal));

      if (!IsVal){
        e.IsValid = false;
        e.ErrorType = ErrorType.Critical;
        e.ErrorContent = "Значение параметра должно быть в границах:\n " +
                         MinVal.ToString() + " и " + MaxVal.ToString();
      }
         
    }

    /* 
    private void SetProbeDataGrid2(String SimpleId, String ProbeId, int SampleState)
    {
      gcSampleData.View.AllowEditing = ((SampleState == 0) && accCmdEditSample);
      gcProbeData.View.AllowEditing = ((SampleState == 0) && accCmdEditProbe);
      
      
      if (string.IsNullOrEmpty(SimpleId)){
        foreach (GridColumn gc in gcSampleData.Columns)
          gc.Visible = false;
        return;
      }
      
      foreach (GridColumn gc in gcProbeData.Columns)
        gc.Visible = false;
      
                 
      dsMagLab.MlUset.DefaultView.RowFilter = "IsSample = 0";
      dsMagLab.MlUset.DefaultView.Sort = "Ftag";
      
            
      if (dsMagLab.MlUset.DefaultView.Count != 0){
        foreach (GridColumn gc in gcProbeData.Columns){
          int tagCol = Convert.ToInt32(gc.Tag);
          int i = dsMagLab.MlUset.DefaultView.Find(tagCol);
          gc.Visible = (i != -1);
        }
        dsMagLab.MlDataProbe.LoadData(ProbeId, currentTstep);
      } else
          dsMagLab.MlDataProbe.Clear();       
    }
    */
    private void SetSampleDataGrid(String SimpleId, int UnitType, int SampleState)
    {
      
      dsMagLab.MlData.Clear();
      //foreach (DataColumn dc in dsMagLab.MlData.Columns)
        //dc.Caption = UnitType.ToString();
      dsMagLab.MlData.Utype = UnitType;
      dsMagLab.MlData.MesDevice = (int) MlMeasureDevice.Ui5099;

      gcSampleData.View.AllowEditing = ((SampleState == 0) && accCmdEditSample && (isQm == 0));
      gcSampleData.Columns.Clear();
      
      dsMagLab.MlUset.DefaultView.RowFilter = "IsSample = 1 AND Utype = " + UnitType.ToString();
      dsMagLab.MlUset.DefaultView.Sort = "Ftag";
      
      gcSampleData.Columns.BeginUpdate();
      
      int vi = 0;  
      foreach (DataRowView drv in dsMagLab.MlUset.DefaultView){
        int tagColRow = Convert.ToInt32(drv.Row["Ftag"]);

        foreach (GridColumn gc in samplDataGridColArr){
          int tagCol = Convert.ToInt32(gc.Tag);

          if (tagCol == tagColRow){
            gc.Visible = false; 
            gc.VisibleIndex = vi;
            gc.Visible = true; 
            gcSampleData.Columns.Add(gc);
            vi++;
          }
        }

      }

      foreach (GridColumn gCol in gcSampleData.Columns)
        gCol.Validate += ValidateValueCell;

      gcSampleData.Columns.EndUpdate();

      dsMagLab.MlData.LoadData(SimpleId, UnitType);
    }

    private void SetProbeDataGrid(String SimpleId, String ProbeId, int SampleState)
    {
      dsMagLab.MlDataProbe.Clear();
      gcProbeData.View.AllowEditing = ((SampleState == 0) && accCmdEditProbe && (isQm == 0));
      gcProbeData.Columns.Clear();
      
      dsMagLab.MlUset.DefaultView.RowFilter = "IsSample = 0";
      dsMagLab.MlUset.DefaultView.Sort = "Ftag";
      
      gcProbeData.Columns.BeginUpdate();
      
      
      int vi = 0;  
      foreach (DataRowView drv in dsMagLab.MlUset.DefaultView){
        int tagColRow = Convert.ToInt32(drv.Row["Ftag"]);

        foreach (GridColumn gc in probeDataGridColArr){
          int tagCol = Convert.ToInt32(gc.Tag);

          if (tagCol == tagColRow){
            gc.Visible = false; 
            gc.VisibleIndex = vi;
            gc.Visible = true;
            gcProbeData.Columns.Add(gc);
            vi++;
          }
        }
      }
      
      /*Для тестовых целей
      foreach (GridColumn gc in probeDataGridColArr){
        gcProbeData.Columns.Add(gc);
        gc.Visible = true;
      }
      */

      gcProbeData.Columns.EndUpdate();
      dsMagLab.MlDataProbe.LoadData(ProbeId, currentTstep);
    }

    /* 
    private void SetSampleDataGrid2(String SimpleId, int UnitType, int SampleState)
    {
      
      gcSampleData.View.AllowEditing = ((SampleState == 0) && accCmdEditSample);
      
                  
      //if (string.IsNullOrEmpty(SimpleId)){
      //  foreach (DevExpress.Xpf.Grid.GridColumn gc in this.gcSampleData.Columns)
      //    gc.Visible = false;
      //  return; 
      //}
        
      foreach (GridColumn gc in gcSampleData.Columns)
        gc.Visible = false;

      dsMagLab.MlData.Clear();
              
      dsMagLab.MlUset.DefaultView.RowFilter = "IsSample = 1 AND Utype = " + UnitType.ToString(); 
      dsMagLab.MlUset.DefaultView.Sort = "Ftag";

      //iPointCount = Convert.ToUInt32(dsMagLab.MlUset.DefaultView.Count);
      //
      //foreach(DevExpress.Xpf.Grid.GridColumn gc in this.gcSampleData.Columns){
      //  int tagCol = Convert.ToInt32(gc.Tag);
      //  int i = dsMagLab.MlUset.DefaultView.Find(tagCol);  
      //  gc.Visible = (i != -1);
      //}

      int iv = 0;
      foreach(DataRowView drv in dsMagLab.MlUset.DefaultView){
        int tagColRow = Convert.ToInt32(drv.Row["Ftag"]);

        foreach (GridColumn gc in gcSampleData.Columns){
          int tagCol = Convert.ToInt32(gc.Tag);
          if (tagCol == tagColRow){
            (gcSampleData.View as GridViewBase).MoveColumnTo(gc, iv, HeaderPresenterType.Headers, HeaderPresenterType.Headers);
            gc.Visible = true;

            //gc.VisibleIndex = iv;

            
            if (iv == 0)
              //(this.gcSampleData.View as GridViewBase).FocusedColumn = gc;
              gcSampleData.CurrentColumn = gc;

            iv++;
            break;
          }
        }

      }
      
      dsMagLab.MlData.LoadData(SimpleId, UnitType);
    }
    */
    private void MeasureIsol(DataRow row)
    {
      List<Boolean> lstVisible = LabAction.GetVisibleElements(currentSampleId, currentMeasureUnit, 1, 50);
      if (lstVisible == null) return;

      /*Временная доделка 3 точки*/

      //Проверяем сколько точек дает настройка 3 или 5      
      Boolean is5 = ((lstVisible[3]) && (lstVisible[4]));

      //В случае 5 точек проверяем включено ли уменьшение до 3 точек
      if ((is5) && (qntIsolPointsPerSide == 3)) {
        lstVisible[3] = false;  
        lstVisible[4] = false;
        lstVisible[8] = false;
        lstVisible[9] = false;
      }
      
      /*Временная доделка 3 точки*/

      var lstMeasureVal = new List<decimal?>();
      if (row["Iup1"] == DBNull.Value)
         lstMeasureVal.Add(null);
      else
         lstMeasureVal.Add((decimal?)row["Iup1"]);

      if (row["Iup2"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Iup2"]);

      if (row["Iup3"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Iup3"]);

      if (row["Iup4"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Iup4"]);

      if (row["Iup5"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Iup5"]);

      if (row["Idown1"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Idown1"]);

      if (row["Idown2"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Idown2"]);

      if (row["Idown3"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Idown3"]);

      if (row["Idown4"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Idown4"]);

      if (row["Idown5"] == DBNull.Value)
        lstMeasureVal.Add(null);
      else
        lstMeasureVal.Add((decimal?)row["Idown5"]);

      Window vIsol = new ViewMeasureIsol(lstMeasureVal, lstVisible);
      bool? dlgRes = vIsol.ShowDialog();
      if ((dlgRes == null) || (dlgRes == false)) return;

      //==================
      row.BeginEdit();
      if (lstVisible[0]) row["Iup1"] = lstMeasureVal[0];
      if (lstVisible[1]) row["Iup2"] = lstMeasureVal[1];
      if (lstVisible[2]) row["Iup3"] = lstMeasureVal[2];
      if (lstVisible[3]) row["Iup4"] = lstMeasureVal[3];
      if (lstVisible[4]) row["Iup5"] = lstMeasureVal[4];
      if (lstVisible[5]) row["Idown1"] = lstMeasureVal[5];
      if (lstVisible[6]) row["Idown2"] = lstMeasureVal[6];
      if (lstVisible[7]) row["Idown3"] = lstMeasureVal[7];
      if (lstVisible[8]) row["Idown4"] = lstMeasureVal[8];
      if (lstVisible[9]) row["Idown5"] = lstMeasureVal[9];

      /*Временная доделка 3 точки*/ 
     
      if ((is5) && (qntIsolPointsPerSide == 3)){

        if ((row["Iup1"] != DBNull.Value) && (row["Iup2"] != DBNull.Value))
          row["Iup4"] = (Convert.ToDecimal(row["Iup1"]) + Convert.ToDecimal(row["Iup2"])) / 2;

        if ((row["Iup2"] != DBNull.Value) && (row["Iup3"] != DBNull.Value))
          row["Iup5"] = (Convert.ToDecimal(row["Iup2"]) + Convert.ToDecimal(row["Iup3"])) / 2;

        if ((row["Idown1"] != DBNull.Value) && (row["Idown2"] != DBNull.Value))
          row["Idown4"] = (Convert.ToDecimal(row["Idown1"]) + Convert.ToDecimal(row["Idown2"])) / 2;

        if ((row["Idown2"] != DBNull.Value) && (row["Idown3"] != DBNull.Value))
          row["Idown5"] = (Convert.ToDecimal(row["Idown2"]) + Convert.ToDecimal(row["Idown3"])) / 2; 
      }
     
      /*Временная доделка 3 точки*/

      row.EndEdit();

      //this.dsMagLab.MlData.SaveData();
      //row.AcceptChanges();
      //row.Table.AcceptChanges();
      dsMagLab.MlData.SaveData();
    }

    private void MeasureListApMpg200D(DataRow SampleRow, int uType)
    {
      Dictionary<string, decimal> resData = new Dictionary<string, decimal>();
      var mesVal = new string[] { "B100", "B800", "B2500", "P1550", "P1750" };

      dsMagLab.MlMpg200d.LoadData(uType);
      var dlgMpg200D = new ViewBrockhausMpg200D(uType, Convert.ToDecimal(SampleRow["ThickNessNominal"]), Convert.ToString(SampleRow["SampleNum"]), dsMagLab.MlMpg200d, resData);

      if (!dlgMpg200D.ShowDialog().GetValueOrDefault())
        return;

      Object obj = gcSampleData.GetFocusedRow();
      if (obj == null) return;
      DataRow rowSampleData = (obj as DataRowView).Row;
      dsMagLab.MlData.MesDevice = (int) MlMeasureDevice.Mpg200D;

      rowSampleData.BeginEdit();

      for (int j = 0; j < mesVal.Length; j++)
      {
        //Здесь происходит корректировка измеренных значений
        dsMagLab.MlMesurCof.DefaultView.ApplyDefaultSort = true;
        int i = dsMagLab.MlMesurCof.DefaultView.Find(new Object[]{SampleRow["Md"], mesVal[j], uType, (int) MlMeasureDevice.Mpg200D});

        if ((i != -1) && (Convert.ToChar(dsMagLab.MlMesurCof.DefaultView[i]["TypCor"]) == 'D'))
          rowSampleData[mesVal[j]] = resData[mesVal[j]] + Convert.ToDecimal(dsMagLab.MlMesurCof.DefaultView[i]["Corr"]);
        else
          rowSampleData[mesVal[j]] = resData[mesVal[j]];
      }

      rowSampleData["Massa"] = resData["Weight"];
      rowSampleData.EndEdit();
      dsMagLab.MlData.SaveData();
      dsMagLab.AcceptChanges();

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Ui5099; //Установка УИ5099
    }

    private void MeasureListApMk4a(DataRow SampleRow, int Utype)
    {
      decimal? sLen = null;
      decimal? sWid = null;
      decimal? sDen = null;

      var ftag = new int[(gcSampleData.View as GridViewBase).VisibleColumns.Count];

      dsMagLab.MlMk4au.LoadData(Convert.ToString(SampleRow["SteelType"]), Convert.ToDecimal(SampleRow["ThickNessNominal"]), Utype);
      if (dsMagLab.MlMk4au.Rows.Count != 0){
        sLen = Convert.ToDecimal(dsMagLab.MlMk4au.Rows[0]["Lsimple"]);
        sWid = Convert.ToDecimal(dsMagLab.MlMk4au.Rows[0]["Wsimple"]);
        sDen = Convert.ToDecimal(dsMagLab.MlMk4au.Rows[0]["Density"]);
      }

      var i = 0;
      foreach (GridColumn gc in (gcSampleData.View as GridViewBase).VisibleColumns){
        ftag[i] = Convert.ToInt32(gc.Tag);
        i++;
      } 
      
      dsMagLab.MlMk4ap.LoadData(Utype, ftag);

      //Что б не было корректировки по Эпштейну
      /*
      foreach (DataColumn dc in dsMagLab.MlData.Columns)
      dc.Caption = ;
      */

      Window vLstAp = new ViewMeasureListAp(Utype, dsMagLab.MlMk4ap, null, sLen, sWid, sDen, currentMd, (int)MlMeasureDevice.Mk4A, dsMagLab.MlMesurCof);
      if (!vLstAp.ShowDialog().GetValueOrDefault())
        return;

      Object obj = gcSampleData.GetFocusedRow();
      if (obj == null) return;
      DataRow rowSampleData = (obj as DataRowView).Row;

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Mk4A; //Установка МК4Э

      rowSampleData.BeginEdit();
      foreach(DataRow row in dsMagLab.MlMk4ap.Rows)
        rowSampleData[Convert.ToString(row["MeasP"])] = row["OutVal"];

      rowSampleData["Massa"] = vLstAp.Tag;
      rowSampleData.EndEdit();
      dsMagLab.MlData.SaveData(); 
      dsMagLab.AcceptChanges();

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Ui5099; //Установка УИ5099
    }

    private void SetPropUnitSelector()
    {
      int state = Convert.ToInt32(currentSampleDataRow["State"]); 
     
      if (state != 0){
        for (int i = 0; i < UnitColor.Count; i++){
          UnitColor[i] = Brushes.Black;
          UnitEnable[i] = true; 
        }  

        return;
      }
 
      dsMagLab.MlUtypeInfo.LoadData(currentSampleId);
      if (dsMagLab.MlUtypeInfo.Rows.Count == 0) return;

      foreach(DataRow row in dsMagLab.MlUtypeInfo.Rows){
        int idx = Convert.ToInt32(row["Utype"]);
        Boolean isEnable = (Convert.ToInt32(row["IsEnable"]) == 1);
        Boolean isFull = (Convert.ToInt32(row["IsFull"]) == 1);

        UnitEnable[idx - 1] = isEnable;
        if (isEnable){
          if (isFull)
            UnitColor[idx - 1] = Brushes.Black;
          else
            UnitColor[idx - 1] = Brushes.Red;
        } else
          UnitColor[idx - 1] = Brushes.Black;
      }
    }     

    private void SaveDataToDb()
    {
      dsMagLab.MlData.SaveData();
      dsMagLab.MlDataProbe.SaveData();

      //Остаточное напряжение берем с образцов этого же стенда Убрано 16.08.2018
      /*
      Boolean isCol = gcProbeData.Columns.Any(col => ((col.FieldName == "OstNapr") && col.Visible));
      if (isCol){
        LabAction.SetOstNaprForStend(currentSampleId);
        dsMagLab.MlDataProbe.LoadData(currentProbeId, currentTstep);
      }
      */

      SetPropUnitSelector();
      dsMagLab.MlListApInfo.LoadData(currentSampleId); 
      GetKesi();
    }

    private void CopyPropToAnotherSample()
    {
      string sample2 = "";
      Boolean drez = ExecDlg.InputQuery("Копирование свойств","Введите № образца получателя", ref sample2, false);
      if (!drez) return;
      if (string.IsNullOrEmpty(sample2.Trim())) return;

      if (!LabAction.CopyPropToAnotherSample(currentSampleId, sample2)) return;
      dsMagLab.MlSamples.SerchBySampleId(sample2);
      DXMessageBox.Show(Application.Current.Windows[0], "Свойство перенесены успешно!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }
    
    private void CopyApstPropToAnotherSample()
    {
      string sample2 = "";
      Boolean drez = ExecDlg.InputQuery("Копирование х-тик Эпштейна", "Введите № образца получателя", ref sample2, false);
      if (!drez) return;
      if (string.IsNullOrEmpty(sample2.Trim())) return;

      if (!LabAction.CopyApstPropToAnotherSample(currentSampleId, sample2)) return;
      dsMagLab.MlSamples.SerchBySampleId(sample2);
      DXMessageBox.Show(Application.Current.Windows[0], "Х-тики Эпштейна перенесены успешно!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void CopyApstCmpCoil()
    {
      int cntSamples = LabAction.GetCntSamples(Convert.ToString(currentSampleDataRow["MatLocalNumber"]), Convert.ToString(currentSampleDataRow["Tstep"]));

      if (cntSamples <= 2){
        DXMessageBox.Show(Application.Current.Windows[0], "У материала два или меньше образца! Он не является составным!","Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }

      if (LabAction.CopyApstProp4CmpCoil(currentSampleId))
        DXMessageBox.Show(Application.Current.Windows[0], "Копирование выполнено!", "Процедура копирования", MessageBoxButton.OK, MessageBoxImage.Information);
    }


    private void GetKesi()
    {
      var lstKesi = LabAction.GetKesi(currentSampleId);
      kesiValue[0] = lstKesi[0];
      kesiValue[1] = lstKesi[1];
      kesiValue[2] = lstKesi[2];
      kesiValue[3] = lstKesi[3];
      kesiValue[4] = lstKesi[4];
      kesiValue[5] = lstKesi[5];
    }

    private void NavigateRowGrid(GridControl grid, string FieldName, string FieldValue)
    {
      //Здесь точное позиционирование на запись в случае поиска по нoмеру образца
      if (grid.VisibleRowCount > 0){
        int rowHandle = grid.View.FocusedRowHandle + 1;

        while (Convert.ToString(grid.GetCellValue(rowHandle, FieldName)) != FieldValue && grid.IsValidRowHandle(rowHandle))
          rowHandle++;

        if (grid.IsValidRowHandle(rowHandle)){
          //this.gcSamples.View.FocusedColumn = grid.Columns["UnitPrice"];
          grid.View.FocusedRowHandle = rowHandle;
        }
      }
    }

    private void DeleteSample()
    {
      string psw = ""; 
      Boolean drez = ExecDlg.InputQuery("Ввод пароля", "Введите пароль для удаления образца", ref psw, true);
      if (!drez) return;
      if (string.IsNullOrEmpty(psw.Trim())) return;
      if (psw != "159"){
       DXMessageBox.Show(Application.Current.Windows[0],"Пароль не верен!","Ошибка",MessageBoxButton.OK,MessageBoxImage.Error);
       return;
      }
      
      MessageBoxResult mbr = DXMessageBox.Show(Application.Current.Windows[0],"Внимание!\r\nТекущий образец будет удален и в ЛИМС и в PSI!\r\nПродолжить?",
                                               "Внимание",MessageBoxButton.OKCancel, MessageBoxImage.Warning);  

      if (mbr == MessageBoxResult.Cancel) return;

      if (LabAction.DeleteSample(currentSampleId)){
        (gcSamples.View as GridViewBase).DeleteRow(gcSamples.View.FocusedRowHandle);
        dsMagLab.MlSamples.AcceptChanges();
        DXMessageBox.Show(Application.Current.Windows[0], "Образец успешно удален!", "Удаление образца", MessageBoxButton.OK, MessageBoxImage.Information); 
      }
    }

    private void ChangeValidationModeSample()
    {
      string psw = "";
      Boolean drez = ExecDlg.InputQuery("Ввод пароля", "Введите пароль для изменения режима проверки заполнения характеристик образца", ref psw, true);
      if (!drez) return;
      if (string.IsNullOrEmpty(psw.Trim())) return;

      if (psw != "159"){
        DXMessageBox.Show(Application.Current.Windows[0], "Пароль не верен!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }

      int CurrValMode = Convert.ToInt32(currentSampleDataRow["IsValidate"]);
      CurrValMode = (CurrValMode == 0) ? 1 : 0;

      if (!LabAction.ChangeValidationMode(currentSampleId, CurrValMode))
        return;

      currentSampleDataRow.BeginEdit();
      currentSampleDataRow["IsValidate"] = CurrValMode;
      currentSampleDataRow.EndEdit();
      currentSampleDataRow.AcceptChanges();
      currentSampleDataRow.Table.AcceptChanges();
      //DevExpress.Xpf.Core.DXMessageBox.Show(Application.Current.Windows[0], "Образец переведен в состояние редактирования.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void ChangeStatFlag()
    {
      string psw = "";
      Boolean drez = ExecDlg.InputQuery("Ввод пароля", "Введите пароль для изменения признака статистики у рулона", ref psw, true);
      if (!drez) return;
      if (string.IsNullOrEmpty(psw.Trim())) return;

      if (psw != "999"){
        DXMessageBox.Show(Application.Current.Windows[0], "Пароль не верен!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }

      int CurrStatFlag = Convert.ToInt32(currentSampleDataRow["StFlg"]);
      CurrStatFlag = (CurrStatFlag == 0) ? 1 : 0;
     

      if (!LabAction.ChangeStatFlag(Convert.ToString(currentSampleDataRow["MatLocalNumber"])))
        return;

      currentSampleDataRow.BeginEdit();
      currentSampleDataRow["StFlg"] = CurrStatFlag;
      currentSampleDataRow.EndEdit();
      currentSampleDataRow.AcceptChanges();
      currentSampleDataRow.Table.AcceptChanges();
    }

    private void CheckS2L()
    {
      string matLocId = "";
      Boolean drez = ExecDlg.InputQuery("Локальный номер материала после АВО", "Введите локальный номер материала", ref matLocId, false);
      if (!drez) return;
      if (string.IsNullOrEmpty(matLocId.Trim())) return;

      if (!LabAction.CheckS2L(matLocId))
        return;
      DXMessageBox.Show(Application.Current.Windows[0], "Материал помечен!", "S2L", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void UnCheckS2L()
    {
      string matLocId = "";
      Boolean drez = ExecDlg.InputQuery("Локальный номер материала после АВО", "Введите локальный номер материала", ref matLocId, false);
      if (!drez) return;
      if (string.IsNullOrEmpty(matLocId.Trim())) return;

      if (!LabAction.UnCheckS2L(matLocId))
        return;
      DXMessageBox.Show(Application.Current.Windows[0], "Пометка снята!", "S2L", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void CopyPropS2L()
    {
      if (!LabAction.CopyPropS2L(Convert.ToString(currentSampleDataRow["MatLocalNumber"])))
        return;
      DXMessageBox.Show(Application.Current.Windows[0], "Св-ва успешно скопированы!", "S2L", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void CurrentItemSampleChanged(object sender, CurrentItemChangedEventArgs args)
    {
      if (args.NewItem == null)
        return;

      currentSampleDataRow = (args.NewItem as DataRowView).Row; 
      isQm = Convert.ToInt32(currentSampleDataRow["IsQm"]);
      sampleState = Convert.ToInt32(currentSampleDataRow["State"]);
      currentSampleId = Convert.ToString(currentSampleDataRow["SampleId"]);
      currentTstep = Convert.ToString(currentSampleDataRow["Tstep"]);
      currentProbeId = Convert.ToString(currentSampleDataRow["MatLocalNumber"]);
      currentSteelType = Convert.ToString(currentSampleDataRow["SteelType"]);
      currentMd = Convert.ToString(currentSampleDataRow["Md"]);
      
      dsMagLab.MlUset.LoadData(currentSampleId);
      SetSampleDataGrid(currentSampleId, currentMeasureUnit, sampleState);
      SetProbeDataGrid(currentSampleId, currentProbeId, sampleState);
      SetPropUnitSelector();
      
      
      dsMagLab.MlListApInfo.LoadData(currentSampleId);
      GetKesi();
        
    }

    private int GetFirstSelectedMeasureDevice()
    {
      if (dsMagLab.MlDeviceLst.Rows.Count == 0)
        return -1;

      return Convert.ToInt32(dsMagLab.MlDeviceLst.Rows[0]["Id"]);
    }

    #endregion Private Method

    #region Constructor
    internal ViewModelMagLab(UserControl control)
    {
      usrControl = control;
      DateSamplesBegin = DateTime.Today;
      DateSamplesEnd = DateTime.Today;

      gcSampleData = LogicalTreeHelper.FindLogicalNode(usrControl, "GcSampleData") as GridControl;
      gcProbeData = LogicalTreeHelper.FindLogicalNode(usrControl, "GcProbeData") as GridControl;
      gcSamples = LogicalTreeHelper.FindLogicalNode(usrControl, "GcSamples") as GridControl;
      gcSamples.CurrentItemChanged += CurrentItemSampleChanged;

      CreateFindModeTable();
      dsMagLab.MlValData.LoadData();
      dsMagLab.MlMesurCof.LoadData();
      dsMagLab.MlDeviceLst.LoadData(1);
      SelectedMeasureDevice = GetFirstSelectedMeasureDevice();

      /* 
      foreach (DevExpress.Xpf.Grid.GridColumn gc in this.gcSampleData.Columns)
        gc.Visible = true;
      */

      samplDataGridColArr = new GridColumn[gcSampleData.Columns.Count];
      gcSampleData.Columns.CopyTo(samplDataGridColArr,0);
      gcSampleData.Columns.Clear();

      probeDataGridColArr = new GridColumn[gcProbeData.Columns.Count];
      gcProbeData.Columns.CopyTo(probeDataGridColArr, 0);
      gcProbeData.Columns.Clear();

      dsMagLab.MlData.ColumnChanging += Column_MlDataChanging;

      /*
      foreach (DevExpress.Xpf.Grid.GridColumn gCol in this.gcSampleData.Columns)
        gCol.Validate += ValidateValueCell;
      */

      /*    
      foreach (DevExpress.Xpf.Grid.GridColumn gCol in this.gcProbeData.Columns)
        gCol.Validate += ValidateValueCell; 
      */
      accCmdViewSamples = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdViewSamples, ModuleConst.ModuleId);
      accCmdFindSamples = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdFindSamples, ModuleConst.ModuleId);
      accCmdSaveChangedData = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdSaveChangedData, ModuleConst.ModuleId);
      accCmdMeasureSample = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdMeasureSample, ModuleConst.ModuleId);
      accCmdSampleToMes = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdSampleToMes, ModuleConst.ModuleId);
      accCmdSampleToMeasure = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdSampleToMeasure, ModuleConst.ModuleId);
      accCmdCopyProp = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdCopyProp, ModuleConst.ModuleId);
      accCmdViewPropProbe = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdViewPropProbe, ModuleConst.ModuleId);
      accCmdEditSample = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdEditSample, ModuleConst.ModuleId);
      accCmdEditProbe = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdEditProbe, ModuleConst.ModuleId);
      accCmdDeleteSample = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdDeleteSample, ModuleConst.ModuleId);
      accCmdValidateSample = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdValidateSample, ModuleConst.ModuleId);
      accCmdChangeStatFlag = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdChangeStatFlag, ModuleConst.ModuleId);
      accCmdCheckS2L = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdCheckS2L, ModuleConst.ModuleId);
      accCmdUnCheckS2L = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdUnCheckS2L, ModuleConst.ModuleId);
      accCmdCopyPropS2L = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdCopyPropS2L, ModuleConst.ModuleId);
      accCmdFr20Import = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdFr20Import, ModuleConst.ModuleId);
      accCmdStZap = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdStZap, ModuleConst.ModuleId);
      accCmdMesurCof = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdMesurCof, ModuleConst.ModuleId);
      accCmdCopyApstProp = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdCopyApstProp, ModuleConst.ModuleId);
      accCmdSiemensSample = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdSiemensSample, ModuleConst.ModuleId);
      accCmdAdgRpt = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdAdgRpt, ModuleConst.ModuleId);
      accCmdCopyApstCmpCoil = Permission.GetPermissionForModuleUif(ModuleConst.AccCmdCopyApstCmpCoil, ModuleConst.ModuleId); 

      //Читаем параметры конфиг файла МЛ
      try
      {
        string cfg = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.MagLabConfig, "ApplyCorrectCoeff");
        if (!string.IsNullOrEmpty(cfg))
          applyCorrectCoeff = (string.Equals(cfg, "Y", StringComparison.Ordinal));

        cfg = Smv.App.Config.ConfigParam.ReadAppSettingsParamValue(Etc.StartPath + ModuleConst.MagLabConfig, "QntIsolPointsPerSide");
        if (!string.IsNullOrEmpty(cfg))
          qntIsolPointsPerSide = Convert.ToUInt32(cfg);
      }
      catch (Exception){
        DXMessageBox.Show(Application.Current.Windows[0], "Ошибка при чтении конфигурационных параметров МЛ!", "Ошибка конфигурации", MessageBoxButton.OK, MessageBoxImage.Error);
      }

    }
    #endregion Constructor

    #region Commands
    private DelegateCommand<Object> selectUnitCommand;
    private DelegateCommand<Object> getListSamplesCommand;
    private DelegateCommand<Object> findSamplesCommand;
    //private DelegateCommand<Object> samplesRowChangedCommand;
    private DelegateCommand<Object> saveDataCommand;
    private DelegateCommand<Object> measureSampleDataCommand;
    private DelegateCommand<Object> toMesCommand;
    private DelegateCommand<Object> toBackToEditCommand;
    private DelegateCommand<Object> manualSetPropUnitSelectorCommand;
    private DelegateCommand<Object> showCalcValuesCommand;
    private DelegateCommand<Object> copyPropCommand;
    private DelegateCommand<Object> deleteSampleCommand;
    private DelegateCommand<Object> validateCommand;
    private DelegateCommand<Object> fr20ImportCommand;
    private DelegateCommand<Object> changeStatFlagCommand;
    private DelegateCommand<Object> checkS2LCommand;
    private DelegateCommand<Object> uncheckS2LCommand;
    private DelegateCommand<Object> copyPropS2LCommand;
    private DelegateCommand<Object> dlgStZapCommand;
    private DelegateCommand<Object> dlgMesurCofCommand;
    private DelegateCommand<Object> copyApstPropCommand;
    private DelegateCommand<Object> dlgSiemensCommand;
    private DelegateCommand<Object> adgRptCommand;
    private DelegateCommand<Object> copyApstCmpCoilCommand;

    public ICommand SelectUnitCommand
    {
      get{ return selectUnitCommand ?? (selectUnitCommand = new DelegateCommand<Object>(ExecuteSelectUnit, CanExecuteSelectUnit)); }
    }

    private void ExecuteSelectUnit(Object parameter)
    {
      try{
        currentMeasureUnit = Convert.ToInt32(parameter);
        SetSampleDataGrid(currentSampleId, currentMeasureUnit, sampleState);

        if ((currentMeasureUnit == 1) | (currentMeasureUnit == 2)){
          dsMagLab.MlDeviceLst.LoadData(currentMeasureUnit);
          SelectedMeasureDevice = GetFirstSelectedMeasureDevice();
        }
        else{
          dsMagLab.MlDeviceLst.Rows.Clear();
          SelectedMeasureDevice = -1;
        }

      }
      catch (Exception e){
        MessageBox.Show(e.Message + "\r\n" + e.Source + "r\n" + e.StackTrace);
      }
    }

    private bool CanExecuteSelectUnit(Object parameter)
    {
      return (dsMagLab.MlSamples.Rows.Count != 0);
    }

    public ICommand GetListSamplesCommand
    {
      get{return getListSamplesCommand ?? (getListSamplesCommand = new DelegateCommand<Object>(ExecuteGetListSamples, CanExecuteGetListSamples));}
    }

    private void ExecuteGetListSamples(Object parameter)
    {
      dsMagLab.MlData.Clear();
      dsMagLab.MlDataProbe.Clear();
      dsMagLab.MlSamples.GetListSample(dateSamplesBegin, dateSamplesEnd, selectedShift);
      //if (dsMagLab.MlSamples.RecordCount == 0)
        //dsMagLab.MlData.Clear();
    }

    private bool CanExecuteGetListSamples(Object parameter)
    {
      return accCmdViewSamples;
    }
        
    public ICommand FindSamplesCommand
    {
      get{ return findSamplesCommand ?? (findSamplesCommand = new DelegateCommand<Object>(ExecuteFindSamples, CanExecuteFindSamples));}
    }

    private void ExecuteFindSamples(Object parameter)
    {
      switch (findSampleMode){
          case 1:
            dsMagLab.MlSamples.SerchBySampleId(findSampleText);
            NavigateRowGrid(gcSamples, "SampleNum", findSampleText);
            break;
          case 2:
            dsMagLab.MlSamples.SerchByMatLocalNum(findSampleText);
            break;
          case 3:
            dsMagLab.MlSamples.SerchByMatMarkNum(findSampleText);
            break;
          default:
             break;
        }
        
        if (dsMagLab.MlSamples.Rows.Count == 0){
          dsMagLab.MlData.Clear();
          dsMagLab.MlDataProbe.Clear();

          foreach (GridColumn gc in gcProbeData.Columns)
            gc.Visible = false;

          foreach (GridColumn gc in gcSampleData.Columns)
            gc.Visible = false;  
        }
    }

    private bool CanExecuteFindSamples(Object parameter)
    {
      return ((!String.IsNullOrEmpty(findSampleText)) && accCmdFindSamples);
    }
 
    /*
    public ICommand SamplesRowChangedCommand
    {
      get{ return samplesRowChangedCommand ?? (samplesRowChangedCommand = new DelegateCommand<Object>(ExecuteSamplesRowChanged, CanExecuteSamplesRowChanged));}
    }

    private void ExecuteSamplesRowChanged(Object parameter)
    {
      currentSampleDataRow = (parameter as DataRowView).Row;
      sampleState = Convert.ToInt32(currentSampleDataRow["State"]);
      currentSampleId = Convert.ToString(currentSampleDataRow["SampleId"]);
      currentTstep = Convert.ToString(currentSampleDataRow["Tstep"]);
      currentProbeId = Convert.ToString(currentSampleDataRow["MatLocalNumber"]);
      currentSteelType = Convert.ToString(currentSampleDataRow["SteelType"]);
      dsMagLab.MlUset.LoadData(currentSampleId);
      SetSampleDataGrid(currentSampleId, currentMeasureUnit, sampleState);
      SetProbeDataGrid(currentSampleId, currentProbeId, sampleState);
      SetPropUnitSelector();

      dsMagLab.MlListApInfo.LoadData(currentSampleId);
      GetKesi(); 
    }

    private bool CanExecuteSamplesRowChanged(Object parameter)
    {
      return true;
    }
    */

    public ICommand SaveDataCommand
    {
      get{ return saveDataCommand ?? (saveDataCommand = new DelegateCommand<Object>(ExecuteSaveData, CanExecuteSaveData));}
    }

    private void ExecuteSaveData(Object parameter)
    {
      SaveDataToDb();
    }

    private bool CanExecuteSaveData(Object parameter)
    {
      return (dsMagLab.HasChanges() && accCmdSaveChangedData && (isQm == 0));
    }


    public ICommand MeasureSampleDataCommand
    {
      get{ return measureSampleDataCommand ?? (measureSampleDataCommand = new DelegateCommand<Object>(ExecuteMeasureSampleData, CanExecuteMeasureSampleData));}
    }

    private void ExecuteMeasureSampleData(Object parameter)
    {
      Object obj = gcSampleData.GetFocusedRow();
      if (obj == null) 
        return;
      DataRow row = (obj as DataRowView).Row;

      switch(currentMeasureUnit){
        case 1:
          if (SelectedMeasureDevice == (int)MlMeasureDevice.Mk4A)
            MeasureListApMk4a(currentSampleDataRow, currentMeasureUnit);
          else if (SelectedMeasureDevice == (int)MlMeasureDevice.Mpg200D)
            MeasureListApMpg200D(currentSampleDataRow, currentMeasureUnit);
          else
            DXMessageBox.Show(Application.Current.Windows[0], "Для выбранного устройства возможен только ручной ввод данных.", "Измерения", MessageBoxButton.OK, MessageBoxImage.Information);

          break;
        case 2:
          if (SelectedMeasureDevice == (int)MlMeasureDevice.Mk4A)
            MeasureListApMk4a(currentSampleDataRow, currentMeasureUnit);
          else if (SelectedMeasureDevice == (int)MlMeasureDevice.Mpg200D)
            MeasureListApMpg200D(currentSampleDataRow, currentMeasureUnit);
          else
            DXMessageBox.Show(Application.Current.Windows[0], "Для выбранного устройства возможен только ручной ввод данных.", "Измерения", MessageBoxButton.OK, MessageBoxImage.Information);

          break;
        case 3:
          MeasureIsol(row);
          break;
        default:
          break;
      }
    }

    private bool CanExecuteMeasureSampleData(Object parameter)
    {
      Boolean b = ((currentMeasureUnit == 1) | (currentMeasureUnit == 2) | (currentMeasureUnit == 3));
      return (b && (sampleState == 0) && accCmdMeasureSample && (isQm == 0));
    }

    public ICommand ToMesCommand
    {
      get{ return toMesCommand ?? (toMesCommand = new DelegateCommand<Object>(ExecuteToMes, CanExecuteToMes));}
    }

    private void ExecuteToMes(Object parameter)
    {
      MessageBoxResult mbRes;
      var tstType = Convert.ToString(currentSampleDataRow["TestType"]);

      //Здесь происходит контроль толщины
      var chkThickness = LabAction.CheckThickness(currentSampleId, true);

      if (string.IsNullOrEmpty(chkThickness))

        DXMessageBox.Show(Application.Current.Windows[0], "Ошибка в функции контроля толщины.\r\nСообщите разработчикам ПО.", "Контроль толщины", MessageBoxButton.OK, MessageBoxImage.Warning);

      else if (chkThickness == "N"){

        mbRes = DXMessageBox.Show(Application.Current.Windows[0], "Внимание!\r\nВыявлено отклонение по толщине.\r\nПродолжить отправку х-тик образца в MES?", "Контроль толщины", MessageBoxButton.OK, MessageBoxImage.Error);

        //if (mbRes == MessageBoxResult.No)
          return;
      }

      //Здесь начало отправеи образца в MES
      if (tstType == "ОП"){

        if (!LabAction.MatToMes(currentSampleId, currentTstep))
          return;
      }
      else{

        if (!LabAction.SampleToMes(currentSampleId, currentTstep))
          return;
      }

      /*
      this.currentSampleDataRow.BeginEdit();
      this.currentSampleDataRow["State"] = 10;
      this.currentSampleDataRow.EndEdit();
      this.currentSampleDataRow.AcceptChanges();
      this.currentSampleDataRow.Table.AcceptChanges();
      */


      dsMagLab.MlSamples.SerchByMatLocalNum(currentProbeId);
      DXMessageBox.Show(Application.Current.Windows[0], "Характеристики образца поставлены в очередь отправки в MES.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private bool CanExecuteToMes(Object parameter)
    {
      return ((sampleState == 0) && (dsMagLab.MlSamples.Rows.Count > 0) && accCmdSampleToMes && (isQm == 0));
    }

    public ICommand ToBackToEditCommand
    {
      get{ return toBackToEditCommand ?? (toBackToEditCommand = new DelegateCommand<Object>(ExecuteToBackToEdit, CanExecuteToBackToEdit));}
    }

    private void ExecuteToBackToEdit(Object parameter)
    {
      if (!LabAction.SampleToEdit(currentSampleId))
        return;
       
      currentSampleDataRow.BeginEdit();
      currentSampleDataRow["State"] = 0;
      currentSampleDataRow.EndEdit();
      currentSampleDataRow.AcceptChanges();
      currentSampleDataRow.Table.AcceptChanges();

      DXMessageBox.Show(Application.Current.Windows[0], "Образец переведен в состояние редактирования.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private bool CanExecuteToBackToEdit(Object parameter)
    {
      return ((sampleState == 20) && (dsMagLab.MlSamples.Rows.Count > 0) && accCmdSampleToMeasure && (isQm == 0));
    }

    public ICommand ManualSetPropUnitSelectorCommand
    {
      get{ return manualSetPropUnitSelectorCommand ?? (manualSetPropUnitSelectorCommand = new DelegateCommand<Object>(ExecuteManualSetPropUnitSelector, CanExecuteManualSetPropUnitSelector));}
    }

    private void ExecuteManualSetPropUnitSelector(Object parameter)
    {
      SetPropUnitSelector(); 
      if (dsMagLab.HasChanges())
        SaveDataToDb();
    }

    private bool CanExecuteManualSetPropUnitSelector(Object parameter)
    {
      return (dsMagLab.MlSamples.Rows.Count > 0);
    }

    public ICommand ShowCalcValuesCommand
    {
      get{ return showCalcValuesCommand ?? (showCalcValuesCommand = new DelegateCommand<Object>(ExecuteShowCalcValues, CanExecuteShowCalcValues));}
    }

    private void ExecuteShowCalcValues(Object parameter)
    {
      var dlgVal = new ViewSampleProp(currentProbeId, dsMagLab.MlSamples);
      dlgVal.ShowDialog();
    }

    private bool CanExecuteShowCalcValues(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && (sampleState == 20) && accCmdViewPropProbe);
    }

    public ICommand CopyPropCommand
    {
      get{ return copyPropCommand ?? (copyPropCommand = new DelegateCommand<Object>(ExecuteCopyProp, CanExecuteCopyProp));}
    }

    private void ExecuteCopyProp(Object parameter)
    {
      CopyPropToAnotherSample();  
    }

    private bool CanExecuteCopyProp(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && accCmdCopyProp);
    }
    
    public ICommand CopyApstPropCommand
    {
      get { return copyApstPropCommand ?? (copyApstPropCommand = new DelegateCommand<Object>(ExecuteCopyApstProp, CanExecuteCopyApstProp)); }
    }

    private void ExecuteCopyApstProp(Object parameter)
    {
      CopyApstPropToAnotherSample();
    }

    private bool CanExecuteCopyApstProp(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && accCmdCopyApstProp);
    }
    
    public ICommand CopyApstCmpCoilCommand
    {
      get { return copyApstCmpCoilCommand ?? (copyApstCmpCoilCommand = new DelegateCommand<Object>(ExecuteCopyApstCmpCoil, CanExecuteCopyApstCmpCoil)); }
    }

    private void ExecuteCopyApstCmpCoil(Object parameter)
    {
      CopyApstCmpCoil();
    }

    private bool CanExecuteCopyApstCmpCoil(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && (sampleState == 0) && (accCmdCopyApstCmpCoil));
    }
    
    public ICommand DeleteSampleCommand
    {
      get{ return deleteSampleCommand ?? (deleteSampleCommand = new DelegateCommand<Object>(ExecuteDeleteSample, CanExecuteDeleteSample));}
    }

    private void ExecuteDeleteSample(Object parameter)
    {
      DeleteSample();
    }

    private bool CanExecuteDeleteSample(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && (accCmdDeleteSample) && (isQm == 0));
    }

    public ICommand ValidateCommand
    {
      get{ return validateCommand ?? (validateCommand = new DelegateCommand<Object>(ExecuteValidate, CanExecuteValidate));}
    }

    private void ExecuteValidate(Object parameter)
    {
      ChangeValidationModeSample();
    }

    private bool CanExecuteValidate(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && (sampleState == 0) && (accCmdValidateSample));
    }

    public ICommand Fr20ImportCommand
    {
      get { return fr20ImportCommand ?? (fr20ImportCommand = new DelegateCommand<Object>(ExecuteFr20Import, CanExecuteFr20Import)); }
    }

    private void ExecuteFr20Import(Object parameter)
    {
      //var dlg = new View.ViewMatGnl(dsMagLab);
      //dlg.ShowDialog();      
      LabAction.Fr20DataImport(Etc.StartPath + ModuleConst.Fr20UsbIsolMeasureUnitConfig);
    }

    private bool CanExecuteFr20Import(Object parameter)
    {
      return accCmdFr20Import;
    }

    public ICommand ChangeStatFlagCommand
    {
      get { return changeStatFlagCommand ?? (changeStatFlagCommand = new DelegateCommand<Object>(ExecuteChangeStatFlag, CanExecuteChangeStatFlag)); }
    }

    private void ExecuteChangeStatFlag(Object parameter)
    {
      ChangeStatFlag();
    }

    private bool CanExecuteChangeStatFlag(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && (accCmdChangeStatFlag));
    }

    public ICommand CheckS2LCommand
    {
      get { return checkS2LCommand ?? (checkS2LCommand = new DelegateCommand<Object>(ExecuteCheckS2L, CanExecuteCheckS2L)); }
    }

    private void ExecuteCheckS2L(Object parameter)
    {
      CheckS2L();
    }

    private bool CanExecuteCheckS2L(Object parameter)
    {
      return (accCmdCheckS2L);
    }

    public ICommand UnCheckS2LCommand
    {
      get { return uncheckS2LCommand ?? (uncheckS2LCommand = new DelegateCommand<Object>(ExecuteUnCheckS2L, CanExecuteUnCheckS2L)); }
    }

    private void ExecuteUnCheckS2L(Object parameter)
    {
      UnCheckS2L();
    }

    private bool CanExecuteUnCheckS2L(Object parameter)
    {
      return (accCmdUnCheckS2L);
    }

    public ICommand CopyPropS2LCommand
    {
      get { return copyPropS2LCommand ?? (copyPropS2LCommand = new DelegateCommand<Object>(ExecuteCopyPropS2L, CanExecuteCopyPropS2L)); }
    }

    private void ExecuteCopyPropS2L(Object parameter)
    {
      CopyPropS2L();
    }

    private bool CanExecuteCopyPropS2L(Object parameter)
    {
      return ((dsMagLab.MlSamples.Rows.Count > 0) && (Convert.ToInt32(currentSampleDataRow["State"]) == 0) && (Convert.ToInt32(currentSampleDataRow["SlFlg"]) == 2) && accCmdCopyPropS2L);
    }

    public ICommand DlgStZapCommand
    {
      get { return dlgStZapCommand ?? (dlgStZapCommand = new DelegateCommand<Object>(ExecuteDlgStZap, CanExecuteDlgStZap)); }
    }

    private void ExecuteDlgStZap(Object parameter)
    {
      var dlg = new ViewDlgStZap(dsMagLab);
      dlg.ShowDialog();
    }

    private bool CanExecuteDlgStZap(Object parameter)
    {
      return (accCmdStZap);
    }
    
    public ICommand DlgMesurCofCommand
    {
      get { return dlgMesurCofCommand ?? (dlgMesurCofCommand = new DelegateCommand<Object>(ExecuteDlgMesurCof, CanExecuteDlgMesurCof)); }
    }

    private void ExecuteDlgMesurCof(Object parameter)
    {
      var dlg = new ViewDlgMesurCof(dsMagLab);
      dlg.ShowDialog();
    }

    private bool CanExecuteDlgMesurCof(Object parameter)
    {
      return (accCmdMesurCof);
    }
    
    public ICommand DlgSiemensCommand
    {
      get { return dlgSiemensCommand ?? (dlgSiemensCommand = new DelegateCommand<Object>(ExecuteDlgSiemens, CanExecuteDlgSiemens)); }
    }

    private void ExecuteDlgSiemens(Object parameter)
    {
      var dlg = new ViewDlgSiemensSample(dsMagLab);
      dlg.ShowDialog();
    }

    private bool CanExecuteDlgSiemens(Object parameter)
    {
      return accCmdSiemensSample;
    }
     
    public ICommand AdgRptCommand
    {
      get { return adgRptCommand ?? (adgRptCommand = new DelegateCommand<Object>(ExecuteAdgRpt, CanExecuteAdgRpt)); }
    }

    private void ExecuteAdgRpt(Object parameter)
    {
      LabAction.AdgRptRpt(DateSamplesBegin, DateSamplesEnd, Etc.StartPath + ModuleConst.AdgRptScript1, Etc.StartPath + ModuleConst.AdgRptScript2, Etc.StartPath + ModuleConst.AdgRptScript3);
    }

    private bool CanExecuteAdgRpt(Object parameter)
    {
      return accCmdAdgRpt;
    }




    #endregion Commands

  }
}
