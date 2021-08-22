using System;
using System.Data;
using System.Windows.Controls;
using System.Windows;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Grid;
using Microsoft.Win32;
using Smv.Utils;
using Viz.WrkModule.MagLab.Db;
using Viz.WrkModule.MagLab.Db.DataSets;
using Viz.MagLab.MeasureUnits;
using System.Collections.Generic;

namespace Viz.WrkModule.MagLab
{
  public class ViewModelDlgSiemensSample
  {
    #region Fields
    private readonly Control view;
    private readonly DsMgLab dsMagLab;
    private readonly GridControl dbgSiemensSmp;
    private DataRow currentSmpDataRow = null;
    private Window oldParentWindow;
    #endregion Fields

    #region Public Property
    public virtual DateTime DateFrom { get; set; } = DateTime.Today;
    public virtual DateTime DateTo { get; set; } = DateTime.Today;
    public virtual DataTable Shift => dsMagLab.MlShift;
    public virtual Int32 SelectedShift { get; set; } = 3;
    public virtual Int32 SelectedMeasureDevice { get; set; } = 3;
    public DataTable DeviceLst => dsMagLab.MlDeviceLst;
    public virtual DataTable SiemensSmp => dsMagLab.MlSiemensSmp;
    public virtual string FindLocNum { get; set; }
    #endregion Public Property

    #region Private Method
    private void CurrentItemChanged(object sender, CurrentItemChangedEventArgs args)
    {
      currentSmpDataRow = (args.NewItem as DataRowView)?.Row;
      if (currentSmpDataRow is null)
        return;

      dbgSiemensSmp.View.AllowEditing = (Convert.ToInt32(currentSmpDataRow["State"]) == 0);
    }

    private void CreateSmp4LocNum()
    {
      this.dsMagLab.MlSiemensSmp.SearchByLocNum(FindLocNum);
      if (this.dsMagLab.MlSiemensSmp.Rows.Count > 0)
        return;

      LabAction.CreateSiemensSamples(FindLocNum);
      this.dsMagLab.MlSiemensSmp.SearchByLocNum(FindLocNum);
    }

    private void DeleteSimensSample()
    {
      string psw = "";
      Boolean drez = ExecDlg.InputQuery("Ввод пароля", "Введите пароль для удаления образца", ref psw, true);

      if (!drez) return;

      if (string.IsNullOrEmpty(psw.Trim())) return;

      if (psw != "159"){
        DXMessageBox.Show((view as Window), "Пароль не верен!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }

      MessageBoxResult mbr = DXMessageBox.Show((view as Window), "Внимание!\r\nТекущий образец будет удален!\r\nПродолжить?", "Внимание", MessageBoxButton.OKCancel, MessageBoxImage.Warning);

      if (mbr == MessageBoxResult.Cancel) return;

      if (!LabAction.ChangeSimensSampleState(Convert.ToInt64(currentSmpDataRow["Id"]), SampleState.Deleted))
        return;

      (dbgSiemensSmp.View as GridViewBase)?.DeleteRow(dbgSiemensSmp.View.FocusedRowHandle);
      dsMagLab.MlSiemensSmp.AcceptChanges();
      DXMessageBox.Show((view as Window), "Образец успешно удален!", "Удаление образца", MessageBoxButton.OK,MessageBoxImage.Information);
    }

    private void ChangeSimensSampleState(SampleState state)
    {
      //Convert.ToInt32(currentSmpDataRow["State"])
      currentSmpDataRow.BeginEdit();
      currentSmpDataRow["State"] = state;

      if (state == SampleState.Closed)
      {
        var vTmp = LabAction.GetSimensSampleDpp1750(Convert.ToInt64(currentSmpDataRow["Id"]));
        currentSmpDataRow["Dpp1750"] = vTmp ?? Convert.DBNull;
      }

      currentSmpDataRow.EndEdit();
      currentSmpDataRow.AcceptChanges();
      currentSmpDataRow.Table.AcceptChanges();
     
      LabAction.ChangeSimensSampleState(Convert.ToInt64(currentSmpDataRow["Id"]), state);
      DXMessageBox.Show((view as Window), state == SampleState.Closed ? "Образец отправлен в MES." : "Образец доступен для редактирования.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void MeasureListMpg200D(DataRow smpDataRow)
    {
      const int uType = 1;
      
      Dictionary<string, decimal> resData = new Dictionary<string, decimal>();
      var mesVal = new string[] { "B100", "B800", "B2500", "P1550", "P1750" };

      dsMagLab.MlMpg200d.LoadData(uType);
      var dlgMpg200D = new ViewBrockhausMpg200D(uType, Convert.ToDecimal(smpDataRow["ThickNessNominal"]), Convert.ToString(smpDataRow["Id"]), dsMagLab.MlMpg200d, resData);

      if (!dlgMpg200D.ShowDialog().GetValueOrDefault())
        return;

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Mpg200D;

      smpDataRow.BeginEdit();

      for (int j = 0; j < mesVal.Length; j++)
      {
        //Здесь происходит корректировка измеренных значений
        dsMagLab.MlMesurCof.DefaultView.ApplyDefaultSort = true;
        int i = dsMagLab.MlMesurCof.DefaultView.Find(new Object[] { smpDataRow["Md"], mesVal[j], uType, (int)MlMeasureDevice.Mpg200D });

        if ((i != -1) && (Convert.ToChar(dsMagLab.MlMesurCof.DefaultView[i]["TypCor"]) == 'D'))
          smpDataRow[mesVal[j]] = resData[mesVal[j]] + Convert.ToDecimal(dsMagLab.MlMesurCof.DefaultView[i]["Corr"]);
        else
          smpDataRow[mesVal[j]] = resData[mesVal[j]];
      }

      //rowSampleData["Massa"] = resData["Weight"];
      smpDataRow.EndEdit();
      dsMagLab.MlSiemensSmp.SaveData();
      dsMagLab.AcceptChanges();

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Ui5099; //Установка УИ5099
    }

    private void MeasureListMk4a(DataRow smpDataRow)
    {
      const int uType = 1;

      decimal? sLen = null;
      decimal? sWid = null;
      decimal? sDen = null;

      var ftag = new int[]{1,2,3,4,6,7,9,10};

      dsMagLab.MlMk4au.LoadData(Convert.ToString(smpDataRow["SteelType"]), Convert.ToDecimal(smpDataRow["ThickNessNominal"]), uType);
      if (dsMagLab.MlMk4au.Rows.Count != 0){
        sLen = Convert.ToDecimal(dsMagLab.MlMk4au.Rows[0]["Lsimple"]);
        sWid = Convert.ToDecimal(dsMagLab.MlMk4au.Rows[0]["Wsimple"]);
        sDen = Convert.ToDecimal(dsMagLab.MlMk4au.Rows[0]["Density"]);
      }

      dsMagLab.MlMk4ap.LoadData(uType, ftag);


      Window vLstAp = new ViewMeasureListAp(uType, dsMagLab.MlMk4ap, null, sLen, sWid, sDen, Convert.ToString(smpDataRow["Md"]), (int)MlMeasureDevice.Mk4A, dsMagLab.MlMesurCof);
      WindowsOption.ActveWindow = vLstAp;

      if (!vLstAp.ShowDialog().GetValueOrDefault()){
        WindowsOption.ActveWindow = (view as Window);
        return;
      }

      WindowsOption.ActveWindow = (view as Window);

      var obj = dbgSiemensSmp.GetFocusedRow();
      if (obj == null) return;
      var rowSampleData = (obj as DataRowView)?.Row;

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Mk4A; //Установка МК4Э

      rowSampleData?.BeginEdit();

      foreach (DataRow row in dsMagLab.MlMk4ap.Rows)
        if ((rowSampleData != null) && (Convert.ToString(row["MeasP"]) != "B3") && (Convert.ToString(row["MeasP"]) != "B30"))
          rowSampleData[Convert.ToString(row["MeasP"])] = row["OutVal"];

      //rowSampleData["Massa"] = vLstAp.Tag;

      rowSampleData?.EndEdit();
      dsMagLab.MlSiemensSmp.SaveData();
      dsMagLab.AcceptChanges();

      dsMagLab.MlData.MesDevice = (int)MlMeasureDevice.Ui5099; //Установка УИ5099
    }

    private void DataGridToExcel()
    {
      var sfd = new SaveFileDialog
      {
        OverwritePrompt = false,
        AddExtension = true,
        DefaultExt = ".xslx",
        Filter = "xlsx file (.xlsx)|*.xlsx"
      };

      if (sfd.ShowDialog().GetValueOrDefault() != true)
        return;

      dbgSiemensSmp.View.ExportToXlsx(sfd.FileName);
    }

    #endregion Private Method

    #region Constructor
    public ViewModelDlgSiemensSample(Control control, DsMgLab dsMagLab)
    {
      this.view = control;
      this.dsMagLab = dsMagLab;
      oldParentWindow = WindowsOption.ActveWindow;
      WindowsOption.ActveWindow = (this.view as Window);

      dbgSiemensSmp = LogicalTreeHelper.FindLogicalNode(this.view, "GcSiemensSample") as GridControl;
      if (this.dbgSiemensSmp != null)
        this.dbgSiemensSmp.CurrentItemChanged += CurrentItemChanged;

      (view as Window).Closing += (sender, args) => WindowsOption.ActveWindow = oldParentWindow;

      dsMagLab.MlDeviceLst.LoadData(1);
      SelectedMeasureDevice = Convert.ToInt32(dsMagLab.MlDeviceLst.Rows[1]["Id"]);
    }

    #endregion Constructor

    #region Command
    public void GetSample()
    {
      dsMagLab.MlSiemensSmp.GetListSimensSample(DateFrom, DateTo, SelectedShift);
    }

    public bool CanGetSample()
    {
      return true;
    }

    public void SaveChanges()
    {
      dsMagLab.MlSiemensSmp.SaveData();
    }

    public bool CanSaveChanges()
    {
      return dsMagLab.HasChanges();
    }

    public void CloseDialog()
    {
      this.dsMagLab.MlSiemensSmp.Clear();
      (view as Window)?.Close();
    }

    public bool CanCloseDialog()
    {
      return true;
    }

    public void SearchByLocNum()
    {
      this.dsMagLab.MlSiemensSmp.SearchByLocNum(FindLocNum);
    }

    public bool CanSearchByLocNum()
    {
      return !string.IsNullOrEmpty(FindLocNum);
    }

    public void CreateSamples4LocNum()
    {
      CreateSmp4LocNum();
    }

    public bool CanCreateSamples4LocNum()
    {
      return !string.IsNullOrEmpty(FindLocNum);
    }

    public void DeleteSample()
    {
      DeleteSimensSample();
    }

    public bool CanDeleteSample()
    {
      return (dsMagLab.MlSiemensSmp.Rows.Count > 0);
    }

    public void ToMes()
    {
      ChangeSimensSampleState(SampleState.Closed);
    }

    public bool CanToMes()
    {
      return ((dsMagLab.MlSiemensSmp.Rows.Count > 0) && (currentSmpDataRow.RowState != DataRowState.Detached) && (Convert.ToInt32(currentSmpDataRow["State"]) == (int)SampleState.Edited));
    }

    public void ToEdit()
    {
      ChangeSimensSampleState(SampleState.Edited);
    }

    public bool CanToEdit()
    {
      return ((dsMagLab.MlSiemensSmp.Rows.Count > 0) && (currentSmpDataRow.RowState != DataRowState.Detached) && (Convert.ToInt32(currentSmpDataRow["State"]) == (int)SampleState.Closed));
    }

    public void SaveRpt()
    {
      LabAction.SaveSimensRpt(DateFrom, DateTo);
    }

    public bool CanSaveRpt()
    {
      return true;
    }


    public void MeasureListParam()
    {
      if (SelectedMeasureDevice == (int) MlMeasureDevice.Mk4A)
        MeasureListMk4a(currentSmpDataRow);
      else if (SelectedMeasureDevice == (int) MlMeasureDevice.Mpg200D)
        MeasureListMpg200D(currentSmpDataRow);
      else
        DXMessageBox.Show((view as Window),  "Измерительное устр-во не определено!", "Измерить", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    public bool CanMeasureListParam()
    {
      return ((dsMagLab.MlSiemensSmp.Rows.Count > 0) && (currentSmpDataRow.RowState != DataRowState.Detached) && (Convert.ToInt32(currentSmpDataRow["State"]) == (int)SampleState.Edited));
    }

    public void ExportDataGridToExcel()
    {
      DataGridToExcel();
    }

    public bool CanExportDataGridToExcel()
    {
      return (dsMagLab.MlSiemensSmp.Rows.Count > 0);
    }


    #endregion Command

  }


}
