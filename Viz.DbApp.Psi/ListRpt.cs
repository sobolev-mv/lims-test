using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Documents;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.DbApp.Psi
{
  public static class ListRpt
  {
    private static DataTable rptRef;
    private static OracleDataAdapter adpRptRef;

    private static void GetRptData(string ModuleId)
    {
      if (rptRef == null){
        adpRptRef = new OracleDataAdapter {SelectCommand = new OracleCommand {Connection = Odac.DbConnection}};
        rptRef = new DataTable {TableName = "RptRef"};
        DataColumn col = null;

        col = new DataColumn("Id", typeof (Int32), null, MappingType.Element) {AllowDBNull = false};
        rptRef.Columns.Add(col);
        col = new DataColumn("Caption", typeof (string), null, MappingType.Element);
        rptRef.Columns.Add(col);
        col = new DataColumn("RptGoal", typeof (string), null, MappingType.Element);
        rptRef.Columns.Add(col);
        col = new DataColumn("RptIn", typeof (string), null, MappingType.Element);
        rptRef.Columns.Add(col);
        col = new DataColumn("RemTxt", typeof (string), null, MappingType.Element);
        rptRef.Columns.Add(col);

        rptRef.Constraints.Add(new UniqueConstraint("Pk_RptRef", new[] {rptRef.Columns["Id"]}, true));
        rptRef.Columns["Id"].Unique = true;

        adpRptRef.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("V_ASC_RPT_REF", "RptRef");
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("CNTRL_CAPTION", "Caption");
        dtm.ColumnMappings.Add("RPT_GOAL", "RptGoal");
        dtm.ColumnMappings.Add("PRM_IN", "RptIn");
        dtm.ColumnMappings.Add("REM_TXT", "RemTxt");
        adpRptRef.TableMappings.Add(dtm);
      }
      
      //выставляем переменную сессии 
      DbVar.SetString(ModuleId);   

      //Заполняем таблицу
      adpRptRef.SelectCommand.CommandText = "SELECT ID, CNTRL_CAPTION, RPT_GOAL, PRM_IN, REM_TXT FROM LIMS.V_ASC_RPT_REF";
      adpRptRef.SelectCommand.Parameters.Clear();
      adpRptRef.SelectCommand.CommandType = CommandType.Text;
      Odac.LoadDataTable(rptRef, adpRptRef, true, null);  
    }

    private static void CreateListRptContent(FlowDocument fd)
    {
      foreach (DataRow row in rptRef.Rows){

        Paragraph par = new Paragraph();
        par.TextAlignment = TextAlignment.Justify;

        InlineUIContainer ilUIconteiner = new InlineUIContainer();
        ilUIconteiner.BaselineAlignment = BaselineAlignment.Center;
        Button btn = new Button();
        btn.Content = Convert.ToString(row["Caption"]);
        btn.Margin = new Thickness(0, 0, 5, 0);
        ilUIconteiner.Child = btn;
        par.Inlines.Add(ilUIconteiner);

        Bold boldNameRpt = new Bold();
        Run runBoldNameRpt = new Run();
        runBoldNameRpt.Text = Convert.ToString(row["RptGoal"]);
        runBoldNameRpt.FontSize = 20;
        boldNameRpt.Inlines.Add(runBoldNameRpt);
        par.Inlines.Add(boldNameRpt);
        par.Inlines.Add(new LineBreak());

        Underline undrInParam = new Underline();
        Run runUndrInParam = new Run();
        runUndrInParam.Text = "Входные параметры: ";
        runUndrInParam.FontSize = 18;
        runUndrInParam.FontWeight = FontWeights.Bold;
        runUndrInParam.Foreground = Brushes.DarkBlue;
        undrInParam.Inlines.Add(runUndrInParam);
        par.Inlines.Add(undrInParam);


        Run runInTextParamRpt = new Run();
        runInTextParamRpt.Text = Convert.ToString(row["RptIn"]);
        runInTextParamRpt.FontSize = 18;
        par.Inlines.Add(runInTextParamRpt);
        par.Inlines.Add(new LineBreak());

        if (!String.IsNullOrEmpty(Convert.ToString(row["RemTxt"]))){

          Underline undrRemTxt = new Underline();
          Run runUndrRemTxt = new Run();
          runUndrRemTxt.Text = "Примечание: ";
          runUndrRemTxt.FontSize = 18;
          runUndrRemTxt.FontWeight = FontWeights.Bold;
          runUndrRemTxt.Foreground = Brushes.DarkBlue;
          undrRemTxt.Inlines.Add(runUndrRemTxt);
          par.Inlines.Add(undrRemTxt);

          Run runInTextRemTxt = new Run();
          runInTextRemTxt.Text = Convert.ToString(row["RemTxt"]);
          runInTextRemTxt.FontSize = 18;
          par.Inlines.Add(runInTextRemTxt);
          par.Inlines.Add(new LineBreak());
        }

        fd.Blocks.Add(par);  
      }      
    }

    public static void ShowListRpt(string ModuleId, Image ModuleImg)
    {
      GetRptData(ModuleId);
      if (rptRef.Rows.Count == 0)
        return;

      ListRptWindow wnd = new ListRptWindow();
      wnd.Title = "Список отчетов";
      FlowDocument fd = wnd.FlowDocHlp;

      //Формирование заголовка документа
      InlineUIContainer HdrIlUIconteiner = new InlineUIContainer();
      HdrIlUIconteiner.BaselineAlignment = BaselineAlignment.Center;
      HdrIlUIconteiner.Child = ModuleImg;

      Bold boldHeader = new Bold();
      Run runBoldHeader = new Run();
      runBoldHeader.Text = "Описание отчетов модуля \"" + ModuleInfo.GetModuleDescription(ModuleId) + "\"";
      boldHeader.Inlines.Add(runBoldHeader);

      Paragraph prHeader = new Paragraph();
      prHeader.FontSize = 28;
      prHeader.Foreground = Brushes.Blue;
      prHeader.Inlines.Add(HdrIlUIconteiner);
      prHeader.Inlines.Add(boldHeader);
      //prHeader.Inlines.Add(new LineBreak());

      fd.Blocks.Add(prHeader);
      CreateListRptContent(fd);
      wnd.ShowDialog();
    }
  }
}
