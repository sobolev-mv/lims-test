using System;
using System.Data;

namespace Smv.Data
{
  public class SmvDataTable : DataTable
  {
    public class FillTableEventArgs : EventArgs
    {
      public FillTableEventArgs(int rowCount)
      {
        this.RowCount = rowCount;
      }
      public int RowCount { get; set; }
    }

    public event EventHandler<FillTableEventArgs> FilledTable;
    protected void OnFilledTable()
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<FillTableEventArgs> temp = FilledTable;
      if (temp != null)
        temp(this, new FillTableEventArgs(this.Rows.Count));
    }

    public object GetValByKey(string keyFieldName, Int64 keyFieldValue,  string valFieldName)
    {
      this.DefaultView.Sort = keyFieldName;
      int i = this.DefaultView.Find(keyFieldValue);
      return i == -1 ? null : this.DefaultView[i][valFieldName];
    }
    

  }
}
