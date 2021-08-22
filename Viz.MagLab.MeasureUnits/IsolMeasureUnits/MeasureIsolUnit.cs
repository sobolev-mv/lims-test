using System;
using System.Collections.Generic;
using System.Text;


namespace Viz.MagLab.MeasureUnits
{

  internal class MeasureEventArgs : EventArgs
  {
    private decimal measureValue;
    private int indexMeasureValue;
    
    public MeasureEventArgs(decimal ValueData, int Index)
    {
      measureValue = ValueData;
      indexMeasureValue = Index;
    }

    public decimal MeasureValue
    {
      get { return measureValue; }
      set { measureValue = value; }
    }

    public int IndexMeasureValue
    {
      get { return indexMeasureValue; }
      set { indexMeasureValue = value; }
    }

  }

  interface IMeasureIsolUnit
  {
    event EventHandler<MeasureEventArgs> MeasuredValue;
    void StartMeasure();
    void StopMeasure();
    void Close();
    int IndexMeasureValue {get; set;}
    Boolean IsError {get; set;}
  }

  internal abstract class MeasureIsolUnit : IMeasureIsolUnit
  {

    private int indexMeasureValue = 0;
    protected uint mCount = 0;
    protected String soundFile = null;
    // Declare an event of delegate type EventHandler of MyEventArgs.
    public event EventHandler<MeasureEventArgs> MeasuredValue;

    protected void OnMeasuredValue(decimal val)
    {
      //Copy to a temporary variable to be thread-safe.
      EventHandler<MeasureEventArgs> temp = MeasuredValue;

      if (temp != null){
        temp(this, new MeasureEventArgs(val, this.indexMeasureValue));

        if (this.indexMeasureValue >= (mCount - 1))
          indexMeasureValue = 0;
        else
          this.indexMeasureValue++;
      }
    }
        
    public  int IndexMeasureValue
    {
      get{ return indexMeasureValue; }
      set{ indexMeasureValue = value; }
    }

    public Boolean IsError { get; set; }  
    public abstract void  StartMeasure();
    public abstract void  StopMeasure();
    public abstract void  Close();

  }
}
