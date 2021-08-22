using System;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.DbModule.Psi.DataSets
{
    
    
  public partial class DsApp 
  {
     
    public partial class vModulesDataTable
    {
      public override void EndInit()
      {
        //call base method DataTable
        base.EndInit(); 
        //this.Connection = Odac.DbConnection;
      }

      public int LoadData()
      {
        return Odac.LoadDataTable(this, true, null);
      }

      public string GetModuleVersion(string ModuleId)
      {
        this.DefaultView.Sort = "Id";
        int i = this.DefaultView.Find(ModuleId);
        if (i == -1)
          return null;
        else
          return Convert.ToString(this.DefaultView[i]["Ver"]);
      }

      public string GetModuleNameDescr(string ModuleId)
      {
        this.DefaultView.Sort = "Id";
        int i = this.DefaultView.Find(ModuleId);
        if (i == -1)
          return null;
        else
          return Convert.ToString(this.DefaultView[i]["Descr"]);
      }




    }
  }


}
