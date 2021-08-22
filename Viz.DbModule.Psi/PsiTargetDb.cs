using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.ComponentModel.Composition;
using Smv.Mef.Contracts;
using Smv.Data.Oracle;

namespace Viz.DbModule.Psi
{

  [Export(typeof(Smv.Mef.Contracts.IDbModuleContract))]
  public sealed class PsiTargetDb : Smv.Mef.Contracts.IDbModuleContract
  {
    DataSets.DsApp dsApp = null;
    Boolean connectionResult;
 
    #region IDbConnectContract Members

    public bool Connect()
    {
      connectionResult = Odac.Connect(true);
                  
      if (this.connectionResult){
        dsApp.vModules.Connection = Odac.DbConnection;
        dsApp.vModules.LoadData();
      }

      return this.connectionResult;
    }

    public void Disconnect(Boolean IsDispose)
    {
       Odac.Disconnect(IsDispose);     
    }

    public string GetStatusInfo1(string inf)
    {
      return String.IsNullOrEmpty(inf) ? "Copyright © NLMK-IT Ltd 2008-" + DateTime.Today.Year.ToString(CultureInfo.InvariantCulture) : inf;
    }

    public string GetStatusInfo2(string inf)
    {
      return String.IsNullOrEmpty(inf) ? Odac.DbConnection.UserId.ToUpper(CultureInfo.InvariantCulture).Replace("_","*") : inf; 
    }

    public string GetStatusInfo3(string inf)
    {
      return String.IsNullOrEmpty(inf) ? "OraClient:" + Odac.GetClientVersion() : inf; 
    }

    public string GetStatusInfo4(string inf)
    {
      return String.IsNullOrEmpty(inf) ? Odac.GetDbAlias() : inf;
    }

    public string GetActualModuleVersion(string ModuleId)
    {
      return dsApp.vModules.GetModuleVersion(ModuleId);
    }

    public string GetModuleNameDescr(string ModuleId)
    {
      return dsApp.vModules.GetModuleNameDescr(ModuleId);      
    }


    #endregion

    public PsiTargetDb()
    {
      Odac.Init(new OdacUtils());
      dsApp = new DataSets.DsApp();
    }

  }
}
