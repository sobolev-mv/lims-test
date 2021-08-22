using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using Devart.Data.Oracle;
using Smv.Data.Oracle;

namespace Viz.DbModule.Psi
{

  public sealed class OdacUtils : IOdacUtils
  {
    public void ShowErrorInfo(string errorsTitle, string errorsMsg)
    {
      Smv.Utils.DxInfo.ShowDxBoxInfo(errorsTitle, errorsMsg, MessageBoxImage.Stop);
    }

    public string LoginName
    {
      get => Convert.ToString(Smv.App.Config.ConfigParam.ReadPrivateConfigParam("ConnectOdacParam", "LoginName"));
      set => Smv.App.Config.ConfigParam.WritePrivateConfigParam("ConnectOdacParam", "LoginName", value);
    }

    public string DbName 
    {
      get => Convert.ToString(Smv.App.Config.ConfigParam.ReadPrivateConfigParam("ConnectOdacParam", "DbName"));
      set => Smv.App.Config.ConfigParam.WritePrivateConfigParam("ConnectOdacParam", "DbName", value);
    }

    public bool IsUnicode
    {
      get
      {
        var str = Convert.ToString(Smv.App.Config.ConfigParam.ReadPrivateConfigParam("ConnectOdacParam", "IsUnicode"));
        return (!string.IsNullOrEmpty(str)) &&  Boolean.Parse(str);
      }

      set => Smv.App.Config.ConfigParam.WritePrivateConfigParam("ConnectOdacParam", "IsUnicode", value.ToString());
    }

    public Boolean GetLogonInfo(ref string Login, ref string DbName, ref string Pass, ref Boolean isUnicode)
    {
      ConnectWindow wnd = new ConnectWindow();
      wnd.tbLogin.Text = Login;
      wnd.tbBase.Text = DbName;

      if (isUnicode)
        wnd.rbUnicode.IsChecked = true;
      else
        wnd.rbAnsi.IsChecked = true;

      bool ?dlgResult =  wnd.ShowDialog();

      if (dlgResult == false) 
        return false;  
      else{
        Login = wnd.tbLogin.Text;
        DbName = wnd.tbBase.Text;
        Pass = wnd.pbPassword.Password;
        isUnicode = (bool) wnd.rbUnicode.IsChecked;
        return true;
      }

    }
  }
}
