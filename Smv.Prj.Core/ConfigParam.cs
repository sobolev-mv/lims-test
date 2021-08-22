using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Windows;


namespace Smv.App.Config
{
  
  public static class ConfigParam
  {

    public static string ReadAppSettingsParamValue(string ExeConfigFile, string KeyParam)
    {
      ExeConfigurationFileMap filemap = new ExeConfigurationFileMap();
      filemap.ExeConfigFilename = ExeConfigFile;
      Configuration config = ConfigurationManager.OpenMappedExeConfiguration(filemap, ConfigurationUserLevel.None);
      return config.AppSettings.Settings[KeyParam].Value;
    }

    public static string ReadConnectionStringParamValue(string ExeConfigFile, string KeyParam)
    {
      ExeConfigurationFileMap filemap = new ExeConfigurationFileMap();
      filemap.ExeConfigFilename = ExeConfigFile;
      Configuration config = ConfigurationManager.OpenMappedExeConfiguration(filemap, ConfigurationUserLevel.None);
      return config.ConnectionStrings.ConnectionStrings[KeyParam].ConnectionString;
    }

    public static object ReadPrivateConfigParam(System.String SectionName, System.String KeyName)
    {
      //открываем настройки пользователя
      Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
      try
      {
        if (config.Sections[SectionName] == null) return null;
        if (((AppSettingsSection)config.Sections[SectionName]).Settings.Count != 0)
          return ((AppSettingsSection)config.Sections[SectionName]).Settings[KeyName].Value;
      }
      catch (Exception ex)
      {
        DevExpress.Xpf.Core.DXMessageBox.Show(ex.Message, "Ошибка чтения кофигур. параметра", MessageBoxButton.OK, MessageBoxImage.Stop);
      }
      return null;
    }

    public static void WritePrivateConfigParam(System.String SectionName, System.String KeyName, System.Object KeyValue)
    {
      System.String Key = "";
      System.Boolean KeyExists = false;

      //открываем настройки пользователя
      Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
      try
      {
        //добавляем секцию наших настроек, если ее нет.
        if (config.Sections[SectionName] == null)
        {
          AppSettingsSection section = new AppSettingsSection();
          section.SectionInformation.AllowExeDefinition = ConfigurationAllowExeDefinition.MachineToLocalUser;
          config.Sections.Add(SectionName, section);
        }


        if (((AppSettingsSection)config.Sections[SectionName]).Settings.Count == 0){
          ((AppSettingsSection)config.Sections[SectionName]).Settings.Add(KeyName, KeyValue.ToString());
        }
        else{
          int len = ((AppSettingsSection)config.Sections[SectionName]).Settings.AllKeys.Length;
          for (int i = 0; i < len; i++){
            Key = ((AppSettingsSection)config.Sections[SectionName]).Settings.AllKeys[i];
            if (Key.CompareTo(KeyName) == 0){
              KeyExists = true;
              break;
            }
          }

          if (KeyExists)
            ((AppSettingsSection)config.Sections[SectionName]).Settings[KeyName].Value = KeyValue.ToString();
          else
            ((AppSettingsSection)config.Sections[SectionName]).Settings.Add(KeyName, KeyValue.ToString());
        }

        //Исправлена ошибка записи конфигурации для Win7
        //config.Save(ConfigurationSaveMode.Full);
        config.SaveAs(System.IO.Path.GetDirectoryName(config.FilePath) + "\\qwert.tmp", ConfigurationSaveMode.Full);
        System.IO.File.Delete(System.IO.Path.GetDirectoryName(config.FilePath) + "\\user.config");
        System.IO.File.Move(System.IO.Path.GetDirectoryName(config.FilePath) + "\\qwert.tmp", System.IO.Path.GetDirectoryName(config.FilePath) + "\\user.config");
      }
      catch (Exception ex){
        DevExpress.Xpf.Core.DXMessageBox.Show(ex.Message, "Ошибка записи кофигур. параметра", MessageBoxButton.OK, MessageBoxImage.Stop);
      }
    }




  
  }
  

}
