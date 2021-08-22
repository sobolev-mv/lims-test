using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace Viz.Lims.Sh
{

  public sealed class ComPortSectionHandler : ConfigurationSection
  {

    //First constructor
    public ComPortSectionHandler() { }

    //Second constructor
    public ComPortSectionHandler(string SectionName, string Port, String StrPortParam, int BaudRate, string Parity, String StopBits, string HandShake)
    {
      this.SectionName = SectionName;
      this.Port = Port;
      this.StrPortParam = StrPortParam;
      this.BaudRate = BaudRate;
      this.Parity = Parity;
      this.StopBits = StopBits;
      this.HandShake = HandShake;
      
    }
    
    //First property: SectionName
    private string _SectionName;

    public string SectionName
    {
      get{return _SectionName;}
      set{_SectionName = value;}
    }

    //Second property: The port Key
    [ConfigurationProperty("port", DefaultValue = "COM1:", IsRequired = true)]
    public string Port
    {
      get { return (string)this["port"]; }
      set { this["port"] = value; }
    }

    //Second property: The build string port param
    [ConfigurationProperty("strportparam", DefaultValue = "", IsRequired = true)]
    public string StrPortParam
    {
      get { return (string)this["strportparam"]; }
      set { this["strportparam"] = value; }
    }

    //Third property: The port serial
    [ConfigurationProperty("baudrate", DefaultValue = 9600, IsRequired = true)]
    public int BaudRate
    {
      get{return (int)this["baudrate"];}
      set{this["baudrate"] = value;}
    }

    [ConfigurationProperty("parity", DefaultValue = "none", IsRequired = true)]
    public string Parity
    {
      get { return (string)this["parity"]; }
      set { this["parity"] = value; }
    }

    [ConfigurationProperty("stopbits", DefaultValue = "One", IsRequired = true)]
    public string StopBits
    {
      get { return (string)this["stopbits"]; }
      set { this["stopbits"] = value; }
    }

    [ConfigurationProperty("handshake", DefaultValue = "None", IsRequired = true)]
    public string HandShake
    {
      get { return (string)this["handshake"]; }
      set { this["handshake"] = value; }
    }

    //Third property: The port serial
    [ConfigurationProperty("syncreadtimeout", DefaultValue = 1000, IsRequired = true)]
    public int SyncReadTimeOut
    {
      get { return (int)this["syncreadtimeout"]; }
      set { this["syncreadtimeout"] = value; }
    }

    //Third property: The port serial
    [ConfigurationProperty("syncwritetimeout", DefaultValue = 1000, IsRequired = true)]
    public int SyncWriteTimeOut
    {
      get { return (int)this["syncwritetimeout"]; }
      set { this["syncwritetimeout"] = value; }
    }


  }

}
