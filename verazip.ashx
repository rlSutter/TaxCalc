<%@ WebHandler Language="C#" Class="verazip" %>

using System;
using System.Web;
using System.Runtime.InteropServices;
using System.Web.Configuration;
using System.IO;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Text;

public class verazip : IHttpHandler {

    // Globals
    System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();

    // Structures
    [StructLayout(LayoutKind.Sequential)]
    public struct InputZip {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public byte[] StateCode;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 5)]
        public byte[] ZipCode;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
        public byte[] ZipExt;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
        public byte[] CityName;
    } 
    
    [StructLayout(LayoutKind.Sequential)]
    public struct LinkTable {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public byte[] StateCode;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 5)]
        public byte[] Zip1;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 5)]
        public byte[] Zip2;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public byte[] Geo;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
        public byte[] ZipExt1;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
        public byte[] ZipExt2;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
        public byte[] CityName;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public byte[] CntyCode;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
        public byte[] CntyName;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] OutsideCity;  
    } 
    
    [DllImport("C:\\Program Files (x86)\\Taxware\\utl\\avpzip.dll")]
    public static extern bool ZIPGEN(IntPtr inputzip, short CompletionCode, IntPtr linktable);

    //[DllImport("C:\\Program Files (x86)\\Taxware\\utl\\avpzip.dll")]
    //public static extern bool ZIPGENALIGNED(IntPtr inputzip, short CompletionCode, byte output);

    [DllImport("C:\\Program Files (x86)\\Taxware\\utl\\avpzip.dll")]
    public static extern bool ZipOpen(short Action);
    
    [DllImport("C:\\Program Files (x86)\\Taxware\\utl\\avpzip.dll")]
    public static extern bool ZipClose();
    
    public void ProcessRequest(HttpContext context)
    {
        //
        // This web service calls the VeraZip application with the provided data, and returns the geocoded information
        // in an XML document
        //
        // Based on G:\Software\Server Support\Taxware Update\Version 4.1\c_4.1_sut_st_Documentation\vzip.pdf
        //

        // ============================================
        // Declarations
        //  Generic   
        string mypath = "";
        string errmsg = "";
        string temp = "";
        
        // Parameters
        string StateCode = "";
        string ZipCode = "";
        string ZipExt = "";
        string CityName = "";
        
        // Structures
        InputZip inputzip = new InputZip();
        LinkTable linktable = new LinkTable();
        
        // Results
        short CompletionCode = 0;
        System.Xml.XmlDocument odoc = new System.Xml.XmlDocument();
        
        //  Logging
        string logfile = "";
        DateTime dt = DateTime.Now;
        string LogStartTime = dt.ToString();
        string Debug = "N";

        // Log4Net configuration
        log4net.Config.XmlConfigurator.Configure();
        log4net.ILog eventlog = log4net.LogManager.GetLogger("EventLog");
        log4net.ILog debuglog = log4net.LogManager.GetLogger("VZDebugLog");

        // ============================================
        // Debug Setup
        mypath = HttpRuntime.AppDomainAppPath;
        try
        {
            temp = WebConfigurationManager.AppSettings["VeraZip_debug"];
            Debug = temp;
        }
        catch { }

        // ============================================
        // Get Parameters from query string
        //  e.g. ?State=&Zip=&City=
        //  Assume zip in the format 01234 or 01234-5678
        if (context.Request.QueryString.HasKeys())
        {
            StateCode = context.Request.QueryString["State"];
            StateCode = StateCode.ToUpper();
            ZipCode = context.Request.QueryString["Zip"];
            if (ZipCode.Length > 5)
            {
                ZipCode = ZipCode.Substring(0, 5);
                ZipExt = ZipCode.Substring(6, 4);
            }
            CityName = context.Request.QueryString["City"];
            CityName = CityName.ToUpper();
        }
        else
        {
            // No parameters - test mode - Default test values
            StateCode = "VA";
            ZipCode = "22209";
            ZipExt = "2414";
            CityName = "Arlington";
        }

        // Convert strings to chars
        byte[] cStateCode = Encoding.ASCII.GetBytes(StateCode);
        byte[] cZipCode = Encoding.ASCII.GetBytes(ZipCode);
        byte[] cZipExt = Encoding.ASCII.GetBytes(ZipExt);
        byte[] cCityName = Encoding.ASCII.GetBytes(CityName);

        // Pad strings if applicable
        byte[] cCityName_f = new byte[26];
        int cCityName_l = cCityName.Length;
        if (cCityName_l < 26)
        {
            for (int i = 0; i < cCityName_l; i++)
            {
                cCityName_f[i] = cCityName[i];
            }
        }
        else
        {
            Array.Copy(cCityName, cCityName_f, 26);
        }
                        
        // Prepare input struct
        inputzip.StateCode = cStateCode;
        inputzip.ZipCode = cZipCode;
        inputzip.ZipExt = cZipExt;
        inputzip.CityName = cCityName_f;
        
        // Get points to arrays
        IntPtr pSource = StructToPtr(inputzip);
        IntPtr pDest = StructToPtr(linktable);
        
        // ============================================
        // Open log file if applicable
        if (Debug == "Y")
        {
            logfile = "C:\\Logs\\VeraZip.log";
            try
            {
                log4net.GlobalContext.Properties["VZLogFileName"] = logfile;
                log4net.Config.XmlConfigurator.Configure();
            }
            catch (Exception e)
            {
                errmsg = errmsg + "Error opening debug Log: " + e.ToString();
            }

            if (Debug == "Y")
            {
                debuglog.Debug("----------------------------------");
                debuglog.Debug("Trace Log Started " + LogStartTime);
                debuglog.Debug("Parameters-");
                debuglog.Debug("  StateCode: " + StateCode);
                debuglog.Debug("  ZipCode: " + ZipCode);
                debuglog.Debug("  ZipExt: " + ZipExt);
                debuglog.Debug("  CityName: " + CityName);
                debuglog.Debug(" ");
            }
        }

        // ============================================
        // Call VeraZip to process the 
        // Open ZIP Master File
        try
        {
            if (Debug == "Y") { debuglog.Debug("Calling ZipOpen \r\n");  }
            ZipOpen(1);
        }
        catch (Exception e)
        {
            errmsg = errmsg + "Error opening VeraZip: " + e.ToString();
        }

        // Lookup geocode
        try
        {
            if (Debug == "Y") { debuglog.Debug("Calling ZipGen \r\n"); }
            ZIPGEN(pSource, CompletionCode, pDest);
            if (Debug == "Y") { debuglog.Debug("  CompletionCode: " + CompletionCode.ToString() + "\r\n"); }
            
            // Convert linktable back to strings
            string rStateCode = Encoding.ASCII.GetString(linktable.StateCode);
            string rZip1 = Encoding.ASCII.GetString(linktable.Zip1);
            string rZip2 = Encoding.ASCII.GetString(linktable.Zip2);
            string rGeo = Encoding.ASCII.GetString(linktable.Geo);
            string rZipExt1 = Encoding.ASCII.GetString(linktable.ZipExt1);
            string rZipExt2 = Encoding.ASCII.GetString(linktable.ZipExt2);
            string rCityName = Encoding.ASCII.GetString(linktable.CityName);
            string rCntyCode = Encoding.ASCII.GetString(linktable.CntyCode);
            string rCntyName = Encoding.ASCII.GetString(linktable.CntyName);
            string rOutsideCity = Encoding.ASCII.GetString(linktable.OutsideCity);

            using (XmlWriter writer = odoc.CreateNavigator().AppendChild())
                try
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("results");
                    writer.WriteElementString("StateCode", rStateCode);
                    writer.WriteElementString("Zip1", rZip1);
                    writer.WriteElementString("Zip2", rZip2);
                    writer.WriteElementString("Geo", rGeo);
                    writer.WriteElementString("ZipExt1", rZipExt1);
                    writer.WriteElementString("ZipExt2", rZipExt2);
                    writer.WriteElementString("CityName", rCityName);
                    writer.WriteElementString("CntyCode", rCntyCode);
                    writer.WriteElementString("CntyName", rCntyName);
                    writer.WriteElementString("OutsideCity", rOutsideCity);
                    writer.WriteElementString("errmsg", errmsg);
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                catch (Exception e)
                {
                    errmsg = errmsg + ", " + e.ToString();
                }
        }
        catch (Exception e)
        {
            errmsg = errmsg + "Error opening VeraZip: " + e.ToString();
        }        
        
        // Close ZIP Master File
        try
        {
            if (Debug == "Y") { debuglog.Debug("Calling ZipClose \r\n"); }
            ZipClose();
        }
        catch (Exception e)
        {
            errmsg = errmsg + "Error closing VeraZip: " + e.ToString();
        }

        // Error Results
        if (errmsg != "")
        {
            using (XmlWriter writer = odoc.CreateNavigator().AppendChild())
                try
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("results");
                    writer.WriteElementString("errmsg", errmsg);
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                catch (Exception e)
                {
                    errmsg = errmsg + ", " + e.ToString();
                }
        }
        
        // ============================================
        // Generate response
        context.Response.ContentType = "text/xml";
        context.Response.Write(odoc.InnerXml);
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

    private static IntPtr StructToPtr(object obj)
    {
        var ptr = Marshal.AllocHGlobal(Marshal.SizeOf(obj));
        Marshal.StructureToPtr(obj, ptr, false);
        return ptr;
    }
    
    // ============================================
    // DEBUG FUNCTION
    private bool writeoutputfs(ref FileStream fs, String instring)
    {
        // This function writes a line to a previously opened filestream, and then flushes it
        // promptly.  This assists in debugging services
        Boolean result;
        try
        {
            instring = instring + "\r\n";
            //byte[] bytesStream = new byte[instring.Length];
            Byte[] bytes = encoding.GetBytes(instring);
            fs.Write(bytes, 0, bytes.Length);
            result = true;
        }
        catch
        {
            result = false;
        }
        fs.Flush();
        return result;
    }

}