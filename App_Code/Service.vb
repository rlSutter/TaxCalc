Option Explicit On

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Collections
Imports System.Configuration
Imports System.Math
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Text
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.Diagnostics
Imports System.Net
Imports System.Net.Sockets
Imports Microsoft.VisualBasic.Strings
Imports log4net
Imports System.Runtime.InteropServices

<WebService(Namespace:="http://accounting.hq.local/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    Public Const UTL_OUTPUT_DATA_SIZE = 1704
    Public Const OUT_PARAM_SIZE = 284
    Public Const OUT_FIELD_COUNT = 42
    Public Const OUTPUT_LENGTH_RECORD_COUNT = 6

    ' Tax declaration
    Declare Function CapiTaxRoutine Lib "C:\Program Files (x86)\Taxware\utl\taxcommono.dll" Alias "CapiTaxRoutine" (ByVal InputData As String, ByVal OutputData As String, ByVal InputLength As Integer) As Integer
    'Declare Function CapiTaxRoutine Lib "\\hciaccounting.hq.local\Taxware\utl\taxcommono.dll" Alias "CapiTaxRoutine" (ByVal InputData As String, ByVal OutputData As String, ByVal InputLength As Integer) As Integer

    ' Verazip declaration
    'Public Shared Function ZIPGENALIGNED(ByVal InputParmPtr As InputZip, ComplCode As Long, ByRef OutputParmPtr As OutputZip) As Byte
    'Public Shared Function ZIPGENALIGNED(ByVal InPtr As IntPtr, ComplCode As Int64, ByVal OutputParmPtr As OutputParmPtrOutput) As Boolean
    '<DllImport("\\hciaccounting.hq.local\taxware\toolkit\AVPZIP.DLL", EntryPoint:="ZIPGENALIGNED", CharSet:=CharSet.Ansi, ExactSpelling:=False, CallingConvention:=CallingConvention.StdCall)> _
    'Public Shared Function ZIPGENALIGNED(ByRef InPtr As IntPtr, ComplCode As Int64, <MarshalAs(UnmanagedType.LPStr, SizeConst:=78)> <InAttribute(), Out()> OutData As String) As Boolean
    'Public Shared Function ZIPGENALIGNED(ByRef InPtr As IntPtr, ByVal ComplCode As Int64, <MarshalAs(UnmanagedType.LPStr, SizeConst:=23400)> ByRef OutputParmPtrOutput As String) As Boolean
    'End Function

    'Declare Function ZIPGENALIGNED_ Lib "C:\Program Files (x86)\Taxware\toolkit\AVPZIP.dll" Alias "ZIPGENALIGNED_" (ByRef InPtr As IntPtr, ByVal ComplCode As Int64, <MarshalAs(UnmanagedType.LPStr, SizeConst:=23400)> ByVal OutputParmPtr As String) As Boolean
    'Declare Function ZIPGENALIGNED_ Lib "\\hciaccounting.hq.local\taxware\toolkit\AVPZIP.dll" Alias "ZIPGENALIGNED" (ByRef InPtr As InputZipAligned, ComplCode As Long, ByVal OutputParmPtr As String) As Byte
    Declare Function ZIPGENALIGNED Lib "\\hciaccounting.hq.local\taxware\toolkit\avpzip.dll" (ByRef InputParmPtr As InputZip, ByRef ComplCode As Long, ByVal OutputParmPtr As String) As Byte
    Declare Function ZIPOPEN Lib "\\hciaccounting.hq.local\taxware\toolkit\avpzip.dll" (ByVal Action As Integer) As Byte

    Public Declare Function VerifyZipIntPtr Lib "C:\Program Files (x86)\Taxware\utl\avpzip.dll" Alias "VERIFY_ZIP" (ByVal zipInBuffer As String, ByVal zipOutBuffer As IntPtr) As Integer
    '<DllImport("C:\Program Files (x86)\Taxware\utl\avpzip.dll", EntryPoint:="VERIFY_ZIP", CharSet:=CharSet.Ansi)>
    'Private Shared Function VerifyZipIntPtr(ByVal zipInBuffer As String, ByVal zipOutBuffer As IntPtr) As Integer

    'Declare Auto Function ZIPGENALIGNED Lib "C:\Program Files (x86)\Taxware\toolkit\AVPZIP.DLL" Alias "ZipGenAligned" (InputParmPtr As InputZip, ComplCode As Long, ByVal OutputParmPtr As String) As Byte
    'Declare Auto Function ZIPOPEN Lib "C:\Program Files (x86)\Taxware\toolkit\AVPZIP.DLL" Alias "ZipOpen" (ByVal Action As Long) As Byte
    'Declare Auto Function ZIPCLOSE Lib "C:\Program Files (x86)\Taxware\toolkit\AVPZIP.DLL" Alias "ZipClosed" () As Byte
    'Declare Auto Function ZIPGENALIGNED Lib "\\hciaccounting.hq.local\taxware\toolkit\AVPZIP.DLL" Alias "ZipGenAligned" (InputParmPtr As InputZip, ComplCode As Long, ByVal OutputParmPtr As String) As Byte
    'Declare Auto Function ZIPOPEN Lib "\\hciaccounting.hq.local\Taxware\toolkit\AVPZIP.DLL" Alias "ZipOpen" (ByVal Action As Long) As Byte
    'Declare Auto Function ZIPCLOSE Lib "\\hciaccounting.hq.local\Taxware\toolkit\AVPZIP.DLL" Alias "ZipClosed" () As Byte

    ' Logging objects
    Private myeventlog As log4net.ILog
    Private mydebuglog As log4net.ILog

    ' Enumerate objct used for validation
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
    End Enum

    Private ReadOnly key() As Byte = {90, 23, 49, 1, 17, 32, 77, 10, 74, 3, 102, 87, 13, 92, 200, 122, 145, 6, 65, 44, 59, 14, 90, 212}
    Private ReadOnly iv() As Byte = {8, 7, 6, 5, 4, 3, 2, 1}

    ' Instantiate the TripleDES encryption class with the arrays
    Private des As New cTripleDES(key, iv)

    '<StructLayout(LayoutKind.Sequential, Pack:=1, Size:=37, CharSet:=CharSet.Ansi)> _
    <StructLayout(LayoutKind.Sequential, Pack:=1, Size:=37, CharSet:=CharSet.Ansi)> _
    Public Structure InputZipAligned
        '   State Code      = 2 LETTERS
        '   Zip Code        = 5 DIGITS
        '   Zip Ext.        = 4 DIGITS
        '   City/Prov       = 26 LETTERS: City or Province name 
        '<FieldOffset(0)> Public State(2) As Char
        '<FieldOffset(2)> Public Zipcode(5) As Char
        '<FieldOffset(7)> Public ZipExt(4) As Char
        '<FieldOffset(11)> Public City(26) As Char

        '<MarshalAs(UnmanagedType.LPStr, SizeConst:=2)> Public State As String
        '<MarshalAs(UnmanagedType.LPStr, SizeConst:=5)> Public Zipcode As String
        '<MarshalAs(UnmanagedType.LPStr, SizeConst:=4)> Public ZipExt As String
        '<MarshalAs(UnmanagedType.LPStr, SizeConst:=26)> Public City As String

        '<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=2)> Public State As String
        '<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=5)> Public Zipcode As String
        '<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=4)> Public ZipExt As String
        '<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=26)> Public City As String

        Public State As String
        Public Zipcode As String
        Public ZipExt As String
        Public City As String

    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=0, Size:=37, CharSet:=CharSet.Ansi)> _
    Structure InputZip
        Dim StateCode As String
        Dim ZipCode As String
        Dim ZipExt As String
        Dim CityName As String

        Public Sub New(ByVal unusedParam As Integer)
            StateCode = New StringBuilder(Space(2), 2).ToString()
            ZipCode = New StringBuilder(Space(5), 5).ToString()
            ZipExt = New StringBuilder(Space(4), 4).ToString()
            CityName = New StringBuilder(Space(26), 26).ToString()
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1, Size:=23400, CharSet:=CharSet.Ansi)> _
    Public Structure OutputParmPtrOutput
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=23400)> Public Outdata As String
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1, Size:=78, CharSet:=CharSet.Ansi)> _
    Public Class OutputZip
        '   State Code      = 2 LETTERS
        '   Zip Code        = 5 DIGITS
        '   Zip Code Range  = 5 DIGITS
        '   Geocode         = 2 DIGITS
        '   Zip Ext. First  = 4 DIGITS
        '   Zip Ext. Last   = 4 DIGITS
        '   City/Prov       = 26 LETTERS : City or Province name
        '   County Code     = 3 DIGITS: FIPS county code
        '   County          = 26 LETTERS: County name
        '   In/Outsize city = 1 DIGIT: 0 inside city limits, 1 outside city limits, 2 outside but in police juris.
        '<FieldOffset(0)> Public State(2) As Char
        '<FieldOffset(2)> Public Zipcode(5) As Char
        '<FieldOffset(7)> Public ZipcodeRange(5) As Char
        '<FieldOffset(12)> Public Geocode(2) As Char
        '<FieldOffset(14)> Public ZipExtFst(4) As Char
        '<FieldOffset(18)> Public ZipExpLast(4) As Char
        '<FieldOffset(22)> Public City(26) As Char
        '<FieldOffset(48)> Public CtyCode(3) As Char
        '<FieldOffset(51)> Public County(26) As Char
        '<FieldOffset(77)> Public InOut(1) As Char
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=2)> Public State As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=5)> Public Zipcode As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=5)> Public ZipcodeRange As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=2)> Public Geocode As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=4)> Public ZipExtFst As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=4)> Public ZipExpLast As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=26)> Public City As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=3)> Public CtyCode As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=26)> Public County As String
        <MarshalAs(UnmanagedType.LPStr, SizeConst:=1)> Public InOut As String
    End Class

    '<WebMethod(Description:="Call Verazap on an address")> _
    Private Function GetVeraZipGeoCode(ByVal ZipCode As String, ByVal City As String, ByVal State As String, ByVal Debug As String, ByRef mydebuglog As log4net.ILog) As String

        ' ============================================
        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        Const outputLength As Integer = 21277
        Dim input As String = ""
        City = UCase(Left(City, 26))
        input = input & State.PadRight(2)
        input = input & ZipCode.PadRight(5)
        input = input & City.PadRight(60)

        Dim geoCode As String
        Try
            Dim output As IntPtr = Marshal.AllocHGlobal(outputLength)
            Dim code As Integer = VerifyZipIntPtr(input, output)
            Dim outputData As Byte() = New Byte(outputLength - 1) {}
            Marshal.Copy(output, outputData, 0, outputLength)
            Dim complCode As String = ASCIIEncoding.ASCII.GetString(outputData, 6, 2) 'Completion Code
            Dim retMsg As String = ASCIIEncoding.ASCII.GetString(outputData, 8, 200) 'Completion Code

            ' Compute geocode
            geoCode = ASCIIEncoding.ASCII.GetString(outputData, 220, 2) 'Geo Code Only
            'geoCode = ASCIIEncoding.ASCII.GetString(outputData, 0, 294) 'All results
            Marshal.FreeHGlobal(output)

            ' Logging
            If Debug = "Y" Then
                mydebuglog.Debug("  >> input: " & input)
                mydebuglog.Debug("  >> geoCode Completion Code: " & complCode)
                mydebuglog.Debug("  >> geoCode Returned Msg: " & Left(retMsg, 30))
                mydebuglog.Debug("  >> geoCode: " & geoCode)
            End If
            'myeventlog.Info("GetVeraZipGeoCode : Geo Code: " & geoCode & " for (ZipCode:" & ZipCode & ", City: " & City & ", State: " & State & ")")

            Return geoCode
        Catch ex As Exception
            If Debug = "Y" Then
                mydebuglog.Debug("GetVeraZipGeoCode failed. Error:" & ex.Message & "; (ZipCode:" & ZipCode & ", City: " & City & ", State: " & State & ")")
            End If
            'myeventlog.Info("GetVeraZipGeoCode failed. Error:" & ex.Message & " for (ZipCode:" & ZipCode & ", City: " & City & ", State: " & State & ")")
            Return "GetVeraZipGeoCode failed. Error:" & ex.Message
        End Try
    End Function

    <WebMethod(Description:="Geocode An Address")> _
    Public Function GeoCodeAddress(ByVal ZipCode As String, ByVal City As String, _
        ByVal State As String, ByVal Debug As String) As String

        ' ============================================
        ' Declarations
        Dim logfile, temp, SGEOCODE, geoCode As String
        Dim LogStartTime As String = Now.ToString
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        geoCode = "00"

        ' ============================================
        ' Get system defaults
        temp = System.Configuration.ConfigurationManager.AppSettings.Get("SourceGeocode")
        If temp <> "" Then SGEOCODE = temp Else SGEOCODE = "00"

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Then
            Dim path As String
            path = HttpRuntime.AppDomainAppPath    ' Path to root
            logfile = "C:\Logs\GeoCodeAddress.log"
            log4net.GlobalContext.Properties("LogFileName") = logfile
            log4net.Config.XmlConfigurator.Configure()
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Format(Now))
            mydebuglog.Debug(vbCrLf & "Parameters: ")
            mydebuglog.Debug(" > ZipCode: " & ZipCode)
            mydebuglog.Debug(" > City: " & City)
            mydebuglog.Debug(" > State: " & State)            
        End If

        ' ============================================
        ' Perform Geocoding
        Try
            geoCode = GetVeraZipGeoCode(ZipCode, City, State, Debug, mydebuglog)
            If Debug = "Y" Then
                mydebuglog.Debug("  << geoCode: " & geoCode)
            End If
        Catch ex As Exception
            If Debug = "Y" Then
                mydebuglog.Debug("  << geoCode error: " & ex.Message)
            End If
        End Try
        If geoCode = "" Then geoCode = "00"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "geoCode: " & geoCode & vbCrLf)

        ' ============================================
        ' Close the log file if any
        If Debug = "Y" Then
            mydebuglog.Debug("Trace Log Ended " & Format(Now))
            mydebuglog.Debug("----------------------------------")
        End If
        Return geoCode

    End Function

    <WebMethod(Description:="Perform tax computation on an order")> _
    Public Function TaxCalc(ByVal OrderID As String, ByVal Source As String, _
        ByVal UpdateFlg As String, ByVal AuditFlg As String, ByVal debug As String) As XmlDocument

        ' This service computes tax for the order specified.  If UpdateFlg is set to "Y" it will
        ' update the order in the database specified by Source.  The tax specifics are returned to
        ' the calling program as an XML document in the following form:

        '   <tax>
        '   <order_id>1-86Z75</order_id>
        '   <invoice_num>222140</invoice_num>
        '   <source>siebeldb</source>
        '   <error>No error - debugging</error>
        '   <state>IA</state>
        '   <city>Cedar Falls</city>
        '   <zipcode>506133400</zipcode>
        '   <freight_taxable>N</freight_taxable>
        '   <audit_flg>N</audit_flg>
        '   <product_amt>260</product_amt>
        '   <freight_amt>12.1</freight_amt>
        '   <order_net_amt>272.1</order_net_amt>
        '   <product_tax_amt>13</product_tax_amt>
        '   <freight_tax_amt>0</freight_tax_amt>
        '   <total_tax_amt>13</total_tax_amt>
        '   <order_total_amt>285.1</order_total_amt>
        '   <tax_rate_pct>5</tax_rate_pct>
        '   <order_items>1</order_items>
        '   <id_1>1-86Z79</id_1>
        '   <tax_1>13</tax_1>
        '   <rate_1>0.0500</rate_1>
        '   <total_1>260</total_1>
        '   </tax>

        ' The calling parameters are as follows:
        '   OrderID     -   The ROW_ID of the S_ORDER record
        '   Source      -   The database where the S_ORDER table is stored on the hcidb2 database
        '   UpdateFlg   -   If set to "Y" then update the order record with the shipping amount
        '   AuditFlg    -   If set to "Y" then send the tax to the audit file
        '   debug       -   If set to "Y" then write a lot to C:\TaxCalc.log on the server

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String

        ' Logging declarations
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Dim VersionNum As String = "101"

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Variable declarations
        Dim errmsg As String                 ' Error message (if any)
        Dim strTaxWareInputs As String
        Dim strTaxWareTemp As String
        Dim strReturn As String
        Dim strTaxRate As String
        Dim dblTotalAmount As Double
        Dim strFreight As String
        Dim dblFreight As Double
        Dim strGrossAmount As String
        Dim dblGrossAmount As Double
        Dim FreightTaxability As Boolean
        Dim strFreightTaxability As String
        Dim strDate As String
        Dim strOrderStatus As String
        Dim strCreditOrderFlag As String
        Dim psInputs1 As New SimpleDictionary(1000)       ' List to hold inputs to tax system
        Dim psItem As DictionaryEntry
        Dim TaxOutputs As New SimpleDictionary(1000)      ' List to hold outputs from tax system
        Dim LineNumber As Integer
        Dim NumberLines As Integer
        Dim returnV As Integer
        Dim decRateTotal As Double
        Dim FreightTaxAmount, ProductTax, ProductTotal As Double
        Dim ROW_ID, TAX_AMT, TAX_PERCENT As String
        Dim TaxRate As Double
        Dim strTotalTax, strFreightTaxAmount, curAuditFlg, InvoiceNum As String
        Dim strProductTax, strProductTotal, strProductTaxRate As String
        Dim strTotalAmount, strOrderTotal As String

        Dim DISCNT_AMT, DISCNT_PERCENT, UNIT_PRI, BASE_UNIT_PRI, QTY_REQ As Double
        Dim LINE_TOTAL, L_DISCNT_AMT, L_DISCNT_PERCENT, LINE_QTY_REQ As Double
        Dim REG_FLG, CITY, ZIPCODE, SHIP_ADDR_ID, SHIP_PER_ADDR_ID, COUNTY, strNewDate As String
        Dim rROW_ID As String
        Dim OU_ID As String
        Dim STATE, DGEOCODE, SGEOCODE As String

        ' ============================================
        ' Correct parameters
        OrderID = Trim(UCase(OrderID))
        debug = UCase(Left(debug, 1))
        If debug = "" Then debug = "N"
        AuditFlg = UCase(Left(AuditFlg, 1))
        If AuditFlg = "" Then AuditFlg = "N"
        UpdateFlg = UCase(Left(UpdateFlg, 1))
        If UpdateFlg = "" Then UpdateFlg = "N"
        Source = LCase(Trim(Source))
        If Source = "" Then Source = "siebeldb"

        ' ============================================
        ' Handle NAGIOS test
        If debug = "T" Then
            OrderID = "1-L7LTD"
            AuditFlg = "N"
            UpdateFlg = "N"
            Source = "siebeldb"
        End If

        ' ============================================
        ' Set default values
        errmsg = "No error"
        If debug = "Y" Then errmsg = errmsg & " - debugging"
        If Source = "" Then Source = "siebeldb"
        SqlS = ""
        CITY = ""
        ZIPCODE = ""
        OU_ID = ""
        STATE = ""
        SHIP_ADDR_ID = ""
        REG_FLG = ""
        COUNTY = ""
        strOrderStatus = ""
        strTaxWareInputs = ""
        strGrossAmount = ""
        strTotalTax = ""
        strFreightTaxability = ""
        strFreight = ""
        strFreightTaxAmount = ""
        strProductTaxRate = ""
        returnV = 1
        curAuditFlg = ""
        InvoiceNum = ""
        rROW_ID = ""
        strProductTotal = ""
        strFreight = ""
        strTotalAmount = ""
        strProductTax = ""
        strFreightTaxAmount = ""
        strTotalTax = ""
        strOrderTotal = ""
        strTaxRate = ""
        temp = ""
        DGEOCODE = ""
        SGEOCODE = ""

        ' ============================================
        ' Get system defaults
        ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
        If ConnS = "" Then ConnS = "server="
        temp = System.Configuration.ConfigurationManager.AppSettings.Get("TaxCalc_debug")
        If temp = "Y" And debug <> "T" Then debug = "Y"
        temp = System.Configuration.ConfigurationManager.AppSettings.Get("SourceGeocode")
        If temp <> "" Then SGEOCODE = temp Else SGEOCODE = "00"

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Then
            Dim path As String
            path = HttpRuntime.AppDomainAppPath    ' Path to root
            logfile = "C:\Logs\TaxCalc.log"
            log4net.GlobalContext.Properties("LogFileName") = logfile
            log4net.Config.XmlConfigurator.Configure()
            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Trace Log Started " & Format(Now))
            mydebuglog.Debug(vbCrLf & "Parameters: ")
            mydebuglog.Debug(" > UpdateFlg: " & UpdateFlg)
            mydebuglog.Debug(" > AuditFlg: " & AuditFlg)
            mydebuglog.Debug(" > Source: " & Source)
            mydebuglog.Debug(" > OrderID: " & OrderID)
        End If

        ' ============================================
        ' Validate basic parameter
        If Trim(OrderID) = "" Then
            errmsg = "No order id was specified"
            GoTo CloseOut2
        End If

        ' ============================================
        ' Get the order basics 
        If Source = "reports2" Then
            SqlS = "SELECT TOP 1 (SELECT CASE WHEN O.DISCNT_AMT IS NULL THEN 0 ELSE O.DISCNT_AMT END), " & _
            "(SELECT CASE WHEN O.DISCNT_PERCENT IS NULL THEN 0 ELSE O.DISCNT_PERCENT END), " & _
            "(SELECT CASE WHEN I.UNIT_PRI IS NULL THEN 0 ELSE I.UNIT_PRI END), " & _
            "(SELECT CASE WHEN I.BASE_UNIT_PRI IS NULL THEN 0 ELSE I.BASE_UNIT_PRI END), " & _
            "(SELECT CASE WHEN I.QTY_REQ IS NULL THEN 0 ELSE I.QTY_REQ END), " & _
            "(SELECT CASE WHEN O.FRGHT_AMT IS NULL THEN 0 ELSE O.FRGHT_AMT END), A.STATE, A.CITY, A.ZIPCODE, O.ROW_ID, " & _
            "E.X_ACCOUNT_NUM, E.BU_ID, E.NAME, O.CREATED, O.STATUS_CD, O.X_REGISTRATION_FLG, CA.CITY, " & _
            "CA.STATE, CA.ZIPCODE, E.ROW_ID, A.ROW_ID, PA.ROW_ID, PA.STATE, PA.CITY, PA.ZIPCODE, " & _
                "O.X_AUDIT_FLG, O.X_INVOICE_NUM, O.ROW_ID, (SELECT CASE WHEN O.REQ_SHIP_DT IS NULL THEN O.CREATED ELSE O.REQ_SHIP_DT END), " & _
                "A.COUNTY, CA.COUNTY, PA.COUNTY " & _
                "FROM reports.dbo.S_ORDER O " & _
                "LEFT OUTER JOIN reports.dbo.S_ORDER_ITEM I ON I.ORDER_ID=O.ROW_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_OFFR C ON C.ROW_ID=I.CRSE_OFFR_ID " & _
                "LEFT OUTER JOIN reports.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.X_HELD_ADDRESS_ID " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT E ON E.ROW_ID=O.ACCNT_ID " & _
                "LEFT OUTER JOIN reports.dbo.S_ADDR_ORG A ON A.ROW_ID=O.SHIP_ADDR_ID " & _
                "LEFT OUTER JOIN reports.dbo.S_ADDR_PER PA ON PA.ROW_ID=O.SHIP_PER_ADDR_ID " & _
                "WHERE O.ROW_ID='" & OrderID & "'"
            Source = "reports"
        Else
            SqlS = "SELECT TOP 1 (SELECT CASE WHEN O.DISCNT_AMT IS NULL THEN 0 ELSE O.DISCNT_AMT END), " & _
            "(SELECT CASE WHEN O.DISCNT_PERCENT IS NULL THEN 0 ELSE O.DISCNT_PERCENT END), " & _
            "(SELECT CASE WHEN I.UNIT_PRI IS NULL THEN 0 ELSE I.UNIT_PRI END), " & _
            "(SELECT CASE WHEN I.BASE_UNIT_PRI IS NULL THEN 0 ELSE I.BASE_UNIT_PRI END), " & _
            "(SELECT CASE WHEN I.QTY_REQ IS NULL THEN 0 ELSE I.QTY_REQ END), " & _
            "(SELECT CASE WHEN O.FRGHT_AMT IS NULL THEN 0 ELSE O.FRGHT_AMT END), A.STATE, A.CITY, A.ZIPCODE, O.ROW_ID, " & _
            "E.X_ACCOUNT_NUM, E.BU_ID, E.NAME, O.CREATED, O.STATUS_CD, O.X_REGISTRATION_FLG, CA.CITY, " & _
            "CA.STATE, CA.ZIPCODE, E.ROW_ID, A.ROW_ID, PA.ROW_ID, PA.STATE, PA.CITY, PA.ZIPCODE, " & _
            "O.X_AUDIT_FLG, O.X_INVOICE_NUM, O.ROW_ID, (SELECT CASE WHEN O.REQ_SHIP_DT IS NULL THEN O.CREATED ELSE O.REQ_SHIP_DT END), " & _
            "A.COUNTY, CA.COUNTY, PA.COUNTY " & _
        "FROM " & Source & ".dbo.S_ORDER O " & _
        "LEFT OUTER JOIN " & Source & ".dbo.S_ORDER_ITEM I ON I.ORDER_ID=O.ROW_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_OFFR C ON C.ROW_ID=I.CRSE_OFFR_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.X_HELD_ADDRESS_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT E ON E.ROW_ID=O.ACCNT_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=O.SHIP_ADDR_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=O.SHIP_PER_ADDR_ID " & _
        "WHERE O.ROW_ID='" & OrderID & "'"
        End If
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Order basics: " & vbCrLf & SqlS & vbCrLf)

        con = New SqlConnection(ConnS)
        con.Open()
        cmd = New SqlCommand(SqlS, con)
        dr = cmd.ExecuteReader()
        If dr Is Nothing Then
            errmsg = "Database error"
            GoTo CloseOut
        End If
        While dr.Read()
            ' Order amount fields
            DISCNT_AMT = dr(0)
            DISCNT_PERCENT = dr(1)
            UNIT_PRI = dr(2)
            BASE_UNIT_PRI = dr(3)
            QTY_REQ = dr(4)
            dblFreight = dr(5)
            REG_FLG = Trim(dr(15).ToString)

            ' Address fields
            OU_ID = Trim(dr(19).ToString)
            If REG_FLG = "Y" Then
                CITY = Trim(dr(16).ToString)
                STATE = Trim(dr(17).ToString)
                ZIPCODE = Trim(dr(18).ToString)
                COUNTY = Trim(dr(30).ToString)
            Else
                STATE = Trim(dr(6).ToString)
                CITY = Trim(dr(7).ToString)
                ZIPCODE = Trim(dr(8).ToString)
                COUNTY = Trim(dr(29).ToString)
            End If
            SHIP_ADDR_ID = Trim(dr(20).ToString)
            SHIP_PER_ADDR_ID = Trim(dr(21).ToString)
            If STATE = "" And SHIP_PER_ADDR_ID <> "" Then
                STATE = Trim(dr(22).ToString)
                CITY = Trim(dr(23).ToString)
                ZIPCODE = Trim(dr(24).ToString)
                COUNTY = Trim(dr(31).ToString)
            End If
            curAuditFlg = Trim(dr(25).ToString)         ' Get current audit status
            If curAuditFlg = "" Then curAuditFlg = "N"
            InvoiceNum = Trim(dr(26).ToString)

            ' Store fields to the collection
            If (AuditFlg = "Y" And InvoiceNum <> "") Or InvoiceNum <> "" Then
                psInputs1.Add("Quote Number", InvoiceNum)
            Else
                psInputs1.Add("Quote Number", Trim(dr(9).ToString))
            End If
            If Trim(dr(10).ToString) = "" Then
                psInputs1.Add("Account Num", "1179")
            Else
                psInputs1.Add("Account Num", Trim(dr(10).ToString))
            End If
            If Trim(dr(11).ToString) = "" Then
                psInputs1.Add("Organization Id", "0-R9NH")
            Else
                psInputs1.Add("Organization Id", Trim(dr(11).ToString))
            End If
            psInputs1.Add("Account", Trim(dr(12).ToString))
            strDate = Format(dr(13), "MM/dd/yyyy")
            strNewDate = Mid(strDate, 7, 4) & Mid(strDate, 1, 2) & Mid(strDate, 4, 2)
            psInputs1.Add("Created", strNewDate)
            If debug = "Y" Then mydebuglog.Debug("* Created date: " & strNewDate)
            strOrderStatus = Trim(dr(14).ToString)
            rROW_ID = Trim(dr(27).ToString)
            ' Original invoice date
            If strOrderStatus = "Credit" Then
                strDate = Format(dr(28), "MM/dd/yyyy")
                strNewDate = Mid(strDate, 7, 4) & Mid(strDate, 1, 2) & Mid(strDate, 4, 2)
                psInputs1.Add("Invoice Date", strNewDate)
                If debug = "Y" Then mydebuglog.Debug("* Invoice date: " & strNewDate)
            Else
                psInputs1.Add("Invoice Date", strNewDate)
            End If
        End While
        dr.Close()

        If debug = "Y" Then mydebuglog.Debug("* Order Status: " & strOrderStatus)

        ' Validate whether an order was found or not
        If rROW_ID = "" Then
            errmsg = "The specified order was not found in the database " & Source
            GoTo CloseOut
        End If

        ' Validate audit status - if already sent to audit file, do not recalc
        If curAuditFlg = "Y" And debug <> "Y" Then
            errmsg = "Order already sent to the audit file - may not be recalculated"
            GoTo CloseOut
        End If

        ' Correct address for accounts
        If OU_ID <> "" And STATE = "" And REG_FLG <> "Y" Then
            SqlS = "SELECT TOP 1 SA.ROW_ID, SA.STATE, SA.CITY, SA.ZIPCODE, BA.ROW_ID, " & _
            "BA.STATE, BA.CITY, BA.ZIPCODE, PA.ROW_ID, PA.STATE, PA.CITY, PA.ZIPCODE, SA.COUNTY, BA.COUNTY, PA.COUNTY " & _
            "FROM siebeldb.dbo.S_ORG_EXT O " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG SA ON SA.ROW_ID=O.PR_SHIP_ADDR_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG BA ON BA.ROW_ID=O.PR_BL_ADDR_ID " & _
            "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG PA ON PA.ROW_ID=O.PR_ADDR_ID " & _
            "WHERE O.ROW_ID='" & OU_ID & "'"
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: New Address: " & vbCrLf & SqlS & vbCrLf)
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If dr Is Nothing Then
                errmsg = "Database error"
                GoTo CloseOut
            End If
            While dr.Read()
                SHIP_ADDR_ID = Trim(dr(0).ToString)
                STATE = Trim(dr(1).ToString)
                CITY = Trim(dr(2).ToString)
                ZIPCODE = Trim(dr(3).ToString)
                COUNTY = Trim(dr(12).ToString)
                If CITY = "" Then
                    SHIP_ADDR_ID = Trim(dr(4).ToString)
                    STATE = Trim(dr(5).ToString)
                    CITY = Trim(dr(6).ToString)
                    ZIPCODE = Trim(dr(7).ToString)
                    COUNTY = Trim(dr(13).ToString)
                    If CITY = "" Then
                        SHIP_ADDR_ID = Trim(dr(8).ToString)
                        STATE = Trim(dr(9).ToString)
                        CITY = Trim(dr(10).ToString)
                        ZIPCODE = Trim(dr(11).ToString)
                        COUNTY = Trim(dr(14).ToString)
                    End If
                End If
            End While
            dr.Close()
        End If

        ' Save address
        If CITY = "" Or STATE = "" Or ZIPCODE = "" Then
            errmsg = "No Address Specified"
            GoTo CloseOut
        End If
        If Len(ZIPCODE) > 5 Then ZIPCODE = Left(ZIPCODE, 5)
        psInputs1.Add("Account Ship To State", STATE)
        psInputs1.Add("Account Ship To City", CITY)
        psInputs1.Add("Account Ship To Postal Code", ZIPCODE)
        psInputs1.Add("Account Ship To County", COUNTY)
        If debug = "Y" Then
            mydebuglog.Debug("* Account Ship To State: " & STATE)
            mydebuglog.Debug("* Account Ship To City: " & CITY)
            mydebuglog.Debug("* Account Ship To Postal Code: " & ZIPCODE)
            mydebuglog.Debug("* Account Ship To County: " & COUNTY)
        End If

        ' ===========================================
        ' GET DESTINATION GEOCODE - disabled for the time being
        ' If COUNTY <> "" And CITY <> "" And ZIPCODE <> "" Then
        'SqlS = "SELECT TOP 1 z.GEO_CODE " & _
        '"FROM siebeldb.dbo.TAXWARE_CNTY_FIPS f  " & _
        '"join siebeldb.dbo.TAXWARE_ZIP_MASTER z on z.CNTY_FIPS = f.FIPS_CNTY and z.STATE_CODE = f.FIPS_ST " & _
        '"WHERE f.CNTY_NAME = UPPER(left('" & COUNTY & "', 26)) " & _
        '"and z.CITY_NAME = UPPER(left('" & CITY & "', 26)) " & _
        '"and z.ZIP5 = LEFT('" & ZIPCODE & "',5)"
        'If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Geocode lookup: " & vbCrLf & SqlS & vbCrLf)
        'cmd.CommandText = SqlS
        'dr = cmd.ExecuteReader()
        'If dr Is Nothing Then
        'errmsg = "Database error"
        'GoTo CloseOut
        'End If
        'While dr.Read()
        'DGEOCODE = Trim(dr(0).ToString)
        'End While
        'dr.Close()
        'Else
        'DGEOCODE = "00"
        'End If
        'psInputs1.Add("Account Ship To Geocode", DGEOCODE)
        'psInputs1.Add("Shipping Source Geocode", SGEOCODE)
        'If debug = "Y" Then
        'mydebuglog.Debug("* Shipping Source Geocode: " & SGEOCODE)
        'mydebuglog.Debug("* Account Ship To Geocode: " & DGEOCODE)
        'End If

        ' ===========================================
        ' SET GLOBAL MISC FIELD VALUES
        ' Audit File Flag:   "1", send to audit file, "2" do not send to audit file
        ' NOTE: Order status updates associated with audit file writing must be performed externally
        If debug = "Y" Then
            mydebuglog.Debug("* Order Status: " & strOrderStatus)
            mydebuglog.Debug("* Current Audit File Flag: " & curAuditFlg)
            mydebuglog.Debug("* Parameter Audit File Flag: " & AuditFlg)
        End If
        psInputs1.Add("Audit File Flag", "2")
        If AuditFlg = "Y" And curAuditFlg = "N" Then
            psInputs1("Audit File Flag") = "1"
        End If
        strCreditOrderFlag = "0"
        If strOrderStatus = "Credit" Then strCreditOrderFlag = "1"
        psInputs1.Add("Credit Order Flag", strCreditOrderFlag)              ' Validate status

        ' ===========================================
        ' DETERMINE FREIGHT TAXABILITY
        FreightTaxability = False
        If dblFreight <> 0 Then
            SqlS = "SELECT TOP 1 HIGH FROM siebeldb.dbo.S_LST_OF_VAL " & _
            "WHERE TYPE='FREIGHT_TAXABILITY' AND NAME='" & STATE & "'"
            If debug = "Y" Then mydebuglog.Debug("QUERY: Freight tax: " & vbCrLf & SqlS)
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If dr Is Nothing Then
                errmsg = "Database error"
                GoTo CloseOut
            End If
            While dr.Read()
                If Trim(dr(0).ToString) = "Y" Then FreightTaxability = True
            End While
            dr.Close()
        End If
        If FreightTaxability Then strFreightTaxability = "Y" Else strFreightTaxability = "N"
        If debug = "Y" Then
            mydebuglog.Debug("* Freight taxable: " & strFreightTaxability)
            mydebuglog.Debug("* Freight amount: " & Str(dblFreight))
        End If

        ' ===========================================
        ' GET NUMBER OF LINE ITEMS AND ADJUST FOR FREIGHT TAXABILITY
        NumberLines = 0
        SqlS = "SELECT COUNT(*) FROM " & Source & ".dbo.S_ORDER_ITEM " & _
        "WHERE ORDER_ID='" & OrderID & "' AND " & _
        "((SELECT CASE WHEN BASE_UNIT_PRI IS NULL OR BASE_UNIT_PRI=0 THEN UNIT_PRI ELSE BASE_UNIT_PRI END)-" & _
        "(SELECT CASE WHEN DISCNT_AMT IS NULL THEN 0 ELSE DISCNT_AMT END))*QTY_REQ<>0"
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Count lines: " & vbCrLf & SqlS & vbCrLf)
        cmd.CommandText = SqlS
        dr = cmd.ExecuteReader()
        If dr Is Nothing Then
            errmsg = "Database error"
            GoTo CloseOut
        End If
        While dr.Read()
            NumberLines = dr(0)
        End While
        dr.Close()
        If FreightTaxability Then NumberLines = NumberLines + 1
        If debug = "Y" Then mydebuglog.Debug("* Number Line Items: " & NumberLines)
        If NumberLines < 1 Then
            errmsg = "No line items found"
            GoTo CloseOut
        End If

        ' ===========================================
        ' GET LINE ITEMS AND PREPARE TAXWARE INPUT STRING
        dblTotalAmount = 0
        SqlS = "SELECT (SELECT CASE WHEN I.UNIT_PRI IS NULL THEN I.BASE_UNIT_PRI ELSE I.UNIT_PRI END)*I.QTY_REQ AS ITEM_TOTAL, " & _
        "(SELECT CASE WHEN I.DISCNT_AMT IS NULL THEN 0 ELSE I.DISCNT_AMT END)*I.QTY_REQ, " & _
        "(SELECT CASE WHEN I.DISCNT_PERCENT IS NULL THEN 0 ELSE I.DISCNT_PERCENT END), " & _
        "P.PART_NUM, " & _
        "(SELECT CASE WHEN I.BASE_UNIT_PRI IS NULL THEN 0 ELSE I.BASE_UNIT_PRI END), " & _
        "(SELECT CASE WHEN I.UNIT_PRI IS NULL THEN 0 ELSE I.UNIT_PRI END), " & _
        "L.HIGH AS TAXPROD, I.ROW_ID, " & _
        "(SELECT CASE WHEN I.QTY_REQ IS NULL THEN 0 ELSE I.QTY_REQ END) AS QTY_REQ " & _
        "FROM " & Source & ".dbo.S_ORDER_ITEM I " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_PROD_INT P ON P.ROW_ID=I.PROD_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_LST_OF_VAL L ON L.TYPE='TAXWARE' AND L.NAME=P.PART_NUM " & _
        "WHERE ORDER_ID='" & OrderID & "' AND " & _
        "((SELECT CASE WHEN BASE_UNIT_PRI IS NULL OR BASE_UNIT_PRI=0 THEN UNIT_PRI ELSE BASE_UNIT_PRI END)-" & _
        "(SELECT CASE WHEN DISCNT_AMT IS NULL THEN 0 ELSE DISCNT_AMT END))*QTY_REQ<>0 " & _
        "ORDER BY I.LN_NUM"
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Line items: " & vbCrLf & SqlS & vbCrLf)
        cmd.CommandText = SqlS
        dr = cmd.ExecuteReader()
        If dr Is Nothing Then
            errmsg = "No taxable records found"
            GoTo CloseOut
        End If
        Dim PART_NUM, strTaxWareProductCode As String
        LineNumber = 0
        psInputs1.Add("Total Amount", "")
        psInputs1.Add("Product", "")
        psInputs1.Add("Freight", 0)
        psInputs1.Add("Quantity Requested", 0)
        If debug = "Y" Then mydebuglog.Debug("=============START PROCESSING LINE ITEMS")
        While dr.Read()
            LineNumber = LineNumber + 1

            ' Get part numbers
            PART_NUM = Trim(dr(3).ToString)                             ' Get first product id				
            strTaxWareProductCode = Trim(dr(6).ToString)                ' Taxware product code

            ' Get prices and apply discount to get LINE_TOTAL
            BASE_UNIT_PRI = dr(4)
            UNIT_PRI = dr(5)
            LINE_QTY_REQ = dr(8)
            If BASE_UNIT_PRI <> 0 Or UNIT_PRI <> 0 Then
                LINE_TOTAL = dr(0)
                L_DISCNT_AMT = dr(1)                                    ' Line Item $ Discount
                If Not dr(1) Is Nothing And L_DISCNT_AMT <> 0 Then
                    If debug = "Y" Then mydebuglog.Debug("LINE ITEM AMOUNT DISCOUNT APPLIED: " & Str(L_DISCNT_AMT))
                    LINE_TOTAL = LINE_TOTAL - L_DISCNT_AMT
                Else
                    L_DISCNT_PERCENT = dr(2)                            ' Line Item % Discount
                    If L_DISCNT_PERCENT <> 0 Then
                        If debug = "Y" Then mydebuglog.Debug("LINE ITEM PERCENTAGE DISCOUNT APPLIED: " & Str(L_DISCNT_PERCENT))
                        L_DISCNT_PERCENT = L_DISCNT_PERCENT / 100       ' Convert to fraction
                        LINE_TOTAL = Round(LINE_TOTAL - (LINE_TOTAL * L_DISCNT_PERCENT), 2)
                    End If
                End If
                dblTotalAmount = dblTotalAmount + LINE_TOTAL            ' Total amount of order
            End If
            If debug = "Y" Then mydebuglog.Debug("LINE #" & Str(LineNumber) & vbCrLf & "  > Unit Price: " & UNIT_PRI & vbCrLf & _
            "  > Base Unit Price: " & BASE_UNIT_PRI & vbCrLf & "  > Line total: " & Format(LINE_TOTAL, "Currency") & vbCrLf & "  > Discount: " & _
            Format(L_DISCNT_AMT, "Currency") & vbCrLf & "  > Order Total: " & Format(dblTotalAmount, "Currency"))

            ' Set fields for tax calculation
            psInputs1("Total Amount") = Trim(Str(Math.Abs(LINE_TOTAL)))
            psInputs1("Product") = strTaxWareProductCode
            psInputs1("Freight") = 0
            psInputs1("Quantity Requested") = LINE_QTY_REQ                  ' Quantity Ordered

            ' Prepare list fields for outputs 
            TaxOutputs.Add("id_" & Trim(Str(LineNumber)), dr(7))            ' Record ID
            TaxOutputs.Add("tax_" & Trim(Str(LineNumber)), 0)               ' Tax amount
            TaxOutputs.Add("rate_" & Trim(Str(LineNumber)), 0)              ' Tax rate
            TaxOutputs.Add("total_" & Trim(Str(LineNumber)), Trim(Str(Math.Abs(LINE_TOTAL))))    ' Line total   

            ' Generate string to pass to taxware
            strTaxWareTemp = Calc_Tax_Line(mydebuglog, NumberLines, LineNumber, psInputs1, debug)
            If debug = "Y" Then mydebuglog.Debug("  To Taxware string: size=" & Str(Len(strTaxWareTemp)) & vbCrLf & "  " & strTaxWareTemp)
            strTaxWareInputs = strTaxWareInputs & strTaxWareTemp
        End While
        If debug = "Y" Then mydebuglog.Debug("=============END PROCESSING LINE ITEMS")
        dr.Close()

        ' ===========================================
        ' GENERATE LINE ITEM FOR FREIGHT TAX IF NECESSARY
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "=============START PROCESSING FREIGHT")
        strFreight = Trim(Str(dblFreight))
        If FreightTaxability Then
            dblGrossAmount = dblTotalAmount + dblFreight

            ' Prepare input fields
            psInputs1("Total Amount") = 0
            psInputs1("Product") = "43504"
            psInputs1("Freight") = Trim(Str(Math.Abs(dblFreight)))
            psInputs1("Quantity Requested") = 1
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "* Freight Amount Taxable: " & strFreight)

            ' Prepare output fields
            TaxOutputs.Add("id_" & Trim(Str(NumberLines)), "freight")      ' Record ID = blank for freight
            TaxOutputs.Add("tax_" & Trim(Str(NumberLines)), 0)             ' Tax amount
            TaxOutputs.Add("rate_" & Trim(Str(NumberLines)), 0)            ' Tax rate
            TaxOutputs.Add("total_" & Trim(Str(NumberLines)), Trim(Str(Math.Abs(dblFreight))))  ' Line total

            ' Generate string to pass to taxware
            LineNumber = NumberLines                                              ' The line number makes it the last line
            strTaxWareTemp = Calc_Tax_Line(mydebuglog, NumberLines, LineNumber, psInputs1, debug)
            strTaxWareInputs = strTaxWareInputs & strTaxWareTemp
            If debug = "Y" Then mydebuglog.Debug("* To Taxware string freight: size=" & Str(Len(strTaxWareTemp)) & vbCrLf & strTaxWareTemp)
        Else
            If debug = "Y" Then mydebuglog.Debug("* Freight Amount NOT taxable: " & strFreight)
            dblGrossAmount = dblTotalAmount
        End If
        If debug = "Y" Then mydebuglog.Debug("=============END PROCESSING FREIGHT" & vbCrLf)

        strGrossAmount = Trim(Str(Math.Abs(dblGrossAmount)))
        If debug = "Y" Then mydebuglog.Debug("* Total Taxable amount: " & strGrossAmount & vbCrLf)

        ' Debug output of list values
        If debug = "Y" Then
            mydebuglog.Debug(">>Tax calculation output values BEFORE Tax calculation" & vbCrLf)
            For Each psItem In TaxOutputs
                mydebuglog.Debug(" > " & psItem.Key & " = " & psItem.Value)
            Next
        End If

        ' ===========================================
        ' CALCULATE INPUT PARAMETERS
        Dim linsize As Integer
        Dim intcount As String
        strReturn = ""
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "* Taxware input string: " & vbCrLf & strTaxWareInputs)
        linsize = Len(strTaxWareInputs)
        intcount = Mid(strTaxWareInputs, 1, OUTPUT_LENGTH_RECORD_COUNT)
        Dim sR2 As String
        sR2 = ""
        strReturn = sR2.PadLeft((UTL_OUTPUT_DATA_SIZE * Val(intcount)) + OUT_PARAM_SIZE, " ")
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & ">>TAXWARE INPUT STRING PARAMETERS")
            mydebuglog.Debug(" * Data size: " & Str(UTL_OUTPUT_DATA_SIZE))
            mydebuglog.Debug(" * Parameter size: " & Str(OUT_PARAM_SIZE))
            mydebuglog.Debug(" * Precomputed return string size: " & Str(Len(strReturn)))
            mydebuglog.Debug(" * Input size: " & Str(Len(strTaxWareInputs)))
            mydebuglog.Debug(" * Lines: " & Str(intcount) & vbCrLf)
        End If

        ' ===========================================
        ' INVOKE TAXWARE
        Try
            Call CapiTaxRoutine(strTaxWareInputs, strReturn, linsize)
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Taxware error: " & ex.ToString)
            errmsg = "Tax calculation error"
            GoTo CloseOut
        End Try
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "========" & vbCrLf & "Taxware output string: Size: " & Str(Len(strReturn)) & vbCrLf & strReturn)
        If Left(strReturn, 10) = "0000000000" Then
            errmsg = "Tax calculation error"
            GoTo CloseOut
        End If

        ' ===========================================
        ' PARSE RESULTS
        decRateTotal = Parse_Tax_Rate(mydebuglog, TaxOutputs, strReturn, dblGrossAmount, debug)
        If strCreditOrderFlag = "1" Then decRateTotal = decRateTotal * -1
        strTotalTax = Trim(Str(decRateTotal))
        If debug = "Y" Then mydebuglog.Debug("* Returned tax amount: " & Str(decRateTotal))

        ' Compute tax rate
        If decRateTotal > 0 Then
            TaxRate = Math.Round((decRateTotal / dblGrossAmount) * 100, 2)
        Else
            TaxRate = 0
        End If
        strTaxRate = Trim(Str(TaxRate))
        If debug = "Y" Then mydebuglog.Debug("* Formatted tax rate: " & strTaxRate & "%")

        ' Debug output of list values
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & ">>Tax calculation output values AFTER Tax calculation")
            For Each psItem In TaxOutputs
                mydebuglog.Debug(" > " & psItem.Key & " = " & psItem.Value)
            Next
        End If

        ' ===========================================
        ' UPDATE ORDER LINE ITEMS AS APPLICABLE
        ' Update Line Items
        FreightTaxAmount = 0           ' Freight tax total
        ProductTax = 0                 ' Total tax charged on products
        ProductTotal = 0               ' Total amount charged for products
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "ORDER UPDATES:" & vbCrLf & "* NumberLines: " & NumberLines.ToString)
        For LineNumber = 1 To NumberLines
            ' Calculate rate total
            ROW_ID = TaxOutputs("id_" & Trim(Str(LineNumber)))
            TAX_AMT = Trim(TaxOutputs("tax_" & Trim(Str(LineNumber))))
            LINE_TOTAL = Trim(TaxOutputs("total_" & Trim(Str(LineNumber))))
            If strCreditOrderFlag = "1" Then
                LINE_TOTAL = LINE_TOTAL * -1
                TaxOutputs("total_" & Trim(Str(LineNumber))) = LINE_TOTAL
                TAX_AMT = TAX_AMT * -1
                TaxOutputs("tax_" & Trim(Str(LineNumber))) = TAX_AMT
            End If

            If Val(TAX_AMT) <> 0 And Val(LINE_TOTAL) <> 0 Then
                TAX_PERCENT = Trim(Format(Val(TAX_AMT) / Val(LINE_TOTAL), "0.0000"))
            Else
                TAX_PERCENT = "0.00"
            End If
            TaxOutputs("rate_" & Trim(Str(LineNumber))) = TAX_PERCENT
            If ROW_ID = "freight" Then
                FreightTaxAmount = TAX_AMT
            Else
                ProductTax = ProductTax + Val(TAX_AMT)
                ProductTotal = ProductTotal + LINE_TOTAL
            End If

            ' Update line item record
            If UpdateFlg = "Y" And curAuditFlg = "N" Then
                If ROW_ID <> "freight" Then
                    SqlS = "UPDATE " & Source & ".dbo.S_ORDER_ITEM SET X_TAX_AMT=" & TAX_AMT & ", X_TAX_RATE=" & TAX_PERCENT & " WHERE ROW_ID='" & ROW_ID & "'"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Update Line item: " & vbCrLf & SqlS)
                    cmd.CommandText = SqlS
                    If UpdateFlg = "Y" Or AuditFlg = "Y" Then returnV = cmd.ExecuteNonQuery()
                    If returnV = 0 Then
                        errmsg = "The item was not found to update"
                    End If

                    ' Set freight taxability
                    SqlS = "UPDATE " & Source & ".dbo.S_ORDER_ITEM_X SET ATTRIB_08='" & strFreightTaxability & "' WHERE PAR_ROW_ID='" & ROW_ID & "'"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Update Line item_x: " & vbCrLf & SqlS)
                    cmd.CommandText = SqlS
                    If UpdateFlg = "Y" Or AuditFlg = "Y" Then returnV = cmd.ExecuteNonQuery()
                    If returnV = 0 Then
                        errmsg = errmsg & "The item extension was not found to update"
                    End If

                End If
            End If
        Next
        strFreightTaxAmount = Trim(Str(FreightTaxAmount))
        strTotalAmount = Trim(Str(ProductTotal + dblFreight))
        strOrderTotal = Trim(Str(ProductTotal + ProductTax + dblFreight + FreightTaxAmount))

        ' ===========================================
        ' UPDATE ORDER AS APPLICABLE
        strProductTax = Trim(Str(ProductTax))
        strProductTotal = Trim(Str(ProductTotal))
        If ProductTax > 0 And ProductTotal > 0 Then
            strProductTaxRate = Trim(Str(Math.Round((ProductTax / ProductTotal) * 100, 2)))
        Else
            strProductTaxRate = "0"
        End If
        If debug = "Y" Then
            mydebuglog.Debug("* strProductTax: " & strProductTax)
            mydebuglog.Debug("* strProductTotal: " & strProductTotal)
            mydebuglog.Debug("* strProductTaxRate: " & strProductTaxRate)
        End If
        If UpdateFlg = "Y" And curAuditFlg = "N" Then
            SqlS = "UPDATE " & Source & ".dbo.S_ORDER " & _
            "SET SHIP_ADDR_ID='" & SHIP_ADDR_ID & "', TAX_AMT=" & strProductTax & ", " & _
            "TAX_AMT_CURCY_CD='USD', TAX_AMT_EXCH_DT=GETDATE(), X_FREIGHT_TAXABILITY='" & strFreightTaxability & "', " & _
            "TAX_PERCENT=" & strProductTaxRate & ", X_FREIGHT_TAX_AMOUNT=" & strFreightTaxAmount & ", X_AUDIT_FLG='" & AuditFlg & "' " & _
            "WHERE ROW_ID='" & OrderID & "'"
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "QUERY: Update order: " & vbCrLf & SqlS)
            cmd.CommandText = SqlS
            If UpdateFlg = "Y" Or AuditFlg = "Y" Then returnV = cmd.ExecuteNonQuery()
            If returnV = 0 Then
                errmsg = "The order was not found to update"
            End If
        End If

        If FreightTaxability Then
            decRateTotal = decRateTotal + FreightTaxAmount
        End If

        ' RESTORE VARIABLES FOR OUTPUT
        STATE = psInputs1("Account Ship To State")
        CITY = psInputs1("Account Ship To City")
        ZIPCODE = psInputs1("Account Ship To Postal Code")

CloseOut:
        ' ============================================
        ' Close database connections
        Try
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Error closing objects: " & ex.ToString
        End Try

CloseOut2:
        ' Log transaction to syslog
        If debug <> "T" Then myeventlog.Info("TaxCalc : Order Id: " & OrderID & ", Tax $:" & strProductTax & ", Tax %:" & strProductTaxRate)

        ' ============================================
        ' Return the freight amount to the service consumer as an XML document
        Dim results As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = results.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        results.InsertBefore(resultsDeclare, results.DocumentElement)
        resultsRoot = results.CreateElement("tax")
        results.InsertAfter(resultsRoot, resultsDeclare)

        If debug = "T" Then
            If errmsg = "No error" Then
                AddXMLChild(results, resultsRoot, "results", "Success")
            Else
                AddXMLChild(results, resultsRoot, "results", "Failure")
                AddXMLChild(results, resultsRoot, "error", Trim(errmsg))
            End If
        Else
            ' Add result items - send what was submitted for debugging purposes on the calling end
            AddXMLChild(results, resultsRoot, "order_id", OrderID)
            AddXMLChild(results, resultsRoot, "invoice_num", InvoiceNum)
            AddXMLChild(results, resultsRoot, "source", Source)
            If errmsg <> "" Then AddXMLChild(results, resultsRoot, "error", Trim(errmsg))

            ' Summary order info
            AddXMLChild(results, resultsRoot, "state", STATE)
            AddXMLChild(results, resultsRoot, "city", CITY)
            AddXMLChild(results, resultsRoot, "zipcode", ZIPCODE)
            AddXMLChild(results, resultsRoot, "freight_taxable", strFreightTaxability)  ' Taxable flag
            AddXMLChild(results, resultsRoot, "audit_flg", AuditFlg)                    ' Send to the audit file

            ' Tax detail
            AddXMLChild(results, resultsRoot, "product_amt", strProductTotal)           ' Products total $
            AddXMLChild(results, resultsRoot, "freight_amt", strFreight)                ' Freight $
            AddXMLChild(results, resultsRoot, "order_net_amt", strTotalAmount)          ' Order Total $ before tax
            AddXMLChild(results, resultsRoot, "product_tax_amt", strProductTax)         ' Product tax $
            AddXMLChild(results, resultsRoot, "freight_tax_amt", strFreightTaxAmount)   ' Freight tax $
            AddXMLChild(results, resultsRoot, "total_tax_amt", strTotalTax)             ' Total tax $
            AddXMLChild(results, resultsRoot, "order_total_amt", strOrderTotal)         ' Order Total $
            AddXMLChild(results, resultsRoot, "tax_rate_pct", strTaxRate)               ' Tax percent
            AddXMLChild(results, resultsRoot, "order_items", Trim(Str(NumberLines)))    ' Number line items

            ' Line items
            For Each psItem In TaxOutputs
                AddXMLChild(results, resultsRoot, psItem.Key, Trim(psItem.Value))
            Next
        End If

        ' ============================================
        ' Close the log file if any
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "* errmsg: " & Trim(errmsg) & vbCrLf)
            mydebuglog.Debug("Trace Log Ended " & Format(Now))
            mydebuglog.Debug("----------------------------------")
        End If
        If errmsg.Substring(0, 8) <> "No error" Then myeventlog.Error("TaxCalc : " & errmsg)

        ' Log Performance Data
        If debug <> "T" Then LoggingService.LogPerformanceData2Async(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)

        ' Drop objects
        psInputs1 = Nothing
        psItem = Nothing
        TaxOutputs = Nothing

        Return results

    End Function

    Private Function Calc_Tax_Line(ByRef mydebuglog As ILog, ByVal NumberLines As Integer, _
        ByVal LineNumber As Integer, ByVal Inputs As SimpleDictionary, ByVal debug As String) As String
        ' THIS IS CALLED BY CALC_TAX
        ' INPUTS:		
        '		NumberLines			Number of order line items
        '		LineNumber			The number of line that is being written
        '		Inputs 					A List of field values, in the format:
        '										name = field name
        '										value	= field value
        '		debug					Debug flag

        ' OUTPUTS:
        '		Calc_Tax_Line		String of concatenated values in the form expected by Taxware	

        'On Error Resume Next
        Dim strGenInput As String
        Dim intGenInputLen As Integer
        Dim strLineInput As String
        Dim intLineInputLen As Integer
        Dim tempTotalAmt As String
        Dim tempFreight As String
        Dim tempQtyReq As String
        Dim psItem As DictionaryEntry
        Dim tf2, tt2, tt3 As String
        Dim tempZip, tempState, tempCity, geoCode As String
        Dim ShipToState As String

        intGenInputLen = 60
        intLineInputLen = 2592
        geoCode = "00"

        strGenInput = Space(intGenInputLen)
        strLineInput = Space(intLineInputLen)
        ShipToState = Inputs("Account Ship To State")

        ' Debug input of list values
        If debug = "Y" Then
            mydebuglog.Debug("  ==========" & vbCrLf & "  ENTERING Tax_Calc_Line")
            For Each psItem In Inputs
                mydebuglog.Debug("   > " & psItem.Key & " = " & psItem.Value)
            Next
        End If

        Dim TAX_RECORD_NUMBER As Integer
        TAX_RECORD_NUMBER = 0
        Dim FLD_LEN_RECORD_NUMBER As Integer
        FLD_LEN_RECORD_NUMBER = 6

        Dim TAX_PROCESSING_INDICATOR As Integer
        TAX_PROCESSING_INDICATOR = 9
        Dim FLD_LEN_PROCESSING_INDICATOR As Integer
        FLD_LEN_PROCESSING_INDICATOR = 6

        Dim TAX_COMPANY_ID As Integer
        TAX_COMPANY_ID = 4
        Dim FLD_LEN_COMPANY_ID As Integer
        FLD_LEN_COMPANY_ID = 20

        Dim TAX_SF_COUNTRY_CODE As Integer
        TAX_SF_COUNTRY_CODE = 44
        Dim FLD_LEN_SF_COUNTRY_CODE As Integer
        FLD_LEN_SF_COUNTRY_CODE = 3

        Dim TAX_SF_STATE_PROVINCE As Integer
        TAX_SF_STATE_PROVINCE = 50
        Dim FLD_LEN_SF_STATE_PROVINCE As Integer
        FLD_LEN_SF_STATE_PROVINCE = 26

        Dim TAX_SF_COUNTY_NAME As Integer
        TAX_SF_COUNTY_NAME = 76
        Dim FLD_LEN_SF_COUNTY_NAME As Integer
        FLD_LEN_SF_COUNTY_NAME = 26

        Dim TAX_SF_CITY As Integer
        FLD_LEN_SF_COUNTY_NAME = 105
        Dim FLD_LEN_SF_CITY As Integer
        FLD_LEN_SF_CITY = 26

        Dim TAX_SF_POSTAL_CODE As Integer
        TAX_SF_POSTAL_CODE = 131
        Dim FLD_LEN_SF_POSTAL_CODE As Integer
        FLD_LEN_SF_POSTAL_CODE = 9

        Dim TAX_SF_GEOCODE As Integer
        TAX_SF_GEOCODE = 144
        Dim FLD_LEN_SF_GEOCODE As Integer
        FLD_LEN_SF_GEOCODE = 2

        Dim TAX_ST_COUNTRY_CODE As Integer
        TAX_ST_COUNTRY_CODE = 146
        Dim FLD_LEN_ST_COUNTRY_CODE As Integer
        FLD_LEN_ST_COUNTRY_CODE = 3

        Dim TAX_ST_STATE_PROVINCE As Integer
        TAX_ST_STATE_PROVINCE = 152
        Dim FLD_LEN_ST_STATE_PROVINCE As Integer
        FLD_LEN_ST_STATE_PROVINCE = 26

        Dim TAX_ST_COUNTY_NAME As Integer
        TAX_ST_COUNTY_NAME = 178
        Dim FLD_LEN_ST_COUNTY_NAME As Integer
        FLD_LEN_ST_COUNTY_NAME = 26

        Dim TAX_ST_CITY As Integer
        TAX_ST_CITY = 207
        Dim FLD_LEN_ST_CITY As Integer
        FLD_LEN_ST_CITY = 26

        Dim TAX_ST_POSTAL_CODE As Integer
        TAX_ST_POSTAL_CODE = 233
        Dim FLD_LEN_ST_POSTAL_CODE As Integer
        FLD_LEN_ST_POSTAL_CODE = 9

        Dim TAX_ST_GEOCODE As Integer
        TAX_ST_GEOCODE = 246
        Dim FLD_LEN_ST_GEOCODE As Integer
        FLD_LEN_ST_GEOCODE = 2

        Dim TAX_POO_COUNTRY_CODE As Integer
        TAX_POO_COUNTRY_CODE = 344
        Dim FLD_LEN_POO_COUNTRY_CODE As Integer
        FLD_LEN_POO_COUNTRY_CODE = 3

        Dim TAX_POO_STATE_PROVINCE As Integer
        TAX_POO_STATE_PROVINCE = 350
        Dim FLD_LEN_POO_STATE_PROVINCE As Integer
        FLD_LEN_POO_STATE_PROVINCE = 26

        Dim TAX_POO_COUNTY_NAME As Integer
        TAX_POO_COUNTY_NAME = 376
        Dim FLD_LEN_POO_COUNTY_NAME As Integer
        FLD_LEN_POO_COUNTY_NAME = 26

        Dim TAX_POO_CITY As Integer
        TAX_POO_CITY = 405
        Dim FLD_LEN_POO_CITY As Integer
        FLD_LEN_POO_CITY = 26

        Dim TAX_POO_POSTAL_CODE As Integer
        TAX_POO_POSTAL_CODE = 431
        Dim FLD_LEN_POO_POSTAL_CODE As Integer
        FLD_LEN_POO_POSTAL_CODE = 9

        Dim TAX_POO_GEOCODE As Integer
        TAX_POO_GEOCODE = 444
        Dim FLD_LEN_POO_GEOCODE As Integer
        FLD_LEN_POO_GEOCODE = 2

        Dim TAX_POA_COUNTRY_CODE As Integer
        TAX_POA_COUNTRY_CODE = 446
        Dim FLD_LEN_POA_COUNTRY_CODE As Integer
        FLD_LEN_POA_COUNTRY_CODE = 3

        Dim TAX_POA_STATE_PROVINCE As Integer
        TAX_POA_STATE_PROVINCE = 452
        Dim FLD_LEN_POA_STATE_PROVINCE As Integer
        FLD_LEN_POA_STATE_PROVINCE = 26

        Dim TAX_POA_COUNTY_NAME As Integer
        TAX_POA_COUNTY_NAME = 478
        Dim FLD_LEN_POA_COUNTY_NAME As Integer
        FLD_LEN_POA_COUNTY_NAME = 26

        Dim TAX_POA_CITY As Integer
        TAX_POA_CITY = 507
        Dim FLD_LEN_POA_CITY As Integer
        FLD_LEN_POA_CITY = 26

        Dim TAX_POA_POSTAL_CODE As Integer
        TAX_POA_POSTAL_CODE = 533
        Dim FLD_LEN_POA_POSTAL_CODE As Integer
        FLD_LEN_POA_POSTAL_CODE = 9

        Dim TAX_POA_GEOCODE As Integer
        TAX_POA_GEOCODE = 546
        Dim FLD_LEN_POA_GEOCODE As Integer
        FLD_LEN_POA_GEOCODE = 2

        Dim TAX_POINT_OF_TITLE_PASSAGE As Integer
        TAX_POINT_OF_TITLE_PASSAGE = 650
        Dim FLD_LEN_POINT_OF_TITLE_PASSAGE As Integer
        FLD_LEN_POINT_OF_TITLE_PASSAGE = 1

        Dim TAX_CALCULATION_MODE As Integer
        TAX_CALCULATION_MODE = 652
        Dim FLD_LEN_CALCULATION_MODE As Integer
        FLD_LEN_CALCULATION_MODE = 1

        Dim TAX_TRANSACTION_TYPE As Integer
        TAX_TRANSACTION_TYPE = 653
        Dim FLD_LEN_TRANSACTION_TYPE As Integer
        FLD_LEN_TRANSACTION_TYPE = 1

        Dim TAX_INVOICE_DATE As Integer
        TAX_INVOICE_DATE = 657
        Dim FLD_LEN_INVOICE_DATE As Integer
        FLD_LEN_INVOICE_DATE = 8

        Dim TAX_DELIVERY_DATE As Integer
        TAX_DELIVERY_DATE = 665
        Dim FLD_LEN_DELIVERY_DATE As Integer
        FLD_LEN_DELIVERY_DATE = 8

        Dim TAX_COMMODITY_PRODUCT_CODE As Integer
        TAX_COMMODITY_PRODUCT_CODE = 675
        Dim FLD_LEN_COMMODITY_PRODUCT_CODE As Integer
        FLD_LEN_COMMODITY_PRODUCT_CODE = 25

        Dim CREDIT_ORDER_CODE As Integer
        CREDIT_ORDER_CODE = 700
        Dim FLD_LEN_CREDIT_ORDER_CODE As Integer
        FLD_LEN_CREDIT_ORDER_CODE = 1

        Dim TAX_INVOICE_NUMBER As Integer
        TAX_INVOICE_NUMBER = 819
        Dim FLD_LEN_INVOICE_NUMBER As Integer
        FLD_LEN_INVOICE_NUMBER = 20

        Dim TAX_AUDIT_FILE_INDICATOR As Integer
        TAX_AUDIT_FILE_INDICATOR = 845
        Dim FLD_LEN_AUDIT_FILE_INDICATOR As Integer
        FLD_LEN_AUDIT_FILE_INDICATOR = 1

        Dim TAX_EXEMPTION_USE_FLAG As Integer
        TAX_EXEMPTION_USE_FLAG = 846
        Dim FLD_LEN_EXEMPTION_USE_FLAG As Integer
        FLD_LEN_EXEMPTION_USE_FLAG = 1

        Dim TAX_EXEMPTION_CRITERION_FLAG As Integer
        TAX_EXEMPTION_CRITERION_FLAG = 847
        Dim FLD_LEN_EXEMPTION_CRITERION_FLAG As Integer
        FLD_LEN_EXEMPTION_CRITERION_FLAG = 1

        Dim TAX_EXEMPTION_PROCESS_FLAG As Integer
        TAX_EXEMPTION_PROCESS_FLAG = 848
        Dim FLD_LEN_EXEMPTION_PROCESS_FLAG As Integer
        FLD_LEN_EXEMPTION_PROCESS_FLAG = 1

        Dim TAX_CUSTOMER_NUMBER As Integer
        TAX_CUSTOMER_NUMBER = 849
        Dim FLD_LEN_CUSTOMER_NUMBER As Integer
        FLD_LEN_CUSTOMER_NUMBER = 20

        Dim TAX_BUSINESS_LOCATION_CODE As Integer
        TAX_BUSINESS_LOCATION_CODE = 889
        Dim FLD_LEN_BUSINESS_LOCATION_CODE As Integer
        FLD_LEN_BUSINESS_LOCATION_CODE = 13

        Dim TAX_CUSTOMER_NAME As Integer
        TAX_CUSTOMER_NAME = 1113
        Dim FLD_LEN_CUSTOMER_NAME As Integer
        FLD_LEN_CUSTOMER_NAME = 20

        Dim TAX_TAXSEL_PARM As Integer
        TAX_TAXSEL_PARM = 1133
        Dim FLD_LEN_TAXSEL_PARM As Integer
        FLD_LEN_TAXSEL_PARM = 1

        Dim TAX_FISCAL_DATE As Integer
        TAX_FISCAL_DATE = 1154
        Dim FLD_LEN_FISCAL_DATE As Integer
        FLD_LEN_FISCAL_DATE = 8

        Dim TAX_STATE_TAX_TYPE As Integer
        TAX_STATE_TAX_TYPE = 1266
        Dim FLD_LEN_STATE_TAX_TYPE As Integer
        FLD_LEN_STATE_TAX_TYPE = 1

        Dim TAX_SERVICE_INDICATOR As Integer
        TAX_SERVICE_INDICATOR = 1257
        Dim FLD_LEN_SERVICE_INDICATOR As Integer
        FLD_LEN_SERVICE_INDICATOR = 1

        Dim TAX_PRODUCT_CODE_CONVERSION As Integer
        TAX_PRODUCT_CODE_CONVERSION = 1262
        Dim FLD_LEN_PRODUCT_CODE_CONVERSION As Integer
        FLD_LEN_PRODUCT_CODE_CONVERSION = 1

        Dim TAX_NUMBER_OF_ITEMS As Integer
        TAX_NUMBER_OF_ITEMS = 2233
        Dim FLD_LEN_NUMBER_OF_ITEMS As Integer
        FLD_LEN_NUMBER_OF_ITEMS = 7

        Dim TAX_INVOICE_LINE_NUMBER As Integer
        TAX_INVOICE_LINE_NUMBER = 2240
        Dim FLD_LEN_INVOICE_LINE_NUMBER As Integer
        FLD_LEN_INVOICE_LINE_NUMBER = 5

        Dim TAX_LINE_ITEM_AMOUNT As Integer
        TAX_LINE_ITEM_AMOUNT = 2245
        Dim FLD_LEN_LINE_ITEM_AMOUNT As Integer
        FLD_LEN_LINE_ITEM_AMOUNT = 14

        Dim TAX_TAX_AMOUNT As Integer
        TAX_TAX_AMOUNT = 2259
        Dim FLD_LEN_TAX_AMOUNT As Integer
        FLD_LEN_TAX_AMOUNT = 14

        Dim TAX_DISCOUNT_AMOUNT As Integer
        TAX_DISCOUNT_AMOUNT = 2273
        Dim FLD_LEN_DISCOUNT_AMOUNT As Integer
        FLD_LEN_DISCOUNT_AMOUNT = 14

        Dim TAX_FREIGHT_AMOUNT As Integer
        TAX_FREIGHT_AMOUNT = 2287
        Dim FLD_LEN_FREIGHT_AMOUNT As Integer
        FLD_LEN_FREIGHT_AMOUNT = 14

        ' If first line, write the general fields
        Dim sNumberLines As String
        If LineNumber = 1 Then
            Dim sN2 As String
            sN2 = Trim(Str(NumberLines))
            sNumberLines = sN2.PadLeft(6, "0")
            Mid(strGenInput, TAX_RECORD_NUMBER + 1, FLD_LEN_RECORD_NUMBER) = sNumberLines
            Mid(strGenInput, TAX_PROCESSING_INDICATOR + 1, FLD_LEN_PROCESSING_INDICATOR) = "1"
        End If

        ' Compute geocode
        tempZip = Inputs("Account Ship To Postal Code")
        tempState = Inputs("Account Ship To State")
        tempCity = Inputs("Account Ship To City")
        tempCity = tempCity.ToUpper()
        If debug = "Y" Then
            mydebuglog.Debug("  << tempZip: " & tempZip)
            mydebuglog.Debug("  << tempState: " & tempState)
            mydebuglog.Debug("  << tempCity: " & tempCity)
        End If
        Try
            geoCode = GetVeraZipGeoCode(tempZip, tempCity, tempState, debug, mydebuglog)
            If debug = "Y" Then
                mydebuglog.Debug("  << geoCode: " & geoCode)
            End If
        Catch ex As Exception
            If debug = "Y" Then
                mydebuglog.Debug("  << geoCode error: " & ex.Message)
            End If
        End Try
        If geoCode = "" Then geoCode = "00"

        Mid(strLineInput, TAX_COMPANY_ID + 1, FLD_LEN_COMPANY_ID) = Inputs("Organization Id")
        Mid(strLineInput, TAX_SF_COUNTRY_CODE + 1, FLD_LEN_SF_COUNTRY_CODE) = "US "
        Mid(strLineInput, TAX_SF_STATE_PROVINCE + 1, FLD_LEN_SF_STATE_PROVINCE) = ""
        Mid(strLineInput, TAX_SF_COUNTY_NAME + 1, FLD_LEN_SF_COUNTY_NAME) = ""
        Mid(strLineInput, TAX_SF_CITY + 1, FLD_LEN_SF_CITY) = ""
        Mid(strLineInput, TAX_SF_POSTAL_CODE + 1, FLD_LEN_SF_POSTAL_CODE) = ""
        Mid(strLineInput, TAX_ST_COUNTRY_CODE + 1, FLD_LEN_ST_COUNTRY_CODE) = "US"
        Mid(strLineInput, TAX_ST_STATE_PROVINCE + 1, FLD_LEN_ST_STATE_PROVINCE) = Inputs("Account Ship To State")
        Mid(strLineInput, TAX_ST_COUNTY_NAME + 1, FLD_LEN_ST_COUNTY_NAME) = Left(Inputs("Account Ship To County"), 26)
        Mid(strLineInput, TAX_ST_CITY + 1, FLD_LEN_ST_CITY) = Left(Inputs("Account Ship To City"), 26)
        If Len(Inputs("Account Ship To Postal Code")) = 10 Then
            Inputs("Account Ship To Postal Code") = Left(Inputs("Account Ship To Postal Code"), 5) & Right(Inputs("Account Ship To Postal Code"), 4)
        End If
        Mid(strLineInput, TAX_ST_POSTAL_CODE + 1, FLD_LEN_ST_POSTAL_CODE) = Left(Inputs("Account Ship To Postal Code"), 9)
        Mid(strLineInput, TAX_SF_GEOCODE + 1, FLD_LEN_SF_GEOCODE) = Left(geoCode, 2)
        Mid(strLineInput, TAX_POO_COUNTRY_CODE + 1, FLD_LEN_POO_COUNTRY_CODE) = ""
        Mid(strLineInput, TAX_POO_STATE_PROVINCE + 1, FLD_LEN_POO_STATE_PROVINCE) = ""
        Mid(strLineInput, TAX_POO_COUNTY_NAME + 1, FLD_LEN_POO_COUNTY_NAME) = ""
        Mid(strLineInput, TAX_POO_CITY + 1, FLD_LEN_POO_CITY) = ""
        Mid(strLineInput, TAX_POO_POSTAL_CODE + 1, FLD_LEN_POO_POSTAL_CODE) = ""
        Mid(strLineInput, TAX_POA_COUNTRY_CODE + 1, FLD_LEN_POA_COUNTRY_CODE) = ""
        Mid(strLineInput, TAX_POA_STATE_PROVINCE + 1, FLD_LEN_POA_STATE_PROVINCE) = ""
        Mid(strLineInput, TAX_POA_COUNTY_NAME + 1, FLD_LEN_POA_COUNTY_NAME) = ""
        Mid(strLineInput, TAX_POA_CITY + 1, FLD_LEN_POA_CITY) = ""
        Mid(strLineInput, TAX_POA_POSTAL_CODE + 1, FLD_LEN_POA_POSTAL_CODE) = ""
        Mid(strLineInput, TAX_POINT_OF_TITLE_PASSAGE + 1, FLD_LEN_POINT_OF_TITLE_PASSAGE) = "O"
        Mid(strLineInput, TAX_CALCULATION_MODE + 1, FLD_LEN_CALCULATION_MODE) = "G"
        Mid(strLineInput, TAX_TRANSACTION_TYPE + 1, FLD_LEN_TRANSACTION_TYPE) = "2"
        Mid(strLineInput, TAX_STATE_TAX_TYPE + 1, FLD_LEN_STATE_TAX_TYPE) = "S"
        'Mid(strLineInput, TAX_INVOICE_DATE + 1, FLD_LEN_INVOICE_DATE) = Inputs("Created")
        If Inputs("Credit Order Flag") = "1" Then
            Mid(strLineInput, TAX_INVOICE_DATE + 1, FLD_LEN_INVOICE_DATE) = Inputs("Invoice Date")
            Mid(strLineInput, TAX_DELIVERY_DATE + 1, FLD_LEN_DELIVERY_DATE) = Inputs("Invoice Date")
        Else
            Mid(strLineInput, TAX_INVOICE_DATE + 1, FLD_LEN_INVOICE_DATE) = Inputs("Created")
            Mid(strLineInput, TAX_DELIVERY_DATE + 1, FLD_LEN_DELIVERY_DATE) = Inputs("Created")
        End If
        Mid(strLineInput, TAX_COMMODITY_PRODUCT_CODE + 1, FLD_LEN_COMMODITY_PRODUCT_CODE) = Inputs("Product")
        If Inputs("Credit Order Flag") = "1" Then
            Mid(strLineInput, CREDIT_ORDER_CODE + 1, FLD_LEN_CREDIT_ORDER_CODE) = Inputs("Credit Order Flag")
        Else
            Mid(strLineInput, CREDIT_ORDER_CODE + 1, FLD_LEN_CREDIT_ORDER_CODE) = " "
        End If
        Mid(strLineInput, TAX_INVOICE_NUMBER + 1, FLD_LEN_INVOICE_NUMBER) = Inputs("Quote Number")
        Mid(strLineInput, TAX_AUDIT_FILE_INDICATOR + 1, FLD_LEN_AUDIT_FILE_INDICATOR) = Inputs("Audit File Flag")
        Mid(strLineInput, TAX_EXEMPTION_USE_FLAG + 1, FLD_LEN_EXEMPTION_USE_FLAG) = "Y"
        Mid(strLineInput, TAX_EXEMPTION_CRITERION_FLAG + 1, FLD_LEN_EXEMPTION_CRITERION_FLAG) = "R"
        Mid(strLineInput, TAX_EXEMPTION_PROCESS_FLAG + 1, FLD_LEN_EXEMPTION_PROCESS_FLAG) = "1"
        Mid(strLineInput, TAX_CUSTOMER_NUMBER + 1, FLD_LEN_CUSTOMER_NUMBER) = Inputs("Account Num")
        If ShipToState <> "CO" Then Mid(strLineInput, TAX_BUSINESS_LOCATION_CODE + 1, FLD_LEN_BUSINESS_LOCATION_CODE) = "HCI"
        'Mid(strLineInput, TAX_CUSTOMER_NUMBER + 1, FLD_LEN_CUSTOMER_NUMBER) = Inputs("Account Num")
        'Mid(strLineInput, TAX_BUSINESS_LOCATION_CODE + 1, FLD_LEN_BUSINESS_LOCATION_CODE) = "HCI"
        Mid(strLineInput, TAX_CUSTOMER_NAME + 1, FLD_LEN_CUSTOMER_NAME) = Left(Inputs("Account"), 20)
        Mid(strLineInput, TAX_TAXSEL_PARM + 1, FLD_LEN_TAXSEL_PARM) = "3"
        Mid(strLineInput, TAX_FISCAL_DATE + 1, FLD_LEN_FISCAL_DATE) = ""
        Mid(strLineInput, TAX_SERVICE_INDICATOR + 1, FLD_LEN_SERVICE_INDICATOR) = ""
        Mid(strLineInput, TAX_PRODUCT_CODE_CONVERSION + 1, FLD_LEN_PRODUCT_CODE_CONVERSION) = " "
        tempQtyReq = Trim(Str(Inputs("Quantity Requested")))
        tt3 = tempQtyReq.PadLeft(FLD_LEN_NUMBER_OF_ITEMS, "0")
        Mid(strLineInput, TAX_NUMBER_OF_ITEMS + 1, FLD_LEN_NUMBER_OF_ITEMS) = tt3
        'Mid(strLineInput, TAX_NUMBER_OF_ITEMS + 1, FLD_LEN_NUMBER_OF_ITEMS) = "0000001"
        Mid(strLineInput, TAX_INVOICE_LINE_NUMBER + 1, FLD_LEN_INVOICE_LINE_NUMBER) = ""
        If Inputs("Credit Order Flag") = "1" Then
            Inputs("Total Amount") = Str(Math.Abs(Val(Inputs("Total Amount"))))
            Inputs("Freight") = Str(Math.Abs(Val(Inputs("Freight"))))
        End If
        tempTotalAmt = Trim(Str(Val(Inputs("Total Amount")) * 100))
        tt2 = tempTotalAmt.PadLeft(FLD_LEN_LINE_ITEM_AMOUNT, "0")
        If debug = "Y" Then mydebuglog.Debug("  << tempTotalAmt: " & tt2)
        Mid(strLineInput, TAX_LINE_ITEM_AMOUNT + 1, FLD_LEN_LINE_ITEM_AMOUNT) = tt2
        Mid(strLineInput, TAX_TAX_AMOUNT + 1, FLD_LEN_TAX_AMOUNT) = ""
        Mid(strLineInput, TAX_DISCOUNT_AMOUNT + 1, FLD_LEN_DISCOUNT_AMOUNT) = ""
        tempFreight = Trim(Str(Val(Inputs("Freight")) * 100))
        tf2 = tempFreight.PadLeft(FLD_LEN_FREIGHT_AMOUNT, "0")
        If debug = "Y" Then mydebuglog.Debug("  << tempFreight: " & tf2)
        Mid(strLineInput, TAX_FREIGHT_AMOUNT + 1, FLD_LEN_FREIGHT_AMOUNT) = tf2
        If debug = "Y" Then mydebuglog.Debug("  LEAVING Calc_Tax_Line FUNCTION" & vbCrLf & "  ========")

        ' If first line, return the general inputs and the first line item
        If LineNumber = 1 Then
            Calc_Tax_Line = strGenInput & strLineInput

            ' If 2nd through N lines, return just the line item
        Else
            Calc_Tax_Line = strLineInput
        End If

        ' Output inputs
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & " ====Inputs====")
            For Each psItem In Inputs
                mydebuglog.Debug(" == " & psItem.Key & " = " & psItem.Value)
            Next
            mydebuglog.Debug(" ====Inputs====" & vbCrLf)
        End If

    End Function

    Private Function Parse_Tax_Rate(ByRef mydebuglog As ILog, ByVal TaxOutputs As SimpleDictionary, _
        ByVal strReturn As String, ByVal dblGrossAmount As Double, ByVal debug As String) As Double
        ' THIS IS CALLED BY CALC_TAX
        ' INPUTS:	strReturn			Information returned by Taxware
        '			dblGrossAmount		Gross taxable order amount
        '			debug				Debug flag
        ' OUTPUTS:	Parse_Tax_Rate		The percentage tax amount

        On Error Resume Next
        Dim UTL_OUTPUT_UTL_COMPLETION_CODE As Integer
        Dim OUTPUT_LENGTH_UTL_COMPLETION_CODE As Integer
        Dim intRecordCount, intOutputSize, intCounter As Integer
        Dim strOutputData, strDataRow, strGenDesc, strLineTotal As String
        Dim UTL_OUTPUT_RECORD_COUNT As Integer
        Dim UTL_OUTPUT_GENERAL, OUTPUT_LENGTH_GENERAL As Integer
        Dim LENGTH_TAX_RATE, UTL_OUTPUT_TAX_RATE_COUNTRY, UTL_SEC_TAX_RATE_STATE As Integer
        Dim UTL_OUTPUT_TAX_RATE_COUNTY, UTL_OUTPUT_TAX_RATE_STATE, UTL_OUTPUT_TAX_RATE_CITY As Integer
        Dim UTL_SEC_TAX_RATE_COUNTY, UTL_SEC_TAX_RATE_CITY, UTL_OUTPUT_LINE_ITEM_AMOUNT As Integer
        Dim OUTPUT_LENGTH_UTL_LINE_AMOUNT, UTL_ST_TAX_AMOUNT, UTL_CT_TAX_AMOUNT As Integer
        Dim UTL_CY_TAX_AMOUNT, SEC_ST_TAX_AMOUNT, SEC_CT_TAX_AMOUNT, SEC_CY_TAX_AMOUNT As Integer
        Dim OUTPUT_LEN_TAX_AMOUNT As Integer
        Dim decTaxRate As Decimal
        Dim strTaxRate As String

        ' Set defaults
        UTL_OUTPUT_RECORD_COUNT = 0
        UTL_OUTPUT_GENERAL = 1
        OUTPUT_LENGTH_GENERAL = 4
        UTL_OUTPUT_UTL_COMPLETION_CODE = 73
        OUTPUT_LENGTH_UTL_COMPLETION_CODE = 200
        LENGTH_TAX_RATE = 7
        UTL_OUTPUT_TAX_RATE_COUNTRY = 1585
        UTL_OUTPUT_TAX_RATE_STATE = 1599
        UTL_OUTPUT_TAX_RATE_COUNTY = 1606
        UTL_OUTPUT_TAX_RATE_CITY = 1613
        UTL_SEC_TAX_RATE_STATE = 1620
        UTL_SEC_TAX_RATE_COUNTY = 1627
        UTL_SEC_TAX_RATE_CITY = 1634
        UTL_OUTPUT_LINE_ITEM_AMOUNT = 1319
        OUTPUT_LENGTH_UTL_LINE_AMOUNT = 14
        UTL_ST_TAX_AMOUNT = 1361
        UTL_CT_TAX_AMOUNT = 1375
        UTL_CY_TAX_AMOUNT = 1389
        SEC_ST_TAX_AMOUNT = 1403
        SEC_CT_TAX_AMOUNT = 1417
        SEC_CY_TAX_AMOUNT = 1431
        OUTPUT_LEN_TAX_AMOUNT = 14
        strTaxRate = ""
        Parse_Tax_Rate = 0

        ' Calculate initial variables
        If debug = "Y" Then mydebuglog.Debug("========" & vbCrLf & "ENTERING Parse_Tax_Rate FUNCTION")
        intRecordCount = Val(Mid(strReturn, UTL_OUTPUT_RECORD_COUNT + 1, OUTPUT_LENGTH_RECORD_COUNT))
        intOutputSize = UTL_OUTPUT_DATA_SIZE * intRecordCount
        strOutputData = Mid(strReturn, OUT_PARAM_SIZE + 1, intOutputSize)
        intCounter = 1
        If debug = "Y" Then
            mydebuglog.Debug(" > Size of output record: " & Str(Len(strReturn)) & " - " & Mid(strReturn, UTL_OUTPUT_RECORD_COUNT + 1, OUTPUT_LENGTH_RECORD_COUNT))
            mydebuglog.Debug(" > Output length record count: " & Str(OUTPUT_LENGTH_RECORD_COUNT))
            mydebuglog.Debug(" > Number of records: " & Str(intRecordCount))
            mydebuglog.Debug(" > Size of records: " & Str(intOutputSize))
            mydebuglog.Debug(" > Gross order amount: " & Str(dblGrossAmount))
        End If

        ' Check for address error
        strDataRow = Mid(strOutputData, 1, UTL_OUTPUT_DATA_SIZE)
        strGenDesc = Mid(strDataRow, UTL_OUTPUT_UTL_COMPLETION_CODE + 1, OUTPUT_LENGTH_UTL_COMPLETION_CODE)
        If debug = "Y" Then mydebuglog.Debug(" > Completion code: " & strGenDesc)
        If Left(strGenDesc, 19) = "No zip code passed." Then
            ' If the zipcode is invalid, cannot compute tax amount
            Parse_Tax_Rate = 0
            Exit Function
        End If

        ' Go through each returned record
        Dim strGenCode, strTaxAmount As String
        Dim decState, decCounty, decCity, decSecState, decSecCounty, decSecCity As Decimal
        Dim decTaxState, decTaxCounty, decTaxCity, decSecTaxState, decSecTaxCounty As Decimal
        Dim decSecTaxCity, decTaxTotal, decLineTotal As Decimal

        Do While intCounter <= intRecordCount
            strDataRow = Mid(strOutputData, ((intCounter - 1) * UTL_OUTPUT_DATA_SIZE) + 1, UTL_OUTPUT_DATA_SIZE)
            strGenCode = Mid(strDataRow, UTL_OUTPUT_GENERAL + 1, OUTPUT_LENGTH_GENERAL)

            decState = Mid(strDataRow, UTL_OUTPUT_TAX_RATE_STATE + 1, LENGTH_TAX_RATE) / 1000000
            decCounty = Mid(strDataRow, UTL_OUTPUT_TAX_RATE_COUNTY + 1, LENGTH_TAX_RATE) / 1000000
            decCity = Mid(strDataRow, UTL_OUTPUT_TAX_RATE_CITY + 1, LENGTH_TAX_RATE) / 1000000

            decSecState = Mid(strDataRow, UTL_SEC_TAX_RATE_STATE + 1, LENGTH_TAX_RATE) / 1000000
            decSecCounty = Mid(strDataRow, UTL_SEC_TAX_RATE_COUNTY + 1, LENGTH_TAX_RATE) / 1000000
            decSecCity = Mid(strDataRow, UTL_SEC_TAX_RATE_CITY + 1, LENGTH_TAX_RATE) / 1000000

            decTaxState = Val(Trim(Mid(strDataRow, UTL_ST_TAX_AMOUNT + 1, OUTPUT_LEN_TAX_AMOUNT))) / 100
            decTaxCounty = Val(Trim(Mid(strDataRow, UTL_CT_TAX_AMOUNT + 1, OUTPUT_LEN_TAX_AMOUNT))) / 100
            decTaxCity = Val(Trim(Mid(strDataRow, UTL_CY_TAX_AMOUNT + 1, OUTPUT_LEN_TAX_AMOUNT))) / 100

            decSecTaxState = Val(Trim(Mid(strDataRow, SEC_ST_TAX_AMOUNT + 1, OUTPUT_LEN_TAX_AMOUNT))) / 100
            decSecTaxCounty = Val(Trim(Mid(strDataRow, SEC_CT_TAX_AMOUNT + 1, OUTPUT_LEN_TAX_AMOUNT))) / 100
            decSecTaxCity = Val(Trim(Mid(strDataRow, SEC_CY_TAX_AMOUNT + 1, OUTPUT_LEN_TAX_AMOUNT))) / 100

            decTaxTotal = decTaxState + decTaxCity + decTaxCounty + decSecTaxState + decSecTaxCity + decSecTaxCounty
            strTaxAmount = Str(decTaxTotal)
            decLineTotal = Val(Mid(strDataRow, UTL_OUTPUT_LINE_ITEM_AMOUNT + 1, OUTPUT_LENGTH_UTL_LINE_AMOUNT)) / 100
            strLineTotal = Str(decLineTotal)

            TaxOutputs("tax_" & Trim(Str(intCounter))) = strTaxAmount
            If TaxOutputs("id_" & Trim(Str(intCounter))) <> "freight" Then
                TaxOutputs("total_" & Trim(Str(intCounter))) = strLineTotal
            End If

            If decLineTotal > 0 And decTaxTotal > 0 Then
                decTaxRate = decTaxTotal / strLineTotal
                strTaxRate = Format(decTaxRate, "0.0000")
                TaxOutputs("rate_" & Trim(Str(intCounter))) = strTaxRate
            Else
                TaxOutputs("rate_" & Trim(Str(intCounter))) = "0"
            End If

            If debug = "Y" Then
                mydebuglog.Debug(" << Line item total: " & strLineTotal)
                mydebuglog.Debug(" << Line tax amount: " & strTaxAmount)
                mydebuglog.Debug(" << Line tax rate: " & strTaxRate)
            End If
            intCounter = intCounter + 1
            Parse_Tax_Rate = Parse_Tax_Rate + decTaxTotal
        Loop
        If debug = "Y" Then mydebuglog.Debug(" << Returned value: " & Str(Parse_Tax_Rate) & vbCrLf & "LEAVING Parse_Tax_Rate FUNCTION" & vbCrLf & "========")
    End Function

    ' =================================================
    ' NETWORKING FUNCTIONS
    Private Function GetData(ByRef oStream As NetworkStream) As String
        Dim bResponse(1024) As Byte
        Dim sResponse As String
        sResponse = ""
        Dim lenStream As Integer = oStream.Read(bResponse, 0, 1024)
        If lenStream > 0 Then
            sResponse = Encoding.ASCII.GetString(bResponse, 0, 1024)
        End If
        Return sResponse
    End Function

    Private Function SendData(ByRef oStream As NetworkStream, ByVal sToSend As String) As String
        Dim sResponse As String
        Dim bArray() As Byte = Encoding.ASCII.GetBytes(sToSend.ToCharArray)
        oStream.Write(bArray, 0, bArray.Length())
        sResponse = GetData(oStream)
        Return sResponse
    End Function

    Private Function ValidResponse(ByVal sResult As String) As Boolean
        Dim bResult As Boolean
        Dim iFirst As Integer
        If sResult.Length > 1 Then
            iFirst = CType(sResult.Substring(0, 1), Integer)
            If iFirst < 3 Then bResult = True
        End If
        Return bResult
    End Function

    Private Function TalkToServer(ByVal oStream As NetworkStream, ByVal sToSend As String) As String
        Dim sresponse As String
        sresponse = SendData(oStream, sToSend)
        Return sresponse
    End Function

    Private Function NSLookup(ByVal sDomain As String) As String
        Dim info As New ProcessStartInfo()
        info.UseShellExecute = False
        info.RedirectStandardInput = True
        info.RedirectStandardOutput = True
        info.FileName = "nslookup"
        info.Arguments = "-type=MX " + sDomain.ToUpper.Trim

        Dim ns As Process
        ns = Process.Start(info)
        Dim sout As StreamReader
        sout = ns.StandardOutput
        Dim reg As Regex = New Regex("mail exchanger = (?<server>[^\\\s]+)")
        Dim mailserver As String = ""
        Dim response As String = ""

        Do While (sout.Peek() > -1)
            response = sout.ReadLine()
            Dim amatch As Match = reg.Match(response)
            Debug.WriteLine(response)
            If (amatch.Success) Then
                mailserver = amatch.Groups("server").Value
                Exit Do
            End If
        Loop
        Return mailserver
    End Function

    ' =================================================
    ' NUMERIC
    Public Function Round(ByVal nValue As Double, ByVal nDigits As Integer) As Double
        Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
    End Function

    ' =================================================
    ' XML DOCUMENT MANAGEMENT
    Private Sub AddXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
        root.AppendChild(resultsItem)
    End Sub

    Private Function GetNodeValue(ByVal sNodeName As String, ByVal oParentNode As XmlNode) As String
        ' Generic function to return the value of a node in an XML document
        Dim oNode As XmlNode = oParentNode.SelectSingleNode(".//" + sNodeName)
        If oNode Is Nothing Then
            Return String.Empty
        Else
            Return oNode.InnerText
        End If
    End Function

    ' =================================================
    ' COLLECTIONS 
    ' This class implements a simple dictionary using an array of DictionaryEntry objects (key/value pairs).
    Public Class SimpleDictionary
        Implements IDictionary

        ' The array of items
        Dim items() As DictionaryEntry
        Dim ItemsInUse As Integer = 0

        ' Construct the SimpleDictionary with the desired number of items.
        ' The number of items cannot change for the life time of this SimpleDictionary.
        Public Sub New(ByVal numItems As Integer)
            items = New DictionaryEntry(numItems - 1) {}
        End Sub

        ' IDictionary Members
        Public ReadOnly Property IsReadOnly() As Boolean Implements IDictionary.IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Function Contains(ByVal key As Object) As Boolean Implements IDictionary.Contains
            Dim index As Integer
            Return TryGetIndexOfKey(key, index)
        End Function

        Public ReadOnly Property IsFixedSize() As Boolean Implements IDictionary.IsFixedSize
            Get
                Return False
            End Get
        End Property

        Public Sub Remove(ByVal key As Object) Implements IDictionary.Remove
            If key = Nothing Then
                Throw New ArgumentNullException("key")
            End If
            ' Try to find the key in the DictionaryEntry array
            Dim index As Integer
            If TryGetIndexOfKey(key, index) Then

                ' If the key is found, slide all the items up.
                Array.Copy(items, index + 1, items, index, (ItemsInUse - index) - 1)
                ItemsInUse = ItemsInUse - 1
            Else

                ' If the key is not in the dictionary, just return. 
            End If
        End Sub

        Public Sub Clear() Implements IDictionary.Clear
            ItemsInUse = 0
        End Sub

        Public Sub Add(ByVal key As Object, ByVal value As Object) Implements IDictionary.Add

            ' Add the new key/value pair even if this key already exists in the dictionary.
            If ItemsInUse = items.Length Then
                Throw New InvalidOperationException("The dictionary cannot hold any more items.")
            End If
            items(ItemsInUse) = New DictionaryEntry(key, value)
            ItemsInUse = ItemsInUse + 1
        End Sub

        Public ReadOnly Property Keys() As ICollection Implements IDictionary.Keys
            Get

                ' Return an array where each item is a key.
                ' Note: Declaring keyArray() to have a size of ItemsInUse - 1
                '       ensures that the array is properly sized, in VB.NET
                '       declaring an array of size N creates an array with
                '       0 through N elements, including N, as opposed to N - 1
                '       which is the default behavior in C# and C++.
                Dim keyArray() As Object = New Object(ItemsInUse - 1) {}
                Dim n As Integer
                For n = 0 To ItemsInUse - 1
                    keyArray(n) = items(n).Key
                Next n

                Return keyArray
            End Get
        End Property

        Public ReadOnly Property Values() As ICollection Implements IDictionary.Values
            Get
                ' Return an array where each item is a value.
                Dim valueArray() As Object = New Object(ItemsInUse - 1) {}
                Dim n As Integer
                For n = 0 To ItemsInUse - 1
                    valueArray(n) = items(n).Value
                Next n

                Return valueArray
            End Get
        End Property

        Default Public Property Item(ByVal key As Object) As Object Implements IDictionary.Item
            Get

                ' If this key is in the dictionary, return its value.
                Dim index As Integer
                If TryGetIndexOfKey(key, index) Then

                    ' The key was found return its value.
                    Return items(index).Value
                Else

                    ' The key was not found return null.
                    Return Nothing
                End If
            End Get

            Set(ByVal value As Object)
                ' If this key is in the dictionary, change its value. 
                Dim index As Integer
                If TryGetIndexOfKey(key, index) Then

                    ' The key was found change its value.
                    items(index).Value = value
                Else

                    ' This key is not in the dictionary add this key/value pair.
                    Add(key, value)
                End If
            End Set
        End Property

        Private Function TryGetIndexOfKey(ByVal key As Object, ByRef index As Integer) As Boolean
            For index = 0 To ItemsInUse - 1
                ' If the key is found, return true (the index is also returned).
                If items(index).Key.Equals(key) Then
                    Return True
                End If
            Next index

            ' Key not found, return false (index should be ignored by the caller).
            Return False
        End Function

        Private Class SimpleDictionaryEnumerator
            Implements IDictionaryEnumerator

            ' A copy of the SimpleDictionary object's key/value pairs.
            Dim items() As DictionaryEntry
            Dim index As Integer = -1

            Public Sub New(ByVal sd As SimpleDictionary)
                ' Make a copy of the dictionary entries currently in the SimpleDictionary object.
                items = New DictionaryEntry(sd.Count - 1) {}
                Array.Copy(sd.items, 0, items, 0, sd.Count)
            End Sub

            ' Return the current item.
            Public ReadOnly Property Current() As Object Implements IDictionaryEnumerator.Current
                Get
                    ValidateIndex()
                    Return items(index)
                End Get
            End Property

            ' Return the current dictionary entry.
            Public ReadOnly Property Entry() As DictionaryEntry Implements IDictionaryEnumerator.Entry
                Get
                    Return Current
                End Get
            End Property

            ' Return the key of the current item.
            Public ReadOnly Property Key() As Object Implements IDictionaryEnumerator.Key
                Get
                    ValidateIndex()
                    Return items(index).Key
                End Get
            End Property

            ' Return the value of the current item.
            Public ReadOnly Property Value() As Object Implements IDictionaryEnumerator.Value
                Get
                    ValidateIndex()
                    Return items(index).Value
                End Get
            End Property

            ' Advance to the next item.
            Public Function MoveNext() As Boolean Implements IDictionaryEnumerator.MoveNext
                If index < items.Length - 1 Then
                    index = index + 1
                    Return True
                End If

                Return False
            End Function

            ' Validate the enumeration index and throw an exception if the index is out of range.
            Private Sub ValidateIndex()
                If index < 0 Or index >= items.Length Then
                    Throw New InvalidOperationException("Enumerator is before or after the collection.")
                End If
            End Sub

            ' Reset the index to restart the enumeration.
            Public Sub Reset() Implements IDictionaryEnumerator.Reset
                index = -1
            End Sub

        End Class

        Public Function GetEnumerator() As IDictionaryEnumerator Implements IDictionary.GetEnumerator

            'Construct and return an enumerator.
            Return New SimpleDictionaryEnumerator(Me)
        End Function


        ' ICollection Members
        Public ReadOnly Property IsSynchronized() As Boolean Implements IDictionary.IsSynchronized
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SyncRoot() As Object Implements IDictionary.SyncRoot
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public ReadOnly Property Count() As Integer Implements IDictionary.Count
            Get
                Return ItemsInUse
            End Get
        End Property

        Public Sub CopyTo(ByVal array As Array, ByVal index As Integer) Implements IDictionary.CopyTo
            Throw New NotImplementedException()
        End Sub

        ' IEnumerable Members
        Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator

            ' Construct and return an enumerator.
            Return Me.GetEnumerator()
        End Function
    End Class

    ' =================================================
    ' CRYPTOGRAPHY CLASS
    Friend Class cTripleDES

        ' define the triple des provider
        Private m_des As New TripleDESCryptoServiceProvider

        ' define the string handler
        Private m_utf8 As New UTF8Encoding

        ' define the local property arrays
        Private m_key() As Byte
        Private m_iv() As Byte

        Public Sub New(ByVal key() As Byte, ByVal iv() As Byte)
            Me.m_key = key
            Me.m_iv = iv
        End Sub

        Public Function Encrypt(ByVal input() As Byte) As Byte()
            Return Transform(input, m_des.CreateEncryptor(m_key, m_iv))
        End Function

        Public Function Decrypt(ByVal input() As Byte) As Byte()
            Return Transform(input, m_des.CreateDecryptor(m_key, m_iv))
        End Function

        Public Function Encrypt(ByVal text As String) As String
            Dim input() As Byte = m_utf8.GetBytes(text)
            Dim output() As Byte = Transform(input, m_des.CreateEncryptor(m_key, m_iv))
            Return Convert.ToBase64String(output)
        End Function

        Public Function Decrypt(ByVal text As String) As String
            Dim input() As Byte = Convert.FromBase64String(text)
            Dim output() As Byte = Transform(input, m_des.CreateDecryptor(m_key, m_iv))
            Return m_utf8.GetString(output)
        End Function

        Private Function Transform(ByVal input() As Byte, _
            ByVal CryptoTransform As ICryptoTransform) As Byte()

            ' create the necessary streams
            Dim memStream As MemoryStream = New MemoryStream
            Dim cryptStream As CryptoStream = New CryptoStream(memStream, CryptoTransform, CryptoStreamMode.Write)

            ' transform the bytes as requested
            cryptStream.Write(input, 0, input.Length)
            cryptStream.FlushFinalBlock()

            ' Read the memory stream and convert it back into byte array
            memStream.Position = 0
            Dim result(CType(memStream.Length - 1, System.Int32)) As Byte
            memStream.Read(result, 0, CType(result.Length, System.Int32))

            ' close and release the streams
            memStream.Close()
            cryptStream.Close()

            ' hand back the encrypted buffer
            Return result
        End Function
    End Class

    ' =================================================
    ' STRING FUNCTIONS
    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        ' Validate email address

        'Dim pattern As String = "^[a-z0-9,!#\$%&'\*\+/=\?\^_`\{\|}~-]+(\.[a-z0-9,!#\$%&'\*\+/=\?\^_`\{\|}~-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*\.([a-z]{2,})$"
        Dim pattern As String = "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
        Dim emailAddressMatch As Match = Regex.Match(LCase(Trim(emailAddress)), pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    Function SqlString(ByVal Instring As String) As String
        ' Make a string safe for use in a SQL query
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            SqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 1) = "'" Then
                outstring = outstring & "''"
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        SqlString = outstring
    End Function

    Function CheckNull(ByVal Instring As String) As String
        ' Check to see if a string is null
        If Instring Is Nothing Then
            CheckNull = ""
        Else
            CheckNull = Instring
        End If
    End Function

    Public Function CheckDBNull(ByVal obj As Object, _
    Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
        ' Checks an object to determine if its null, and if so sets it to a not-null empty value
        Dim objReturn As Object
        objReturn = obj
        If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
            objReturn = ""
        ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
            objReturn = 0
        ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
            objReturn = 0.0
        End If
        Return objReturn
    End Function

    Public Function NumString(ByVal strString As String) As String
        ' Remove everything but numbers from a string
        Dim bln As Boolean
        Dim i As Integer
        Dim iv As String
        NumString = ""

        'Can array element be evaluated as a number?
        For i = 1 To Len(strString)
            iv = Mid(strString, i, 1)
            bln = IsNumeric(iv)
            If bln Then NumString = NumString & iv
        Next

    End Function

    Public Function ToBase64(ByVal data() As Byte) As String
        ' Encode a Base64 string
        If data Is Nothing Then Throw New ArgumentNullException("data")
        Return Convert.ToBase64String(data)
    End Function

    Public Function FromBase64(ByVal base64 As String) As Byte()
        ' Decode a Base64 string
        If base64 Is Nothing Then Throw New ArgumentNullException("base64")
        Return Convert.FromBase64String(base64)
    End Function

    Public Function BitmapToBase64(ByVal image As System.Drawing.Bitmap) As String
        ' Convert a bitmap to a base64 string
        Dim base64 As String
        Dim memory As New System.IO.MemoryStream()
        image.Save(memory, Imaging.ImageFormat.Jpeg)
        base64 = System.Convert.ToBase64String(memory.ToArray)
        memory.Close()
        memory = Nothing
        Return base64
    End Function

    Public Function BitmapFromBase64(ByVal base64 As String) As System.Drawing.Bitmap
        ' Convert a base64 string to a bitmap image
        Dim oBitmap As System.Drawing.Bitmap
        Dim memory As New System.IO.MemoryStream(Convert.FromBase64String(base64))
        oBitmap = New System.Drawing.Bitmap(memory)
        memory.Close()
        memory = Nothing
        Return oBitmap
    End Function

    Function DeSqlString(ByVal Instring As String) As String
        ' Convert a string from SQL query encoded to non-encoded
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        CheckDBNull(Instring, enumObjectType.StrType)
        If Len(Instring) = 0 Then
            DeSqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 2) = "''" Then
                outstring = outstring & "'"
                i = i + 1
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        DeSqlString = outstring
    End Function

    Public Function StringToBytes(ByVal str As String) As Byte()
        ' Convert a random string to a byte array
        ' e.g. "abcdefg" to {a,b,c,d,e,f,g}
        Dim s As Char()
        s = str.ToCharArray
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            b(i) = Convert.ToByte(s(i))
        Next
        Return b
    End Function

    Public Function NumStringToBytes(ByVal str As String) As Byte()
        ' Convert a string containing numbers to a byte array
        ' e.g. "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16" to 
        '  {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}
        Dim s As String()
        s = str.Split(" ")
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            b(i) = Convert.ToByte(s(i))
        Next
        Return b
    End Function

    Public Function BytesToString(ByVal b() As Byte) As String
        ' Convert a byte array to a string
        Dim i As Integer
        Dim s As New System.Text.StringBuilder()
        For i = 0 To b.Length - 1
            Console.WriteLine(b(i))
            If i <> b.Length - 1 Then
                s.Append(b(i) & " ")
            Else
                s.Append(b(i))
            End If
        Next
        Return s.ToString
    End Function

    ' =================================================
    ' DATABASE FUNCTIONS
    Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
        ' Function to open a database connection with extreme error-handling
        ' Returns an error message if unable to open the connection
        Dim SqlS As String
        SqlS = ""
        OpenDBConnection = ""

        Try
            con = New SqlConnection(ConnS)
            con.Open()
            If Not con Is Nothing Then
                Try
                    cmd = New SqlCommand(SqlS, con)
                    cmd.CommandTimeout = 300
                Catch ex2 As Exception
                    OpenDBConnection = "Error opening the command string: " & ex2.ToString
                End Try
            End If
        Catch ex As Exception
            If con.State <> Data.ConnectionState.Closed Then con.Dispose()
            ConnS = ConnS & ";Pooling=false"
            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex2 As Exception
                OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
            End Try
        End Try

    End Function

    ' =================================================
    ' DEBUG FUNCTIONS
    Public Sub writeoutput(ByVal fs As StreamWriter, ByVal instring As String)
        ' This function writes a line to a previously opened streamwriter, and then flushes it
        ' promptly.  This assists in debugging services
        fs.WriteLine(instring)
        fs.Flush()
    End Sub

    Public Sub writeoutputfs(ByVal fs As FileStream, ByVal instring As String)
        ' This function writes a line to a previously opened filestream, and then flushes it
        ' promptly.  This assists in debugging services
        fs.Write(StringToBytes(instring), 0, Len(instring))
        fs.Write(StringToBytes(vbCrLf), 0, 2)
        fs.Flush()
    End Sub

    ' =================================================
    ' Time functions
    Private Sub UpdateDuration(ByVal START_HR As String, ByVal START_MIN As String, _
              ByVal END_HR As String, ByVal END_MIN As String, ByVal DUR_HOUR As String, _
              ByVal DUR_MIN As String, ByVal START_AP As String, ByVal END_AP As String)
        '
        ' This function calculates the duration of the session by first converting to a decimal 24-hour
        ' clock, finding the difference, and then converting the difference back to non-decimal time
        '
        Dim SH, SM, EM, EH As Integer
        Dim START24, END24, dur24, durhr, durmin As Integer

        ' Validate input and convert to integer values
        START_MIN = NumString(START_MIN)
        If (START_MIN = "0" Or CheckDBNull(START_MIN, enumObjectType.StrType) = "0" Or Trim(START_MIN) = "") Then
            SM = 0
        Else
            SM = Val(START_MIN)
        End If
        END_MIN = NumString(END_MIN)
        If (END_MIN = "0" Or CheckDBNull(END_MIN, enumObjectType.StrType) = "0" Or Trim(END_MIN) = "") Then
            EM = 0
        Else
            EM = Val(END_MIN)
        End If
        START_HR = NumString(START_HR)
        If (START_HR = "0" Or CheckDBNull(START_HR, enumObjectType.StrType) = "0" Or Trim(START_HR) = "") Then
            SH = 0
        Else
            SH = Val(START_HR)
        End If
        END_HR = NumString(END_HR)
        If (END_HR = "0" Or CheckDBNull(END_HR, enumObjectType.StrType) = "0" Or Trim(END_HR) = "") Then
            EH = 0
        Else
            EH = Val(START_MIN)
        End If

        ' Convert to 24 hour time
        If SH = 12 Then
            If START_AP = "P" Then
                START24 = 12 + (SM / 60)
            Else
                START24 = (SM / 60)
            End If
        Else
            START24 = IIf(START_AP = "P", 12, 0) + SH + (SM / 60)
        End If

        If EH = 12 Then
            If END_AP = "P" Then
                END24 = 12 + (EM / 60)
            Else
                END24 = (EM / 60)
            End If
        Else
            END24 = IIf(END_AP = "P", 12, 0) + EH + (EM / 60)
        End If

        ' Correct end time if necessary
        If END24 < START24 Then
            END24 = END24 + 24
        End If

        ' Compute time
        dur24 = END24 - START24
        durhr = Int(dur24)
        durmin = Int((dur24 - durhr) * 60)
        DUR_HOUR = Str(durhr)
        DUR_MIN = Str(durmin)
    End Sub

End Class