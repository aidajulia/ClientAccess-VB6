Attribute VB_Name = "WebServiceModule"
Public Type ReturnData
    ErrorCode As Integer
    ErrorDescription As String
    Server As String
    DBname As String
    UID As String
    PWD As String
    ConnectionString As String
End Type

Public Function GetConnection(clientID As Integer, applicationID As Integer) As ReturnData
    Dim R As ReturnData
    Dim RetString As String
    Dim strXml As String

    Const strUrl As String = "http://69.48.141.109/ClientAccess/ClientAccess.asmx"
    Const strSoapAction As String = "http://TaxiMagic.Com/ClientAccess/ValidateApplicationLegacy"
    Const secKey As String = "b1082abe-bf8c-4f91-97c6-c47934cbde01"
   

    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>"
    strXml = strXml & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
    strXml = strXml & "<soap:Body>"
    strXml = strXml & "<ValidateApplicationLegacy xmlns=""http://TaxiMagic.Com/ClientAccess"">"
    strXml = strXml & "<clientid>" & clientID & "</clientid>"
    strXml = strXml & "<applicationid>" & applicationID & "</applicationid>"
    strXml = strXml & "<seckey>" & secKey & "</seckey>"
    strXml = strXml & "</ValidateApplicationLegacy>"
    strXml = strXml & "</soap:Body>"
    strXml = strXml & "</soap:Envelope>"
    RetString = PostWebservice(strUrl, strSoapAction, strXml)
    If Left(RetString, 6) = "ERROR:" Then
        R.ErrorCode = "999"
        R.ErrorDescription = RetString
    Else
        ErrorCode = ReadValue(RetString, "errorcode")
        If Val(ErrorCode) <> 0 Then
            R.ErrorCode = ErrorCode
            R.ErrorDescription = ReadValue(RetString, "errordescription")
        Else
            R.Server = ReadValue(RetString, "dbserverip")
            R.DBname = ReadValue(RetString, "dbname")
            R.UID = ReadValue(RetString, "dbuid")
            R.PWD = ReadValue(RetString, "dbpwd")
            R.ConnectionString = "PROVIDER = MSDASQL; DRIVER={SQL Server}; DATABASE=" & R.DBname & "; SERVER=" & R.Server & "; UID=" & R.UID & "; PWD=" & R.PWD & ";"
        End If
    End If
    GetConnection = R
End Function
Private Function ReadValue(inputstr As String, key As String)
Dim retStart As Integer
Dim retEnd As Integer
If Len(inputstr) = 0 Then
    ReadValue = ""
    Exit Function
End If
retStart = InStr(1, inputstr, "<" & key & ">", vbTextCompare) + Len("<" & key & ">")
retEnd = InStr(1, inputstr, "</" & key & ">", vbTextCompare)
If retStart < 0 Then
    ReadValue = ""
    Exit Function
End If
If retEnd < 0 Then
    ReadValue = ""
    Exit Function
End If

ReadValue = Mid(inputstr, retStart, retEnd - retStart)




End Function


Private Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Load XML
    objDom.async = False
    objDom.loadXML XmlBody

    ' Open the webservice
    objXmlHttp.open "POST", AsmxUrl, False
    
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responseText
    
    ' Close object
    Set objXmlHttp = Nothing
    
    ' Return result
    PostWebservice = strRet
    
Exit Function
Err_PW:
    PostWebservice = "ERROR: " & Err.Number & " - " & Err.Description

End Function

