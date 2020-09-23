Attribute VB_Name = "ModCGI"
Option Explicit
'######################################################
'##  Caedo HTTPd is (C) 2006 Steven Dorman - ARR     ##
'##                                                  ##
'##  This program has been released under the GPL    ##
'##  and is Open-Source Software.                    ##
'##                                                  ##
'######################################################
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'//Make new dosoutput variable
Private dos As New dosoutputs

'//This function will query the PHP Executeable and return its outputs
'//In turn , processing the PHP script
Public Function RunPHP(filename As String, Index As Integer, Optional GetString As String, Optional Cookies As String) As String
On Error GoTo hell:
Dim cmdline As String, fullfile As String

If GetString = "" Then
    cmdline = "" & frmMain.PHPPath & " " & filename & ""
ElseIf GetString <> "" And Cookies = "" Then
    cmdline = "" & frmMain.PHPPath & " " & filename & " -- " & GetString
ElseIf GetString <> "" And Cookies <> "" Then
    cmdline = "" & frmMain.PHPPath & " " & filename & " -- " & GetString & " " & Cookies
End If
        
RunPHP = dos.ExecuteCommand(cmdline)
DoEvents

Exit Function
hell:
Reset
frmMain.LogWrite "Error Processing PHP CGI: " & Err.Description
End Function

'//This function will query the PERL Executeable and return its outputs
'//In turn , processing the PERL script
Public Function RunPERL(filename As String, Index As Integer, Optional GetString As String) As String
On Error GoTo hell:
Dim cmdline As String

cmdline = frmMain.PERLPath & " " & filename
RunPERL = dos.ExecuteCommand(cmdline)
DoEvents

Exit Function
hell:
Reset
frmMain.LogWrite "Error Processing PERL CGI: " & Err.Description
End Function
'//This function will query the Custom CGI Executeable and return its outputs
'//In turn , processing the Custom CGI script
Public Function RunxCGI(filename As String, Index As Integer, Optional GetString As String) As String
On Error GoTo hell:
Dim cmdline As String

cmdline = frmMain.xCGIPath
cmdline = Replace(cmdline, "#File", filename)
cmdline = Replace(cmdline, "#Get", frmMain.webinfo.FormData)
cmdline = Replace(cmdline, "#Post", "Sorry this is not currently Supported")
cmdline = Replace(cmdline, "#Header", GetString)

RunxCGI = dos.ExecuteCommand(cmdline)
DoEvents

Exit Function
hell:
Reset
frmMain.LogWrite "Error Processing Custom CGI: " & Err.Description
End Function
'//Proccess SSI Pages
Public Function SSI(inData As String, page As String, Optional SSIPath As String, Optional Index As Integer) As String
Dim iPart As Long, lPart As Long, iCnt As Long, StrSSI As String, TStr As String _
, tCmd As String, tVarname As String, outData As String, strBuff As Variant, sTemp As String

On Error Resume Next

    outData = inData
    strBuff = Split(inData, vbNewLine)
    For iCnt = LBound(strBuff) To UBound(strBuff)
        iPart = InStr(strBuff(iCnt), "<!--#")       ' Find the start of the ssi tag
        lPart = InStr(iPart, strBuff(iCnt), "-->")  ' Find the end part of the sii tag
        TStr = Mid(strBuff(iCnt), iPart + 5, lPart - iPart - 5)
        tCmd = Trim(UCase$(Mid(TStr, 1, InStr(TStr, " ")))) ' tells us waht command it is

        If tCmd = "ECHO" Then ' So now we now that here it is a ECHO variable
           iPart = InStr(TStr, "=") ' The start of the variable name
            lPart = InStr(iPart + 2, TStr, Chr(34)) ' Get the end of the variable name
            tVarname = Mid(TStr, iPart + 2, lPart - iPart - 2) ' Get the variable name

            StrSSI = "<!--#" & TStr & "-->" ' This is what we need to replace in the file

            Select Case tVarname
                Case "DATE_LOCAL"
                    outData = Replace(outData, StrSSI, Format(Date, "mmm dd yyyy"))
                Case "SERVER_SOFTWARE"
                    outData = Replace(outData, StrSSI, "Caedo HTTPd WebServer V." & App.Major & "." & App.Minor)
                Case "SERVER_PORT"
                    outData = Replace(outData, StrSSI, frmMain.Port)
                Case "SERVER_PROTOCOL"
                    outData = Replace(outData, StrSSI, frmMain.webinfo.HTTPVer)
                Case "SERVER_NAME"
                    outData = Replace(outData, StrSSI, GetSystemName)
                Case "DOCUMENT_NAME"
                    page = frmMain.webinfo.page
                    outData = Replace(outData, StrSSI, page)
                Case "DOCUMENT_URI"
                    page = frmMain.webinfo.Host & frmMain.webinfo.page 'TODO
                    outData = Replace(outData, StrSSI, page)
                Case "HTTP_USER_AGENT"
                    outData = Replace(outData, StrSSI, frmMain.webinfo.Agent)
                Case "HTTP_REFERER"
                    outData = Replace(outData, StrSSI, frmMain.webinfo.URLReferer)
                Case "REQUEST_METHOD"
                    outData = Replace(outData, StrSSI, frmMain.webinfo.RequestType)
                Case "REMOTE_ADDR"
                    outData = Replace(outData, StrSSI, frmMain.sockets(Index).RemoteHostIP)
                Case "CONTENT_TYPE"
                    outData = Replace(outData, StrSSI, "text/html") 'TODO
                Case "CONTENT_LENGTH"
                    outData = Replace(outData, StrSSI, FileLen(page))
                Case "HTTP_ACCEPT_LANGUAGE"
                   outData = Replace(outData, StrSSI, frmMain.webinfo.Language)
                Case "HTTP_ACCEPT_ENCODING"
                    outData = Replace(outData, StrSSI, frmMain.webinfo.Encoding)
                Case "HTTP_CONNECTION"
                    outData = Replace(outData, StrSSI, frmMain.webinfo.Connection)
                Case Else
                    outData = Replace(outData, StrSSI, "Unknown command")
            End Select

        ElseIf tCmd = "INCLUDE" Then
            iPart = InStr(TStr, "=")
            lPart = InStr(iPart + 2, TStr, Chr(34))
            tVarname = Trim(Mid(TStr, iPart + 2, lPart - iPart - 2))

            If Not Left(tVarname, 1) = "\" Then
                sTemp = OpenFile(SSIPath & "\" & tVarname)
            Else
                sTemp = OpenFile(tVarname)
            End If

            StrSSI = "<!--#" & TStr & "-->"
            outData = Replace(outData, StrSSI, sTemp)
            iPart = 0: lPart = 0: sTemp = ""
        ElseIf tCmd = "FSIZE" Then
            iPart = InStr(TStr, "=")
            lPart = InStr(iPart + 2, TStr, Chr(34))
            tVarname = Trim(Mid(TStr, iPart + 2, lPart - iPart - 2))

            If Not Left(tVarname, 1) = "\" Then
                sTemp = SSIPath & "\" & tVarname
            Else
                sTemp = tVarname
           End If

            StrSSI = "<!--#" & TStr & "-->"
            outData = Replace(outData, StrSSI, FileLen(sTemp))
            iPart = 0: lPart = 0
        End If

    Next

    SSI = outData
    StrSSI = ""
   
End Function
Public Function GetSystemName() As String
Dim iRet As Long, CompName As String
    CompName = Space$(128) ' Create a buffer to hold the computer name
    iRet = GetComputerName(CompName, 128) ' Get the computername
    If iRet = 0 Then GetSystemName = "": Exit Function ' Exit the function
    GetSystemName = Left$(CompName, InStr(CompName, Chr(0)) - 1)
    ' The above  trims any spaces form the string and returns the computername
    CompName = "" ' Clean out the buffer
End Function

'//Simple function to read all the data from a file into memory
Public Function OpenFile(lzFile As String) As String
Dim iFile As Long
Dim nByte() As Byte
   
    iFile = FreeFile
    Open lzFile For Binary As #iFile
        ReDim nByte(0 To LOF(iFile) - 1)
        Get #iFile, , nByte
    Close #iFile
    ' Clean up
    OpenFile = StrConv(nByte, vbUnicode)
    Erase nByte
    
End Function

'//Function that creates the temp data for PHP files for POST data
Public Function AddVars(pdata As String, filedata As String) As String
Dim tmp() As String, tmp2 As String, tmp3 As String, I As Integer, out As String
tmp = Split(pdata, "&")
out = "<?PHP " & vbCrLf

For I = 0 To UBound(tmp)
    tmp2 = Split(tmp(I), "=")(0)
    tmp3 = Split(tmp(I), "=")(1)

    out = out & "$_POST['" & tmp2 & "']=""" & tmp3 & """;" & vbCrLf
Next I

out = out & " ?>"
AddVars = out & filedata
End Function

'//Simple function to Decode URL's.
Public Function URLDecode(sEncodedURL As String) As String
    On Error GoTo Catch
    
    Dim iLoop As Integer
    Dim sRtn As String
    Dim sTmp As String
    


    If Len(sEncodedURL) > 0 Then
        ' Loop through each char


        For iLoop = 1 To Len(sEncodedURL)
            sTmp = Mid(sEncodedURL, iLoop, 1)
            sTmp = Replace(sTmp, "+", " ")
            ' If char is % then get next two chars
            ' and convert from HEX to decimal


            If sTmp = "%" And Len(sEncodedURL) + 1 > iLoop + 2 Then
                sTmp = Mid(sEncodedURL, iLoop + 1, 2)
                sTmp = Chr(CDec("&H" & sTmp))
                ' Increment loop by 2
                iLoop = iLoop + 2
            End If
            sRtn = sRtn & sTmp
        Next iLoop
        URLDecode = sRtn
    End If
Finally:
    Exit Function
Catch:
    URLDecode = ""
    Resume Finally
End Function

