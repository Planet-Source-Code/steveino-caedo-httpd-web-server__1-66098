VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "webclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################
'##  Caedo HTTPd is (C) 2006 Steven Dorman - ARR     ##
'##                                                  ##
'##  This program has been released under the GPL    ##
'##  and is Open-Source Software.                    ##
'##                                                  ##
'######################################################

'------------------------------------------------'
' Code from DM HTTP Easy Weberver                '
' By Ben Jones (vbdream2k@yahoo.com)             '
' Many thanks to him for letting me use his code '
'------------------------------------------------'

Private Type w_head
    mAccept As String
    mRequest As String
    mCookie As String
    mReferer As String
    mPage As String
    mAcceptLan As String
    mHost As String
    mEncoding As String
    mHTTPVer As String
    mAgent As String
    mContentLen As Long
    mFormData As String
    mConnection As String
    mAuthorization As String
End Type

Private WebGet As w_head

Public Function PhaseWebHead(Strdata As String) As Boolean
Dim Base64Str As String
On Error Resume Next

Dim iPart As Integer, lPart As Integer
    iPart = InStr(Strdata, " ")
    WebGet.mRequest = Mid(Strdata, 1, iPart - 1)
    ' now we need to get the requested resource
    lPart = InStr(iPart + 1, Strdata, " ")
    WebGet.mPage = Mid(Strdata, iPart + 1, lPart - iPart - 1)
    WebGet.mFormData = Split(WebGet.mPage, "?")(1)
    WebGet.mPage = Split(WebGet.mPage, "?")(0)
 
    ' Now we move on to get the HTTP Verision
    iPart = InStr(lPart + 1, Strdata, vbCrLf)
    WebGet.mHTTPVer = Trim(Mid(Strdata, lPart, iPart - lPart))
    ' Now we will move on to get the clients Agent
    iPart = InStr(1, Strdata, "User-Agent:")
    lPart = InStr(iPart, Strdata, vbCrLf)
    WebGet.mAgent = Trim(Mid(Strdata, iPart + 11, lPart - iPart - 11))
    ' Now we need to get the host name
    iPart = InStr(1, Strdata, "Host:", vbTextCompare)
    lPart = InStr(iPart + 1, Strdata, vbCrLf, vbTextCompare)
    WebGet.mHost = Trim(Mid(Strdata, iPart + 5, lPart - iPart - 5))
    ' Now we now need to get the Referer URL
    iPart = InStr(1, Strdata, "Referer:")
    lPart = InStr(iPart, Strdata, vbCrLf)
    If iPart = 0 Or lPart = 0 Then
        WebGet.mReferer = ""
    Else
        WebGet.mReferer = Trim(Mid(Strdata, iPart + 8, lPart - iPart - 8))
    ' This gets the accept request
    End If
    
    iPart = InStr(1, Strdata, "Accept:")
    lPart = InStr(iPart, Strdata, vbCrLf)
    WebGet.mAccept = Trim$(Mid$(Strdata, iPart + 7, lPart - iPart - 7))
    ' This gets the current language
    iPart = InStr(1, Strdata, "Accept-Language:")
    lPart = InStr(iPart, Strdata, vbCrLf)
    WebGet.mAcceptLan = Trim(Mid(Strdata, iPart + 16, lPart - iPart - 16))
    ' This gets the current encoding type
    iPart = InStr(1, Strdata, "Accept-Encoding:")
    lPart = InStr(iPart, Strdata, vbCrLf)
    WebGet.mEncoding = Trim(Mid(Strdata, iPart + 16, lPart - iPart - 16))
    ' Get the connection state
    iPart = InStr(1, Strdata, "Connection:")
    lPart = InStr(iPart, Strdata, vbCrLf)
    WebGet.mConnection = Trim(Mid(Strdata, iPart + 11, lPart - iPart - 11))
    ' This will now get the Autorizantion value
    iPart = InStr(1, Strdata, "Authorization: ", vbTextCompare)
    lPart = InStr(iPart, Strdata, vbCrLf)
    If iPart > 0 And iPart > 0 Then Base64Str = Trim(Mid(Strdata, iPart + 15, lPart - iPart - 15))
    ' now to finish off we need to extract the base64 string
    iPart = InStr(1, Base64Str, Chr(32), vbTextCompare)
    WebGet.mAuthorization = Trim(Mid(Base64Str, iPart + 1, Len(Base64Str) - iPart))
    Base64Str = ""
    ' This gets any cookie stuff
    iPart = InStr(1, Strdata, "Cookie:")
    lPart = InStr(Strdata, iPart, vbCrLf)
    If iPart = 0 Or lPart = 0 Then
        WebGet.mCookie = ""
    Else
        WebGet.mCookie = Trim$(Mid$(Strdata, iPart + 7, lPart - iPart - 7))
    End If
    
    iPart = 0: lPart = 0
    ' Below is used for getting form stuff this is not used in this ver.
    iPart = InStr(Strdata, "Content-Length: ")

    If iPart = 0 Then
        WebGet.mContentLen = 0
    Else
        ' This gets the forms data length
        WebGet.mContentLen = Val(Mid(Strdata, iPart + 16, lPart - iPart - 16) + 2)
        WebGet.mContentLen = Val(Mid(Strdata, iPart + 16, lPart - iPart - 16) + 2)
     
     
    End If
    iPart = 0: lPart = 0
    
End Function

Public Function DecodeURL(TUrl As String) As String
Dim iPart As Integer
Dim CGI_Str As String
    'This is used for decodeing web URLS
    While (InStr(TUrl, "%") <> 0)
        iPart = InStr(TUrl, "%")
        CGI_Str = Mid(TUrl, iPart + 1, 2)
        TUrl = Replace(TUrl, "%" & CGI_Str, Chr("&H" & CGI_Str))
    Wend
    DecodeURL = Replace(TUrl, "+", " ")
    iPart = 0
    CGI_Str = ""
    
End Function

Public Property Get Authorization() As String
    Authorization = WebGet.mAuthorization
End Property

Public Property Get Host() As String
    Host = WebGet.mHost
End Property

Public Property Get RequestType() As String
    RequestType = WebGet.mRequest
End Property

Public Property Get URLReferer() As String
    URLReferer = WebGet.mReferer
End Property

Public Property Get page() As String
    page = WebGet.mPage
End Property

Public Property Get Language() As String
    Language = WebGet.mAcceptLan
End Property

Public Property Get Encoding() As String
    Encoding = WebGet.mEncoding
End Property

Public Property Get Connection() As String
    Connection = WebGet.mConnection
End Property

Public Property Get HTTPVer() As String
    HTTPVer = WebGet.mHTTPVer
End Property
Public Property Get Accept() As String
    Accept = WebGet.mAccept
End Property

Public Property Get Agent() As String
    Agent = WebGet.mAgent
End Property

Public Property Get ContentLen() As String
    ContentLen = WebGet.mContentLen
End Property
Public Property Get Cookie() As String
    Cookie = WebGet.mCookie
End Property

Public Property Get FormData() As String
    FormData = WebGet.mFormData
End Property
