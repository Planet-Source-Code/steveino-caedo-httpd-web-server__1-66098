VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Caedo HTTPd"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8790
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Breather 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Timer_log 
      Interval        =   60000
      Left            =   1920
      Top             =   120
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   1320
      Top             =   120
   End
   Begin CaedoHTTPd.Socket sockets 
      Index           =   0
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin CaedoHTTPd.Socket sckMain 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox cmdline 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Command Line Interface (type help or ? to more information)"
      Top             =   4560
      Width           =   8775
   End
   Begin VB.TextBox txtlog 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
   Begin VB.Menu mnu01 
      Caption         =   "Main"
      Begin VB.Menu mnu02 
         Caption         =   "ShutDown"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'######################################################
'##  Caedo HTTPd is (C) 2006 Steven Dorman - ARR     ##
'##                                                  ##
'##  This program has been released under the GPL    ##
'##  and is Open-Source Software.                    ##
'##                                                  ##
'######################################################
'#The power of VB isnt that bad now is it .... :)
'#
'#PS: i also used some code from other open source programs
'#    i would like to send out a thanks to anybody whos
'#    code helped make this program possible, Aswell as
'#    the people and friends to helped aswell...

Private oZLIB As New cZLIB
Public webinfo As New webclass

'//Arrays
Private SendQue(1024) As Boolean
Dim Transfer(1024) As Boolean
Dim BIP() As String


'//Strings
Public Conf_Path As String
Public E_PAGE As String
Public HTDocs As String
Public IndexPage As String
Public PHPPath As String
Public PERLPath As String
Public Leech As String
    '//Custom CGI Module
Public xCGIPath As String
Public xCGIExt As String
Public xCGIName As String


'//Integers/Longs
Public Port As Integer
Public MaxC As Integer


'//Booleans
Public SSIOn As Boolean
Public PHPOn As Boolean
Public PERLOn As Boolean
Public Reporting As Boolean
Public GZip As Boolean
    '//Custom CGI Module
Public xCGIOn As Boolean



'//Variants
Public Bytes As Variant
Public Tout As Variant




Private Sub Breather_Timer(Index As Integer)
'//Simple Breathing Code :P
Transfer(Index) = True
Breather(Index).Enabled = False

End Sub

Private Sub cmdline_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
LogWrite cmdline.Text
    If cmdline.Text = "version" Then
        LogWrite "Caedo HTTPd Web Server V." & App.Major & "." & App.Minor & "." & App.Revision & " " & vbCrLf & " was originally written by Steven Dorman. It is (C) Steven Dorman 2006 - all rights reserved." & vbCrLf & "          This program is released as open source under the GPL"
    ElseIf cmdline.Text = "help" Or cmdline.Text = "?" Then
        LogWrite "Caedo HTTPd Web Server V." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & " Version(Shows some version information) " & vbCrLf & " cls (wipes console) " & vbCrLf & " shutdown (kills the server) " & vbCrLf & " leeching (Toggles anti-leeching support)" & vbCrLf & " load_blist (Reloads the banned IP's List into memory)" & vbCrLf & " stats (Used to load VHost stats into the console IE: stats sub.you.com)"
        
    ElseIf cmdline.Text = "cls" Then
        txtlog.Text = "Caedo HTTPd Web Server V." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    ElseIf cmdline.Text = "shutdown" Then
        mnu02_Click
    ElseIf cmdline.Text = "leeching" Then
        If Leech = False Then
            Leech = True
            LogWrite "Anti Leeching is now activated. Remember, to view your main site you must have a virtual host entry for it!"
        Else
            Leech = False
            LogWrite "Anti Leeching is now De-Activated"
        End If
    
        WriteINI "main", "leeching", Leech, Conf_Path
    ElseIf cmdline.Text = "load_blist" Then
        isbanned "0.0.1.0", True
    ElseIf Left(cmdline.Text, 5) = "stats" Then
      LogWrite LoadStats(Split(cmdline.Text, " ")(1))
    ElseIf cmdline.Text = "load_conf" Then
          txtlog.Text = "Caedo HTTPd Web Server V." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
          Load_Conf
          LogWrite "Configuration Files Reloaded!"
    Else
    LogWrite "No Such Command Found!"
    End If

cmdline.Text = ""
End If
End Sub



Private Sub Form_Load()
'//Version Stamp
 txtlog.Text = "Caedo HTTPd Web Server V." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
 
'//Change caption to include version number
Me.Caption = Me.Caption & " V." & App.Major & "." & App.Minor

    '//Set the Configuration files path and error pages
    Conf_Path = App.Path & "\data\conf\httpd.cfg"
    E_PAGE = App.Path & "\data\E_PAGE\"
    
    '//Load the Banned IP Table
    isbanned "abc", True
    
    '//Load all settings from the configuration file
    Load_Conf
    
    '//Load sockets for usage
    Load_Sock

    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'//Code for tray icon
Dim Sys As Long
    Sys = X / Screen.TwipsPerPixelX
    Select Case Sys
    Case WM_LBUTTONDOWN:
        Me.Visible = True
        Me.WindowState = vbNormal
        Me.Show
    End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
'//Make the console text boxes fit the screen
txtlog.Width = Me.Width - 140
    txtlog.Height = Me.Height - cmdline.Height - 700
    cmdline.Width = Me.Width - 140
    cmdline.Top = txtlog.Height
frmMain.Refresh
End If

'//Code for tray icon
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub



Private Sub Form_Terminate()
On Error Resume Next
Dim I As Integer
'//Clear out the sockets...
sckMain.CloseSck
For I = 0 To MaxC
    sockets(I).CloseSck
Next I
End
'//Destroy  shell icon
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnu02_Click()
'//Current code for shutting down the server
Me.Hide
    Unload Me
    Set frmMain = Nothing
End
End Sub

Private Sub cmdline_Click()
'//Clear default command line text
cmdline.Text = ""
End Sub

Private Sub sckMain_CloseSck()
'//Close the socket definitly and listen again
sckMain.CloseSck
sckMain.Listen
End Sub

Private Sub sckMain_ConnectionRequest(ByVal requestID As Long)
'//Check for an open socket and send the request to it

'//Are they banned?
If isbanned(sckMain.RemoteHostIP, False) = True Then
Dim f As Integer

f = FreeFile
Open App.Path & "\data\logs\access.log" For Append As #f
      Print #f, "[" & Now & "][" & sckMain.RemoteHostIP & "] is BANNED and has tried to access the server"
Close #f
sckMain.CloseSck
sckMain.Listen
Exit Sub
End If

Dim I As Long
    For I = 0 To MaxC
        If sockets(I).State = sckClosed Then
            sockets(I).Accept requestID
            DoEvents
            sckMain.CloseSck
            sckMain.Listen
            I = Empty
            Exit For
        End If
    Next I
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'//Explain the error to console, and close the socket
LogWrite "Main Socket Error: " & Description
    sckMain.CloseSck
DoEvents
End Sub

Private Sub sockets_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo hell
Dim Data As String, PostData As String, tmpdata As String, I As Integer, Headz As String

'//Grab the headers from the socket
sockets(Index).GetData Data, , bytesTotal
DoEvents

'//Wait for sending transfer
If SendQue(Index) = True Then
    While SendQue(Index) = True
        DoEvents
    Wend
End If

'//And then BREATH so it dosent "over heat"
If Transfer(Index) = True Then
Transfer(Index) = False
Breather(Index).Enabled = True
    While Transfer(Index) = False
        DoEvents
    Wend
End If


'//And then parse the header information
webinfo.PhaseWebHead Data
DoEvents

'//Gather Post Data
If webinfo.RequestType = "POST" Then
   PostData = Split(Data, vbCrLf & vbCrLf)(1)
   PostData = URLDecode(PostData)
End If


'//Backwards Directory Hole Check
If InStr(1, webinfo.page, "..", vbTextCompare) Then
            Send_Out E_PAGE & "403.E_PAGE", Index
        Exit Sub
End If

    
'//Headers ok?
If webinfo.HTTPVer <> "HTTP/1.1" And webinfo.HTTPVer <> "HTTP/1.0" And webinfo.HTTPVer <> "HTTP/0.9" Then
            Send_Out E_PAGE & "403.E_PAGE", Index
            Exit Sub
End If

'//fix up request line so we can understand it
Dim FileOut As String, Refer As String, tmp As String, f As Integer, ff As Integer, tmp2 As String, tmp3 As String, tmp4 As String, tmp5 As String
f = FreeFile
ff = FreeFile
FileOut = webinfo.page
FileOut = Replace(FileOut, "/", "\")
FileOut = Replace(FileOut, "http://", "")
FileOut = Replace(FileOut, "%20", " ")

Refer = Replace(webinfo.URLReferer, "http://", "")
If Refer <> "" Then
    Refer = Split(Refer, "/")(0)
End If

tmp2 = ReadINI(webinfo.Host, "htdocs", App.Path & "\data\conf\vhost.cfg")
tmp3 = ReadINI(webinfo.Host, "indexpage", App.Path & "\data\conf\vhost.cfg")
tmp4 = ReadINI(Refer, "htdocs", App.Path & "\data\conf\vhost.cfg")
tmp5 = ReadINI(webinfo.Host, "enabled", App.Path & "\data\conf\vhost.cfg")



'//Support for anti-leeching
If Leech = True Then
If Refer <> "" And tmp4 = "" Then
                 f = FreeFile
                Open App.Path & "\data\logs\access.log" For Append As #f
                    Print #f, "[" & Now & "][" & sockets(Index).RemoteHostIP & "] LEECH DETECTED - Denied leeching of the file " & FileOut & " from server " & Refer & ""
                Close #f
                Send_Out E_PAGE & "403.E_PAGE", Index
                DoEvents
Exit Sub
End If
End If

'//Check to see if its root
If tmp2 <> "" Or tmp2 <> Empty Then
'//Check wether Vhost is even enabled
If tmp5 = "1" Then
    
    '//Manage the virtual hosts
    If FileOut = "/" Or FileOut = "\" Then
        FileOut = tmp2 & "\" & tmp3
    Else
        FileOut = tmp2 & FileOut
    End If

Else
    FileOut = E_PAGE & "403.E_PAGE"
End If

Else

    '//Just manage the main site
    If FileOut = "/" Or FileOut = "\" Then
        FileOut = HTDocs & "\" & IndexPage
    Else
        FileOut = HTDocs & FileOut
    End If

End If




'//check to see if file exists
If Dir(FileOut) = "" Then
    '//404 - file was not found
    Send_Out E_PAGE & "404.E_PAGE", Index
    
    '//Write to the access log for 404 error
    Open App.Path & "\data\logs\access.log" For Append As #f
           Print #f, "[" & Now & "][" & sockets(Index).RemoteHostIP & "] 404 ERROR - Requested File was " & FileOut
    Close #f
Else

    '//Write to the access log for file request
    Open App.Path & "\data\logs\access.log" For Append As #f
           Print #f, "[" & Now & "][" & sockets(Index).RemoteHostIP & "] Requested File was " & FileOut
    Close #f
    
'//Execute PHP Scripts
If Right(FileOut, 3) = "php" Then
    If PHPOn = True Then
         tmpdata = ModCGI.OpenFile(FileOut)
         tmpdata = ModCGI.AddVars(PostData, tmpdata)
         
         Open App.Path & "\data\temp\" & App.hInstance & ".dat" For Output As #96
            Print #96, tmpdata
         Close #96
         
         tmp = RunPHP(App.Path & "\data\temp\" & App.hInstance & ".dat", Index, webinfo.FormData, webinfo.Cookie)
         Headz = Split(tmp, vbCrLf & vbCrLf)(0)
         
         tmp = Replace(tmp, Headz, "")
         Headz = Replace(Headz, "Content-type: text/html", "")
        Open App.Path & "\data\temp\" & App.hInstance & ".php" For Output As #96
            Print #96, tmp
        Close #96
        
        FileOut = App.Path & "\data\temp\" & App.hInstance & ".php"
    Else
          FileOut = E_PAGE & "cgi.E_PAGE"
    End If

End If

'//Execute PERL Scripts (Not Currently Fully Supported)
If Right(FileOut, 3) = ".pl" Then
    If PERLOn = True Then
        tmp = RunPERL(FileOut, Index, webinfo.FormData)
        Open App.Path & "\data\temp\" & App.hInstance & ".pl" For Output As #f
            Print #f, tmp
        Close #f
        FileOut = App.Path & "\data\temp\" & App.hInstance & ".pl"
   Else
    FileOut = E_PAGE & "cgi.E_PAGE"
    End If
   
End If

'//Execute Custom CGI Code
If Right(FileOut, Len(xCGIExt)) = xCGIExt Then
    If xCGIOn = True Then
        tmp = RunxCGI(FileOut, Index, Data)
        Open App.Path & "\data\temp\" & App.hInstance & xCGIExt For Output As #f
            Print #f, tmp
        Close #f
        FileOut = App.Path & "\data\temp\" & App.hInstance & xCGIExt
   Else
    FileOut = E_PAGE & "cgi.E_PAGE"
    End If
   
End If

'//Execute SSI Code
If Right(FileOut, 5) = "shtml" Then
    If SSIOn = True Then
        
        If tmp2 = "" Then
        tmp = SSI(ModCGI.OpenFile(FileOut), FileOut, HTDocs, Index)
        Else
        tmp = SSI(ModCGI.OpenFile(FileOut), FileOut, tmp2, Index)
        End If
        
        Open App.Path & "\data\temp\" & App.hInstance & ".shtml" For Output As #99
            Print #99, tmp
        Close #99
        DoEvents
        FileOut = App.Path & "\data\temp\" & App.hInstance & ".shtml"
   Else
    FileOut = E_PAGE & "cgi.E_PAGE"
    End If
   
End If
        
    '//Files seems to get ahead of themself...try to que it up
   While SendQue(Index) = True
   DoEvents
   Wend
    
    '//Officially send the file
    If Headz = "" Then
        Send_Out FileOut, Index
    Else
        Send_Out FileOut, Index, Headz
    End If
    
End If
Exit Sub
hell:
For I = 0 To MaxC
sockets(I).CloseSck
Next I

For I = 0 To MaxC
TimeOut(I).Enabled = False
Next I

Reset
SendQue(Index) = False
LogWrite "Request Parsing Error :" & Err.Description
End Sub

Private Sub sockets_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'//Error in sockets so we will notify console and kill it
sockets(Index).CloseSck
Reset
LogWrite "Socket " & Index & " Error: " & Description
End Sub

Private Sub TimeOut_Timer(Index As Integer)
'//Check to see if the equivelant socket to timer is idle...
If sockets(Index).State = sckConnected Then
    sockets(Index).CloseSck
    TimeOut(Index).Enabled = False
    SendQue(Index) = False
End If
End Sub

Private Sub Timer_log_Timer()
'//This timer simply saves the latest up-to-date copy
'//Of the console log file
Dim f As Integer
f = FreeFile
Open App.Path & "\data\logs\console.log" For Output As #f
Print #f, txtlog.Text
Close #f
End Sub

Private Sub txtlog_Change()
'//Jump to the end of the log
txtlog.SelLength = Len(txtlog)
End Sub

Public Sub LogWrite(Strdata As String)
'// Write out to the console with style
    txtlog.Text = txtlog.Text & vbCrLf & "Caedo />  " & Strdata & vbCrLf
    txtlog.SelStart = Len(txtlog.Text)
    
End Sub

Private Sub Load_Conf()
'//Start Loading all settings from the configuration file
    HTDocs = ReadINI("main", "htdocs", Conf_Path)
    HTDocs = Replace(HTDocs, "<apppath>", App.Path)
    
    IndexPage = ReadINI("main", "indexpage", Conf_Path)
    Port = ReadINI("main", "port", Conf_Path)
    MaxC = ReadINI("main", "maxc", Conf_Path)
    Bytes = ReadINI("main", "bytes", Conf_Path)
    Leech = ReadINI("main", "leeching", Conf_Path)
    Tout = ReadINI("main", "tout", Conf_Path)
    Tout = Tout * 100
    
    PHPPath = ReadINI("cgi", "php", Conf_Path)
    PERLPath = ReadINI("cgi", "perl", Conf_Path)
    SSIOn = ReadINI("cgi", "ssi", Conf_Path)
    Reporting = ReadINI("main", "reporting", Conf_Path)
    GZip = ReadINI("main", "gzip", Conf_Path)
    
    '//CustomCGI Module
    xCGIOn = ReadINI("main", "enabled", App.Path & "\data\conf\customcgi_mod.cfg")
    xCGIPath = ReadINI("main", "path", App.Path & "\data\conf\customcgi_mod.cfg")
    xCGIExt = ReadINI("main", "ext", App.Path & "\data\conf\customcgi_mod.cfg")
    xCGIName = ReadINI("main", "name", App.Path & "\data\conf\customcgi_mod.cfg")
     
        
     '//Check for CGI Executeables//
    If Dir(PHPPath) = "" Then
        PHPOn = False
    Else
        PHPOn = True
        LogWrite "PHP Support ON"
    End If
    
    If Dir(PERLPath) = "" Then
        PERLOn = False
    Else
        PERLOn = True
        LogWrite "PERL Support ON"
    End If
    
    If SSIOn = True Then
        LogWrite "SSI Support On"
    End If
    
    If xCGIOn = True Then
        LogWrite xCGIName & " Support On"
    End If
    
End Sub

Private Sub Load_Sock()
Dim I As Integer
 '//Load Connection Sockets & timeout events
    For I = 1 To MaxC
        load sockets(I)
        DoEvents
    Next I
    
    For I = 1 To MaxC
        load TimeOut(I)
        TimeOut(I).Interval = Tout
        DoEvents
    Next I

       For I = 1 To MaxC
        load Breather(I)
        Breather(I).Interval = 100
        DoEvents
    Next I
 '//Now Bind and Listen on configured port
    sckMain.CloseSck
    sckMain.LocalPort = Port
    sckMain.Listen
    DoEvents
    LogWrite "Caedo is running on port " & Port
    
    
End Sub

Private Sub Send_Out(strFile As String, Index As Integer, Optional e_head As String)
'On Error GoTo hell
Dim Compressed As Boolean
Dim f As Long, I As Integer, FileLen As Long, BinOut As String, tmp2 As String, W As Long, ff As Integer, snt As Long, csite As String, tmpb As Long
 ff = FreeFile
    SendQue(Index) = True
           
    f = FreeFile
    BinOut = Space(Bytes)
    
    '//IS GZIP for Plain text/HTML Documents enabled?
    If GZip = True Then
    If Left(GetFileInfo(strFile), 4) = "text" Or Left(GetFileInfo(strFile), 5) = "image" Then
    If Right(strFile, 3) <> "php" Then
    '//Simply preload the file and compress it..
    Dim tmps As String
        Compressed = True
        tmps = ModCGI.OpenFile(strFile)
       
        DoEvents
        Call oZLIB.CompressString(tmps, Z_GZIP)
     
        DoEvents
        Open App.Path & "\data\temp\" & App.hInstance & ".dat" For Output As #75
        Print #75, tmps
        FileLen = LOF(75)
        Close #75

    '//Send headers early so it knows what kind of file it is
    sockets(Index).SendData SendHeaders(FileLen, strFile, e_head)
    DoEvents
    
    strFile = App.Path & "\data\temp\" & App.hInstance & ".dat"
    tmps = Empty
   End If
   End If
   End If
   
   
    
    Open strFile For Binary As #f
    FileLen = LOF(f)
    W = 0
    

    
    '//Log usage per Host (if virtual host entry exists)
If Reporting = True And webinfo.Host <> "" Then

    Dim HostnameFix As String, BW As String, kb As String, Pather As String
    HostnameFix = webinfo.Host
    HostnameFix = Replace(HostnameFix, ".", "_")
    HostnameFix = Replace(HostnameFix, ":", "_")
    Pather = modini.ReadINI(webinfo.Host, "htdocs", App.Path & "\data\conf\vhost.cfg")
  
    '//Log Bandwidth OUT
    BW = modini.ReadINI("stats", "bandwidth", App.Path & "\data\logs\" & HostnameFix & ".log")
    If BW = "" Then
    BW = FileLen
    
    Else
          BW = BW + FileLen
    End If
    kb = BW / 1024
    
    modini.WriteINI "stats", "Bandwidth", BW, App.Path & "\data\logs\" & HostnameFix & ".log"
    modini.WriteINI "stats", "Bandwidth_KB", kb, App.Path & "\data\logs\" & HostnameFix & ".log"
    modini.WriteINI "stats", "Bandwidth_MB", kb / 1024, App.Path & "\data\logs\" & HostnameFix & ".log"
    
    '//Log Disk Usage
    If Pather <> "" Then
        BW = GetFolderSize(Pather, True)
        kb = BW / 1024
        modini.WriteINI "stats", "Disk_UsageKB", kb, App.Path & "\data\logs\" & HostnameFix & ".log"
        modini.WriteINI "stats", "Disk_UsageMB", kb / 1024, App.Path & "\data\logs\" & HostnameFix & ".log"
    End If
    
    '//Log Hits to VHost
    BW = modini.ReadINI("stats", "Requests", App.Path & "\data\logs\" & HostnameFix & ".log")
    If BW = "" Then
    BW = 1
    
    Else
          BW = BW + 1
    End If
    modini.WriteINI "stats", "Requests", BW, App.Path & "\data\logs\" & HostnameFix & ".log"
    
End If
    '//END LOGGING
    

    Do Until sockets(Index).State = sckConnected
    DoEvents
    Loop
    
    '//IF the file isnt compressed, send headers
    If Compressed = False Then
        sockets(Index).SendData SendHeaders(FileLen, strFile, e_head)
        DoEvents
    End If
    
    '//Below is the basic file sending routene
       Do
           Get #f, , BinOut
           W = W + Len(BinOut)
           If W > FileLen Then
           
        '//And then BREATH so it dosent "over heat"
            Transfer(Index) = False
            Breather(Index).Interval = 256
            Breather(Index).Enabled = True
                While Transfer(Index) = False
                    DoEvents
                Wend
            Breather(Index).Interval = 100
            
       
        
            sockets(Index).SendData Mid(BinOut, 1, Len(BinOut) - (W - FileLen))
            
            Else
            sockets(Index).SendData BinOut
            
           End If
           
       Loop Until EOF(f)
       
Close #f

'//Initiate the TimeOut Timer
TimeOut(Index).Enabled = True
SendQue(Index) = False
DoEvents

Exit Sub
hell:
For I = 0 To MaxC
sockets(I).CloseSck
Next I

For I = 0 To MaxC
TimeOut(I).Enabled = False
Next I

SendQue(Index) = False
Reset
LogWrite "Error: " & Err.Description
End Sub

'//Simple Function to create the headers to be sent to the browser
Private Function SendHeaders(howbig As Long, namer As String, Optional e_head As String) As String
Dim theheaders As String
   
theheaders = "HTTP/1.0 200 OK"
    theheaders = theheaders & vbCrLf & "Server: Caedo HTTPd"
    theheaders = theheaders & vbCrLf & "Date:" & Format(Date, "Medium Date", vbMonday, vbFirstJan1)
    theheaders = theheaders & vbCrLf & "Last-Modified:" & Format(Date, "Medium Date", vbMonday, vbFirstJan1)
    
    theheaders = theheaders & vbCrLf & "Accept-Ranges: bytes"
    theheaders = theheaders & vbCrLf & "Content-Length: " & howbig
    theheaders = theheaders & vbCrLf & "Connection: Keep-Alive"
    theheaders = theheaders & vbCrLf & "Content-Type: " & GetFileInfo(namer)
    
    '//If were going to compress, send correct headers
    If Left(GetFileInfo(namer), 4) = "text" Or Left(GetFileInfo(namer), 5) = "image" And GZip = True Then
    If Right(namer, 3) <> "php" Then
        theheaders = theheaders & vbCrLf & "Content-Encoding: gzip"
    End If
    End If
    
    '//Send extra headers from php, if they are there
    If e_head <> "" Then
    theheaders = theheaders & vbCrLf & e_head
    End If
    
    theheaders = theheaders & vbCrLf & ""
    theheaders = theheaders & vbCrLf
SendHeaders = theheaders
   
End Function
'//And whats its mime type?
Private Function GetFileInfo(filename As String) As String
Dim TextFile As Boolean
Dim FileType As String
Dim ParsedFile As String, B As String

If Right(filename, Len(xCGIExt)) = xCGIExt Then
    FileType = "text/html"
    GoTo doner
End If

ParsedFile = Right(filename, 3)
B = modini.ReadINI("types", ParsedFile, App.Path & "\data\mime.types")

If B = "" Then
    ParsedFile = Right(filename, 4)
    B = modini.ReadINI("types", ParsedFile, App.Path & "\data\mime.types")
End If

If B = "" Then
    ParsedFile = Right(filename, 5)
    B = modini.ReadINI("types", ParsedFile, App.Path & "\data\mime.types")
End If

If B = "" Then
    ParsedFile = Right(filename, 6)
    B = modini.ReadINI("types", ParsedFile, App.Path & "\data\mime.types")
End If

If B <> "" Then
    FileType = B
Else
    FileType = "unknown/binary"
End If

doner:
GetFileInfo = FileType
End Function

'//Simple Function to check and see if a user is banned
Public Function isbanned(ip As String, load As Boolean) As Boolean
On Error Resume Next
isbanned = False

If load = True Then
Dim fullfile As String, f As Integer
    f = FreeFile
    Open App.Path & "\data\conf\banned.cfg" For Binary As #f
        fullfile = Input(FileLen(App.Path & "\data\conf\banned.cfg"), #f)
        DoEvents
    Close #f
    f = Empty
        
    BIP = Split(fullfile, vbCrLf)
    LogWrite "Loaded Banned IP's List"
End If

    Dim I As Integer
    
    If BIP(0) = "" Then
        isbanned = False
        Exit Function
    End If

    For I = 0 To UBound(BIP)
   
        If BIP(I) = ip Then
                isbanned = True
                Exit For
        End If
        isbanned = False
    Next I
   
End Function

'//Very simple function for outputting a VHosts stats to the console
Public Function LoadStats(vhost As Variant) As String
Dim tmp As String, HostnameFix As String, out As String
 HostnameFix = vhost
    HostnameFix = Replace(HostnameFix, ".", "_")
    HostnameFix = Replace(HostnameFix, ":", "_")
    out = "Statistics for " & vhost & ":" & vbCrLf & "Type | Amount" & vbCrLf & "-------------" & vbCrLf
        
    tmp = ReadINI("stats", "bandwidth", App.Path & "\data\logs\" & HostnameFix & ".log")
    out = out & "Bandwidth in Bytes | " & tmp & vbCrLf
    
    tmp = ReadINI("stats", "bandwidth_kb", App.Path & "\data\logs\" & HostnameFix & ".log")
    out = out & "Bandwidth in KB | " & tmp & vbCrLf
    
    tmp = ReadINI("stats", "bandwidth_mb", App.Path & "\data\logs\" & HostnameFix & ".log")
    out = out & "Bandwidth in MB | " & tmp & vbCrLf
    
     tmp = ReadINI("stats", "disk_usagekb", App.Path & "\data\logs\" & HostnameFix & ".log")
    out = out & "DiskUsage in KB | " & tmp & vbCrLf
    
      tmp = ReadINI("stats", "disk_usagemb", App.Path & "\data\logs\" & HostnameFix & ".log")
    out = out & "DiskUsage in MB | " & tmp & vbCrLf
    
      tmp = ReadINI("stats", "requests", App.Path & "\data\logs\" & HostnameFix & ".log")
    out = out & "Total Requests | " & tmp & vbCrLf
    
    LoadStats = out
End Function

