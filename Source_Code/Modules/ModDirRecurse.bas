Attribute VB_Name = "ModDirRecurse"
Option Explicit
Option Compare Text
'######################################################
'##  Caedo HTTPd is (C) 2006 Steven Dorman - ARR     ##
'##                                                  ##
'##  This program has been released under the GPL    ##
'##  and is Open-Source Software.                    ##
'##                                                  ##
'######################################################

'Recursion Bas Module - By Deth
'Most of these functions work in the same basic manner
'but are seperated to perform slightly different
'tasks each...optimally you would only c&p the
'sub you require into your own code... a few
'of the functions make calls to other functions
'included here also, so you will want to check
'that if you only use a sub or so in your code.
'However you can easily just add this module to
'your code project and use as is...
'if you find any way of optimizing or find bugs
'please let me know : deth@subdimension.com
'enjoy :)

'this is used to store file information in the dirfileinformation sub
Public Type FileInformation
    Folder As String
    Path As String
    Title As String
    Size As Long
End Type

Private Const Period As String = "."

Public Cancelled As Boolean

'retreives all the files in a folder, does not recurse into subfolders
'full file mask capability! you can use multiple mask list like so "*.exe;*.ocx;*.dll"
Sub GetFiles(Files As Collection, ByVal Path As String, Optional ByVal Mask As String = "*.*")

  Dim sFile As String, MaskArr() As String, X As Long

    If InStr(Mask, ";") Then
        MaskArr = Split(Mask, ";")
      Else
        ReDim MaskArr(0) As String
        MaskArr(0) = Mask
    End If

    Path = FormatPath(Path) & "\"
    On Error Resume Next
        For X = 0 To UBound(MaskArr)
            sFile = Dir$(Path & "\" & MaskArr(X))
            Do While LenB(sFile) > 0 And Not Cancelled
                If (GetAttr(FormatPath(Path & sFile)) And vbDirectory) <> vbDirectory Then
                    Files.Add Path & sFile
                End If
                sFile = Dir$
            Loop
        Next X

End Sub

'returns all folders inside a folder, no recursion
Sub GetFolders(Folders As Collection, ByVal Path As String)

  Dim sFolder As String

    Path = FormatPath(Path) & "\"
    sFolder = Dir$(Path, vbDirectory)
    Do While (Len(sFolder) <> 0) And Not Cancelled
        If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
            If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                Folders.Add Path & sFolder
            End If
        End If
        sFolder = Dir$
    Loop

End Sub

'returns all files and folders including current folder,
'in order of - folder\files... folder\files etc in 1 collection
Sub RecurseAll(Folders As Collection, ByVal Path As String, Optional ByVal Mask As String = "*.*")

    On Error Resume Next
      Dim sFolder As String, colFolders As New Collection

        Path = FormatPath(Path)
        Folders.Add Path
        GetFiles Folders, Path, Mask
        Path = Path & "\"
        DoEvents

        'the patented dethmiester dir recursion engine :o)
        sFolder = Dir$(Path, vbDirectory)
        Do While (Len(sFolder) <> 0) And Not Cancelled
            If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then      'look for the parent dir dots e.g ".."
                If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then 'its a folder
                    colFolders.Add Path & sFolder                                 'temporarily store it
                End If
            End If
            sFolder = Dir$  'check the next one
        Loop

        If colFolders.Count > 0 Then
            Do While (colFolders.Count > 0) And Not Cancelled 'go thru all temp stored folders
                RecurseAll Folders, colFolders(1), Mask       'and recurse it!
                colFolders.Remove 1                           'remove it
            Loop
        End If

End Sub

'returns all files and folders but in separate collections
'1 for files and 1 for folders, current folder/path is included
Sub RecurseSeperate(Files As Collection, Folders As Collection, ByVal Path As String, Optional ByVal Mask As String = "*.*")

    On Error Resume Next
      Dim sFolder As String, colFolders As New Collection

        Path = FormatPath(Path)
        Folders.Add Path
        GetFiles Files, Path, Mask
        Path = Path & "\"
        DoEvents

        sFolder = Dir$(Path, vbDirectory)
        Do While (LenB(sFolder) > 0) And Not Cancelled
            If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
                If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                    colFolders.Add Path & sFolder
                End If
            End If
            sFolder = Dir$
        Loop

        If colFolders.Count > 0 Then
            Do While (colFolders.Count > 0) And Not Cancelled
                RecurseSeperate Files, Folders, colFolders(1), Mask
                colFolders.Remove 1
            Loop
        End If

End Sub

'gets all the folders including subfolders starting from path, and descending
Sub RecurseFolders(Folders As Collection, ByVal Path As String)

  Dim sFolder As String
  Dim colNew As Collection

    On Error Resume Next
        Path = FormatPath(Path)
        Folders.Add Path
        Path = Path & "\"
        Set colNew = New Collection
        DoEvents

        sFolder = Dir$(Path, vbDirectory)
        Do While (LenB(sFolder) > 0) And Not Cancelled
            If Not (InStr(sFolder, String$(Len(sFolder), ".")) > 0) Then
                If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                    colNew.Add Path & sFolder
                End If
            End If
            sFolder = Dir$
        Loop

        If colNew.Count > 0 Then
            Do While colNew.Count > 0 And Not Cancelled
                RecurseFolders Folders, colNew(1)
                colNew.Remove 1
            Loop
        End If

End Sub

'returns all files in current folder including subfolders
Sub RecurseFiles(Files As Collection, ByVal Path As String, Optional ByVal Mask As String = "*.*")

    On Error Resume Next
      Dim sFolder As String, colFolders As New Collection

        Path = FormatPath(Path)
        GetFiles Files, Path, Mask
        Path = Path & "\"
        DoEvents
        
        sFolder = Dir$(Path, vbDirectory)
        Do While (LenB(sFolder) > 0) And Not Cancelled
            If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
                If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                    colFolders.Add Path & sFolder
                End If
            End If
            sFolder = Dir$
        Loop

        If colFolders.Count > 0 Then
            Do While (colFolders.Count > 0) And Not Cancelled
                RecurseFiles Files, colFolders(1), Mask
                colFolders.Remove 1
            Loop
        End If

End Sub

'this function will look thru all the files and calculate/return
'the filelength of all files, not optimal for checking drives/really large folders though :(
'set UseRecursion = False to only do the current folder path an not subfolders
Function GetFolderSize(ByVal Path As String, Optional ByVal UseRecursion As Boolean = True) As Double

    On Error Resume Next
      Dim sFolder As String, sFile As String, colFolders As New Collection

        Path = FormatPath(Path) & "\"

        sFile = Dir$(Path & "*.*")
        Do While (LenB(sFile) > 0) And Not Cancelled
            If (GetAttr(Path & sFile) And vbDirectory) <> vbDirectory Then
                GetFolderSize = GetFolderSize + FileLen(Path & sFile)
            End If
            sFile = Dir$
        Loop

        If UseRecursion Then 'check subfolders
            
            DoEvents
            sFolder = Dir$(Path, vbDirectory)
            Do While (LenB(sFolder) > 0) And Not Cancelled
                If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
                    If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                        colFolders.Add Path & sFolder
                    End If
                End If
                sFolder = Dir$
            Loop

            If colFolders.Count > 0 Then
                Do While (colFolders.Count > 0) And Not Cancelled
                    GetFolderSize = GetFolderSize + GetFolderSize(colFolders(1), UseRecursion)
                    colFolders.Remove 1
                Loop
            End If

        End If

End Function

'this function will look thru all the files and calculate/return in a FileInformation type,
'the file info of all files, not optimal for checking drives/really large folders though :(
'set UseRecursion = False to only do the current folder path an not subfolders
Sub DirFileInformation(File() As FileInformation, ByVal Path As String, Optional ByVal UseRecursion As Boolean = True, Optional ByVal Mask As String = "*.*")

    On Error Resume Next
      Dim sFolder As String, X As Long, I As Long
      Dim colFiles As New Collection, colFolders As New Collection
      
        Path = FormatPath(Path) & "\"
        GetFiles colFiles, Path, Mask
                                
        If colFiles.Count > 0 Then
             For I = 1 To colFiles.Count
                If UBound(File) = 0 Then
                    If Err Then
                         Err.Clear
                         On Error Resume Next
                         X = 0
                     Else
                         X = UBound(File) + 1
                     End If
                Else
                    X = UBound(File) + 1
                End If
                ReDim Preserve File(X)
                With File(X)
                    .Folder = Left$(Path, Len(Path) - 1)
                    .Path = colFiles(I)
                    .Title = Mid$(colFiles(I), InStrRev(colFiles(I), "\") + 1)
                    .Size = FileLen(colFiles(I))
                End With
             Next I
        End If
           

        If UseRecursion Then 'check subfolders

            DoEvents
            sFolder = Dir$(Path, vbDirectory)
            Do While (LenB(sFolder) > 0) And Not Cancelled
                If Not (InStr(sFolder, String$(Len(sFolder), Period)) > 0) Then
                    If (GetAttr(Path & sFolder) And vbDirectory) = vbDirectory Then
                        colFolders.Add Path & sFolder
                    End If
                End If
                sFolder = Dir
            Loop

            If colFolders.Count > 0 Then
                Do While (colFolders.Count > 0) And Not Cancelled
                    DirFileInformation File, colFolders(1), UseRecursion, Mask
                    colFolders.Remove 1
                Loop
            End If

        End If

End Sub

'simple function that does some checking of a folderpath
'and removes trailing slashes
Function FormatPath(ByVal FolderPath As String) As String

    On Error Resume Next
        If Len(FolderPath) > 2 Then
            Do Until Right$(FolderPath, 1) <> "\"
                FolderPath = Left$(FolderPath, Len(FolderPath) - 1)
            Loop
            FolderPath = Replace$(FolderPath, "/", "\")
        End If

        If Len(FolderPath) > 2 Then
            FormatPath = Left$(FolderPath, 2) & Replace$(Mid$(FolderPath, 3), "\\", "\")
          Else
            FormatPath = FolderPath
        End If

End Function

