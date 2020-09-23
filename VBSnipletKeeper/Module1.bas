Attribute VB_Name = "Module1"

' This file is form my TFile32 ActiveX


' TFile32 File Functions
' This is a freeware control for dealing with File Operations
' You may use this control in your own project but please place
' This message in or you can add my Name somewhere in the program
' Thanks

' Copyright © 1997-2000 Ben Jones


Const SHGFI_DISPLAYNAME = &H200
Const SHGFI_TYPENAME = &H400
Const MAX_PATH = 260

Enum TlistOptions
    JustFilename = 0
    FilenameAndPath = 1
End Enum


Enum WinShow
    vsHide = 0
    vsNormal = 1
    vsMinSized = 2
    vsMaxSized = 3
End Enum

Enum TFolderPaths
    WindowsPath = 1
    TempPath = 2
    SystemPath = 3
    ApplactionCurrentPath = 4
End Enum


Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Sniplets_Path As String

Public Const Sniplet_Index = "\Sniplets\Code.idx"
Public Const Sniplet_Code = "\Sniplets\Code.vb"

Function Encode(S As String)
Dim I As Integer
 For I = 1 To Len(S)
  letter = Mid(S, I, 1)
   Mid(S, I, 1) = Chr(Asc(letter) Xor 128)
    Next
    Encode = S
    
End Function
Public Function CenterForm(Frm As Form)
With Frm
    .Top = (Screen.Height - Frm.Height) / 2
    .Left = (Screen.Width - Frm.Width) / 2
End With

End Function
Function AddBackSlash(Pathname As String) As String
Dim TBackSlash As String

If Not Right(Pathname, 1) = "\" Then
    TBackSlash = Pathname & "\"
    Else
    TBackSlash = Pathname
End If
    AddBackSlash = TBackSlash
    
End Function
Function CopyFiles(FileIn, FileOut As String)
Dim DataByte() As Byte
  If FileExists(FileIn) = 0 Then
    MsgBox FileIn & " not found", vbInformation
    Else
    Open FileIn For Binary Shared As #1
    Open FileOut For Binary As #2
        ReDim DataByte(0 To LOF(1))
        Get #1, , DataByte()
        Put #2, , DataByte()
        Close #1
        Close #2
    End If
    
End Function
Public Function DeleteFileName(Filename As String)
If FileExists(Filename) = 0 Then
 Exit Function
 Else
 Kill Filename
End If
 
End Function
Public Function FileExists(ByVal Filename As String) As Integer
If Dir(Filename) = "" Then FileExists = 0 Else FileExists = 1

End Function
Public Function FolderExists(ByVal Foldername As String) As Integer
If Dir(Foldername, vbDirectory) = "" Then FolderExists = 0 Else FolderExists = 1

End Function
Public Function FolderList(Foldername As String, LstBox As ListBox, AddOptions As TlistOptions)

Dim Filename As String

Filename = Dir(AddBackSlash(Foldername))
Do While Filename <> ""
 
 If AddOptions = JustFilename Then
    LstBox.AddItem Filename
 Else
    LstBox.AddItem AddBackSlash(Foldername) & Filename
 End If
    Filename = Dir
    
Loop

End Function
Public Function GetFileSize(Filename As String) As Long

Dim FSize As Long

If FileExists(Filename) = 0 Then
    MsgBox "Can't find file " & Filename
    Else
        FSize = FileLen(Filename)
    End If
    GetFileSize = FSize
    
End Function
Public Function GetFileTime(Filename As String) As Variant
Dim TDate As String
Dim StrBuffer As String
Dim StrPos As Integer

If FileExists(Filename) = 0 Then
    Exit Function
    Else
    
StrBuffer = FileDateTime(Filename)

Startpos = InStr(StrBuffer, " ")
 If Startpos Then
    Ttime = Trim(Mid(StrBuffer, Startpos, Len(StrBuffer)))
    End If
End If
    GetFileTime = Ttime
    
End Function
Public Function GetFileDate(Filename As String) As Variant
Dim TDate As String
Dim StrBuffer As String
Dim StrPos As Integer

If FileExists(Filename) = 0 Then
    Exit Function
    Else
        StrBuffer = FileDateTime(Filename)
        Startpos = InStr(StrBuffer, " ")
    If Startpos Then
        TDate = Trim(Mid(StrBuffer, 1, Startpos))
    End If
End If
GetFileDate = TDate



End Function
Public Function GetShortPath(Filename As String) As String
Dim TRes As Long, Pathname As String
    
    Pathname = String(165, 0)
    TRes = GetShortPathName(Filename, Pathname, 164)
    GetShortPath = Left(Pathname, TRes)
    
End Function
Function GetDriveFromPath(Pathname As String) As String

Dim Root As String
Dim B As Integer

B = InStr(Pathname, ":\")
  If B Then
     Root = Left(Pathname, B + 1)
End If
    GetDriveFromPath = Root

End Function
Public Function GetFilenamePath(Pathname As String) As String
Dim I As Integer
Dim Xpos As Integer

 For I = 1 To Len(Pathname)
 ch = Mid(Pathname, I, 1)
  If ch = "\" Then
    Xpos = I
    End If
    Next
     GetFilenamePath = Mid(Pathname, 1, Xpos)
     
End Function
Function GetFileExtension(Filename As String) As String

Dim TFileExt As String
Dim B As Integer

    If Right(Filename, 1) = "\" Then
       Exit Function
     Else
        B = InStr(Filename, ".")
        If B Then
          TFileExt = Mid(Filename, B + 1, Len(Filename))
        End If
        End If
        GetFileExtension = UCase(TFileExt)
        
End Function
Function GetFileNameFromPath(Pathname As String) As String
Dim Xpos As Integer
 
 If Right(Pathname, 1) = "\" Then
   Exit Function
   Else
   For m_count = 1 To Len(Pathname)
    ch = Mid(Pathname, m_count, 1)
     
     If ch = "\" Then
        Xpos = m_count
     End If
     
    Next
     GetFileNameFromPath = Mid(Pathname, Xpos + 1, Len(Pathname))
     End If
     
End Function
Public Function GetFolderPaths(Foldername As TFolderPaths) As String
Dim StrTemp As String
Dim StrWin As String
Dim StrSys As String
Dim CurPath As String

StrTemp = String(255, Chr(0))
StrWin = String(255, Chr(0))
StrSys = String(255, Chr(0))

    GetTempPath 255, StrTemp
    GetWindowsDirectory StrWin, 255
    GetSystemDirectory StrSys, 255
    
    StrTemp = Left(StrTemp, InStr(StrTemp, Chr(0)))
    StrWin = Left(StrWin, InStr(StrWin, Chr(0)))
    StrSys = Left(StrSys, InStr(StrSys, Chr(0)))
    CurPath = App.Path
    
    If Foldername = ApplactionCurrentPath Then
         GetFolderPaths = CurPath
    ElseIf Foldername = WindowsPath Then
         GetFolderPaths = StrWin
     ElseIf Foldername = TempPath Then
        GetFolderPaths = StrTemp
    ElseIf Foldername = SystemPath Then
        GetFolderPaths = StrSys
    End If
    
End Function
Public Function GetFileTypeA(Filename As String) As String
Dim TFile As SHFILEINFO

SHGetFileInfo Filename, 0, TFile, Len(TFile), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME
GetFileTypeA = RemoveCharZero(TFile.szTypeName)
    
End Function
Function RemoveBackSlash(Pathname As String) As String
Dim TBackSlash As String

If Right(Pathname, 1) = "\" Then
    TBackSlash = Left(Pathname, Len(Pathname) - 1)
    Else
    TBackSlash = Pathname
End If
    RemoveBackSlash = TBackSlash

End Function
Public Function RenameFile(OldFilename As String, NewFilename As String)
If FileExists(OldFilename) = 0 Then
  Exit Function
  Else
  Name OldFilename As NewFilename
End If

End Function
Private Function RemoveCharZero(RemoveRubish As String) As String
Dim I As Integer

 For I = 1 To Len(RemoveRubish)
    ch = Mid(RemoveRubish, I, 1)
        If ch = Chr(0) Then
            B = B + ""
        Else
            B = B + ch
        End If
    Next
    RemoveCharZero = B
    
End Function
Function MakeFolder(Foldername As String)
On Error Resume Next
If FolderExists(Foldername) = 0 Then
    MkDir Foldername
    Else
     MsgBox Foldername & " Allready Exists", vbInformation
    End If

If Err Then
 MsgBox Err.Description & " " & Foldername, vbInformation
End If

End Function
Public Function RunDialog(mHwnd As Long, DialogTitle As String)
Dim Ret As Long
    Ret = SHRunDialog(mHwnd, 0, 0, DialogTitle, vbNullString, 0)
    
End Function
Public Function RunProgran(mHwnd As Long, ProgramNamePath As String, ShowWindow As WinShow)
If FileExists(ProgramNamePath) = 0 Then
    MsgBox "Can't find file " & ProgramNamePath, vbInformation
Else
    ShellExecute mHwnd, vbNullString, ProgramNamePath, vbNullString, vbNullString, ShowWindow
End If

End Function
