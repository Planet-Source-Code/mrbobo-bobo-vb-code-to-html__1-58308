Attribute VB_Name = "CmnDlgOpenSave"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - January 2005
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Worker routines
'Open save dialogs, file IO routines and path functions
Option Explicit
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type CMDialog
    Ownerform As Long
    Filter As String
    Filetitle As String
    FilterIndex As Long
    FileName As String
    DefaultExtension As String
    OverwritePrompt As Boolean
    AllowMultiSelect As Boolean
    Initdir As String
    Dialogtitle As String
    Flags As Long
End Type
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
Public cmndlg As CMDialog
Public Sub ShowOpen()
    Dim OFName As OPENFILENAME
    Dim temp As String
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .Dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.Flags = .Flags Or OFN_EXPLORER Or IIf(.AllowMultiSelect, OFN_ALLOWMULTISELECT, 0)
        If GetOpenFileName(OFName) Then
            .FilterIndex = OFName.nFilterIndex
            If .AllowMultiSelect Then
                temp = Replace(Trim$(OFName.lpstrFile), Chr(0), ";")
                If Right(temp, 2) = ";;" Then temp = Left(temp, Len(temp) - 2)
                .FileName = temp
            Else
                .FileName = StripTerminator(Trim$(OFName.lpstrFile))
                .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            End If
        Else
            .FileName = ""
        End If
    End With
End Sub
Public Sub ShowSave()
    Dim OFName As OPENFILENAME
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .Dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.lpstrDefExt = .DefaultExtension
        OFName.lpstrFile = .FileName & Space$(254 - Len(.FileName))
        OFName.Flags = .Flags Or IIf(.OverwritePrompt, OFN_OVERWRITEPROMPT, 0)
        If GetSaveFileName(OFName) Then
            .FileName = StripTerminator(Trim$(OFName.lpstrFile))
            .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .FilterIndex = OFName.nFilterIndex
        End If
    End With
End Sub
Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Function FileExists(mPath As String) As Boolean
    FileExists = CBool(PathFileExists(mPath))
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    If InStr(1, filepath, ".") <> 0 Then
        ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
        If dot = True Then ExtOnly = "." + ExtOnly
    Else
        ExtOnly = ""
    End If
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = Mid$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Public Sub FileSave(Text As String, filepath As String, Optional mRemoveFirst As Boolean)
    On Error Resume Next
    If mRemoveFirst And FileExists(filepath) Then Kill filepath
    Dim f As Integer
    f = FreeFile
    Open filepath For Binary As #f
    Put #f, , Text
    Close #f
    Exit Sub
End Sub
Public Function OneGulp(Src As String) As String
    On Error Resume Next
    Dim f As Integer, temp As String
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    OneGulp = temp
End Function

