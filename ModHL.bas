Attribute VB_Name = "ModHL"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - January 2005
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Setting the backcolor of a richtextbox
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LF_FACESIZE = 32
Private Const WM_USER = &H400
Private Const EM_SETCHARFORMAT = (WM_USER + 68)
Private Const CFM_BACKCOLOR = &H4000000
Private Const SCF_SELECTION = &H1
Private Const SCF_ALL = &H4
Private Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type
Public HighLightColor As Long
Public Sub HighLight(mRTF As RichTextBox, Optional AlternateColor As Long = -1)
    Dim RTFformat As CHARFORMAT2
    Dim Fch As Boolean
    With RTFformat
        .cbSize = Len(RTFformat)
        .dwMask = CFM_BACKCOLOR
        .crBackColor = IIf(AlternateColor = -1, HighLightColor, AlternateColor)
    End With
    SendMessage mRTF.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
End Sub
Public Sub UnHighLight(mRTF As RichTextBox)
    Dim RTFformat As CHARFORMAT2
    Dim Fch As Boolean
    With RTFformat
        .cbSize = Len(RTFformat)
        .dwMask = CFM_BACKCOLOR
        .crBackColor = mRTF.BackColor
    End With
    SendMessage mRTF.hwnd, EM_SETCHARFORMAT, SCF_ALL, RTFformat
End Sub
