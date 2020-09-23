VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCode 
   Caption         =   "Code to HTML"
   ClientHeight    =   7170
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6855
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
            Text            =   "No file loaded"
            TextSave        =   "No file loaded"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1140
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCode.frx":1976
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frFunctions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   8715
      Begin MSComctlLib.Toolbar TB 
         Height          =   330
         Left            =   3060
         TabIndex        =   4
         Top             =   0
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   582
         ButtonWidth     =   2355
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Selection  "
               Object.ToolTipText     =   "Copy Selection as Plain Text"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Copy as HTML"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Copy as HTML (no header)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Save As HTML"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Procedure  "
               Object.ToolTipText     =   "Copy Procedure as Plain Text"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Copy as HTML"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Copy as HTML (no header)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Save As HTML"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Module Code"
               Object.ToolTipText     =   "Copy All as Plain Text"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Copy as HTML"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Copy as HTML (no header)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Save As HTML"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboProcs 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblProcedure 
         Caption         =   "Procedure"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   780
   End
   Begin RichTextLib.RichTextBox RTFtemp 
      Height          =   735
      Left            =   -1740
      TabIndex        =   1
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   0   'False
      MousePointer    =   1
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmCode.frx":1F10
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmCode.frx":1F90
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSeperators 
         Caption         =   "Procedure Seperators"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileString 
         Caption         =   "Color String Elements"
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuRTF 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRTFCopy 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuRTFCopy 
         Caption         =   "Copy as HTML"
         Index           =   1
      End
      Begin VB.Menu mnuRTFCopy 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRTFCopy 
         Caption         =   "Save As HTML"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - January 2005
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'This project is part of a Code Library application that I still use today
'It has HTML saving options for publishing source code to web sites
'It demonstrates manipulation of the Richtextbox control and string
'parsing routines. Like all code this old, I would do many of the
'functions quite differently nowadays, but there are still some handy
'techniques used here.
'I hope I didn't create too many bugs when I pulled it out of my Code Library app !

Option Explicit
'Just some API to stop those damned scroll bars flying around while we fiddle with the RTF code
Private Type POINTL
    x As Long
    y As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Const WM_USER = &H400
Private Const EM_CHARFROMPOS = (WM_USER + 39)
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
Private Const SEL_OBJECT = &H2
Private Const SEL_TEXT = &H1
Private Const EM_SELECTIONTYPE = (WM_USER + 66)
'Variables...
Dim Procs As Collection ' holds the positions of the procedures for scrolling
Dim DontScroll As Boolean ' stops unnecessary scrolling being invoked
Dim CurCaption As String ' allows changing of caption via timer (instead of showing a progressbar)
Dim DotCount As Long ' used in indicating "busy"
Dim Header As String
'RTF code for a line - anyone know a better way ?
'RTFLine represents RTF.SelRTF so includes the header font table etc - Used for insertion of a line
Private Const RTFLine As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang3081\deflangfe3081{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}}" & vbCrLf & _
            "\uc1\pard\f0\fs17{\pict\wmetafile8\picw1764\pich882\picwgoal15871\pichgoal20" & vbCrLf & _
            "010009000003b700000006001c00000000000400000003010800050000000b0200000000050000" & vbCrLf & _
            "000c027100171b040000002e0118001c000000fb0238ff00000000000090010000000004400012" & vbCrLf & _
            "54696d6573204e657720526f6d616e0000000000000000000000000000000000040000002d0100" & vbCrLf & _
            "000400000002010100050000000902000000020d000000320aac00f4ff01000400f4fff4ff0e1b" & vbCrLf & _
            "650020b95b00030000001e0007000000fc020000808080000000040000002d01010008000000fa" & vbCrLf & _
            "02050000000000ffffff00040000002d0102000e0000002403050000000000000064000d1b6400" & vbCrLf & _
            "0d1b00000000000008000000fa0200000000000000000000040000002d01030007000000fc0200" & vbCrLf & _
            "00ffffff000000040000002d010400040000002701ffff1c000000fb021000070000000000bc02" & vbCrLf & _
            "000000000102022253797374656d00009b99807c9499190094991900a4971900b09719006d10d4" & vbCrLf & _
            "30040000002d010500030000000000" & vbCrLf & _
            "}}"
'RemRTFLine represents the line as it would appear in the RTF file with no header - used for removal of a line
Private Const RemRTFLine As String = "\par \fs20{\pict\wmetafile8\picw1764\pich882\picwgoal15871\pichgoal20 " & vbCrLf & _
            "010009000003b700000006001c00000000000400000003010800050000000b0200000000050000" & vbCrLf & _
            "000c027100171b040000002e0118001c000000fb0238ff00000000000090010000000004400012" & vbCrLf & _
            "54696d6573204e657720526f6d616e0000000000000000000000000000000000040000002d0100" & vbCrLf & _
            "000400000002010100050000000902000000020d000000320aac00f4ff01000400f4fff4ff0e1b" & vbCrLf & _
            "650020b95b00030000001e0007000000fc020000808080000000040000002d01010008000000fa" & vbCrLf & _
            "02050000000000ffffff00040000002d0102000e0000002403050000000000000064000d1b6400" & vbCrLf & _
            "0d1b00000000000008000000fa0200000000000000000000040000002d01030007000000fc0200" & vbCrLf & _
            "00ffffff000000040000002d010400040000002701ffff1c000000fb021000070000000000bc02" & vbCrLf & _
            "000000000102022253797374656d00009b99807c9499190094991900a4971900b09719006d10d4" & vbCrLf & _
            "30040000002d010500030000000000" & vbCrLf & _
            "}\fs20 "
Private Sub cboProcs_Click()
    If cboProcs.ListCount = 0 Then Exit Sub
    'If the combobox was changed by a selection change in the RTF
    'then we dont want to go any further here
    If DontScroll Then Exit Sub
    DontScroll = True
    If cboProcs.ListIndex = 0 Then
        RTF.SelStart = 0
    Else
        'Force the scroll bar to the bottom
        RTF.SelStart = Len(RTF.Text)
        'Now when we scroll the caret will be at the top
        RTF.SelStart = Procs(cboProcs.ListIndex + 1).Position
    End If
    RTF.SetFocus
    DontScroll = False
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    'Simple positioning of controls on the form
    frFunctions.Move 0, 0, ScaleWidth
    cboProcs.Left = lblProcedure.Left + lblProcedure.Width
    cboProcs.Width = ScaleWidth - TB.Width - cboProcs.Left
    TB.Move ScaleWidth - TB.Width, 0
    RTF.Move 0, frFunctions.Height, ScaleWidth, ScaleHeight - cboProcs.Height - SB.Height
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    Dim z As Long
    With cmndlg
        .Dialogtitle = "Open file"
        .Filter = "Visual Basic Source Code |*.frm;*.bas;*.cls;*.ctl;*.dsr;*.pag"
        .Flags = 5 ' no "Open as Readonly" checkbox
        ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        cboProcs.Enabled = True
        RTFtemp.Tag = .FileName
        RTF.Tag = .Filetitle
        DontScroll = True
        cboProcs.Clear
        CleanRTF ' set default font
        'Indicate progress in the title bar
        DotCount = 0
        CurCaption = "Formatting Code"
        Timer1.Enabled = True
        'load into hidden Richtextbox to avoid flickering
        'also grab all procedure names/locations
        RTFtemp.Text = LoadVB(.FileName)
        If RTFtemp.Text = "" Then GoTo woops
        CurCaption = "Applying Syntax Colors"
        KeyColor 'make keywords blue
        FixStrings
        CommentColor 'make comments backcolor green
        Timer1.Enabled = False
        'Load up the combobox
        If Procs.Count > 0 Then
            For z = 1 To Procs.Count
                cboProcs.AddItem Procs(z).Title
            Next
        End If
        'Move to displayed richtextbox
        RTF.TextRTF = RTFtemp.TextRTF
        ShowLines ' if selected
        cboProcs.ListIndex = 0
        'alter display items to match current file
        Caption = RTF.Tag
        SB.Panels(1).Text = RTFtemp.Tag
        SB.Panels(2).Text = IIf(Procs.Count = 1, Procs.Count & " Procedure", Procs.Count & " Procedures")
        TB.Buttons(2).Enabled = True
        TB.Buttons(3).Enabled = True
        mnuFileSaveAs.Enabled = True
        DontScroll = False
        Exit Sub
woops:
        Timer1.Enabled = False
        Caption = RTF.Tag & " - error loading file"
        SB.Panels(1).Text = RTFtemp.Tag
        SB.Panels(2).Text = "0 Procedures"
        cboProcs.Enabled = False
        TB.Buttons(1).Enabled = False
        TB.Buttons(2).Enabled = False
        TB.Buttons(3).Enabled = False
        mnuFileSaveAs.Enabled = False
        DontScroll = False
    End With
End Sub
Private Sub mnuFileSaveAs_Click()
    Dim sfile As String
    Dim temp As String
    Dim mFilter As String
    With cmndlg
        .Dialogtitle = "Save As"
        'We only want to save as original file types for VB saving
        'so build the filter accordingly
        Select Case LCase(ExtOnly(RTF.Tag))
            Case "frm"
                mFilter = "|Visual Basic Form (*.frm)|*.frm"
            Case "bas"
                mFilter = "|Visual Basic Module (*.bas)|*.bas"
            Case "cls"
                mFilter = "|Visual Basic Class (*.cls)|*.cls"
            Case "ctl"
                mFilter = "|Visual Basic Usercontrol (*.ctl)|*.ctl"
            Case "dsr"
                mFilter = "|Visual Basic Designer (*.dsr)|*.dsr"
            Case "pag"
                mFilter = "|Visual Basic Property Page (*.pag)|*.pag"
        End Select
        .Filter = "Rich text (*.rtf)|*.rtf|Web page (*.htm)|*.htm" & mFilter
        .FileName = ChangeExt(FileOnly(RTF.Tag))
        .OverwritePrompt = True
        ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        Refresh
        Select Case .FilterIndex
            Case 1 'rich text
                sfile = ChangeExt(.FileName, "rtf")
                RTF.SaveFile sfile
            Case 2 'HTML
                sfile = ChangeExt(.FileName, "htm")
                DotCount = 0
                CurCaption = "Converting to HTML"
                DoEvents
                Caption = "Converting to HTML"
                Timer1.Enabled = True
                temp = RTFToHTML(RTFtemp.TextRTF, ChangeExt(FileOnly(RTF.Tag)), mnuFileSeperators.Checked)
                FileSave temp, sfile, True
                Timer1.Enabled = False
                DoEvents
                Caption = RTF.Tag
            Case 3 'VB file
                'we seperated the header when we loaded the file
                'time to add it back in before saving
                sfile = ChangeExt(.FileName, LCase(ExtOnly(RTF.Tag)))
                temp = RTF.Text
                temp = Header & vbCrLf & temp
                FileSave temp, sfile, True
        End Select
    End With
End Sub
Private Function RTFToHTML(RTFStr As String, Title As String, Optional UseLines As Boolean) As String
    'This routine is quite fast - but only because we know and control the content
    'of the rich text. It would need modifications to be used as a general
    'RTF to HTML routine
    Dim z As Long
    Dim temp As String
    z = InStr(RTFStr, "\fs20")
    temp = Right(RTFStr, Len(RTFStr) - z - 5) 'RTF minus header
    temp = Left(temp, Len(temp) - 3) 'get rid of the tail
    'Determine the first color used
    z = InStr(RTFStr, "\cf")
    If z > -1 Then
        If IsNumeric(Mid(RTFStr, z + 3, 1)) Then
            z = CInt(Mid(RTFStr, z + 3, 1))
        Else
            z = 0
        End If
    End If
    Select Case z
        Case 1 'blue
            temp = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & Title & "</title>" & vbCrLf & "</head>" & "<body><font color=#000080 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>" & vbCrLf & temp
        Case Else 'black or green
            'if it was green, we still use black as we are only interested in the forecolor here
            temp = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & Title & "</title>" & vbCrLf & "</head>" & "<body><font color=#000000 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>" & vbCrLf & temp
    End Select
    temp = Replace(temp, "\pard", "") 'tail end of RTF
    temp = Replace(temp, "\cf0 ", "</font><font color=#000000 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>") 'blue - stays as blue in HTML
    temp = Replace(temp, "\cf1 ", "</font><font color=#000080 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>") 'black - stays as black in HTML
    temp = Replace(temp, "\cf2 ", "</font><font color=#000000 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>") 'green uses black forecolor
    temp = Replace(temp, RemRTFLine, IIf(UseLines, "<hr>", "")) 'lines
    temp = Replace(temp, "\par ", "<br>") 'change carriage returns to breaks
    temp = Replace(temp, "\tab ", "&nbsp;&nbsp;&nbsp;&nbsp;") 'change tabs to 4 spaces
    temp = Replace(temp, "\tab", "&nbsp;&nbsp;&nbsp;&nbsp;") 'change tabs to 4 spaces - multiple "\tab"'s will not have a space after them
    'this is where we fix up the green bits - which are only green in backcolor
    temp = Replace(temp, "\highlight2", "<span style=" & Chr(34) & "background-color: #00FF00" & Chr(34) & ">")
    temp = Replace(temp, "\highlight0 ", "</span>")
    temp = temp & vbCrLf & "</font></body></html>" 'finish up
    RTFToHTML = temp
End Function
Public Function LoadVB(sFilename As String) As String
    On Error GoTo Buggery
    Dim f As Integer, fg As Long, Searchstr As String, z As Long, temp As String
    Dim RetStr As String
    'Load the file and separate any header elements
    If LCase(Mid$(sFilename, InStrRev(sFilename, ".") + 1)) = "bas" Then
        Searchstr = "Attribute VB_Name = "
    Else
        Searchstr = "Attribute VB_Exposed"
    End If
    f = FreeFile
    Open sFilename For Binary As f
    temp = String(LOF(f), Chr$(0))
    Get f, , temp
    Close f
    fg = InStr(1, temp, Searchstr)
    If fg <> 0 Then
        fg = InStr(fg + 1, temp, vbCrLf)
        If fg <> 0 Then
            RetStr = RTFIndent(Right(temp, Len(temp) - fg - 1))
            Header = Left(temp, fg)
        End If
    Else
        GoTo Buggery
    End If
    LoadVB = RetStr
    Exit Function
Buggery:
    Header = ""
    LoadVB = ""
End Function

Private Function RTFIndent(mCode As String) As String
    'This could be made faster using Regular Expressions
    'but when I wrote this I was not up to speed with complex
    'Regular Expression patterns
    Dim temp As String, temp1 As String, z As Long, mLine As String
    Dim lastTab As Long, q As Long, qz As Long, tmpLine() As String
    Dim mProc As ClProcData
    Dim aLine As String
    Dim aLineLength As Long
    'adjust variables - leaving spaces for line insertion later according to preferences
    aLineLength = 3
    aLine = " "
    'initialize collection of procedures
    Set Procs = New Collection
    Set mProc = New ClProcData
    'first procedure will be declarations
    With mProc
        .FullTitle = "General"
        .Position = 0
        .Title = "General"
    End With
    Procs.Add mProc
    'Split the code into lines
    tmpLine = Split(mCode, vbCrLf)
    
    q = UBound(tmpLine)
    'cycle through each line to determine correct formatting
    'and apply consistent line wrapping
    'and insert lines if selected
    For z = 0 To q
        DoEvents
        temp = Trim(tmpLine(z))
        If Left(temp, 1) = "'" Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            GoTo nextPlease
        End If
        If InStr(temp, "'") Then
            temp1 = CleanComments(temp)
        Else
            temp1 = temp
            'If someone else hasn't already done wrapping ....
            If Right(temp1, 2) = " _" Then GoTo Carryon
            'apply line wrapping - trickier than you'd think
            If Len(temp1) > 90 Then
                Dim tL() As String
                Dim cnt As Long
                Dim c As Long
                Dim Pos As Long
                Dim LineEnd As String
                Dim LineStart As String
                Dim re As New RegExp
                Dim founditems As MatchCollection
                re.IgnoreCase = False
                re.Global = True
                re.Pattern = Chr(34)
                ReDim tL(Int(Len(temp1) / 90)) As String
                cnt = 1
                LineStart = ""
                LineEnd = ""
                For c = 0 To UBound(tL)
                    Pos = InStr(cnt + 90, temp1, Chr(32))
                    If Pos > 0 Then
                        If Pos - cnt > 160 Then Pos = cnt + 90
                    Else
                        Pos = cnt + 90
                    End If
                    If Pos > Len(temp1) Then Pos = Len(temp1)
                    qz = InStr(1, Left(temp1, Pos), Chr(34))
                    If c > 0 Then
                        LineStart = IIf(LineEnd = " _" & vbCrLf, "", "& " & Chr(34))
                    End If
                    If qz > 0 Then
                        Set founditems = re.Execute(Left(temp1, Pos))
                        If (founditems.Count Mod 2) = 0 Then
                            LineEnd = " _" & vbCrLf
                        Else
                            LineEnd = Chr(34) & " _" & vbCrLf
                        End If
                    Else
                        LineEnd = " _" & vbCrLf
                    End If
                    If Pos >= Len(temp1) Then LineEnd = ""
                    tL(c) = LineStart & Mid(temp1, cnt, Pos - cnt + 1) & LineEnd
                    cnt = Pos + 1
                    If Pos >= Len(temp1) Then Exit For
                Next
                temp = Join(tL, String(IIf(lastTab = 0, 1, lastTab), vbTab))
            End If
        End If
Carryon:
        If Left(temp1, 3) = "If " Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            If Right(temp1, 4) = "Then" Then lastTab = lastTab + 1
        ElseIf Left(temp1, 7) = "ElseIf " Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            If Right(temp1, 4) = "Then" Then lastTab = lastTab + 1
        ElseIf temp1 = "Else" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf temp1 = "End If" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf Left(temp1, 3) = "Do " Then
            mLine = mLine + String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf Left(temp1, 6) = "While " Then
            mLine = mLine + String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf Left(temp1, 4) = "Wend" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf temp1 = "Do" Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf Left(temp1, 4) = "Loop" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf Left(temp1, 4) = "For " Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf temp1 = "Next" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf Left(temp1, 5) = "Next " Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf Left(temp1, 12) = "Select Case " Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 2
        ElseIf Left(temp1, 5) = "Case " Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf Left(temp1, 10) = "End Select" Then
            lastTab = lastTab - 2
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf Left(temp1, 5) = "With " Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf Left(temp1, 8) = "End With" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        'Procedures
        ElseIf Left(temp1, 12) = "Private Sub " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 13, InStr(13, temp, "(") - 13)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 17) = "Private Function " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 18, InStr(18, temp, "(") - 18)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 17) = "Private Property " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 22, InStr(22, temp, "(") - 22) & " (" & Mid(temp, 17, 3) & ")"
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 11) = "Public Sub " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 12, InStr(12, temp, "(") - 12)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 16) = "Public Function " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 17, InStr(17, temp, "(") - 17)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 16) = "Public Property " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 21, InStr(21, temp, "(") - 21) & " (" & Mid(temp, 17, 3) & ")"
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 4) = "Sub " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 5, InStr(5, temp, "(") - 5)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
      ElseIf Left(temp1, 9) = "Function " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 10, InStr(10, temp, "(") - 10)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        ElseIf Left(temp1, 9) = "Property " Then
            Set mProc = New ClProcData
            With mProc
                .FullTitle = temp
                .Position = IIf(Len(mLine) <> 0, Len(mLine) + aLineLength, Len(mLine))
                .Title = Mid(temp, 14, InStr(14, temp, "(") - 14)
            End With
            Procs.Add mProc
            lastTab = 1
            mLine = mLine & aLine & vbCrLf & temp & vbCrLf
        'now anything else remaining
        ElseIf Left(temp1, 8) = "Private " Then
            lastTab = 0
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 7) = "Public " Then
            lastTab = 0
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 4) = "Dim " And Procs.Count = 0 Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 7) = "End Sub" Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 12) = "End Property" Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 12) = "End Function" Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 4) = "#If " Then
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            If Right(temp1, 4) = "Then" Then lastTab = lastTab + 1
        ElseIf Left(temp1, 8) = "#ElseIf " Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            If Right(temp1, 4) = "Then" Then lastTab = lastTab + 1
        ElseIf temp1 = "#Else" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
            lastTab = lastTab + 1
        ElseIf temp1 = "#End If" Then
            lastTab = IIf(lastTab < 1, 0, lastTab - 1)
            mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
        ElseIf Left(temp1, 13) = "Private Type " Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 12) = "Public Type " Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 13) = "Private Enum " Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 12) = "Public Enum " Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 5) = "Type " Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 5) = "Enum " Then
            lastTab = 1
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 8) = "End Type" Then
            lastTab = 0
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 8) = "End Enum" Then
            lastTab = 0
            mLine = mLine & temp & vbCrLf
        ElseIf Left(temp1, 10) = "Attribute " Then
            lastTab = 0
            mLine = mLine & temp & vbCrLf
        Else
            If temp1 <> "" Then
                If Right(temp1, 2) = " _" Then
                    mLine = mLine & String(lastTab + 1, vbTab) & temp & vbCrLf
                ElseIf z > 0 Then
                    If Right(Trim(tmpLine(z - 1)), 2) = " _" Then
                        mLine = mLine & String(lastTab + 1, vbTab) & temp & vbCrLf
                    Else
                        mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
                    End If
                Else
                    mLine = mLine & String(lastTab, vbTab) & temp & vbCrLf
                End If
            End If
        End If
nextPlease:
    Next
    Erase tmpLine
    RTFIndent = mLine
End Function
Public Function CleanComments(mLine As String) As String
    Dim cm As Long, qu As Long, z As Long, qCount As Long
    cm = InStr(1, mLine, Chr(39)) 'first comment character
    qu = InStrRev(mLine, Chr(34), cm) 'any quotes ?
    If cm <> 0 Then
        If qu <> 0 Then
            qCount = 0
            'count the quotes - this will "validate" the actions of the comment character
            For z = 1 To cm
                If Mid(mLine, z, 1) = Chr(34) Then qCount = qCount + 1
            Next
            If qCount = 0 Or (qCount Mod 2) = 0 Then
                CleanComments = Trim(Left(mLine, cm - 1))
            Else
                CleanComments = Trim(mLine)
            End If
        Else
            'no quotes - just return the code minus the comment
            CleanComments = Trim(Left(mLine, cm - 1))
        End If
    Else
        'No comments
        CleanComments = Trim(mLine)
    End If
End Function
Public Sub KeyColor()
    'We use regular expressions to quickly find the keywords
    Dim re As New RegExp, d As Long, KeyWords() As String
    Dim Found As Match, founditems As MatchCollection
    Dim temp As String
    KeyWords = GetVBKeyWords
    re.IgnoreCase = False
    re.Global = True
    For d = LBound(KeyWords) To UBound(KeyWords)
        DoEvents
        If Trim(KeyWords(d)) <> "" Then
            re.Pattern = "\b" + KeyWords(d) + "\b"
            Set founditems = re.Execute(RTFtemp.Text)
            For Each Found In founditems
                DoEvents
                'make em blue
                If Found <> "" Then
                    RTFtemp.SelStart = Found.FirstIndex
                    RTFtemp.SelLength = Len(re.Pattern) - 4
                    RTFtemp.SelColor = &H800000
                End If
            Next
        End If
    Next d
End Sub
Private Sub CommentColor()
    Dim st As Long, sl As Long, FT As Long, z As Long
    Dim re As New RegExp
    Dim Found As Match, founditems As MatchCollection
    re.IgnoreCase = False
    re.Global = True
    re.Pattern = Chr(34)
    With RTFtemp
        .SelStart = 0
        FT = .Find(Chr(39)) 'find the comment character
        If FT <> -1 Then
            z = Len(.Text)
            Do Until .SelStart + .SelLength >= Len(.Text)
                DoEvents
                st = .SelStart
                .Span vbCrLf, False, True
                If InStr(1, .SelText, vbCrLf) = 0 Then
                    'count the quotes
                    Set founditems = re.Execute(.SelText)
                    If (founditems.Count Mod 2) = 0 Then
                        .SelStart = st
                        'select to the end of the line
                        .Span vbCrLf, True, True
                        .SelColor = vbBlack
                        'make the backcolor green
                        HighLight RTFtemp, vbGreen
                    End If
                End If
                .SelStart = st + .SelLength + 2
                .SelLength = 0
                FT = .Find(Chr(39), .SelStart)
                If FT = -1 Then Exit Do
            Loop
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelIndent = 150
        .SelStart = 0
        .SelLength = 0
        st = 0
    End With
End Sub
Private Sub FixStrings()
    Dim st As Long, sl As Long, FT As Long, z As Long
    Dim col As Long
    col = IIf(mnuFileString.Checked, vbRed, vbBlack)
    With RTFtemp
        z = Len(.Text)
        If z = 0 Then Exit Sub
        FT = .Find(Chr(34), 0)
        If FT <> -1 Then
            Do Until .SelStart + .SelLength >= Len(.Text)
                DoEvents
                .SelStart = .SelStart + 1
                .Span Chr(34), True, True
                st = .SelStart
                sl = .SelLength
                If InStr(1, .SelText, vbCrLf) = 0 Then
                    .SelStart = st - 1
                    .SelLength = 1
                    .SelLength = sl + 2
                    .SelColor = col
                End If
                .SelLength = 0
                .SelStart = st + sl + 2
                FT = .Find(Chr(34), .SelStart)
                If FT = -1 Then Exit Do
            Loop
        End If
        .SelStart = 0
    End With
End Sub
Public Function GetVBKeyWords() As String()
    GetVBKeyWords = Split("#Const|#Else|#ElseIf|#End|#If|Alias|Alias|And|As|Attribute|Base|Binary|Boolean|Byte|ByVal|Call|Case|CBool|CByte|CCur|CDate|CDbl|CDec|CInt|CLng|Close|Compare|Const|CSng|CStr|Currency|CVar|CVErr|Decimal|Declare|DefBool|DefByte|DefCur|DefDate|DefDbl|DefDec|DefInt|DefLng|DefObj|DefSng|DefStr|DefVar|Dim|Do|Double|Each|Else|ElseIf|End|Enum|Eqv|Erase|Error|Exit|Explicit|False|For|Function|Get|Global|GoSub|GoTo|If|Imp|In|Input|Input|Integer|Is|LBound|Let|Lib|Like|Line|Lock|Long|Loop|LSet|Name|New|Next|Not|Object|On|Open|Option|Or|Output|Print|Private|Property|Public|Put|Random|Read|ReDim|Resume|Return|RSet|Seek|Select|Set|Single|Spc|Static|String|Stop|Sub|Tab|Then|Then|True|Type|UBound|Unlock|Variant|Wend|While|With|Xor|Nothing|To", "|")
End Function
Public Sub CleanRTF()
    RTF.TextRTF = ""
    RTFtemp.TextRTF = ""
    Set RTF.Font = Me.Font
    Set RTFtemp.Font = Me.Font
    RTF.RightMargin = 200000
    RTFtemp.RightMargin = 200000
End Sub

Private Sub mnuFileSeperators_Click()
    mnuFileSeperators.Checked = Not mnuFileSeperators.Checked
    ShowLines
End Sub
Private Sub ShowLines()
    Dim z As Long
    Dim s As Long
    Dim P As POINTL
    If Procs Is Nothing Then Exit Sub
    If Procs.Count < 2 Then Exit Sub
    DontScroll = True
    Me.MousePointer = vbHourglass
    s = RTF.SelStart
    LockWindowUpdate RTF.hwnd
    SendMessage RTF.hwnd, EM_GETSCROLLPOS, 0, P
    RTFtemp.TextRTF = RTF.TextRTF
    If mnuFileSeperators.Checked Then
        For z = Procs.Count To 2 Step -1
            RTFtemp.SelStart = Procs(z).Position - 3
            RTFtemp.SelLength = 1
            RTFtemp.SelRTF = RTFLine
        Next
    Else
        RTFtemp.TextRTF = Replace(RTFtemp.TextRTF, RemRTFLine, "\par  ")
    End If
    RTF.TextRTF = RTFtemp.TextRTF
    SendMessage RTF.hwnd, EM_SETSCROLLPOS, 0, P
    RTF.SelStart = s
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
    DontScroll = False
End Sub

Private Sub mnuFileString_Click()
    mnuFileString.Checked = Not mnuFileString.Checked
End Sub

Private Sub mnuRTFCopy_Click(Index As Integer)
    Select Case Index
        Case 0
            Clipboard.SetText RTF.SelText, vbCFText
        Case 1
            Clipboard.SetText RTFToHTML(RTF.SelRTF, ChangeExt(FileOnly(RTF.Tag)), True)
        Case 3
            Dim sfile As String
            With cmndlg
                .Dialogtitle = "Save as HTML"
                .Filter = "Web page (*.htm)|*.htm"
                .FileName = ChangeExt(FileOnly(RTF.Tag))
                .OverwritePrompt = True
                ShowSave
                If Len(.FileName) = 0 Then Exit Sub
                sfile = ChangeExt(.FileName, "htm")
                FileSave RTFToHTML(RTF.SelRTF, ChangeExt(FileOnly(RTF.Tag)), True), sfile, True
            End With
    End Select
End Sub

Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)
    'Readonly richtextbox
    If Not (Shift = 2 And KeyCode = vbKeyC) Then
        KeyCode = 0
    Else
        'you can copy - but even then we need to filter out the lines - so  vbCFText
        Clipboard.Clear
        Clipboard.SetText RTF.SelText, vbCFText
        KeyCode = 0
    End If
End Sub

Private Sub RTF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Hiding selection of a line
    Dim MY_SEL As Long
    Dim P As POINTL, CurPos As Long
    P.x = x / Screen.TwipsPerPixelX
    P.y = y / Screen.TwipsPerPixelY
    RTFtemp.SelStart = SendMessage(RTF.hwnd, EM_CHARFROMPOS, 0, P)
    RTFtemp.SelLength = 1
    MY_SEL = SendMessage(RTFtemp.hwnd, EM_SELECTIONTYPE, 0, ByVal 0)
    If MY_SEL And SEL_OBJECT Then
        LockWindowUpdate RTF.hwnd
        RTF.SelStart = RTFtemp.SelStart - 1
        RTF.SelLength = 0
    End If
End Sub

Private Sub RTF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Hiding selection of a line
    Dim MY_SEL As Long
    MY_SEL = SendMessage(RTF.hwnd, EM_SELECTIONTYPE, 0, ByVal 0)
    If MY_SEL And SEL_OBJECT Then
        If Not MY_SEL And SEL_TEXT Then
            RTF.SelLength = 0
        End If
    End If
    LockWindowUpdate 0
    If Button = 2 Then
        If RTF.SelLength > 0 Then
            Me.PopupMenu mnuRTF, , x, y
        End If
    End If
End Sub

Private Sub RTF_SelChange()
    'adjust the combobox to match the procedure selected
    Dim z As Long
    If Procs Is Nothing Then Exit Sub
    If Procs.Count = 1 Then Exit Sub
    If DontScroll Then Exit Sub
    For z = 1 To Procs.Count - 1
        If RTF.SelStart >= Procs(z).Position And RTF.SelStart < Procs(z + 1).Position Then
            Exit For
        End If
    Next
    DontScroll = True
    cboProcs.ListIndex = z - 1
    DontScroll = False
    TB.Buttons(1).Enabled = RTF.SelLength > 0
    
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
    'place desired elements on the clipboard
    Clipboard.Clear
    Select Case Button.Index
        Case 1
            Clipboard.SetText RTF.SelText, vbCFText
        Case 2
            If cboProcs.ListCount > 1 Then
                Select Case cboProcs.ListIndex
                    Case 0
                        Clipboard.SetText Left(RTF.Text, Procs(cboProcs.ListIndex + 2).Position), vbCFText
                    Case cboProcs.ListCount - 1
                        Clipboard.SetText Right(RTF.Text, Len(RTF.Text) - Procs(cboProcs.ListIndex + 1).Position), vbCFText
                    Case Else
                        Clipboard.SetText Mid(RTF.Text, Procs(cboProcs.ListIndex + 1).Position + 1, Procs(cboProcs.ListIndex + 2).Position - Procs(cboProcs.ListIndex + 1).Position - 2), vbCFText
                End Select
            Else
                Clipboard.SetText RTF.Text, vbCFText
            End If
        Case 3
            Clipboard.SetText RTF.Text, vbCFText
    End Select
End Sub


Private Sub TB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim temp As String
    Dim Title As String
    
    Select Case ButtonMenu.Parent.Index
        Case 1
            Title = ChangeExt(FileOnly(RTF.Tag))
            temp = RTFToHTML(RTF.SelRTF, Title, True)
        Case 2
            If cboProcs.ListCount > 1 Then
                Title = Procs(cboProcs.ListIndex + 1).Title
                Select Case cboProcs.ListIndex
                    Case 0
                        RTFtemp.SelStart = 0
                        RTFtemp.SelLength = Procs(cboProcs.ListIndex + 2).Position
                        temp = RTFToHTML(RTFtemp.SelRTF, Title, mnuFileSeperators.Checked)
                    Case cboProcs.ListCount - 1
                        RTFtemp.SelStart = Procs(cboProcs.ListIndex + 1).Position
                        RTFtemp.SelLength = Len(RTF.Text) - Procs(cboProcs.ListIndex + 1).Position
                        temp = RTFToHTML(RTFtemp.SelRTF, Title, mnuFileSeperators.Checked)
                    Case Else
                        RTFtemp.SelStart = Procs(cboProcs.ListIndex + 1).Position
                        RTFtemp.SelLength = Procs(cboProcs.ListIndex + 2).Position - Procs(cboProcs.ListIndex + 1).Position
                        temp = RTFToHTML(RTFtemp.SelRTF, Title, mnuFileSeperators.Checked)
                End Select
            Else
                Title = ChangeExt(FileOnly(RTF.Tag))
                temp = RTFToHTML(RTF.TextRTF, Title, mnuFileSeperators.Checked)
            End If
        Case 3
            Title = ChangeExt(FileOnly(RTF.Tag))
            temp = RTFToHTML(RTF.TextRTF, ChangeExt(FileOnly(RTF.Tag)), mnuFileSeperators.Checked)
    End Select
    Select Case ButtonMenu.Index
        Case 1
            'place desired elements on the clipboard as HTML
            Clipboard.Clear
            Clipboard.SetText temp
        Case 2
            'place desired elements on the clipboard as HTML - remove header/footer
            temp = Replace(temp, "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & Title & "</title>" & vbCrLf & "</head><body>", "")
            temp = Replace(temp, "</body></html>", "")
            Clipboard.Clear
            Clipboard.SetText temp
        Case 4
            'save desired elements as HTML
            Dim sfile As String
            With cmndlg
                .Dialogtitle = "Save as HTML"
                .Filter = "Web page (*.htm)|*.htm"
                .FileName = Title
                .OverwritePrompt = True
                ShowSave
                If Len(.FileName) = 0 Then Exit Sub
                sfile = ChangeExt(.FileName, "htm")
                FileSave temp, sfile, True
            End With
    End Select
End Sub

Private Sub Timer1_Timer()
    DotCount = DotCount + 1
    Caption = CurCaption & String(DotCount, ".")
    If DotCount >= 5 Then DotCount = 0
End Sub
