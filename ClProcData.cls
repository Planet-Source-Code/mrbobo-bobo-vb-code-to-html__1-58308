VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClProcData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - January 2005
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Simple class to hold procedure data
Option Explicit

Public FullTitle As String
Public Title As String
Public Position As Long

