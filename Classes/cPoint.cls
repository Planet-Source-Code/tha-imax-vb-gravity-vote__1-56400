VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sName As String  ' Point Name
Public X As Single      ' Position X
Public Y As Single      ' Position Y
Public fSy As Single    ' Fall Speed Y
Public fSx As Single    ' Fall Speed X
Public g As Single      ' Gramm of Ball
Public Col As Long      ' Color of Ball
