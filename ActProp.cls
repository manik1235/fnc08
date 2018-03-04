VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Title As String
Public ActType As Long
Public Data0 As String
Public Data1 As String
Public Data2 As String
Public Data3 As String
Public Selected As String
Public ErrNum As Integer
Public ErrText As String
Public ActName As String 'The name of the action assigned to this object (Replace Character, Switch, etc)

