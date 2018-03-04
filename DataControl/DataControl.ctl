VERSION 5.00
Begin VB.UserControl ActProp 
   BackColor       =   &H000000FF&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ScaleHeight     =   390
   ScaleWidth      =   375
End
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
