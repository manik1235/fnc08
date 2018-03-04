VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debugger"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCount 
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtCount 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      Height          =   735
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Text"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtMods 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox lstMods 
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Times in list"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function dModuleCount(ModuleName As String, Optional ReturnCount As Boolean = False) As Integer
On Error GoTo errh

Dim c As Integer
Dim i As Integer

dModuleCount = -1

If Not ReturnCount Then
  'Add the module name to the list
  lstMods.AddItem Left(Right(Time, 5), 2) & ": " & ModuleName, 0
  Me.Caption = "Debugger. Items in list = " & lstMods.ListCount
Else
  'Return the number of times the name is present
  i = 0
  For c = 0 To lstMods.ListCount - 1
    If InStr(1, lstMods.List(c), ModuleName) Then i = i + 1
  Next c
End If

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "frmDebug::dModuleCount"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function

Private Sub cmdClear_Click()
On Error GoTo errh

lstMods.Clear

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "frmDebug::cmdClear_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdCopy_Click()
On Error GoTo errh

Dim c As Integer

txtMods.Text = ""
For c = 0 To lstMods.ListCount - 1
  txtMods.Text = txtMods.Text & lstMods.List(c) & vbCrLf
Next c

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdCopy_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdCount_Click()
txtCount(1).Text = dModuleCount(txtCount(0).Text, True)
End Sub


