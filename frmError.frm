VERSION 5.00
Begin VB.Form frmError 
   Caption         =   "Filename Changer Error"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Copy to Clipboard then Close"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Close Program"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtError 
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdButton_Click(Index As Integer)
On Error GoTo errh

If Index = 0 Then
  'Close Program
  #If Debugging = 1 Then
    Stop
    Unload Me
  #Else
    End
  #End If
ElseIf Index = 1 Then
  'Copy to clipboard then close
  Clipboard.Clear
  Clipboard.SetText txtError.Text
  #If Debugging = 1 Then
    Unload Me
  #Else
    End
  #End If
ElseIf Index = 2 Then
  'Copy to clipboard only
  Clipboard.Clear
  Clipboard.SetText txtError.Text
End If

Exit Sub
errh:
MsgBox "Copy Failed. " & Err & ": " & Err.Description
Exit Sub
End Sub

Public Sub MsgBoxError(Prompt As String, Optional Buttons As fncErrorButtons = fncErrorMode, Optional Title As String = "", Optional HelpFile As Integer = 0, Optional Context As Integer = 0)
Load frmError

'Set the buttons
SetButtons Buttons

frmError.txtError.Text = Prompt
If Title <> "" Then frmError.Caption = Title
frmError.Show 1, frmMain
End Sub

Private Sub Form_Resize()
On Error GoTo errh

Dim c As Integer

txtError.Width = Me.ScaleWidth
txtError.Height = Me.ScaleHeight - cmdButton(0).Height - 105

For c = 0 To cmdButton.Count - 1
  cmdButton(c).Top = Me.ScaleHeight - cmdButton(c).Height - 105
Next c

Exit Sub
errh:
If Err = 380 Then Resume Next 'invalid property
MsgBox Err & ": " & Err.Description & " Module: " & "frmError::Form_Resize"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Public Sub AddStatus(ByVal Text As String, Optional Buttons As fncErrorButtons = fncStatusMode, Optional TextHeader As String = "", Optional Title As String = "", Optional HelpFile As Integer = 0)
'This is run when this form is used as a status updater. Add the text to the text box
On Error GoTo errh

'Set the buttons
SetButtons Buttons

'if a header is given, add it before the text
If TextHeader <> "" Then Text = TextHeader & vbCrLf & Text

'if a title is given, set the form's caption to it
If Title = "" Then
  Me.Caption = "Status: Filename Changer"
Else
  Me.Caption = Title
End If
  

'Add it to the box, with the time in front to help separation
'txtError.Text = Time() & vbCrLf & Text & vbCrLf & String(10, "*") & vbCrLf & vbCrLf & txtError.Text
txtError.Text = Time() & vbCrLf & Text & vbCrLf & vbCrLf & txtError.Text

Me.Show

Exit Sub
errh:
MsgBox Err & ": " & Err.Description & " Module: " & "frmError::AddError"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Public Sub SetButtons(Buttons As fncErrorButtons)
On Error GoTo errh

If Buttons = fncCloseCopyThenClose Or Buttons = fncCloseOnly Then cmdButton(0).Visible = True Else cmdButton(0).Visible = False
If Buttons = fncCloseCopyThenClose Or Buttons = fncCopyThenCloseOnly Then cmdButton(1).Visible = True Else cmdButton(1).Visible = False
If Buttons = fncCopyOnly Then cmdButton(2).Visible = True Else cmdButton(2).Visible = False

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "SetButtons"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errh

'if this is a status form, just hide it
If InStr(1, Me.Caption, "Status") Then
  Me.Visible = False 'hide the form, so that the status text stays
  Cancel = 1
Else
  'otherwise STOP if debugging, or go through otherwise
  #If Debugging Then
    Stop
  #End If
End If
  


Exit Sub
errh:
MsgBoxError Err & ": " & Err.Description & " Module: " & "Form_Unload"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub
