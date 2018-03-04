VERSION 5.00
Begin VB.UserControl ActionContainer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin FilenameChanger.ActProp ActProp 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
End
Attribute VB_Name = "ActionContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ActionCount() As Integer
On Error GoTo errh

ActionCount = ActProp.Count - 1

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::ItemCount"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function

Public Sub Clear()
On Error GoTo errh

Dim c As Integer
Dim tmpTotal As Integer

tmpTotal = ActProp.Count - 1

For c = 1 To tmpTotal
  Unload ActProp(c)
Next c

With ActProp(0)
  .ActType = -1
  .Data0 = ""
  .Data1 = ""
  .Data2 = ""
  .Data3 = ""
  .Title = ""
  .Selected = ""
  ResetError 0
End With

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::Clear"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Public Function AddAction(ActType As fncActionEnum) As Integer
'Gets the type of action to create, and outputs the unique ID for that action object (it never actually uses ActProp of index 0)
On Error GoTo errh

Load ActProp(ActProp.Count)

AddAction = ActProp.Count - 1
ActProp(AddAction).ActType = ActType 'store the action type for this item

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::AddAction"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Sub SetTitle(UID As Integer, Title As String)
On Error GoTo errh

ActProp(UID).Title = Title

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::SetTitle"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Public Sub SetData(UID As Integer, DataIndex As fncDataEnum, Data As Variant)
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Sub

'set the correct item with it's new data
If DataIndex = 0 Then ActProp(UID).Data0 = Data
If DataIndex = 1 Then ActProp(UID).Data1 = Data
If DataIndex = 2 Then ActProp(UID).Data2 = Data
If DataIndex = 3 Then ActProp(UID).Data3 = Data

ResetError UID

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::SetData"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Public Function GetTitle(UID As Integer) As String
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Function

GetTitle = ActProp(UID).Title

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::GetTitle"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Function GetActType(UID As Integer) As fncActionEnum
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Function

GetActType = ActProp(UID).ActType


Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionContainer::GetActType"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Function GetData(UID As Integer, DataIndex As fncDataEnum) As Variant
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Function

Select Case DataIndex
  Case 0
    GetData = ActProp(UID).Data0
  Case 1
    GetData = ActProp(UID).Data1
  Case 2
    GetData = ActProp(UID).Data2
  Case 3
    GetData = ActProp(UID).Data3
  Case Else
    GetData = ""
End Select

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::GetData"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Property Let Selected(UID As Integer, ByVal Index As Integer, Value As Boolean)
'Keeps track of which files in the list are selected and which aren't
On Error GoTo errh

Dim tmpStr As String

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Property

'Index 0 refers to position 1 in the string
Index = Index + 1

If Index > Len(ActProp(UID).Selected) Then
  'if it is setting an index that is greater than the length of the string, add "1"'s to fill it up
  ActProp(UID).Selected = ActProp(UID).Selected & String(Index - Len(ActProp(UID).Selected), "1")
End If

tmpStr = ActProp(UID).Selected
If Value Then
  Mid(tmpStr, Index, 1) = "1"
Else
  Mid(tmpStr, Index, 1) = " "
End If
ActProp(UID).Selected = tmpStr

  

Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::Selected"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

Public Property Get Selected(UID As Integer, ByVal Index As Integer) As Boolean
'Keeps track of which files in the list are selected and which aren't
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Property

'Index 0 refers to position 1 in the string
Index = Index + 1

If Index <= Len(ActProp(UID).Selected) Then
  If Mid(ActProp(UID).Selected, Index, 1) = "1" Then
    Selected = True
  Else
    Selected = False
  End If
Else
  Selected = True
End If

Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::Selected"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

Public Property Get ErrorNumber(UID As Integer) As Integer
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Property

ErrorNumber = ActProp(UID).ErrNum

Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::GetErrorNumber"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

Public Property Get ErrorText(UID As Integer) As String
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Property

ErrorText = ActProp(UID).ErrText

Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::GetErrorText"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

Private Property Let ErrorNumber(UID As Integer, ErrNum As Integer)
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Property

ActProp(UID).ErrNum = ErrNum

Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::LetErrorNumber"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

Private Property Let ErrorText(UID As Integer, ErrText As String)
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Property

ActProp(UID).ErrText = ErrText

Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::LetErrorText"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

Private Sub ResetError(UID As Integer)
On Error GoTo errh

If UID < 0 Or UID > ActProp.Count - 1 Then Exit Sub

ErrorNumber(UID) = 0
ErrorText(UID) = ""

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::ResetError"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Public Property Get ActName(UID As Integer) As String
On Error GoTo errh



Exit Property
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActContainer::ActName"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Property

