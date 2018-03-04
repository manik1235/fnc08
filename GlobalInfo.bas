Attribute VB_Name = "GlobalInfo"
Option Explicit

Const cTOTALACTIONS = 7

Public Enum fncActionEnum
  fncBlank = 0
  fncDeleteBetween = 1
  fncCapitalization = 2
  fncSwitchCharacters = 3
  fncIncludeExcludeRules = 4
  fncConcatenate = 5
  fncReplaceCharacters = 6
  fncFilemode = 7
End Enum

Public Enum fncDataEnum
  fncConcatText = 0
  fncConcatPosition
  fncConcatLeftRight
  fncCapsOption = 0
  fncCapsPosition
  fncRulesAddRule = 0
  fncSwitch1Sel_Start = 0
  fncSwitch1SelLen
  fncSwitch2Sel_Start
  fncSwitch2SelLen
  fncDelBetLeft = 0
  fncDelBetRight
  fncDelBetLeftDir    'The left item of the Delete between action, the data in this item determines the search direction, to the left(0) or right(1)
  fncDelBetDelFirst   'Delete All (False) or Delete First (True)
  fncReplaceOld = 0
  fncReplaceNew
  fncFileModePath = 0
End Enum

Public Enum fncSampleEnum
  'Holds the Enums for txtSample
  fncSampleConcat = 0
  fncSampleSwitchR1 = 1
  fncSampleSwitchR2 = 2
End Enum

Public Enum fncFileLists
  'Holds the Enums for the file lists
  fncOrgNames     'The file list in the selection thing
  fncActPrev      'The file list in the active preview frame
  fncSampleList   'The file list that holds the Sample Textbox names
'  fncExtList      'The file list that holds the File Extentions when preserving them
End Enum

'Holds enums for the Settings function
Public Enum fncSettingsEnum
  fncExitEarly = 0  'if the settings function is already running, it will exit w/o running again. this is what it returns when it does that
  fncSaveSuccessful 'save was successful
  fncLoadSuccessful 'load was successful
  fncSaveFailed     'save failed
  fncLoadFailed     'load failed
  fncConverted      'an old save file was converted, so nothing was loaded
End Enum
'state of saving or loading
Public Enum fncSaveLoadEnum
  fncSave = 0
  fncLoad = 1
End Enum
'enums for the different settings you can save or load. right now the only ones that can be saved is the Queue
Public Enum fncSaveLoadNameEnum
  fncSLAll = 0
  fncSLQueue = 1
End Enum
'value for what to use. checkboxes will have a true/false option
Public Enum fncSaveLoadValueEnum
  fncSLUID = 0
End Enum

'used for the frmError
Public Enum fncErrorButtons
  fncCloseCopyThenClose = 0 'Main two buttons for error mode, "Close Program" and "Copy to Clipboard then close"
  fncCopyThenCloseOnly = 1  'Only "Copy to Clipboard then close"
  fncCloseOnly = 2          'Only "Close Program"
  fncCopyOnly = 3           'Only "Copy Text"
  fncErrorMode = fncCloseCopyThenClose
  fncStatusMode = fncCopyOnly
End Enum

Public Type ValidateType
  NumInvalids As Integer
  Text As String
End Type

Public Enum fncExtentionEnum
  fncGetExt = 0 'pull the extention off the filenames
  fncPutExt     'put the extentions back on the filenames
End Enum

Public Enum fncIgnoreCaseEnum
  fncTrue = vbTextCompare
  fncFalse = vbBinaryCompare
End Enum



Public Function ActionDetails(Optional ActionID As fncActionEnum = -1, Optional ActionName As String = "", Optional GetTotalActions As Boolean = False) As Variant
'Returns the name of the action if given the ID, or the action number if given a name, defaults to giving the name if both are supplied
'If true is send to GetTotalActions, return then total number of actions to add to the list (as an integer).
On Error GoTo errh

If GetTotalActions Then
  ActionDetails = cTOTALACTIONS
  Exit Function
End If

If ActionID > -1 Then
  'return the Name of this action
  Select Case ActionID
    Case fncBlank:
      ActionDetails = "Blank"
    Case fncDeleteBetween:
      ActionDetails = "Delete Between"
    Case fncCapitalization:
      ActionDetails = "Capitalization"
    Case fncSwitchCharacters:
      ActionDetails = "Switch Characters"
    'Case fncIncludeExcludeRules:
    '  ActionDetails = "Rules"
    Case fncConcatenate:
      ActionDetails = "Concatenation"
    Case fncReplaceCharacters:
      ActionDetails = "Replace Characters"
    Case fncFilemode:
      ActionDetails = "Filemode"
    Case Else
      ActionDetails = "*No Action Assigned to " & CStr(ActionID)
  End Select
Else
  Select Case ActionName
    Case "Blank":
      ActionDetails = fncBlank
    Case "Delete Between":
      ActionDetails = fncDeleteBetween
    Case "Capitalization":
      ActionDetails = fncCapitalization
    Case "Switch Characters":
      ActionDetails = fncSwitchCharacters
    'Case "Rules":
    '  ActionDetails = fncIncludeExcludeRules
    Case "Concatenation":
      ActionDetails = fncConcatenate
    Case "Replace Characters":
      ActionDetails = fncReplaceCharacters
    Case "Filemode":
      ActionDetails = fncFilemode
    Case Else
      ActionDetails = CLng(-1) 'needs to return a long
  End Select
End If

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ActionDetails"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function
