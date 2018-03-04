Attribute VB_Name = "FNCMod"
Option Explicit


'Enums for certain properties in these functions
Public Enum fncChangeCapsEnum
  fncCapNthLetter = 0
  fncToggleCaps = 1
End Enum

Public Enum fncDelBetStartEnum
  fncDBLeft = 0 'Start looking from the left of the current position
  fncDBRight    'Start looking from the right of the current position
End Enum
'Public Enum fncDelBetBoundEnum
'  fncDBBeginning = 1    'the bound is the beginning of the string
'  fncDBWholeString = -1 'the bound is the whole string
'End Enum


Public Function ReplaceChr(ByVal String1 As String, ByVal BadString As String, ByVal NewString As String, Optional ByVal Start As Long = 1, Optional ByVal ReplaceAll As Boolean = True, Optional IgnoreCase As fncIgnoreCaseEnum = fncTrue) As String
On Err GoTo errh

Dim textFore As String
Dim textAft As String

'Don't do anything if BadString is empty
If BadString = "" Then
  ReplaceChr = String1
  Exit Function
End If

' only does the one closest to where it started (to the right)
Do
  If InStr(Start, String1, BadString, IgnoreCase) Then
    textFore = Mid(String1, 1, InStr(Start, String1, BadString, IgnoreCase) - 1)
    textAft = Mid(String1, InStr(Start, String1, BadString, IgnoreCase) + Len(BadString))
    String1 = textFore & NewString & textAft
    ReplaceChr = textFore & NewString & textAft
    Start = Len(textFore) + Len(NewString) + 1
  End If
Loop While InStr(Start, String1, BadString, IgnoreCase) And ReplaceAll = True 'If they only want to replace one chr, it'll go thru only once

  
  

If ReplaceChr = "" Then ReplaceChr = String1

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: ReplaceChr"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
On Error GoTo 0
End Function

Public Function CheckIfFile(ByVal FilePath As String) As Boolean
  On Error GoTo errh
  
  ' assign to free file number
  Dim FileNumber As Integer
  FileNumber = FreeFile
  
  ' attempt to open the file. if it opens, it is a file, if not, its a folder, or its not there.
  Open FilePath For Input As #FileNumber
  
  CheckIfFile = True
    
  On Error GoTo 0
  Exit Function
errh:
If Err = 53 Or Err = 75 Or Err = 76 Then
  CheckIfFile = False
  Err.Clear
  On Error GoTo 0
  Exit Function
End If
frmError.MsgBoxError Err & ": " & Err.Description
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function

Public Function FindLast(ByVal String1 As String, ByVal FindChr As String, Optional ByVal Start As Integer = -1, Optional ByVal IgnoreCase As fncIgnoreCaseEnum = fncTrue) As Integer
On Error GoTo errh

Dim LastPos As Integer
Dim NewPos As Integer
Dim c As Integer

If Start = -1 Then
  'by default start at the end of the string
  Start = Len(String1)
End If

For c = Start To 1 Step -1
  LastPos = InStr(1, Mid(String1, c), FindChr, IgnoreCase)
  If LastPos Then Exit For
Next c
FindLast = c

On Error GoTo 0
Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function

Public Function DeleteBetween(ByVal String1 As String, ByVal LeftString As String, ByVal RightString As String, _
                              Optional ByVal StartSide As fncDelBetStartEnum = fncDBLeft, _
                              Optional ByVal DeleteAll As Boolean = True, Optional ByVal IgnoreCase As fncIgnoreCaseEnum = fncTrue) As String
On Error GoTo errh

Dim LeftPos As Integer
Dim RightPos As Integer
Dim swap As Integer
Dim RightStart As Integer 'The starting point for the Right term can change depending on the mode

Dim preDeleteBetween As String

DeleteBetween = String1

'Make sure LeftString and RightString exist, if they don't return the unmodified string
If LeftString = "" Or RightString = "" Then Exit Function

Do
  'If the string is empty, there's no reason to continue, everything was deleted. exit as such
  If DeleteBetween = "" Then Exit Function
  
  'allows me to check for changes
  preDeleteBetween = DeleteBetween
  
  'Find LeftString moving to the Left or Right as determined by StartSide
  If StartSide = fncDBRight Then
    'Search starting at the Right
    LeftPos = FindLast(DeleteBetween, LeftString, Len(DeleteBetween), IgnoreCase)
  Else
    'Search starting at the Left
    LeftPos = InStr(1, DeleteBetween, LeftString, IgnoreCase)
  End If
  
  
  
  
  'Only check for the RightString if LeftString was found
  If LeftPos > 0 Then
    'Determine where to start looking based on the mode (there is only one mode right now)
    If True Then
      'Continue from where LeftPos left off (adjust for the length of the right string so there is no overlap)
      RightPos = InStr(LeftPos + Len(LeftString), DeleteBetween, RightString, IgnoreCase)
    Else
      'Start from the right most spot in the string
      RightPos = FindLast(DeleteBetween, RightString, Len(DeleteBetween), IgnoreCase)
    End If
  End If


  'make sure leftpos is left of rightpos, and vice versa (don't swap if the locations =0, that means nothing was found
  If LeftPos > RightPos And LeftPos > 0 And RightPos > 0 Then
    swap = LeftPos
    LeftPos = RightPos
    RightPos = swap
  End If
  
  'delete as appropriate
  If LeftPos > 0 And RightPos > 0 Then DeleteBetween = Mid(DeleteBetween, 1, LeftPos - 1) & Mid(DeleteBetween, RightPos + Len(RightPos) - 1)
Loop While (Not preDeleteBetween = DeleteBetween) And DeleteAll


Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "DeleteBetween"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function



Public Function ChangeCaps(ByVal String1 As String, ByVal Mode As fncChangeCapsEnum, ByVal Letter As Integer) As String
On Error GoTo errh

'Mode 0 is "Capitalize the nth letter of each word"
'Mode 1 is "Toggle all capitalization"

Dim c As Integer
Dim l As Integer

If Mode = 0 Then
  'Mode 0 is "Capitalize the nth letter of each word"
  l = 0
  For c = 1 To Len(String1)
    l = l + 1
    If Mid(String1, c, 1) = " " Then
      'if we run into a space, reset the letter count
      l = 0
    ElseIf l = Letter Then
      'if we are on the nth letter of the current word, Caps it.
      Mid(String1, c, 1) = UCase(Mid(String1, c, 1))
    End If
  Next c
ElseIf Mode = 1 Then
  'Mode 1 is "Toggle all capitalization"
  For c = 1 To Len(String1)
    If Mid(String1, c, 1) = LCase(Mid(String1, c, 1)) Then
      Mid(String1, c, 1) = UCase(Mid(String1, c, 1))
    Else
      Mid(String1, c, 1) = LCase(Mid(String1, c, 1))
    End If
  Next c
End If

ChangeCaps = String1

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ChangeCaps"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Function SwitchRange(ByVal String1 As String, ByVal Range1Start As Integer, ByVal Range1Length As Integer, ByVal Range2Start As Integer, ByVal Range2Length As Integer) As String
On Error GoTo errh

'if any test fails, get out of function and return the same string.
If InStr(1, CheckSwitch(Range1Start, Range1Length, Range2Start, Range2Length, False, String1), "Error") Then
  SwitchRange = String1
  Exit Function
End If

'These will hold each part of the string
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim s4 As String
Dim s5 As String

'Define s1
If Range1Start > 0 Then
  s1 = Mid(String1, 1, Range1Start)
Else
  s1 = ""
End If

'Define s2
s2 = Mid(String1, Range2Start + 1, Range2Length)

'Define s3
If Range1Start + Range1Length < Range2Start Then
  s3 = Mid(String1, Range1Start + Range1Length + 1, Range2Start - (Range1Start + Range1Length))
Else
  s3 = ""
End If

'Define s4
s4 = Mid(String1, Range1Start + 1, Range1Length)

'Define s5
If Range2Start + Range2Length < Len(String1) Then
  s5 = Mid(String1, Range2Start + Range2Length + 1)
Else
  s5 = ""
End If

'Put them all together
SwitchRange = s1 & s2 & s3 & s4 & s5

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "SwitchChr"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Function CheckSwitch(ByVal R1S As Integer, ByVal R1L As Integer, ByVal R2S As Integer, ByVal R2L As Integer, Optional ByVal DisplayOutput As Boolean = False, Optional ByVal Example As String = "") As String
'If DisplayOutput is set to true, Returns descriptive errors
'If DisplayOutput is set to false, Returns only basic errors
On Error GoTo errh

'Make sure 0 <= Range1Start, Range1Length > 0, Range2Start > Range1Start + Range1Length, Range2Start + Range2Length <= Len(string1)
CheckSwitch = ""
If R1S < 0 Then
  If DisplayOutput Then
    CheckSwitch = CheckSwitch & "Error A: Range 1 must begin at at least 0." & vbCrLf
  Else
    CheckSwitch = CheckSwitch & "Error A:"
  End If
End If
If R1L < 1 Then
  If DisplayOutput Then
    CheckSwitch = CheckSwitch & "Error B: Range 1's length must be at least 1." & vbCrLf
  Else
    CheckSwitch = CheckSwitch & "Error B:"
  End If
End If
If R2S < R1S + R1L Then
  If DisplayOutput Then
    CheckSwitch = CheckSwitch & "Error C: Range 2 must begin after Range 1 ends. (R1 ends at " & Str(R1S + R1L) & ")" & vbCrLf
  Else
    CheckSwitch = CheckSwitch & "Error C:"
  End If
End If
If R2L < 1 Then
  If DisplayOutput Then
    CheckSwitch = CheckSwitch & "Error D: Range 2's length must be at least 1." & vbCrLf
  Else
    CheckSwitch = CheckSwitch & "Error D:"
  End If
End If
If R2S + R2L > Len(Example) Then
  If DisplayOutput Then
    CheckSwitch = CheckSwitch & "Warning! (Error E): The sample text isn't long enough to show the switch. Add" & Str(R2S + R2L - Len(Example)) & " more characters to see the output. This switch will still work on filename that are at least" & Str(R2S + R2L) & " characters long." & vbCrLf
  Else
    CheckSwitch = CheckSwitch & "Error E:"
  End If
End If

'If an error occured, leave now
If InStr(1, CheckSwitch, "Error") Then Exit Function
  
'If an example was provided, show the output if it is valid
If Not Example = "" And DisplayOutput Then
  CheckSwitch = "Switched:" & vbCrLf & SwitchRange(Example, R1S, R1L, R2S, R2L) & vbCrLf & "Original:" & vbCrLf & Example
End If
    
  
Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "CheckSwitch"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Function Concat(ByVal String1 As String, ByVal StringAdd As String, ByVal Position As Integer, Optional ByVal LeftRight As Long = 0) As String
'If LeftRight = 0, add from the Left. If it is 1, add from the Right
'If BeforeAfter = 0, Add before. if its 1, add after
On Error GoTo errh

Dim s1 As String
Dim s2 As String

If LeftRight = 0 Then
  'Counting from Left
  If Position >= 1 Then
    s1 = Left(String1, Position)
  Else
    s1 = ""
  End If
  If Position + 1 < Len(String1) Then
    If Position + 1 <= 0 Then
      'if the position is 0 or less, take the whole string
      s2 = String1
    Else
      s2 = Mid(String1, Position + 1)
    End If
  Else
    s2 = ""
  End If
Else
  'Counting from Right
  If Position >= 1 Then
    s2 = Right(String1, Position)
  Else
    s2 = ""
  End If
  If Position < Len(String1) Then
    If Position <= 0 Then
      'if the position is 0 or less, take the whole string
      s1 = String1
    Else
      s1 = Left(String1, Len(String1) - Position)
    End If
  Else
    s1 = ""
  End If
End If

Concat = s1 & StringAdd & s2


Exit Function
errh:
If Err = 5 Then
  'Invalid procedure call or argument, the Right or Left is out of range
  Concat = "Error 5: Concatenation position out of range"
  #If Debugging = 1 Then
    Stop
  #End If
  Exit Function
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Concat"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Public Function FileMode(NewF As ListBox, Optional ByVal FileIndex As Integer = -1, Optional ByVal DefaultName As String = "", Optional ByVal FMPath As String = "") As String
'NewF is a string of filenames, separated with vbCrLf
'Returns the filename 'FileIndex' files down a list
'If a path is passed through FMPath, create the listbox of items from that file
'DefaultName is the string to return if the requested index is past the number of entries in the list
On Error GoTo errh

Dim FileNum As Integer
Dim Data As String
Dim c As Integer
  
If FMPath <> "" Then
  FileNum = FreeFile()
  
  NewF.Clear
  Open FMPath For Input As #FileNum
    While Not EOF(FileNum)
      Line Input #FileNum, Data
      NewF.AddItem Data
    Wend
  Close #FileNum
End If

If FileIndex < NewF.ListCount Then
  FileMode = NewF.List(FileIndex)
Else
  FileMode = DefaultName
End If

Exit Function
errh:
If Err = 76 Then
  'Path not found
  FileMode = DefaultName
  'Stop
  #If Debugging = 1 Then
    Stop
  #Else
    Exit Function
  #End If
ElseIf Err = 53 Then
  'File not found
  FileMode = DefaultName
  #If Debugging = 1 Then
    Stop
  #Else
    Exit Function
  #End If
ElseIf Err = 75 Then
  'Path/File access error
  FileMode = DefaultName
  #If Debugging = 1 Then
    Stop
  #Else
    Exit Function
  #End If
End If


frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "FileMode"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function



