VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Filename Changer"
   ClientHeight    =   12915
   ClientLeft      =   3180
   ClientTop       =   2310
   ClientWidth     =   20085
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   20085
   Begin VB.ListBox lstSorted 
      Height          =   255
      Index           =   0
      Left            =   10800
      Sorted          =   -1  'True
      TabIndex        =   109
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Global"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   103
      ToolTipText     =   "Global (Match the Active Preview checks to the Global list in the Path window)"
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "None"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   102
      ToolTipText     =   "Select None"
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "All"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   101
      ToolTipText     =   "Select All"
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Toggle"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   100
      ToolTipText     =   "Toggle Selected"
      Top             =   6360
      Width           =   735
   End
   Begin VB.ListBox lstFiles 
      Height          =   285
      Index           =   2
      Left            =   9360
      OLEDropMode     =   1  'Manual
      Style           =   1  'Checkbox
      TabIndex        =   97
      ToolTipText     =   "Complete List of Files"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   9360
      Top             =   0
   End
   Begin VB.Frame framePath 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Path (The ""Global"" file list is in here)"
      Height          =   975
      HelpContextID   =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.ListBox lstFiles 
         Height          =   4110
         Index           =   0
         Left            =   3840
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   82
         ToolTipText     =   "Complete List of Files"
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CommandButton cmdPathPaste 
         Height          =   300
         HelpContextID   =   1
         Left            =   7440
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Paste Path, Directory or File Location"
         Top             =   240
         Width           =   300
      End
      Begin VB.CheckBox chkHidSys 
         Caption         =   "Show System Files"
         Height          =   255
         HelpContextID   =   1
         Index           =   1
         Left            =   1800
         TabIndex        =   24
         Top             =   600
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkHidSys 
         Caption         =   "Show Hidden Files"
         Height          =   255
         HelpContextID   =   1
         Index           =   0
         Left            =   120
         MaskColor       =   &H8000000F&
         TabIndex        =   23
         Top             =   600
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.FileListBox File1 
         Height          =   285
         HelpContextID   =   1
         Hidden          =   -1  'True
         Left            =   3840
         OLEDropMode     =   1  'Manual
         System          =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.DirListBox Dir1 
         Height          =   4140
         HelpContextID   =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   21
         Top             =   1440
         Width           =   3615
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         HelpContextID   =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CommandButton cmdPath 
         Height          =   300
         HelpContextID   =   1
         Left            =   7800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Expand or Collapse File Window"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         HelpContextID   =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   240
         Width           =   7215
      End
      Begin VB.Image imgPathDown 
         Height          =   225
         Left            =   4680
         Picture         =   "frmMain.frx":0856
         Top             =   600
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgPathUp 
         Height          =   225
         Left            =   4320
         Picture         =   "frmMain.frx":0B68
         Top             =   600
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11880
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin FilenameChanger.ActionContainer ActCont 
      Left            =   8280
      Top             =   360
      _ExtentX        =   1720
      _ExtentY        =   1085
   End
   Begin VB.Frame framePreview 
      Caption         =   "Active Preview"
      Height          =   2415
      Left            =   4200
      TabIndex        =   74
      Top             =   5760
      Width           =   5055
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   2400
         TabIndex        =   84
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkActivePreview 
         Caption         =   "Enable Active Preview"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.ListBox lstFiles 
         Height          =   1635
         Index           =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   75
         ToolTipText     =   "Complete List of Files"
         Top             =   600
         Width           =   9015
      End
   End
   Begin VB.Frame frameProperties 
      Caption         =   "Action Properties"
      Height          =   10815
      Left            =   9360
      TabIndex        =   26
      Top             =   1080
      Width           =   10335
      Begin VB.TextBox txtSample 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         HideSelection   =   0   'False
         Index           =   3
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   3375
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0 Blank"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   600
         Width           =   3375
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00FFC0C0&
         Caption         =   "7 Filemode"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   7
         Left            =   5880
         TabIndex        =   79
         Top             =   2640
         Width           =   3375
         Begin VB.ListBox lstFileMode 
            Height          =   2010
            Left            =   120
            TabIndex        =   94
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox txtFileModePath 
            Height          =   285
            Left            =   720
            TabIndex        =   81
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Path"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00C0FFFF&
         Caption         =   "6 Replace Characters"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   6
         Left            =   5520
         TabIndex        =   28
         Top             =   2280
         Width           =   3375
         Begin VB.CommandButton cmdReplace 
            Caption         =   "Use ^"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   39
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton cmdReplace 
            Caption         =   "Use ^"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   38
            Top             =   720
            Width           =   615
         End
         Begin VB.ListBox lstReplace 
            Height          =   1815
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox txtReplace 
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   35
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtReplace 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "with"
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00FFFFC0&
         Caption         =   "5 Concat"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   5
         Left            =   5280
         TabIndex        =   33
         Top             =   1920
         Width           =   3375
         Begin VB.TextBox txtConcat 
            Height          =   285
            Index           =   2
            Left            =   1920
            TabIndex        =   93
            Top             =   1560
            Width           =   375
         End
         Begin VB.Frame frameConcatSub 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   1200
            TabIndex        =   89
            Top             =   2160
            Width           =   855
            Begin VB.OptionButton optConcat1 
               Caption         =   "Right"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   91
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optConcat1 
               Caption         =   "Left"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   90
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.TextBox txtSample 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   87
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtConcat 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   59
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtConcat 
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   58
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label lblConcat 
            Alignment       =   2  'Center
            Caption         =   "between letters"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   92
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblConcat 
            Caption         =   "Start counting from the"
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   88
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblConcat 
            Caption         =   "Sample Name"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblConcat 
            Alignment       =   2  'Center
            Caption         =   "and"
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   60
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblConcat 
            Caption         =   "Add"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00C0FFC0&
         Caption         =   "4 Rules"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   4
         Left            =   4800
         TabIndex        =   32
         Top             =   1320
         Width           =   3375
         Begin VB.OptionButton optRules 
            Caption         =   "E&xclude"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   62
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optRules 
            Caption         =   "&Include"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ListBox lstRules 
            Height          =   2010
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   3135
         End
         Begin VB.CommandButton cmdRules 
            Caption         =   "Add"
            Height          =   255
            Left            =   2640
            TabIndex        =   55
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtRules 
            Height          =   285
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00C0E0FF&
         Caption         =   "3 Switch"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   3
         Left            =   4440
         TabIndex        =   31
         Top             =   960
         Width           =   3375
         Begin VB.TextBox txtSwitch 
            Height          =   855
            Index           =   6
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   2280
            Width           =   3135
         End
         Begin VB.TextBox txtSwitch 
            Height          =   285
            Index           =   3
            Left            =   2040
            TabIndex        =   72
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtSwitch 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   70
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtSwitch 
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   68
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtSwitch 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   65
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtSample 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            HideSelection   =   0   'False
            Index           =   2
            Left            =   840
            TabIndex        =   53
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtSample 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            HideSelection   =   0   'False
            Index           =   1
            Left            =   840
            TabIndex        =   52
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Result (or Errors)"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   73
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Range Length"
            Height          =   495
            Index           =   1
            Left            =   1440
            TabIndex        =   71
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Begin at character"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Range Length"
            Height          =   495
            Index           =   4
            Left            =   1440
            TabIndex        =   67
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Range 2"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   66
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Begin at character"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblSwitch 
            Caption         =   "Range 1"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00FFC0FF&
         Caption         =   "2 Capitalization"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   2
         Left            =   4080
         TabIndex        =   30
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtCap 
            Height          =   285
            Left            =   1440
            TabIndex        =   48
            Text            =   "1"
            Top             =   360
            Width           =   375
         End
         Begin VB.OptionButton optCap 
            Caption         =   "Toggle all capitalization"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton optCap 
            Caption         =   "Capitalize the"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "letter of each word"
            Height          =   255
            Left            =   2040
            TabIndex        =   50
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblOrdinal 
            Caption         =   "st"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   49
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame frameProps 
         BackColor       =   &H00C0C0FF&
         Caption         =   "1 DeleteBetween"
         Enabled         =   0   'False
         Height          =   3375
         Index           =   1
         Left            =   360
         TabIndex        =   29
         Top             =   5760
         Width           =   3375
         Begin VB.Frame frameDelBet 
            BackColor       =   &H00FF80FF&
            Caption         =   "Delete"
            Height          =   825
            Index           =   1
            Left            =   1680
            TabIndex        =   117
            Top             =   1800
            Width           =   1455
            Begin VB.OptionButton optDelBet 
               BackColor       =   &H00C0C0FF&
               Caption         =   "All found"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   119
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optDelBet 
               BackColor       =   &H00C0C0FF&
               Caption         =   "First found"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   118
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkDelBet 
            BackColor       =   &H0080C0FF&
            Caption         =   "Continue from Last"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   2880
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame frameDelBet 
            BackColor       =   &H00FF80FF&
            Caption         =   "Begin from the..."
            Height          =   825
            Index           =   0
            Left            =   120
            TabIndex        =   113
            Top             =   1800
            Width           =   1455
            Begin VB.OptionButton optDelBet 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Right"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   115
               Top             =   480
               Width           =   735
            End
            Begin VB.OptionButton optDelBet 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Left"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   114
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Use ^"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   45
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Use ^"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   44
            Top             =   480
            Width           =   615
         End
         Begin VB.ListBox lstDelete 
            Height          =   840
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox txtDelete 
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   42
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox txtDelete 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "and"
            Height          =   255
            Left            =   1560
            TabIndex        =   41
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin VB.Frame frameRun 
      Caption         =   "Execute Queue"
      Height          =   1095
      Left            =   0
      TabIndex        =   18
      Top             =   4560
      Width           =   3975
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Validate Filenames"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   110
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute Queue"
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   63
         Top             =   240
         Width           =   1695
      End
      Begin VB.ListBox lstExecute 
         Height          =   255
         Left            =   2520
         TabIndex        =   83
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame frameRecycle 
      Caption         =   "Recycled Actions (Double Click to Restore)"
      Height          =   3015
      Left            =   9360
      TabIndex        =   16
      Top             =   5160
      Width           =   3615
      Begin VB.ListBox lstRecycle 
         Height          =   2595
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdHelp 
      Height          =   375
      Left            =   8880
      Picture         =   "frmMain.frx":0E7A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Displays the help topic associated with the object you click."
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Frame frameQ 
      Caption         =   "Queue"
      Height          =   3375
      HelpContextID   =   3
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
      Begin VB.CommandButton cmdQ 
         Height          =   255
         Index           =   4
         Left            =   4680
         Picture         =   "frmMain.frx":0F7C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Undo Delete"
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton cmdQ 
         Height          =   495
         Index           =   2
         Left            =   4440
         Picture         =   "frmMain.frx":107E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Copy Action"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdQ 
         Height          =   495
         Index           =   3
         Left            =   4440
         Picture         =   "frmMain.frx":1180
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete Action"
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdQ 
         Height          =   495
         HelpContextID   =   3
         Index           =   1
         Left            =   4440
         Picture         =   "frmMain.frx":15C2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move Action Down"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdQ 
         Height          =   495
         HelpContextID   =   3
         Index           =   0
         Left            =   4440
         Picture         =   "frmMain.frx":1A04
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move Action Up"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.ListBox lstQ 
         Height          =   2985
         HelpContextID   =   3
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frameActions 
      Caption         =   "Actions"
      Height          =   3375
      HelpContextID   =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
      Begin VB.Frame frameAdd 
         Caption         =   "Add to Queue"
         Height          =   2175
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   1575
         Begin VB.CommandButton cmdAddAction 
            Caption         =   "At the &End"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddAction 
            Caption         =   "&After Selected"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddAction 
            Caption         =   "In the &Front"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddAction 
            Caption         =   "&Before Selected"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.ListBox lstActions 
         Height          =   2985
         HelpContextID   =   2
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "     Options"
      Height          =   2655
      Left            =   0
      TabIndex        =   95
      Top             =   5760
      Width           =   4140
      Begin VB.CheckBox chkOptions 
         Caption         =   "Ignore Case"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   112
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame frameOptionsSubDir 
         Caption         =   "Output Subdirectory"
         Height          =   2295
         Left            =   1800
         TabIndex        =   104
         Top             =   240
         Width           =   2295
         Begin VB.ListBox lstSubDir 
            Height          =   1425
            Left            =   120
            TabIndex        =   108
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtSubDir 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   105
            Text            =   "\"
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Keep Extention"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   98
         Top             =   2280
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdOptions 
         Height          =   300
         Index           =   0
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":1E46
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Expand or Collapse the Options Window"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.Frame frameOptionsSelected 
         Caption         =   "Selected"
         Height          =   1695
         Left            =   0
         TabIndex        =   99
         Top             =   360
         Width           =   1695
         Begin VB.OptionButton optOptionsSel 
            Caption         =   "Active Preview Only"
            Height          =   615
            Index           =   1
            Left            =   720
            TabIndex        =   107
            Top             =   840
            Width           =   950
         End
         Begin VB.OptionButton optOptionsSel 
            Caption         =   "Change Global"
            Height          =   495
            Index           =   0
            Left            =   720
            TabIndex        =   106
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.Image imgArrowLeft 
         Height          =   225
         Left            =   1440
         Picture         =   "frmMain.frx":2158
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgArrowRight 
         Height          =   225
         Left            =   1080
         Picture         =   "frmMain.frx":246A
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Queue"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Queue"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Queue"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuShowStatus 
         Caption         =   "Show &Status Window"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuHelpRoot 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#Const Debugging = 1

'When you click the help button, change this so everything knows its in help mode
Dim InHelpMode As Boolean

'Keeps track of the Z-position of the Path frame
Dim PathZOrder As Integer

'Set the Large and Small sizes of the Resizable frames (in twips)
Const FRAMEPATHSMALL = 975
Const FRAMEPATHLARGE = 5775
Const FRAMEOPTIONSSMALL = 300
Const FRAMEOPTIONSLARGE = 4140

'Keep track of the number of IDs
Dim IDCount As Integer

'Keep track of the UID who's properties are being shown
Dim UIDIndex As Integer

'keep track of which control has focus in the txtSwitch property so that the currently edited control isn't updated while editing
Dim curSwitchIndex As Integer
'keep track of which control has focus in the txtSample property so that the currently edited control isn't updated while editing
Dim curSampleIndex As fncSampleEnum
Dim curConcatIndex As Integer

'True if the path is being sent to other objects, so if an object's path changes it shouldn't resend this new path, False if its ok to send
Dim CascadingPath As Boolean

'Collections used throughout the program to move groups of objects and whatnot
Dim colPathObjs As New Collection

'true if Settings are being saved or loaded, so that other functions don't call it too
Dim InSettings As Boolean

'This is a copy of frmError, but needs to be able to be new
Dim frmStatus As New frmError

'These are the object references to use throughout the program
Dim clstOrgNames As Object 'The list of original file names
Dim clstActPrev As Object  'The list in the active preview window
Dim clstSample As Object   'The list that holds the changes for the sample text boxes (one less than the selected change)
Dim clstSorted As Object   'The list that holds the sorted data for duplicate checking

'This is the current version of the save files
Const cSAVE_VER = 2


Private Sub chkActivePreview_Click()

On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help chkActivePreview.HelpContextID
  Exit Sub
End If


UpdateFilter lstQ, clstOrgNames, clstActPrev


Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "chkActivePreview_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub chkHidSys_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help chkHidSys(Index).HelpContextID
  Exit Sub
End If

File1.Hidden = chkHidSys(0).Value
File1.System = chkHidSys(1).Value

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "chkHidSys_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub






Private Sub chkOptions_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help chkOptions(Index).HelpContextID
  Exit Sub
End If

If Index = 0 Or Index = 1 Then
  'Ignore Case OR Keep extention
  MakeChanges lstQ, clstOrgNames, clstActPrev, lstQ.ListIndex
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "chkOptions_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdAddAction_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdAddAction(Index).HelpContextID
  Exit Sub
End If

If lstActions.ListIndex = -1 Or lstActions.ListCount = 0 Then Exit Sub 'don't continue if there are no items selected/in list

Select Case Index
  Case 0
    'Add to the front of the list
    lstQ.AddItem lstActions.List(lstActions.ListIndex), 0
    lstQ.ItemData(lstQ.NewIndex) = AssignID(lstActions.ItemData(lstActions.ListIndex), clstOrgNames)
  Case 1
    If lstQ.ListIndex = -1 Then
      'if nothing is selected, add to the beginning
      lstQ.AddItem lstActions.List(lstActions.ListIndex), 0
      lstQ.ItemData(lstQ.NewIndex) = AssignID(lstActions.ItemData(lstActions.ListIndex), clstOrgNames)
    Else
      'add before selection in queue
      lstQ.AddItem lstActions.List(lstActions.ListIndex), lstQ.ListIndex
      lstQ.ItemData(lstQ.NewIndex) = AssignID(lstActions.ItemData(lstActions.ListIndex), clstOrgNames)
    End If
  Case 2
    If lstQ.ListIndex = -1 Then
      'if nothing is selected, add to the end
      lstQ.AddItem lstActions.List(lstActions.ListIndex), lstQ.ListCount
      lstQ.ItemData(lstQ.NewIndex) = AssignID(lstActions.ItemData(lstActions.ListIndex), clstOrgNames)
    Else
      'add after selected
      lstQ.AddItem lstActions.List(lstActions.ListIndex), lstQ.ListIndex + 1
      lstQ.ItemData(lstQ.NewIndex) = AssignID(lstActions.ItemData(lstActions.ListIndex), clstOrgNames)
    End If
  Case 3
    'add to the end
    lstQ.AddItem lstActions.List(lstActions.ListIndex), lstQ.ListCount
    lstQ.ItemData(lstQ.NewIndex) = AssignID(lstActions.ItemData(lstActions.ListIndex), clstOrgNames)
End Select
lstQ.ListIndex = lstQ.NewIndex 'always set the newly added item active

Update_lstQ

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdAddAction_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub cmdDelete_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdDelete(Index).HelpContextID
  Exit Sub
End If

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub


Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdDelete_Click"
Resume

End Sub

Private Sub cmdExecute_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdExecute(Index).HelpContextID
  Exit Sub
End If

Dim strErrors As String

If Index = 0 Then 'Execute button
  'Make sure all the files are updated, and go for it!
  MakeChanges lstQ, clstOrgNames, clstActPrev, -1, True  '-1 will go through all of lstQ no matter what, and Active preview needs to be true too
  frmStatus.AddStatus "Execution Started. Changing " & clstActPrev.SelCount & " files." & vbCrLf & String(30, "*")
  doChanges clstOrgNames, clstActPrev, File1
ElseIf Index = 1 Then 'Validate button
  strErrors = ValidateFilenames(clstActPrev, clstSorted)
  If strErrors = "" Then strErrors = "No invalid or duplicate file names detected."
  strErrors = strErrors & vbCrLf & String(30, "*") 'add the separation stuff to the end
  frmStatus.AddStatus strErrors 'Send the results to the Status form
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdExecute_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Function ValidateFilenames(F As ListBox, SortedList As ListBox) As String
'Returns "" if everything is valid
'SortedList is a listbox that has the Sorted property as true
On Error GoTo errh

Dim strInvalidErrors As String 'holds the names of any Invalid files and the symbols that violate
Dim strDupErrors As String 'holds the names of any duplicate files
Dim strErrors As String 'holds the Invalid and Dup errors strings
Dim tmpErrors As String 'holds the different symbols that violate
Dim c As Integer
Dim d As Integer
Dim r As Integer
Dim strIC As String 'invalid characters to check for
Dim InvalidFileCount As Integer 'number of files that have invalid characters
Dim DupFileCount As Integer 'number of names that are duplicates


strIC = "\/:*?<>|" & Chr(34)
InvalidFileCount = 0

'Stop

If SortedList.Sorted Then
  'List is sorted, this will work, so copy the one we're validating to the sorted list
  SortedList.Clear
  For c = 0 To F.ListCount - 1
    SortedList.AddItem F.List(c)
  Next c
Else
  'List is not sorted, add to the status and don't check for duplicates
  frmStatus.AddStatus "Duplicate checking failed to complete. Skipping pre-check.", fncCloseOnly
End If
  
'Check for problems in the filenames
For c = 0 To F.ListCount - 1
  'Check for invalid characters in the filenames
  tmpErrors = ""
  For d = 1 To Len(strIC)
    If InStr(1, F.List(c), Mid(strIC, d, 1)) Then
      tmpErrors = tmpErrors & Mid(strIC, d, 1) & " " 'add the offending symbol to the string
    End If
  Next d
  'if invalid characters were found, add them to the list
  If tmpErrors <> "" Then
    strInvalidErrors = strInvalidErrors & "'" & F.List(c) & "'" & " contains " & "'" & tmpErrors & "'" & vbCrLf
    InvalidFileCount = InvalidFileCount + 1
  End If
  
  'Check for duplicate names, if given a sorted list AND we aren't on the last item
  If SortedList.Sorted And c < F.ListCount - 1 Then
    If SortedList.List(c) = SortedList.List(c + 1) Then
      DupFileCount = DupFileCount + 1
      If InStr(1, strDupErrors, SortedList.List(c)) = 0 Then
        'if the name appears multiple times, don't add it again
        strDupErrors = strDupErrors & "'" & SortedList.List(c) & "'" & vbCrLf
      End If
    End If
  End If
Next c



  
'return the errors found
If strInvalidErrors <> "" Then
  strErrors = "The following " & CStr(InvalidFileCount) & " files contain invalid characters:" & vbCrLf & strInvalidErrors & vbCrLf
End If
If strDupErrors <> "" Then
  strErrors = strErrors & "The following file names are duplicates and therefore invalid:" & vbCrLf & strDupErrors
End If
ValidateFilenames = strErrors

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ValidateFilenames"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Private Sub doChanges(OrgFile As ListBox, ActPrev As ListBox, SystemFileList As Object)
'SystemFileList is File1, which is updated if the contents of the folder change while the program is running. Returns the errors
On Error GoTo errh
'change the file at index c of OrgFile to the name if ActPrev

Dim c As Integer
Dim r As Integer
Dim FromPath As String 'where to get the file names from
Dim ToPath As String   'where to put them (only different if a subDir is set)
Dim strErrors As String 'holds the errors returned by Validate Filenames


strErrors = ""   'holds the list of errors that happen, or if it is successful, returns just that

'Check the validity of the path
FromPath = txtPath.Text
If Dir(FromPath, vbDirectory) = "" Then
  MsgBox "Invalid Path Selected: " & vbCrLf & FromPath, vbCritical, "Aborting Execution"
  Exit Sub
End If

'Make sure the folder contents haven't changed since the program has updated them
If FilesListChanged(OrgFile, SystemFileList) Then
  r = MsgBox("The contents of the folder have changed since it was last checked. This may affect the changes you plan to make. Select 'Yes' to abort and refresh, or 'No' to continue anyway.", vbYesNo, "Folder Contents Changed")
  If r = vbYes Then
    'CopyList SystemFileList, OrgFile
    'CopyList SystemFileList, ActPrev
    CascadePath "txtPath", txtPath.Text 'This function will update the lstboxes, and using the path text box seems most logical
    Err.Raise -333, , "Execution aborted by user; Folder contents changed." 'random, custom error used to abort and refresh the lists
    Exit Sub
  End If
End If

'Make sure each filename is actually a valid filename
strErrors = ValidateFilenames(ActPrev, clstSorted)
If strErrors <> "" Then
  'errors were found
  frmStatus.AddStatus strErrors, , "Filename Validation Results"
  frmStatus.AddStatus "Execution Aborted."
  Exit Sub
End If

'add the ending slash to the ToPath if there isn't one
If Right(FromPath, 1) <> "\" Then
  FromPath = FromPath & "\"
End If
'add the ending slash to the FromPath if there isn't one
If Left(txtSubDir.Text, 1) = "\" Then
  ToPath = FromPath & Mid(txtSubDir.Text, 2)
Else
  ToPath = FromPath & txtSubDir.Text
End If
If Right(ToPath, 1) <> "\" Then
  ToPath = ToPath & "\"
End If

'Duplicate checking is done on the Name line by the file system
For c = 0 To OrgFile.ListCount - 1
  If (FromPath & OrgFile.List(c)) <> (ToPath & ActPrev.List(c)) Then 'only rename if the new and old names are different, including path just incase they're being moved to a new subdirectory
    Name FromPath & OrgFile.List(c) As ToPath & ActPrev.List(c)
  End If
Next c

'Update the  file lists
CopyList SystemFileList, OrgFile
CopyList SystemFileList, ActPrev

'doChanges = doChanges & "Execution Complete."
frmStatus.AddStatus "Execution Complete."

Exit Sub
errh:
If Err = 58 Then
  'File already exists
  Stop
End If
If Err = -333 Then
  'Error number defined by me, means the contents of the folder have changed, and the run is being aborted so the lists can be refreshed.
  frmStatus.AddStatus "Error " & Err.Number & ": " & Err.Description
  Resume Next
End If
If Err = 75 Then
  'Path/File access error (File open or duplicate name) Record in error log
  frmStatus.AddStatus "Error " & Err.Number & ": " & Err.Description & ". Original File Name: " & Chr(34) & OrgFile.List(c) & Chr(34) & " New File Name:" & Chr(34) & ActPrev.List(c) & vbCrLf
  Resume Next
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "doChanges"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub MakeChanges(Q As ListBox, OrgFile As ListBox, ActPrev As ListBox, Optional ByVal ChangeTo As Integer = -1, Optional ActivePreviewEnabled As Boolean = False)     'Queue Listbox and Files Listbox
On Error GoTo errh

Dim c As Integer
Dim cq As Integer 'for the Queue list
Dim cf As Integer 'for the File list
Dim UID As Integer 'holds the unique ID of the queue item
Dim IgnoreCase As fncIgnoreCaseEnum

'Holds the different data items for sending to the functions
Dim DataA As Variant
Dim DataB As Variant
Dim DataC As Variant
Dim DataD As Variant
Dim DataE As Variant

Dim ActPrevVisState As Boolean 'make the Act Prev window invisible when updating for speed

'Counts the times the module is called
#If Debugging = 1 Then
  frmDebug.dModuleCount "MakeChanges"
#End If

'Holds the Currently selected item in the file lists so they can be reset properly
Dim CurOrgFileIndex As Integer
Dim CurActPrevIndex As Integer
CurOrgFileIndex = OrgFile.ListIndex
CurActPrevIndex = ActPrev.ListIndex

Dim RULES As String 'holds the list of rules that are currently being applied
RULES = ""

'Reset the Active Preview list of files to what is in clstOrgNames
CopyList OrgFile, ActPrev, chkOptions(1).Value

'If ChangeTo = -1 (or if the value is greater than the number of queued actions), do all the items in the queue, otherwise only go as far as ChangeTo
If ChangeTo = -1 Or ChangeTo > Q.ListCount - 1 Then
  ChangeTo = Q.ListCount - 1
End If

'set the listbox being changed's Enable so it's Click routine doesn't go
ActPrev.Tag = "noclicks"
ActPrevVisState = ActPrev.Visible
ActPrev.Visible = False

'If the default of False is sent, change it to the checkmark
If Not ActivePreviewEnabled Then ActivePreviewEnabled = chkActivePreview.Value

'If the Ignore Case check is checked, do that
If chkOptions(0).Value Then
  IgnoreCase = fncTrue
Else
  IgnoreCase = fncFalse
End If

If ActivePreviewEnabled Then
  For cq = 0 To ChangeTo 'move through each Queue item in the list
    UID = Q.ItemData(cq) 'get the Unique ID for the current action item in the queue
    
    'If the action doesn't require file manipulation, handle it first.
    If ActCont.GetActType(UID) = fncIncludeExcludeRules Then 'Include/exclude
      RULES = ActCont.GetData(UID, 0)
    End If
    
    'Create the list of new file names if this action is file mode
    If ActCont.GetActType(UID) = fncFilemode And ActivePreviewEnabled Then
      FileMode lstFileMode, 0, , ActCont.GetData(UID, 0)
    End If
    
    
    'Select or Unselect the files according to the Action's settings
    For cf = 0 To ActPrev.ListCount - 1
      ActPrev.Selected(cf) = ActCont.Selected(Q.ItemData(cq), cf)
    Next cf
  

    'If the action does require file manipulation, handle it here
    For cf = 0 To ActPrev.ListCount - 1 'move through each file in the list
      If ActPrev.Selected(cf) And FollowsRules(ActPrev.List(cf), RULES) Then 'make sure it has a check next to it AND make sure it follows all the current rules
        If ActCont.GetActType(UID) = fncReplaceCharacters Then
          'Replace characters mode
          DataA = CStr(ActPrev.List(cf))
          DataB = CStr(ActCont.GetData(UID, fncReplaceOld))
          DataC = CStr(ActCont.GetData(UID, fncReplaceNew))
          ActPrev.List(cf) = ReplaceChr(DataA, DataB, DataC, , , IgnoreCase)
        ElseIf ActCont.GetActType(UID) = fncDeleteBetween Then
          'Delete everything between
          DataA = CStr(ActPrev.List(cf))                                  'Filename to change
          DataB = CStr(ActCont.GetData(UID, fncDelBetLeft))               'LeftString
          DataC = CStr(ActCont.GetData(UID, fncDelBetRight))              'RightString
          DataD = Val(ActCont.GetData(UID, fncDelBetLeftDir))             'Search Direction
          DataE = Not CBool(Val(ActCont.GetData(UID, fncDelBetDelFirst))) 'DeleteAll: take the opposite of this value because it records if you want just the First, the function asks if you want All. I do it this way so that the default value of false will default to replaceing all of them
          ActPrev.List(cf) = DeleteBetween(DataA, DataB, DataC, DataD, DataE, IgnoreCase)
        ElseIf ActCont.GetActType(UID) = fncCapitalization Then
          'Capitalization
          DataA = CStr(ActPrev.List(cf))
          DataB = Val(ActCont.GetData(UID, 0))
          DataC = Val(ActCont.GetData(UID, 1))
          ActPrev.List(cf) = ChangeCaps(DataA, DataB, DataC)
        ElseIf ActCont.GetActType(UID) = fncSwitchCharacters Then
          'Switch Characters
          DataA = CStr(ActPrev.List(cf))
          DataB = Val(ActCont.GetData(UID, 0))
          DataC = Val(ActCont.GetData(UID, 1))
          DataD = Val(ActCont.GetData(UID, 2))
          DataE = Val(ActCont.GetData(UID, 3))
          ActPrev.List(cf) = SwitchRange(DataA, DataB, DataC, DataD, DataE)
        ElseIf ActCont.GetActType(UID) = fncConcatenate Then
          'Concatenation
          DataA = CStr(ActPrev.List(cf))
          DataB = ActCont.GetData(UID, fncConcatText)
          DataC = Val(ActCont.GetData(UID, fncConcatPosition))
          DataD = Val(ActCont.GetData(UID, fncConcatLeftRight))
          ActPrev.List(cf) = Concat(DataA, DataB, DataC, DataD)
        ElseIf ActCont.GetActType(UID) = fncFilemode Then
          'File Mode
          DataB = CInt(cf)
          DataC = CStr(ActPrev.List(cf))
          ActPrev.List(cf) = FileMode(lstFileMode, DataB, DataC)
        End If
      End If
    Next cf
  Next cq
End If

'Put the ListIndex back to where it was before
If CurOrgFileIndex < OrgFile.ListCount Then OrgFile.ListIndex = CurOrgFileIndex
If CurActPrevIndex < ActPrev.ListCount Then ActPrev.ListIndex = CurActPrevIndex

'set the listbox being changed's tag so it's Click routine can go
ActPrev.Tag = ""
ActPrev.Visible = ActPrevVisState

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "MakeChanges"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub CopyList(FullList As Object, BlankList As Object, Optional RemoveExtention As Boolean = False)
'RemoveExtention pulls the extention off before it goes into BlankList
On Error GoTo errh

Dim c As Integer
Dim BlankVisState As Boolean 'if updating a list that can be seen, get rid of it for speed

'Counts the times the module is called
#If Debugging = 1 Then
  frmDebug.dModuleCount "CopyList " & FullList.Name & " to " & BlankList.Name & BlankList.Index
#End If

'Clear the "BlankList" and then fill it with FullList's items
BlankList.Clear
BlankList.Tag = "noclicks"
BlankVisState = BlankList.Visible
BlankList.Visible = False
FullList.Refresh
'lstFiles(fncExtList).Clear 'clear the extention list here because the Pull and Put Extention functions are horribly written right now.
For c = 0 To FullList.ListCount - 1
  'If RemoveExtention Then
  '  BlankList.AddItem PullExtention(FullList.List(c), c)
  'Else
    BlankList.AddItem FullList.List(c), c
  'End If
  BlankList.ItemData(BlankList.NewIndex) = c
  If FullList.Style = 0 Then
    BlankList.Selected(BlankList.NewIndex) = True
  Else
    BlankList.Selected(BlankList.NewIndex) = FullList.Selected(BlankList.NewIndex)
  End If
Next c
BlankList.Tag = ""
BlankList.Visible = BlankVisState

Exit Sub
errh:
If Err = 438 Then Resume Next 'File1 doesn't have a .style property, so default to checking the boxes
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "CopyList"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Function FollowsRules(ByVal FileName As String, ByVal RULES As String) As Boolean
'Check a file's name to make sure it follows all the rules for inclusion, and doesn't match any for exclusion
On Error GoTo errh

Dim c As Integer

FollowsRules = True

If Not RULES = "" Then
  'Only check if rules are sent
  ParseRules False, lstExecute, RULES
  For c = 0 To lstExecute.ListCount - 1
    If Len(lstExecute.List(c)) > 1 Then 'the length must be at least 2 characters long or its not a real rule
      If Left(lstExecute.List(c), 1) = "+" Then
        'Include this
        If InStr(1, FileName, Mid(lstExecute.List(c), 2)) = 0 Then
          FollowsRules = False 'fails test, get out
          Exit Function
        End If
      ElseIf Left(lstExecute.List(c), 1) = "-" Then
        'Exclude this
        If InStr(1, FileName, Mid(lstExecute.List(c), 2)) > 0 Then
          FollowsRules = False 'fails test, get out
          Exit Function
        End If
      Else
        'nothing should get here, if it does that means a rule is written wrong.
        #If Debugging = 1 Then
          Stop
        #End If
      End If
    End If
  Next c
End If


  

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "FollowsRules"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Function

Private Function ParseRules(ByVal Un As Boolean, prList As ListBox, ByVal RULES As String) As String
'If Un is true, Unparse the listbox data and output a string.
'If Un is false, Parse the String, and output its parts as separate items in the list box
'prList is the ListBox to get/send data from/to.
'Rules is the string in the form of "/+IncludeRule/-ExcludeRule"
On Error GoTo errh

Dim c As Integer
Dim tmpPos As Integer


If Un Then
  'Unparse the listbox, return a string
  RULES = ""
  For c = 0 To prList.ListCount - 1
    RULES = RULES & "/" & prList.List(c)
  Next c
  ParseRules = RULES
Else
  'Parse the string, fill the listbox
  prList.Clear
  While Len(RULES) > 1 'if it's valid it will always be "/" and something else, so >1, not >0
    tmpPos = InStr(2, RULES, "/")
    If tmpPos > 0 Then
      prList.AddItem Mid(RULES, 2, tmpPos - 2)
      RULES = Mid(RULES, tmpPos)
    Else
      'last item, add everything after the slash
      prList.AddItem Mid(RULES, 2)
      RULES = ""
    End If
  Wend
End If

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ParseRules"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function


Private Sub cmdHelp_Click()
On Error GoTo errh

If InHelpMode Then
  'flat and enabled; change that
  cmdHelp.BackColor = &H8000000F
  InHelpMode = False
Else
  '3D and disabled; change that
  cmdHelp.BackColor = vbYellow
  InHelpMode = True
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdHelp_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub Help(HelpID As Integer, Optional Context As String = "", Optional Data As String = "")
On Error GoTo errh

'Provides help for the selected item or function, hopefully
MsgBox "Sorry, no help yet. Hopefully soon." & vbCrLf & "ID: " & HelpID & vbCrLf & "Context: " & Context & vbCrLf & "Data: " & Data


'Get out of help mode
cmdHelp_Click

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Help"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub cmdOptions_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdOptions(Index).HelpContextID
  Exit Sub
End If

Dim c As Integer
Dim ListNum As Integer

If Index = 0 Then
  'The expand/contract button
  If frameOptions.Width = FRAMEOPTIONSSMALL Then
    'enlarge the frame
    ExpandFrame frameOptions
  Else
    'shrink the frame
    ExpandFrame frameOptions, False
  End If
Else
  If optOptionsSel(0).Value Then
    'If "Global" is selected, change the main list
    ListNum = fncOrgNames
  Else
    'If "Active Preview" is selected, change that one
    If lstQ.ListCount = 0 Then
      'if nothing is in the Queue, you can't change the Active Preview checks, so change the option back to global and do that
      'make the optionframe large so the change can be seen
      ExpandFrame frameOptions
      optOptionsSel(0).Value = True
      ListNum = fncOrgNames
    Else
      'Use the Active Preview list
      ListNum = fncActPrev
    End If
  End If
  
  With lstFiles(ListNum)
    .Tag = "noclicks" 'don't let the list's subs run
    If Index = 1 Then
      'The Toggle button
      For c = 0 To clstActPrev.ListCount - 1
        If lstQ.ListCount > 0 Or ListNum = fncOrgNames Then
          'if there are items in the Q, do this
          If .Selected(c) Then
            .Selected(c) = False
            If ListNum = fncActPrev Then ActCont.Selected(UIDIndex, c) = False
          Else
            .Selected(c) = True
            If ListNum = fncActPrev Then ActCont.Selected(UIDIndex, c) = True
          End If
        End If
      Next c
    ElseIf Index = 2 Then
      'The All button
      For c = 0 To clstActPrev.ListCount - 1
        If lstQ.ListCount > 0 Or ListNum = fncOrgNames Then
          'if there are items in the Q, do this
          .Selected(c) = True
          If ListNum = fncActPrev Then ActCont.Selected(UIDIndex, c) = True
        End If
      Next c
    ElseIf Index = 3 Then
      'The None button
      For c = 0 To clstActPrev.ListCount - 1
        If lstQ.ListCount > 0 Or ListNum = fncOrgNames Then
          'if there are items in the Q, do this
          .Selected(c) = False
          If ListNum = fncActPrev Then ActCont.Selected(UIDIndex, c) = False
        End If
      Next c
    End If
    .Tag = "" 'let the list's subs run
  End With
  If Index = 4 Then
  'The Global button (don't put this inside the With block)
    If lstQ.ListCount > 0 Then 'there still needs to be stuff in the Queue
      clstActPrev.Tag = "noclicks"
      For c = 0 To clstActPrev.ListCount - 1
        'if there are items in the Q, do this
        clstActPrev.Selected(c) = clstOrgNames.Selected(c)
        ActCont.Selected(UIDIndex, c) = clstOrgNames.Selected(c)
      Next c
      clstActPrev.Tag = ""
    End If
  End If
  'Update the filtered list of files up to this point.
  UpdateFilter lstQ, clstOrgNames, clstActPrev, lstQ.ListIndex

End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdOptions_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

'Move the button that is being hovered to the top
If Index > 0 Then
  cmdOptions(Index).ZOrder 0
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdOptions_MouseMove"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdPath_Click()
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdPath.HelpContextID
  Exit Sub
End If

If framePath.Height = FRAMEPATHSMALL Then
  'enlarge the frame
  ExpandFrame framePath
Else
  'shrink the frame
  ExpandFrame framePath, False
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdPath_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub ExpandFrame(Frame As Object, Optional Expand As Boolean = True)
On Error GoTo errh

Dim LBlue
Dim LGrey
Dim c As Integer

LBlue = &HFFFFC0    'light blue
LGrey = &H8000000F  'light grey (button face)

'make sure the frame is in front
Frame.ZOrder 0

If Frame.Name = "framePath" Then
  If Expand Then
    'make the path frame big and blue
    Frame.BackColor = LBlue
    chkHidSys(0).BackColor = LBlue
    chkHidSys(1).BackColor = LBlue
    Frame.Height = FRAMEPATHLARGE
    cmdPath.Picture = imgPathUp.Picture
    cmdPath.ToolTipText = "Collapse File Window"
  Else
    'shrink the frame and make it grey
    Frame.BackColor = LGrey
    chkHidSys(0).BackColor = LGrey
    chkHidSys(1).BackColor = LGrey
    Frame.Height = FRAMEPATHSMALL
    cmdPath.Picture = imgPathDown.Picture
    cmdPath.ToolTipText = "Expand File Window"
  End If
ElseIf Frame.Name = "frameOptions" Then
  If Expand Then
    'make the options frame big and blue
    Frame.BackColor = LBlue
    Frame.Width = FRAMEOPTIONSLARGE
    cmdOptions(0).Picture = imgArrowLeft.Picture
    cmdOptions(0).ToolTipText = "Collapse Options Window"
  Else
    'shrink the frame and make it grey
    Frame.BackColor = LGrey
    Frame.Width = FRAMEOPTIONSSMALL
    cmdOptions(0).Picture = imgArrowRight.Picture
    cmdOptions(0).ToolTipText = "Expand Options Window"
  End If
  'make sure the buttons are always on top
  For c = 1 To cmdOptions.Count - 1
    cmdOptions(c).ZOrder 0
  Next c
  Form_Resize
End If


Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "ExpandPathFrame"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdPathPaste_Click()
On Error GoTo errh

'check if the clipboard contains a path to a directory or file
If Dir(Clipboard.GetText, vbDirectory + vbSystem + vbHidden + vbNormal) = "" Then
  'if it doesn't, don't change anything
Else
  'if it does, put it in the path box
  txtPath.Text = Clipboard.GetText
  Drive1.Drive = txtPath.Text
  Dir1.Path = txtPath.Text
End If

Exit Sub
errh:
If Err = 52 Then
  'Bad filename or number from the Dir line
  Resume Next
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "InitPath"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdPathPaste_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

'check if the clipboard contains a path to a directory or file
If Dir(Clipboard.GetText, vbDirectory + vbSystem + vbHidden + vbNormal) = "" Then
  'if it doesn't contain valid data, change the tool tip to reflect that
  If Not cmdPathPaste.ToolTipText = "Clipboard contains the following INVALID Path or Filename: " & Chr(34) & Clipboard.GetText & Chr(34) Then
    cmdPathPaste.ToolTipText = "Clipboard contains the following INVALID Path or Filename: " & Chr(34) & Clipboard.GetText & Chr(34)
  End If
Else
  'if it does contain valid data, change the tool tip to reflect that
  If Not cmdPathPaste.ToolTipText = "Clipboard contains the following VALID Path or Filename: " & Chr(34) & Clipboard.GetText & Chr(34) Then
    cmdPathPaste.ToolTipText = "Clipboard contains the following VALID Path or Filename: " & Chr(34) & Clipboard.GetText & Chr(34)
  End If
End If

Exit Sub
errh:
If Err = 52 Or 53 Then
  'bad filename or number
  Resume Next
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdPathPaste_MouseMove"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub cmdQ_Click(Index As Integer)
' Holds the list of actions to take
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdQ(Index).HelpContextID
  Exit Sub
End If

'Holds the currently selected Action
Dim ListIndex As Integer
ListIndex = lstQ.ListIndex

'Holds the index of the item being sent to the recycle bin
Dim tmpLI As Integer

Select Case Index
  Case 0
    'Move Action item up on the list
    If ListIndex > 0 Then 'can't move up past the top
      lstQ.AddItem lstQ.List(ListIndex), ListIndex - 1
      lstQ.ItemData(lstQ.NewIndex) = lstQ.ItemData(ListIndex + 1) 'bring the item data along too
      lstQ.RemoveItem ListIndex + 1
      ListIndex = ListIndex - 1 'the item is now one slot lower
      lstQ.ListIndex = ListIndex 'reselect the item that moved
    End If
  Case 1
    'Move Action Item down on the list
    If ListIndex < lstQ.ListCount - 1 Then 'can't move down past the bottom
      lstQ.AddItem lstQ.List(ListIndex), ListIndex + 2
      lstQ.ItemData(lstQ.NewIndex) = lstQ.ItemData(ListIndex) 'bring the item data along too
      lstQ.RemoveItem ListIndex
      ListIndex = ListIndex + 1 'the item is now one slot lower
      lstQ.ListIndex = ListIndex 'reselect the item that moved
    End If
  Case 2
    'Copy the current item, place it directly below
    If ListIndex > -1 Then
      lstQ.AddItem lstQ.List(ListIndex), lstQ.ListIndex + 1
      'pull the UID of the item being copied, to get it's action type, and also send the UID itself
      lstQ.ItemData(lstQ.NewIndex) = AssignID(ActCont.GetActType(lstQ.ItemData(lstQ.NewIndex - 1)), clstOrgNames, lstQ.ItemData(lstQ.NewIndex - 1))
    End If
  Case 3
    'Add to the Recycle Bin
    If ListIndex > -1 Then
      lstRecycle.AddItem lstQ.List(ListIndex)
      lstRecycle.ItemData(lstRecycle.NewIndex) = lstQ.ItemData(ListIndex)
      lstQ.RemoveItem ListIndex
      If ListIndex < lstQ.ListCount Then
        lstQ.ListIndex = ListIndex
      Else
        lstQ.ListIndex = lstQ.ListCount - 1
      End If
      If lstQ.ListCount = 0 Then
        'if there are no items in the list, show the blank frame
        With frameProps(0)
          .ZOrder 0 'raise the blank one to the top
          .Top = 240  'move it to the right spot
          .Left = 120 'move it to the right spot
          .Enabled = True 'enable it
        End With
      End If
    Else
      'if nothing is selected, show the blank properties page
      With frameProps(0)
        .ZOrder 0 'raise the blank one to the top
        .Top = 240  'move it to the right spot
        .Left = 120 'move it to the right spot
        .Enabled = True 'enable it
      End With
    End If
  Case 4
    'Undo Delete
    lstRecycle.ListIndex = lstRecycle.ListCount - 1
    lstRecycle_DblClick
End Select

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdQ_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub







Private Sub cmdReplace_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdReplace(Index).HelpContextID
  Exit Sub
End If

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdReplace_Click"
Resume

End Sub

Private Sub cmdRules_Click()
On Error GoTo errh

Dim ie As String

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help cmdRules.HelpContextID
  Exit Sub
End If

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

If optRules(0).Value Then
  'include this rule. "/" separates each item, the + or - is if its included or not
  ie = "/+"
Else
  ie = "/-"
End If

ActCont.SetData UIDIndex, 0, ActCont.GetData(UIDIndex, 0) & ie & txtRules.Text
Update_lstRules ActCont.GetData(UIDIndex, 0)

Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdRules_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub









Private Sub cmdUpdate_Click()
On Error GoTo errh

'Update from the filesystem list

CascadePath "txtPath", txtPath.Text


'Dim c As Integer

'Update the filtered list of files up to this point.
'If chkActivePreview.Value = 1 Then
'  'Only update if the checkbox is enabled
'  UpdateFilter lstQ, clstOrgNames, clstActPrev
'  'select the same file that is selected in the main list box
'  For c = 0 To clstActPrev.ListCount - 1
'    If clstActPrev.ListIndex = clstActPrev.ItemData(c) Then
'      clstActPrev.ListIndex = c
'      Exit For
'    End If
'  Next c
'End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "cmdUpdate_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Command1_Click()
On Error GoTo errh

CopyList clstActPrev, lstSorted




Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Command1_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub





Private Sub Dir1_Change()
On Error GoTo errh

Dim c As Integer


If Not CascadingPath Then CascadePath "Dir1", Dir1.Path

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Dir1_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub Dir1_Click()
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help Dir1.HelpContextID
  Exit Sub
End If

If Not txtPath.Text = Dir1.Path Then Dir1_Change
  
Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Dir1_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub Dir1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

OLEDrop Data, Effect, Button, Shift, X, Y

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Dir1_OLEDragDrop"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub Drive1_Change()
On Error GoTo errh

If Not CascadingPath Then Dir1.Path = Drive1.Drive 'send the drive to Dir1 so it can find the new path
  
Exit Sub
errh:
If Err = 68 Then
  'drive not available
  Exit Sub
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Drive1_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub Drive1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

OLEDrop Data, Effect, Button, Shift, X, Y

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Drive1_OLEDragDrop"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub File1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

OLEDrop Data, Effect, Button, Shift, X, Y

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "File1_OLEDragDrop"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub Form_Load()
On Error GoTo errh

'Counts the times the module is called
#If Debugging = 1 Then
  frmDebug.Show
  Me.Caption = "Filename Changer (Debug Mode)"
#End If

'Set the references for the list box items
InitListBoxRefs

Me.Show

'Create the group of Objects that accept Paths
InitCollections

'I work with the Properties frame extended, and most of the windows small. make them the normal sizes
InitFrameSizes

'Initalize the Path
InitPath lstQ, clstOrgNames, clstActPrev, File1


'Add Items to the Action List
InitActions

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Form_Load"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub InitListBoxRefs()
On Error GoTo errh

Set clstOrgNames = lstFiles(fncOrgNames) 'The list of original file names
Set clstActPrev = lstFiles(fncActPrev)   'The list in the active preview window
Set clstSample = lstFiles(fncSampleList) 'The sample list box
Set clstSorted = lstSorted(0)            'The list that holds the sorted data for duplicate checking

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "InitListBoxRefs"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub InitCollections()
'Creates the group of Objects that accept Paths so I can cycle through them easily
On Error GoTo errh

'add the appropriate items to the Path frame collection
colPathObjs.Add txtPath
colPathObjs.Add Drive1
colPathObjs.Add Dir1
colPathObjs.Add File1

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "InitPathCollection"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub InitFrameSizes()
'I work with the Properties frame extended, and most of the windows small. make them the normal sizes
On Error GoTo errh

Dim c As Integer
Dim FRAMEPROPSWIDTH As Integer
Dim FRAMEPROPSHEIGHT As Integer
'Set the sizes (in twips)
FRAMEPROPSWIDTH = 3375
FRAMEPROPSHEIGHT = 3375

'set the siz of the main frame that holds the others (in twips)
frameProperties.Width = 3615
frameProperties.Height = 4095

'width = 3375
'height = 3375
For c = 0 To frameProps.Count - 1
  frameProps(c).Width = FRAMEPROPSWIDTH
  frameProps(c).Height = FRAMEPROPSHEIGHT
  frameProps(c).BorderStyle = 0 ' get rid of their borders, not need when running
Next c

'set the options frame large by default
ExpandFrame frameOptions

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "InitPropertySizes"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub InitActions()
'Adds the items to the Actions list
On Error GoTo errh

Dim c As Integer
Dim tmpStr As String

With lstActions
  .Clear 'clear the action list before re-adding everything
  For c = 1 To ActionDetails(, , True) 'skip the first one, which is the blank one
    tmpStr = ActionDetails(CLng(c))
    If Left(tmpStr, 1) <> "*" Then
      'Don't add if the ActionDetail starts with a *, that means there is nothing there yet
      .AddItem ActionDetails(CLng(c))
      .ItemData(.NewIndex) = CLng(c)
    End If
  Next c
End With
  

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "InitActions"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub



Private Sub Form_Resize()
On Error GoTo errh

'constants for the sizes of the frames and objects (these are all in twips)
Dim PROPLEFT As Integer
PROPLEFT = 9360 'farthest left the properties window should go
Dim QWIDTH As Integer
QWIDTH = 840 'the width from the right edge of lstQ to the right edge of frameQ
Dim CMDQWIDTH As Integer
CMDQWIDTH = 225 'distance from the left of the buttons in the Queue frame to the right edge of frameQ
Dim RECYCLEBUFFER As Integer
RECYCLEBUFFER = 300 'distance from the bottom of lstrecycle to the bottom of the frameRecycle
Dim FILESHEIGHTBUFFER As Integer
FILESHEIGHTBUFFER = 100 'distance from the bottom of clstActPrev to the bottom of the framePreview
Dim FILESWIDTHBUFFER As Integer
FILESWIDTHBUFFER = 225 'distance from the left of clstActPrev in the Preview frame to the right edge of framePreview

Dim c As Integer

'Minimum form width is 13125
'Minimum form height is 8640
#If Debugging = 0 Then 'only limit the window size if its compiled
  If Me.Width < 13125 Then
    Me.Width = 13125
  End If
  If Me.Height < 8640 Then
    Me.Height = 8640
  End If
#End If

'set size and position of the recycle bin
frameRecycle.Top = frameProperties.Height + frameProperties.Top
frameRecycle.Left = Me.ScaleWidth - frameRecycle.Width
frameRecycle.Height = Me.ScaleHeight - frameRecycle.Top
lstRecycle.Height = frameRecycle.Height - RECYCLEBUFFER

'set size and position of the files frame
framePreview.Left = frameOptions.Width
framePreview.Width = Me.ScaleWidth - frameRecycle.Width - frameOptions.Width
framePreview.Height = Me.ScaleHeight - framePreview.Top
clstActPrev.Width = framePreview.Width - FILESWIDTHBUFFER
clstActPrev.Height = framePreview.Height - FILESHEIGHTBUFFER - clstActPrev.Top

'set the size of the options frame
frameOptions.Height = Me.ScaleHeight - frameOptions.Top

'set size and position of the action properties frame
If frameProperties.Left >= PROPLEFT Then
  frameProperties.Left = Me.ScaleWidth - frameProperties.Width
Else
  frameProperties.Left = PROPLEFT
End If

'set size and position of the Queue frame
frameQ.Width = Me.ScaleWidth - frameProperties.Width - frameQ.Left
lstQ.Width = frameQ.Width - QWIDTH
For c = 0 To cmdQ.Count - 1
  cmdQ(c).Left = lstQ.Width + CMDQWIDTH
Next c

Exit Sub
errh:
If Err = 380 Then Exit Sub 'invalid sizing
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Form_Resize"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errh

End

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Form_Unload"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub framePreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

'Move framePreview above the Option buttons
framePreview.ZOrder 0

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "framePreview_MouseMove"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub frameProps_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help frameProps(Index).HelpContextID
  Exit Sub
End If



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "frameProps_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub





Private Sub Label1_Click()

End Sub

Private Sub lblSwitch_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help lblSwitch(Index).WhatsThisHelpID
  Exit Sub
End If



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lblSwitch_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub lstActions_Click()
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help lstActions.HelpContextID, lstActions.List(lstActions.ListIndex)
  Exit Sub
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstActions_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub


Private Sub lstActions_DblClick()
On Error GoTo errh

'do the default action, which right now is Add After Selected (2)
cmdAddAction_Click 2

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstActions_DblClick"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub



Private Sub lstFiles_Click(Index As Integer)
On Error GoTo errh

Dim c As Integer

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help lstFiles(Index).HelpContextID
  Exit Sub
End If

'get out if the checks are being changed
If lstFiles(Index).Tag <> "" Then Exit Sub

#If Debugging = 1 Then
  frmDebug.dModuleCount "lstFiles_Click(" & CStr(Index) & ")"
#End If

'If an item in the active preview list is selected,
If Index = fncActPrev Then
  If lstQ.ListIndex > 0 Then
    'Fill the SampleList with changed names up to the selected index, less one
    MakeChanges lstQ, clstOrgNames, clstSample, lstQ.ListIndex - 1
  Else
    'if the first item is selected, just copy the OrgNames box
    CopyList clstOrgNames, clstSample
  End If
    
  'go through all the sample text boxes and put the selected file's name in there
  For c = 0 To txtSample.Count - 1
    txtSample(c).Text = clstSample.List(clstActPrev.ListIndex)
  Next c
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstFiles_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub


Private Sub lstFiles_ItemCheck(Index As Integer, Item As Integer)
On Error GoTo errh

Dim tmpIndex As Integer

'Get out immediate if its not supposed to be clicked. (the control will be disabled)
If lstFiles(Index).Tag <> "" Then Exit Sub

#If Debugging = 1 Then
  frmDebug.dModuleCount "lstFiles_ItemCheck(" & CStr(Index) & ")"
#End If

If lstQ.ListCount > 0 And Index = fncActPrev Then
  'This is the Active Preview window, so if something is checked/unchecked, record it in the Action's property (if Q items exist)
  ActCont.Selected(UIDIndex, lstFiles(Index).ItemData(Item)) = lstFiles(Index).Selected(Item)
ElseIf lstQ.ListCount = 0 Then
  If Index = fncActPrev Then
    'If there are no Q items, treat it as clicking the main list box
    clstOrgNames.Tag = "noclicks"
    clstOrgNames.Selected(Item) = clstActPrev.Selected(Item)
    clstOrgNames.Tag = ""
  ElseIf Index = fncOrgNames Then
    'If there are no Q items, treat it as clicking the main list box
    clstActPrev.Tag = "noclicks"
    clstActPrev.Selected(Item) = clstOrgNames.Selected(Item)
    clstActPrev.Tag = ""
  End If
End If


Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstFiles_ItemCheck"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub lstFiles_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

'Move framePreview above the Option buttons
framePreview.ZOrder 0

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstFiles_MouseMove"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub lstFiles_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

OLEDrop Data, Effect, Button, Shift, X, Y

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstFiles_OLEDragDrop"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub lstQ_Click()
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help lstQ.HelpContextID
  Exit Sub
End If

Dim c As Integer

'Update the properties page with the currently selected item's data
UpdateProperties lstQ.ItemData(lstQ.ListIndex)

'Update the filtered list of files up to this point.
UpdateFilter lstQ, clstOrgNames, clstActPrev, lstQ.ListIndex




Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstQ_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub UpdateFilter(Q As ListBox, OrgFile As ListBox, ActPrev As ListBox, Optional ChangeTo As Integer = -2)
'Updates the Active Preview list if ActivePreview is enabled. -2 means nothing was sent to ChangeTo, anything else means something was specified, send that
On Error GoTo errh

If ChangeTo = -2 Then
  'no ChangeTo was sent, default to the Q listindex
  ChangeTo = Q.ListIndex
End If

'If ActivePreview is enabled, update. otherwise just update the check boxes
MakeChanges lstQ, OrgFile, ActPrev, ChangeTo, chkActivePreview.Value



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "UpdateFilter"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub


Private Function AssignID(ActType As Long, OrgFile As ListBox, Optional CopyUID As Integer = -1) As Integer
'if CopyUID is supplied, copy all the data from it, otherwise make a blank one
On Error GoTo errh

Dim r As Integer
Dim c As Long 'this way it can be sent w/o casting

'warn the user if they are getting close to the limit on IDs (about 32000 of them, range from 1 to 32000)
If ActCont.ActionCount > 31750 Then
  MsgBox "Warning, you have created over 31,750 Actions. You will only be able to create about 250 more before the program crashes. Why the hell have you kept the program open this long?!??! Restart the damn thing already!", vbExclamation, "Almost out of Unique Action IDs"
ElseIf ActCont.ActionCount > 31500 Then
  MsgBox "Warning, you have created over 31,500 Actions. You will only be able to create about 500 more before the program crashes. Restart the program to reset the count.", vbExclamation, "Almost out of Unique Action IDs"
ElseIf ActCont.ActionCount > 31000 Then
  MsgBox "Warning, you have created over 31,000 Actions. You will only be able to create about 1000 more before the program crashes. Restart the program to reset the count.", vbExclamation, "Almost out of Unique Action IDs"
End If

'assign a unique ID to the action item that it put in the queue
AssignID = ActCont.AddAction(ActType)

If CopyUID = -1 Then 'Create a blank one
  'Set the default checkmarks to OrgFile's current checks
  For c = 0 To OrgFile.ListCount - 1
    If OrgFile.Selected(c) Then
      ActCont.Selected(AssignID, c) = True
    Else
      ActCont.Selected(AssignID, c) = False
    End If
  Next c
Else
  'if a CopyUID was supplied, get the data from that UID and put it in the new one
  For c = 0 To 3
    ActCont.SetData AssignID, c, ActCont.GetData(CopyUID, c)
  Next c
End If

Exit Function
errh:
MsgBox Err & Err.Description
If Err = 6 Then
  'Overflow error, IDCount is too big
  r = MsgBox("I told you you were going to run out of Unique IDs. Nice going." & vbCrLf & vbCrLf & "Click 'Retry' to restart the count, or 'Cancel' to close the program." & vbCrLf & vbCrLf & "If you choose Retry, you may experience unexpected behavior.", vbCritical + vbRetryCancel, "Unique ID Overflow")
  If r = vbRetry Then
    IDCount = -32001
    Resume Next
  Else
    #If Debugging = 1 Then
      Stop
    #Else
      End
    #End If
  End If
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "AssignID"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Private Sub lstRecycle_Click()
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help lstRecycle.HelpContextID
  Exit Sub
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstRecycle_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub lstRecycle_DblClick()
On Error GoTo errh

'add to the end
lstQ.AddItem lstRecycle.List(lstRecycle.ListIndex), lstQ.ListCount
lstQ.ItemData(lstQ.NewIndex) = lstRecycle.ItemData(lstRecycle.ListIndex)
lstRecycle.RemoveItem lstRecycle.ListIndex

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstRecycle_DblClick"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub InitPath(Q As ListBox, OrgFile As ListBox, ActPrev As ListBox, SystemFileList As Object)
On Error GoTo errh

Dim Path As String
Dim ClipBoardText As String

Path = ""

'Get the data from the clipboard
If Clipboard.GetFormat(vbCFText) Then
  ClipBoardText = Clipboard.GetText
  Path = Dir(ClipBoardText, vbDirectory + vbSystem + vbHidden + vbNormal)
  'Handle unusual paths
  If Path = "" Then
    Path = CurDir()
  ElseIf Path = "." Then
    'The item is located in the current directory
    Path = CurDir() & "\" & ClipBoardText
  ElseIf Path = Right(ClipBoardText, Len(Path)) Then
    'If the name of the directory is returned, its valid
    Path = ClipBoardText
  Else
    'Raise an error for unhandled paths
    If Len(Path) < 3 Then Err.Raise -334
    If InStr(1, Path, "\") = 0 Then Err.Raise -334
  End If
Else
  Path = CurDir()
End If

'put it in the appropriate boxes, and make the frame big
CascadePath "", Path
ExpandFrame framePath

'OBSOLETE - CascadePath handles these
'CopyList SystemFileList, OrgFile 'Add the file names from the current path to the main list
'UpdateFilter Q, OrgFile, ActPrev 'Filter as appropriate

Exit Sub
errh:

If Err = 68 Then
  'Invalid property (drive1.drive is set to something invalid)
  txtPath.Text = "Invalid Drive or Path selected."
  Exit Sub
End If
If Err = 52 Or Err = 53 Then
  'Bad filename or number OR file doesn't exist from the Dir line
  Resume Next
End If
If Err = -334 Then
  'My error, means the path is invalid, and I don't know how to handle it. set the Error Description and let the error handler run
  Err.Description = "Unhandled, Invalid Path: " & Chr(34) & Path & Chr(34) & ". Try clearing the clipboard of text and reload the program."
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "InitPath"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub







Private Sub lstSubDir_Click()
On Error GoTo errh

'If they choose a drive, put it in the textbox
If lstSubDir.ListIndex > 0 Then
  txtSubDir.Text = lstSubDir.List(lstSubDir.ListIndex)
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstSubDir_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub lstSubDir_DblClick()
On Error GoTo errh

Dim NewName As String
NewName = ""

'If they double click on the first item (create new subdir) ask for the name
Do
  NewName = InputBox("Enter the name of the new directory to create", "Directory Name", NewName)
  If InStr(1, NewName, "\") Then MsgBox "Don't put any slashes in the name.", vbCritical, "Invalid Name"
Loop While InStr(1, NewName, "\") > 0 Or Err = 75

If NewName <> "" Then
  MkDir Dir1.Path & "\" & NewName
End If
CascadePath "", Dir1.Path

Exit Sub
errh:
If Err = 75 Then
  'Directory already exists
  MsgBox "Directory already exists.", vbCritical, "Duplicate Name"
  Exit Sub
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "lstSubDir_DblClick"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuAbout_Click()
On Error GoTo errh

frmAbout.Show

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuAbout_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuDebug_Click()
On Error GoTo errh

frmDebug.Show

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuDebug_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuExit_Click()
On Error GoTo errh

End

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuExit_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuHelp_Click()
On Error GoTo errh

MsgBox "No help yet, sorry. If you're getting errors, email me at andrew.kolberg@gmail.com. Please include the error message you recieved, and any information you think might be helpful.", , "Helpless"

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuHelp_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuLoad_Click()
On Error GoTo errh

Settings fncLoad, fncSLQueue

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuLoad_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuNew_Click()
On Error GoTo errh

Dim r As Integer
r = MsgBox("Are you sure you want to clear the Queue and all action items you've created?", vbYesNo + vbQuestion, "Clear Data")
If r = vbNo Then Exit Sub

'Clear the UIDs and the Queue list
ActCont.Clear
lstQ.Clear

frameProps(fncBlank).ZOrder 0

UpdateFilter lstQ, clstOrgNames, clstActPrev

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuNew_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuSave_Click()
On Error GoTo errh

Settings fncSave, fncSLQueue

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuSave_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub mnuShowStatus_Click()
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help mnuShowStatus.HelpContextID
  Exit Sub
End If

frmStatus.SetButtons fncStatusMode
frmStatus.Caption = "Status: Filename Changer"
frmStatus.Show


Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "mnuShowStatus_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub optCap_Click(Index As Integer)
On Error GoTo errh

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help optCap(Index).HelpContextID
End If

ActCont.SetData UIDIndex, 0, Str(Index)

Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "optCap_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub optConcat1_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help optConcat1(Index).HelpContextID
End If

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

ActCont.SetData UIDIndex, fncConcatLeftRight, Index

Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "optConcat1_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub



Private Sub optDelBet_Click(Index As Integer)
On Error GoTo errh

Dim LeftDir As Long
Dim DelFirst As Integer

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

'stop the update timer
tmrUpdate.Enabled = False

'Search Mode
If Index = 0 Or Index = 1 Then
  'Begin from Left Or Right
  '(Index=0) Begin the search for the first term at the left of the string and move right. Then begin the search for the second term at the left of the string and move right
  'OR
  '(Index=1) Begin the search for the first term at the left of the string and move right. Then begin the search for the second term at the right of the string and move left
  If Index = 0 Then LeftDir = fncDBLeft Else LeftDir = fncDBRight
  ActCont.SetData UIDIndex, fncDelBetLeftDir, LeftDir
ElseIf Index = 2 Or Index = 3 Then
  'Delete All (3) or First Found (2)
  If optDelBet(3).Value Then DelFirst = 0 Else DelFirst = 1
  ActCont.SetData UIDIndex, fncDelBetDelFirst, DelFirst
End If

'start the update timer
tmrUpdate.Enabled = True



''Search mode
'If Index = 0 Then
'  'Right, Right
'  'Begin the search for the first term at the left of the string and move right. Then begin the search for the second term at the left of the string and move right
'  LeftDir = fncDBRight  'means move in the -> (right) direction
'  RightDir = fncDBRight 'means move in the -> (right) direction
'ElseIf Index = 1 Then
'  'Right, Left
'  'Begin the search for the first term at the left of the string and move right. Then begin the search for the second term at the right of the string and move left
'  LeftDir = fncDBRight  'means move in the -> (right) direction
'  RightDir = fncDBLeft  'means move in the <- (left) direction
'ElseIf Index = 2 Then
'  'Left, Left
'  'Begin the search for the first term at the right of the string and move left. Then begin the search for the second term at the right of the string and move left
'  LeftDir = fncDBLeft   'means move in the <- (left) direction
'  RightDir = fncDBLeft  'means move in the <- (left) direction
'ElseIf Index = 3 Then
'  'Left, Right
'  'Begin the search for the first term at the right of the string and move left. Then begin the search for the second term at the left of the string and move right
'  LeftDir = fncDBLeft   'means move in the <- (left) direction
'  RightDir = fncDBRight 'means move in the -> (right) direction
'End If
    



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "optDelBetL_Click"
Resume
End Sub

Private Sub optOptionsSel_Click(Index As Integer)
On Error GoTo errh

If Index = 0 Then
  'get rid of the "Global" button
  cmdOptions(4).Visible = False
Else
  'show the "Global" button
  cmdOptions(4).Visible = True
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "optOptionsSel_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub optRules_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help optRules(Index).HelpContextID
  Exit Sub
End If

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

txtRules.SetFocus

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "optRules_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub


Private Sub tmrUpdate_Timer()
On Error GoTo errh
'Periodically update the clstActPrev window so that it's contents are up to date

'Don't update while settings are being updated/saved
If InSettings Then
  tmrUpdate.Enabled = False
  Exit Sub
End If

UpdateFilter lstQ, clstOrgNames, clstActPrev, lstQ.ListIndex

tmrUpdate.Enabled = False

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "tmrUpdate_Timer"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtCap_Change()
On Error GoTo errh

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

ActCont.SetData UIDIndex, 1, txtCap.Text

Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtCap_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtConcat_Change(Index As Integer)
On Error GoTo errh

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

'If one text box changes, update the other one
If curConcatIndex = 2 Then
  txtConcat(1).Text = Val(txtConcat(2).Text) - 1
  ActCont.SetData UIDIndex, fncConcatPosition, txtConcat(2).Text
ElseIf curConcatIndex = 1 Then
  txtConcat(2).Text = Val(txtConcat(1).Text) + 1
  ActCont.SetData UIDIndex, fncConcatPosition, txtConcat(2).Text
ElseIf curConcatIndex = 0 Then
  ActCont.SetData UIDIndex, fncConcatText, txtConcat(0).Text
End If



Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtConcat_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub txtConcat_GotFocus(Index As Integer)
On Error GoTo errh

curConcatIndex = Index

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtConcat_GotFocus"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtDelete_Change(Index As Integer)
On Error GoTo errh

'stop the update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

ActCont.SetData UIDIndex, CLng(Index), txtDelete(Index).Text

Update_lstQ

'start the update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtDelete_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtFileModePath_Change()
On Error GoTo errh

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

ActCont.SetData UIDIndex, 0, txtFileModePath.Text

Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtFileModePath_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub txtPath_Change()
On Error GoTo errh

txtPath.ToolTipText = txtPath.Text

If Dir(txtPath.Text, vbDirectory) = "" Then
  'If its invalid, make the box red
  txtPath.BackColor = vbRed
Else
  'otherwise update the directory
  txtPath.BackColor = vbWhite
  If Not CascadingPath Then CascadePath "txtPath", txtPath.Text
End If

Exit Sub
errh:
If Err = 52 Or Err = 53 Then
  'bad filename or number
  Resume Next
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtPath_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub CascadePath(CallingObjectName As String, Path As String)
'This function will send the current path to each object that needs it, without sending it back to the object that just changed
On Error GoTo errh

'holds each object for the loop
Dim Obj As Variant

CascadingPath = True 'when true, the items in the collection won't try to update the path

'make sure its a valid path, if not, exit right away
If Dir(Path, vbDirectory + vbNormal + vbReadOnly + vbSystem + vbHidden + vbVolume) = "" Then
  CascadingPath = False
  Exit Sub
End If

'goes through each item and updates the path it's showing.
For Each Obj In colPathObjs
  If Obj.Name <> CallingObjectName Then
    Obj.Path = Path 'if the object doesn't support the method, it will Resume Next
    Obj.Drive = Path
    Obj.Text = Path
  End If
  Obj.BackColor = vbWhite
Next Obj
'Update the Sub Dir list
CreateSubDirList lstSubDir

'make sure its a valid path and then update the file lists
If Dir(Path, vbDirectory + vbNormal + vbReadOnly + vbSystem + vbHidden + vbVolume) = "" Then
  'Don't do anything
Else
 'Update the lists
  CopyList File1, clstOrgNames
  UpdateFilter lstQ, clstOrgNames, clstActPrev
  framePreview.Caption = "Active Preview - Displaying " & clstActPrev.ListCount & " files."
End If

CascadingPath = False 'when false, allows the items to update the path

Exit Sub
errh:
If Err = 52 Then
  'Bad file name or number (Dir recieved an invalid file, skip it)
  Resume Next
End If
If Err = 438 Then
  'Object doesn't support this property or method, so just go to the next one
  Resume Next
End If
If Err = 76 Then
  'Bad path, don't continue to cascade
  For Each Obj In colPathObjs
    Obj.BackColor = vbRed
  Next Obj
  CascadingPath = False 'when false, allows the items to update the path
  Exit Sub
End If
If Err = 68 Then
  'Device Unavailable, so it's a bad path
  'Bad path, don't continue to cascade
  For Each Obj In colPathObjs
    Obj.BackColor = vbRed
  Next Obj
  CascadingPath = False 'when false, allows the items to update the path
  Exit Sub
End If
  
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "CascadePath"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub CreateCollections()
On Error GoTo errh
'Create all the collects used throughout the program




Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "CreateCollections"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Function FilesListChanged(FileList As Object, OrgList As Object) As Boolean
On Error GoTo errh

Dim c As Integer

'if a change is found, exit the sub with FLC = true

FilesListChanged = True

'Make sure each list is current
FileList.Refresh
OrgList.Refresh

'check for length
If FileList.ListCount <> OrgList.ListCount Then Exit Function

'check each item
For c = 0 To OrgList.ListCount - 1
  If FileList.List(c) <> OrgList.List(c) Then Exit Function
Next c

'no changes
FilesListChanged = False

Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "FilesListChanged"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume


End Function

Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

OLEDrop Data, Effect, Button, Shift, X, Y

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtPath_OLEDragDrop"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub OLEDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

Dim tmpLast As Integer
Dim Path As String

If Data.GetFormat(vbCFFiles) Then
  'A path, file, or list of files was dropped
  Path = Data.Files(1)
  'Determine which it was
  If Dir(Path, vbNormal + vbHidden + vbSystem) <> "" Then
    'A file was dropped
    Path = Mid(Path, 1, FindLast(Path, "\"))
  End If
  CascadePath "", Path
End If

Exit Sub
errh:
If Err = 76 Then
  'Path not found
  'assume that the problem is that a file was dropped... choose the root folder
  If Len(Data.Files(1)) > 3 Then
    'if the data is longer than just the drive name ("c:\" for example)
    tmpLast = FindLast(Data.Files(1), "\")
    Dir1.Path = Mid(Data.Files(1), 1, tmpLast)
    Resume Next
  Else
    MsgBox "Invalid Drag and Drop path... Ending: " & Data.Files(1) & " Module: OLEDrop"
    Stop
  End If
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "OLEDrop"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtReplace_Change(Index As Integer)
On Error GoTo errh

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

ActCont.SetData UIDIndex, CLng(Index), txtReplace(Index).Text

Update_lstQ

'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtReplace_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub



Private Sub txtRules_Change()
On Error GoTo errh


'stop the active preview update timer
tmrUpdate.Enabled = False


Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtRules_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtRules_GotFocus()
On Error GoTo errh

'Allow "enter" to add the data in this text box
cmdRules.Default = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtRules_GotFocus"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume
End Sub

Private Sub UpdateProperties(UID As Integer)
'Display the correct properties page and update the information accordingly
On Error GoTo errh

Dim c As Integer
Dim tmpInt As Integer

'disable all other frames
For c = 0 To frameProps.Count - 1
  frameProps(c).Enabled = False
Next c

With frameProps(ActCont.GetActType(UID))
  .ZOrder 0 'raise the correct one to the top
  .Top = 600  'move it to the right spot
  .Left = 120 'move it to the right spot
  .Enabled = True 'enable it
  frameProperties.Caption = "Action Properties (" & ActionDetails(ActCont.GetActType(UID)) & ")"
  lstFiles_Click fncActPrev
End With

'Keep track of the UID who's properties are currently being shown
UIDIndex = UID

InSettings = True 'Don't let things update right now

Select Case ActCont.GetActType(UID)
  Case fncReplaceCharacters
    'Replace Characters
    txtReplace(0).Text = ActCont.GetData(UID, fncReplaceOld)
    txtReplace(1).Text = ActCont.GetData(UID, fncReplaceNew)
    Update_lstReplace
  Case fncDeleteBetween
    'Delete everything between
    txtDelete(0).Text = ActCont.GetData(UID, fncDelBetLeft)
    txtDelete(1).Text = ActCont.GetData(UID, fncDelBetRight)
    
    'Determine the correct option to set for Search Direction
    If Val(ActCont.GetData(UID, fncDelBetLeftDir)) = fncDBLeft Then
      'Search Direction = Left
      optDelBet(0).Value = True
    Else
      optDelBet(1).Value = True
    End If
    
    'Determine the correct option for Delete First
    If CBool(Val(ActCont.GetData(UID, fncDelBetDelFirst))) Then 'take the BOOL of the VAL because if the data is "" there is an error
      'Delete only the first term
      optDelBet(2).Value = True
    Else
      'Delete all terms found
      optDelBet(3).Value = True
    End If
    
    Update_lstDelete
  Case fncCapitalization
    'Capitalization
    tmpInt = Val(ActCont.GetData(UID, fncCapsOption))
    optCap(tmpInt).Value = True
    txtCap.Text = ActCont.GetData(UID, fncCapsPosition)
  Case fncSwitchCharacters
    'Switch Characters
    txtSwitch(0).Text = ActCont.GetData(UID, fncSwitch1Sel_Start)
    txtSwitch(1).Text = ActCont.GetData(UID, fncSwitch1SelLen)
    txtSwitch(2).Text = ActCont.GetData(UID, fncSwitch2Sel_Start)
    txtSwitch(3).Text = ActCont.GetData(UID, fncSwitch2SelLen)
  Case fncIncludeExcludeRules
    'Include / Exclude
    Update_lstRules ActCont.GetData(UID, fncRulesAddRule)
  Case fncConcatenate
    'Concatenation
    txtConcat(fncConcatText).Text = ActCont.GetData(UID, fncConcatText)
    txtConcat(fncConcatPosition + 1).Text = Val(ActCont.GetData(UID, fncConcatPosition)) 'the index of fncPosition refers to the data, the textbox that get it is one higher
    txtConcat(fncConcatPosition).Text = Val(txtConcat(fncConcatLeftRight).Text) - 1      'the dummy text box has the index of the data
    tmpInt = Val(ActCont.GetData(UID, fncConcatLeftRight))
    optConcat1(tmpInt).Value = True
  Case fncFilemode
    'Filemode
    txtFileModePath.Text = ActCont.GetData(UID, fncFileModePath)
    FileMode lstFileMode, , ActCont.GetData(UID, fncFileModePath)
  Case Else
End Select

InSettings = False 'let things update again

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "UpdateProperties"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Update_lstReplace()
On Error GoTo errh



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Update_lstReplace"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Update_lstDelete()
On Error GoTo errh



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Update_lstDelete"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Update_lstRules(ListItems As String)
'Parse the list of includes and excludes
On Error GoTo errh

ParseRules False, lstRules, ListItems

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Update_lstRules"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub Update_lstQ()
On Error GoTo errh

Dim c As Integer

Dim curUID As Integer
Dim strTemp As String
Dim strTemp1 As String
Dim MAXLEN As Integer  'maximum length for the soft data so the Titles don't get too long
MAXLEN = 8

'Make sure each item in the Queue has the correct custom title
For c = 0 To lstQ.ListCount - 1
  curUID = lstQ.ItemData(c)
  If ActCont.GetActType(curUID) = fncReplaceCharacters Then
    'Modify the title for the replace characters type
    strTemp = ActCont.GetData(curUID, 0)
    If Len(strTemp) > MAXLEN Then strTemp = Left(strTemp, MAXLEN - 3) & "..." 'truncate the length if needed
    strTemp1 = ActCont.GetData(curUID, 1)
    If Len(strTemp1) > MAXLEN Then strTemp1 = Left(strTemp1, MAXLEN - 3) & "..." 'truncate the length if needed
    ActCont.SetTitle curUID, "Replace " & Chr(34) & strTemp & Chr(34) & " with " & Chr(34) & strTemp1 & Chr(34)
  ElseIf ActCont.GetActType(curUID) = fncDeleteBetween Then
    'Modify the title for the delete between type
    strTemp = ActCont.GetData(curUID, 0)
    If Len(strTemp) > MAXLEN Then strTemp = Left(strTemp, MAXLEN - 3) & "..." 'truncate the length if needed
    strTemp1 = ActCont.GetData(curUID, 1)
    If Len(strTemp1) > MAXLEN Then strTemp1 = Left(strTemp1, MAXLEN - 3) & "..." 'truncate the length if needed
    ActCont.SetTitle curUID, "Delete everything between " & Chr(34) & strTemp & Chr(34) & " and " & Chr(34) & strTemp1 & Chr(34)
  ElseIf ActCont.GetActType(curUID) = fncCapitalization Then
    'Modify the title for the capitalization type
    If Val(ActCont.GetData(curUID, 0)) = 0 Then
      ActCont.SetTitle curUID, "Capitalize the character " & ActCont.GetData(curUID, 1)
    ElseIf Val(ActCont.GetData(curUID, 0)) = 1 Then
      ActCont.SetTitle curUID, "Toggle Capitalization"
    End If
  ElseIf ActCont.GetActType(curUID) = fncConcatenate Then
    'Modify the title for the concat type
    strTemp = ActCont.GetData(curUID, fncConcatText)
    If Len(strTemp) > MAXLEN Then strTemp = Left(strTemp, MAXLEN - 3) & "..." 'truncate the length if needed
    strTemp1 = "Add " & Chr(34) & strTemp & Chr(34) & " between characters "
    strTemp1 = strTemp1 & Val(ActCont.GetData(curUID, fncConcatPosition)) - 1 & " and " & Val(ActCont.GetData(curUID, fncConcatPosition))
    If Val(ActCont.GetData(curUID, fncConcatLeftRight)) = 0 Then
      'Start the count from the left
      strTemp1 = strTemp1 & ", from the Left"
    Else
      'Start the count from the right
      strTemp1 = strTemp1 & ", from the Right"
    End If
    ActCont.SetTitle curUID, strTemp1
  ElseIf ActCont.GetActType(curUID) = fncFilemode Then
    'Modify the title for the Filemode type
    If FindLast(ActCont.GetData(curUID, 0), "\") > 0 Then
      'if a path is set already, use that data
      strTemp = Mid(ActCont.GetData(curUID, 0), FindLast(ActCont.GetData(curUID, 0), "\"))
      If Len(strTemp) > MAXLEN Then strTemp = Left(strTemp, MAXLEN - 3) & "..." 'truncate the length if needed
      ActCont.SetTitle curUID, "Use " & Chr(34) & strTemp & Chr(34) & " to rename files."
    Else
      ActCont.SetTitle curUID, "File Mode (No path set yet)"
    End If
  ElseIf ActCont.GetActType(curUID) = fncSwitchCharacters Then
    'Modify the title for the Switch Chars type
    ActCont.SetTitle curUID, "Switch Characters"
  ElseIf ActCont.GetActType(curUID) = fncIncludeExcludeRules Then
    'Modify the tilte for the Rules type
    ActCont.SetTitle curUID, "Rules"
  End If
  'update the title if it needs to be
  If lstQ.List(c) <> ActCont.GetTitle(curUID) Then lstQ.List(c) = ActCont.GetTitle(curUID)
Next c



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Update_lstQ"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSample_Change(Index As Integer)
On Error GoTo errh

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

'Update the tooltiptext
txtSample(Index).ToolTipText = txtSample(Index).Text

'stop the update timer
tmrUpdate.Enabled = False

If curSampleIndex = fncSampleSwitchR1 Then
  'these are the fields that hold the visible selection (for SWITCH)
  txtSample(fncSampleSwitchR2).Text = txtSample(fncSampleSwitchR1).Text
  txtSample(fncSampleSwitchR2).SelStart = Val(ActCont.GetData(UIDIndex, fncSwitch1Sel_Start))
  txtSample(fncSampleSwitchR2).SelLength = Val(ActCont.GetData(UIDIndex, fncSwitch1SelLen))
  'start the update timer
  tmrUpdate.Enabled = True
ElseIf curSampleIndex = fncSampleSwitchR2 Then
  'these are the fields that hold the visible selection (for SWITCH)
  txtSample(fncSampleSwitchR1).Text = txtSample(fncSampleSwitchR2).Text
  txtSample(fncSampleSwitchR1).SelStart = Val(ActCont.GetData(UIDIndex, fncSwitch1Sel_Start))
  txtSample(fncSampleSwitchR1).SelLength = Val(ActCont.GetData(UIDIndex, fncSwitch1SelLen))
  'start the update timer
  tmrUpdate.Enabled = True
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSample_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSample_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help txtSample(Index).HelpContextID
  Exit Sub
End If

If curSampleIndex = fncSampleConcat Then
  'Set the placement based on where the user clicks (for CONCAT)
  curConcatIndex = 2
  If optConcat1(0).Value Then
    'Counting from the left
    txtConcat(2).Text = txtSample(Index).SelStart
  Else
    'Counting from the right
    txtConcat(2).Text = Len(txtSample(Index).Text) - (txtSample(Index).SelStart)
  End If
  curConcatIndex = -1
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSample_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSample_GotFocus(Index As Integer)
On Error GoTo errh

'keep track of which control has focus so that the currently edited control isn't updated while editing
curSampleIndex = Index

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSample_GotFocus"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSample_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errh

'call these when keys are pressed
txtSample_Click Index
txtSample_MouseMove Index, 1, 0, 0, 0 'send the shift as the button being pressed

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSample_KeyUp"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSample_LostFocus(Index As Integer)
On Error GoTo errh

If Index = 4 Then
  txtSample(fncSampleSwitchR1).SelStart = Val(ActCont.GetData(UIDIndex, fncSwitch1Sel_Start))
  txtSample(fncSampleSwitchR1).SelLength = Val(ActCont.GetData(UIDIndex, fncSwitch1SelLen))
ElseIf Index = 5 Then
  txtSample(fncSampleSwitchR2).SelStart = Val(ActCont.GetData(UIDIndex, fncSwitch2Sel_Start))
  txtSample(fncSampleSwitchR2).SelLength = Val(ActCont.GetData(UIDIndex, fncSwitch2SelLen))
End If

'if the curSampleIndex item is losing focus, then nothing is selected
If curSampleIndex = Index Then
  curSampleIndex = -1
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSample_LostFocus"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub




Private Sub txtSubDir_Change()
On Error GoTo errh

txtSubDir.ToolTipText = txtSubDir.Text

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSubDir_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSwitch_Change(Index As Integer)
On Error GoTo errh

'stop the active preview update timer
tmrUpdate.Enabled = False

'Don't update while settings are being updated/saved
If InSettings Then Exit Sub

If curSwitchIndex >= 0 And curSwitchIndex <= 3 Then
  ActCont.SetData UIDIndex, CLng(curSwitchIndex), txtSwitch(curSwitchIndex)
  If curSampleIndex <> fncSampleSwitchR1 Then
    txtSample(1).SelStart = Val(ActCont.GetData(UIDIndex, 0))
    txtSample(1).SelLength = Val(ActCont.GetData(UIDIndex, 1))
  End If
  If curSampleIndex <> fncSampleSwitchR2 Then
    txtSample(2).SelStart = Val(ActCont.GetData(UIDIndex, 2))
    txtSample(2).SelLength = Val(ActCont.GetData(UIDIndex, 3))
  End If
End If

txtSwitch(6).Text = CheckSwitch(Val(txtSwitch(0).Text), Val(txtSwitch(1).Text), Val(txtSwitch(2).Text), Val(txtSwitch(3).Text), True, txtSample(1).Text)
If InStr(1, txtSwitch(6).Text, "Error A") Then
  'Range 1 Start is off
  txtSwitch(0).BackColor = vbRed
Else
  txtSwitch(0).BackColor = &H80000005
End If
If InStr(1, txtSwitch(6).Text, "Error B") Then
  'Range 1 Start is off
  txtSwitch(1).BackColor = vbRed
Else
  txtSwitch(1).BackColor = &H80000005
End If
If InStr(1, txtSwitch(6).Text, "Error C") Then
  'Range 1 Length is off
  txtSwitch(2).BackColor = vbRed
Else
  txtSwitch(2).BackColor = &H80000005
End If
If InStr(1, txtSwitch(6).Text, "Error D") Then
  'Range 2 Length is off
  txtSwitch(3).BackColor = vbRed
Else
  txtSwitch(3).BackColor = &H80000005
End If
If InStr(1, txtSwitch(6).Text, "Error E") Then
  'Example is too short
  txtSample(1).BackColor = vbYellow
  txtSample(2).BackColor = vbYellow
Else
  txtSample(1).BackColor = &H80000005
  txtSample(2).BackColor = &H80000005
End If



Update_lstQ


'start the active preview update timer
tmrUpdate.Enabled = True

Exit Sub
errh:
If Err = 380 Then
  'invalid property
  Resume Next
End If
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSwitch_Change"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub



Private Sub txtSwitch_Click(Index As Integer)
On Error GoTo errh

If InHelpMode Then
  'If the Help button is activated, display a help topic instead
  Help txtSwitch(Index).HelpContextID
  Exit Sub
End If



Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSwitch_Click"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSwitch_GotFocus(Index As Integer)
On Error GoTo errh

'keep track of which control has focus so that the currently edited control isn't updated while editing
curSwitchIndex = Index

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSwitch_GotFocus"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSwitch_LostFocus(Index As Integer)
On Error GoTo errh

'if the curSwitchIndex item is losing focus, then nothing is selected
If curSwitchIndex = Index Then
  curSwitchIndex = -1
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSwitch_LostFocus"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Sub txtSample_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

If Button = 1 Then
  If curSampleIndex = fncSampleSwitchR1 Then
    'if index 4 is being edited save this data, and update 5 (for SWITCH)
    ActCont.SetData UIDIndex, 0, txtSample(fncSampleSwitchR1).SelStart
    ActCont.SetData UIDIndex, 1, txtSample(fncSampleSwitchR1).SelLength
    txtSwitch(0).Text = ActCont.GetData(UIDIndex, 0)
    txtSwitch(1).Text = ActCont.GetData(UIDIndex, 1)
    txtSample(fncSampleSwitchR2).SelStart = Val(ActCont.GetData(UIDIndex, 2))
    txtSample(fncSampleSwitchR2).SelLength = Val(ActCont.GetData(UIDIndex, 3))
  ElseIf curSampleIndex = fncSampleSwitchR2 Then
    'if index 5 is being edited save this data, and update 4 (for SWITCH)
    ActCont.SetData UIDIndex, 2, txtSample(fncSampleSwitchR2).SelStart
    ActCont.SetData UIDIndex, 3, txtSample(fncSampleSwitchR2).SelLength
    txtSwitch(2).Text = ActCont.GetData(UIDIndex, 2)
    txtSwitch(3).Text = ActCont.GetData(UIDIndex, 3)
    txtSample(fncSampleSwitchR1).SelStart = Val(ActCont.GetData(UIDIndex, 0))
    txtSample(fncSampleSwitchR1).SelLength = Val(ActCont.GetData(UIDIndex, 1))
  End If
End If

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "txtSample_MouseMove"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

Private Function Settings(SaveLoad As fncSaveLoadEnum, Optional SettingName As fncSaveLoadNameEnum = fncSLAll, Optional SettingValue As fncSaveLoadValueEnum) As fncSettingsEnum
'SaveLoad: = fncSave = 0 when saving, fncLoad = 1 when loading
'SettingName: = the name of the setting to save or load. If nothing is supplied, save or load all settings
'SettingValue: = the value to write to file. Only used when saving.
On Error GoTo errh

Dim FileName As String 'holds the file name to save/load to/from
Dim Data As String

'holds the Version pulled from the file
Dim FileVersion As String
Dim OldVerText As String 'Holds all the data from the old version file

Dim FileNum As Integer
Dim c As Integer
Dim c1 As Long 'sent to the ActCont data thing, needs to be a long
Dim r As Integer
Dim QCount As Integer 'holds the number of queue items saved when loading the Queue
Dim curUID As Integer
Dim PRE As String
PRE = Chr(1) 'placed at the beginning of each line, means a new property is starting

'Make sure the function isn't already running
If InSettings Then
  Settings = fncExitEarly 'return value, means the function is already running
  Exit Function
End If

'Don't let other functions call this
InSettings = True

'Get the filename
FileName = InputBox("Enter the path and filename of the save file.", "Enter Filename", App.Path & "\fnc.fnc")
'Open it for the appropriate mode
FileNum = FreeFile()
If SaveLoad = fncLoad Then
  Open FileName For Input As #FileNum
  Line Input #FileNum, FileVersion
  If Left(FileVersion, 1) = PRE Then
    'Seems to be a valid version number, take off the PRE character
    FileVersion = Mid(FileVersion, 2)
  ElseIf InStr(1, FileVersion, ",") > 0 Then
    r = MsgBox("This save file appears to be from an earlier version of Filename Changer that may not be supported. Do you want to try loading anyway?", vbYesNo + vbQuestion, "Old File Version")
    If r = vbNo Then
      Close #FileNum
      InSettings = False
      Exit Function
    End If
    FileVersion = "1"
  End If
ElseIf SaveLoad = fncSave Then
  Open FileName For Output As #FileNum
  Print #FileNum, PRE & LTrim(Str(cSAVE_VER))     'place the version number at the beginning, and get rid of the space that numbers put there
End If

'Read or write the settings
If SaveLoad = fncSave Then
  'Save Settings
  If SettingName = fncSLQueue Or SettingName = fncSLAll Then
    'If the Queue (or All settings) are being saved...

    Print #FileNum, PRE & "QueueItems"
    Print #FileNum, lstQ.ListCount
    For c = 0 To lstQ.ListCount - 1
      curUID = lstQ.ItemData(c)
      With ActCont
        'print the action type to file
        Print #FileNum, .GetActType(curUID)
        'Print each data item to file
        For c1 = 0 To 3
          Print #FileNum, .GetData(curUID, c1)
        Next c1
      End With
    Next c
  End If
  Settings = fncSaveSuccessful
ElseIf SaveLoad = fncLoad Then
  'Check Version and load it appropriately. Otherwise do any conversions needed
  If Val(FileVersion) > cSAVE_VER Then
    'From a newer version
    r = MsgBox("This save file is from a newer version of Filename Changer, and may not load correctly. Would you like to try anyway?" & vbCrLf & vbCrLf & "You can obtain the newest version at http://web.ics.purdue.edu/~akolberg", vbExclamation + vbYesNo, "Newer File Version")
    If r = vbNo Then
      Settings = fncLoadFailed
      InSettings = False
      Exit Function
    End If
  ElseIf Val(FileVersion) = 1 Then
    'Version 1
    'Supported right now, so do nothing
  ElseIf Val(FileVersion) = cSAVE_VER Then
    'Current Version
    'Nothing should ever need to be done, so don't do anything
  Else
    'Version not supported
    MsgBox "The save file you chose has an invalid version number, and may be corrupt. Aborting load.", vbCritical, "Corrupt Save File"
    Settings = fncLoadFailed
    InSettings = False
    Exit Function
  End If

  
  'Load Settings
  While Not EOF(FileNum)
    Line Input #FileNum, Data
    'if Data is the QueueItems header, and you are loading Queue settings, do it
    If Data = PRE & "QueueItems" And (SettingName = fncSLAll Or SettingName = fncSLQueue) Then
      lstQ.Visible = False 'make it invisible for speed
      Line Input #FileNum, Data 'The first thing after this is the number of queue items
      QCount = Val(Data)
      With ActCont
        'Add the appropriate number of actions, and the data associated with it
        For c = 0 To QCount - 1
          Line Input #FileNum, Data
          curUID = .AddAction(CLng(Val(Data)))
          
          'add to the end of the Q list
          lstQ.AddItem "Loaded Item", lstQ.ListCount
          lstQ.ItemData(lstQ.NewIndex) = curUID
          
          'Get each piece of data
          For c1 = 0 To 3
            Line Input #FileNum, Data
            .SetData curUID, c1, Data
          Next c1
        Next c
      End With
      lstQ.Visible = True 'make it visible again
      'Once done loading, update the titles
      Update_lstQ
    End If
  Wend
  Settings = fncLoadSuccessful
End If

Close #FileNum

InSettings = False 'allow settings to be saved/loaded again
Exit Function
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "Settings"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Function

Private Sub CreateSubDirList(SubList As ListBox)
'Create a list of the subdirectories of the current path from Dir1
On Error GoTo errh

Dim c As Integer
Dim LastSlash As Integer

SubList.Visible = False 'invisible for speed
SubList.Clear
Dir1.Refresh

'Add the "Create New Directory"
SubList.AddItem "*Add New Subdirectory"
SubList.AddItem "\"
'Add the directory names
For c = 0 To Dir1.ListCount - 1
  LastSlash = FindLast(Dir1.List(c), "\")
  
  SubList.AddItem Mid(Dir1.List(c), LastSlash)
Next c

SubList.Visible = True

Exit Sub
errh:
frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "CreateSubDirList"
#If Debugging = 1 Then
  Stop
#Else
  End
#End If
Resume

End Sub

'Private Sub PutExtention()
'On Error GoTo errh'
'
'Dim c As Integer
'
'Exit Sub'
'
'clstActPrev.Visible = True
'For c = 0 To clstActPrev.ListCount - 1
'  clstActPrev.List(c) = clstActPrev.List(c) & lstFiles(fncExtList).List(c)
'Next c
'
'Exit Sub
'errh:
'frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "PutExtention"
'#If Debugging = 1 Then
'  Stop
'#Else
'  End
'#End If
'Resume
'
'End Sub
'
'Private Function PullExtention(FileName As String, FileIndex As Integer) As String
''Takes a filename and returns it without the extention
'On Error GoTo errh'
'
'Dim c As Integer
'
''lstFiles(fncExtList).Clear
'c = FindLast(FileName, ".")
'If c - 1 >= 1 Then
'  PullExtention = Mid(FileName, 1, c - 1) 'Get just the name
'  lstFiles(fncExtList).AddItem Mid(FileName, c)
'Else
'  PullExtention = FileName
'  lstFiles(fncExtList).AddItem ""
'End If
'
'Exit Function
'errh:
'frmError.MsgBoxError Err & ": " & Err.Description & " Module: " & "PullExtention"
'#If Debugging = 1 Then
'  Stop
'#Else
'  End
'#End If
'Resume
'
'End Function
