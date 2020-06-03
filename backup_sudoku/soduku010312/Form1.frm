VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form SOD 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "G_SUDOKU "
   ClientHeight    =   5835
   ClientLeft      =   13050
   ClientTop       =   6240
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmhlp 
      Caption         =   "HELP"
      ForeColor       =   &H00808080&
      Height          =   4335
      Left            =   5040
      TabIndex        =   123
      Top             =   4680
      Width           =   3255
      Begin VB.CommandButton CMDHLP 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   125
         ToolTipText     =   "Exit"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txthlp 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   3495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   124
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   121
      ToolTipText     =   "Erase"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdlock 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaskColor       =   &H000000C0&
      Picture         =   "Form1.frx":16AC2
      Style           =   1  'Graphical
      TabIndex        =   119
      ToolTipText     =   "Click when you are done"
      Top             =   360
      Width           =   375
   End
   Begin VB.Timer wmptimer 
      Interval        =   1000
      Left            =   11520
      Top             =   6480
   End
   Begin VB.CheckBox chkmusic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Music On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   0
      TabIndex        =   116
      ToolTipText     =   "Alt+M"
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer timermarquee 
      Interval        =   100
      Left            =   2400
      Top             =   6840
   End
   Begin VB.Timer txttip 
      Left            =   2280
      Top             =   6480
   End
   Begin VB.TextBox txthp 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   8520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   111
      Text            =   "Form1.frx":16F04
      ToolTipText     =   "Press Esc To Exit This hellp Screen"
      Top             =   600
      Width           =   3255
   End
   Begin VB.Timer lblDateTimer 
      Interval        =   1000
      Left            =   1800
      Top             =   6840
   End
   Begin VB.TextBox txtgoogle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   960
      TabIndex        =   103
      ToolTipText     =   "Hit Enter To Search or Press Esc to Cancel"
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton PressNumber 
      Caption         =   "Press No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   102
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "****"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   91
      ToolTipText     =   "Clear All Boxes."
      Top             =   4560
      Width           =   495
   End
   Begin VB.Frame MENU 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   5040
      TabIndex        =   101
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdcustome 
         BackColor       =   &H008080FF&
         Caption         =   "Custome"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Manually Decorate Your Grid"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdclHlp 
         BackColor       =   &H008080FF&
         Caption         =   "Close bOXxx $"
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Close This Boxxxxx......"
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton cmdtip 
         BackColor       =   &H008080FF&
         Caption         =   "Tips"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Tips Of Your Game...."
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdimp 
         BackColor       =   &H008080FF&
         Caption         =   "Impossible"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Impossible Mode Of Your Game...."
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdprof 
         BackColor       =   &H008080FF&
         Caption         =   "Professional"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Professional Mode Of Your Game...."
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdhrd 
         BackColor       =   &H008080FF&
         Caption         =   "Hard"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Hard Mode Of Your Game...."
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdmd 
         BackColor       =   &H008080FF&
         Caption         =   "Medium"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Medium Mode Of Your Game...."
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdeasy 
         BackColor       =   &H008080FF&
         Caption         =   "Easy"
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FF80FF&
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Easy Mode Of Your Game...."
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         Height          =   3015
         Left            =   240
         Top             =   240
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   3375
         Left            =   120
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Timer bg 
      Interval        =   300
      Left            =   1560
      Top             =   6480
   End
   Begin VB.Timer title 
      Interval        =   600
      Left            =   840
      Top             =   6480
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   720
      Top             =   6840
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1200
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   240
      Top             =   6840
   End
   Begin VB.CommandButton cmdchk 
      Caption         =   "done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   94
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CMDEND 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   93
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CMD9 
      Caption         =   "9"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   90
      ToolTipText     =   "Nine"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD8 
      Caption         =   "8"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   89
      ToolTipText     =   "Eight"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD7 
      Caption         =   "7"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   88
      ToolTipText     =   "Seven"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD6 
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   87
      ToolTipText     =   "Six"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD5 
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   86
      ToolTipText     =   "Five"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD4 
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   85
      ToolTipText     =   "Four"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD3 
      Caption         =   "3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   84
      ToolTipText     =   "Three"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   83
      ToolTipText     =   "Two"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   82
      ToolTipText     =   "One"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdstart 
      Appearance      =   0  'Flat
      Caption         =   "STA&RT.... the Game Now"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "Form1.frx":16F0A
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Start The Game "
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Timer lbltimer 
      Interval        =   100
      Left            =   120
      Top             =   6480
   End
   Begin VB.Label lbmistakes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "**No Mistakes Yet."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   135
      TabIndex        =   122
      Top             =   4200
      Width           =   3240
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblmode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Game Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      Left            =   600
      TabIndex        =   120
      Top             =   4515
      Width           =   2355
   End
   Begin WMPLibCtl.WindowsMediaPlayer lblmusic 
      Height          =   735
      Left            =   9480
      TabIndex        =   118
      Top             =   7080
      Width           =   1935
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3413
      _cy             =   1296
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpmusicGobinda 
      Height          =   855
      Left            =   9480
      TabIndex        =   117
      Top             =   6000
      Width           =   1935
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3413
      _cy             =   1508
   End
   Begin VB.Label Label82 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO                       G_ GAME WORLD'..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   705
      Left            =   120
      TabIndex        =   115
      Top             =   0
      Width           =   3480
   End
   Begin VB.Label lblmarquee 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "User Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   -1080
      TabIndex        =   113
      ToolTipText     =   "This is your Machin Name"
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblUserMsg 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   112
      Top             =   360
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbldiv 
      BackColor       =   &H00C00000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbltimer2 
      BackColor       =   &H00C00000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   99
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbltimer1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbllog 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "Time :"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   97
      Top             =   5590
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "#########"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   96
      Top             =   5590
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label83 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Creator: Gobinda Nandi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   885
      MousePointer    =   10  'Up Arrow
      TabIndex        =   95
      ToolTipText     =   "Click Here To Get To Know About The Creator Of This Game.  "
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label LBLCON 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   92
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape SHP 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   375
      Left            =   4680
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label81 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label80 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   24
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label78 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label76 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label75 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label72 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label63 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   53
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   51
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   52
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   44
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   42
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   43
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   35
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   34
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   47
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   46
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   38
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   37
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   28
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   50
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   48
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   49
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   41
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   39
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   40
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   32
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   30
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   31
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   80
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   78
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   79
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   71
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   69
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   70
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   3000
      TabIndex        =   62
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2280
      TabIndex        =   60
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   2640
      TabIndex        =   61
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   73
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   75
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   72
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   65
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   63
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   64
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   840
      TabIndex        =   56
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   480
      TabIndex        =   55
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   77
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   74
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   76
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   68
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   66
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   67
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1920
      TabIndex        =   59
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1200
      TabIndex        =   57
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1560
      TabIndex        =   58
      Top             =   3000
      Width           =   375
   End
   Begin VB.Menu mnNew 
      Caption         =   "New"
      Index           =   2
      NegotiatePosition=   2  'Middle
      WindowList      =   -1  'True
      Begin VB.Menu mnunew1 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuG 
         Caption         =   "Google Search"
         Shortcut        =   ^G
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "Help"
      Index           =   2
      Begin VB.Menu method 
         Caption         =   "How To Play"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuopt 
         Caption         =   "Color Option"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuweb 
      Caption         =   "Web"
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
      End
      Begin VB.Menu mnublog 
         Caption         =   "Blog"
      End
   End
End
Attribute VB_Name = "SOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputNumber As Integer
Dim CheckNumber As Integer
Dim x1 As String
Dim x2 As String
Dim x3 As String
Dim X4 As String
Dim X5 As String
Dim X6 As String
Dim X7 As String
Dim X8 As String
Dim X9 As String
Dim CheckError As String
Dim str As String
Dim cmdstartString As Integer
Dim txthelp As String
Dim txthpi As Integer
Dim FillBoxCount As Integer
Dim mid As Integer
Dim idChk As Integer
Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long
Dim modChk As String
Public ComputerNameN As String
Dim finishI As Integer
Dim FinishB As Boolean
Dim I As Integer
Dim j As Integer
Dim FinishTimeT As Integer
Dim musicI As Integer
Dim FinishScore As Double
Public LabelToolTipText As String
Public lblnmmu As String
Dim startClick As Boolean
Dim PlayMode As String
Public Function ComputerName() As String
  Dim sBuffer As String
  Dim lAns As Long
  sBuffer = Space$(255)
  lAns = GetComputerName(sBuffer, 255)
  If lAns <> 0 Then
        ''''''''''''''''''''''''read from beginning of string to null-terminator
        ComputerNameN = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
   Else
        Err.Raise Err.LastDllError, , _
          "A system call returned an error code of " _
           & Err.LastDllError
   End If
   If ComputerNameN = "" Then
        ComputerNameN = "Set A Computer Name"
   End If
   lblmarquee.Caption = "Player : " & UCase(ComputerNameN)
End Function

Private Sub chk1()
    If Label1.Caption <> "" Then
        x1 = Label1.Caption
        If Label2.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x1 Then
            Label19.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label28.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label28.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label31.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label31.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label34.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label34.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label61.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label61.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label55.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label55.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label58.Caption = x1 Then
            Label1.BackColor = vbWhite
            Label58.BackColor = vbWhite
            CheckError = "1"
        End If
End If
End Sub

Private Sub chk17()
    If Label17.Caption <> "" Then
        X17 = Label17.Caption
        If Label18.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label38.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label38.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label41.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label41.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label44.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label44.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label65.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label65.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label68.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label68.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label71.Caption = X17 Then
            Label17.BackColor = vbWhite
            Label71.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk18()
 If Label18.Caption <> "" Then
    X18 = Label18.Caption
     If Label16.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label16.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label17.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label17.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label13.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label13.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label14.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label14.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label15.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label15.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label10.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label10.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label11.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label11.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label12.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label12.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label26.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label26.BackColor = vbWhite
            CheckError = "1"
    End If
    If Label25.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label25.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label27.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label27.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label7.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label7.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label8.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label8.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label9.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label19.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label39.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label39.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label42.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label42.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label45.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label45.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label66.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label66.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label69.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label69.BackColor = vbWhite
        CheckError = "1"
    End If
    If Label72.Caption = X18 Then
        Label18.BackColor = vbWhite
        Label72.BackColor = vbWhite
        CheckError = "1"
    End If
 End If
End Sub

Private Sub chk19()
    If Label19.Caption <> "" Then
        x19 = Label19.Caption
        If Label23.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label1.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label46.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label73.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label73.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label76.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label76.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label79.Caption = x19 Then
            Label19.BackColor = vbWhite
            Label79.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk2()
    If Label2.Caption <> "" Then
        x2 = Label2.Caption
        If Label1.Caption = x2 Then
            Label1.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x2 Then
            Label19.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label29.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label29.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label32.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label32.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label35.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label35.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label62.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label62.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label56.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label56.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label59.Caption = x2 Then
            Label2.BackColor = vbWhite
            Label59.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk20()
    If Label20.Caption <> "" Then
        x20 = Label20.Caption
        If Label23.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label1.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label74.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label74.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label77.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label77.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label80.Caption = x20 Then
            Label20.BackColor = vbWhite
            Label80.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk21()
    If Label21.Caption <> "" Then
        x21 = Label21.Caption
        If Label23.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label1.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label75.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label75.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label78.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label78.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label81.Caption = x21 Then
            Label21.BackColor = vbWhite
            Label81.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk22()
    If Label22.Caption <> "" Then
        x22 = Label22.Caption
        If Label19.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label46.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label73.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label73.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label76.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label76.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label79.Caption = x22 Then
            Label22.BackColor = vbWhite
            Label79.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk23()
    If Label23.Caption <> "" Then
        x23 = Label23.Caption
        If Label19.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label74.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label74.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label77.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label77.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label80.Caption = x23 Then
            Label23.BackColor = vbWhite
            Label80.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk24()
    If Label24.Caption <> "" Then
        x24 = Label24.Caption
        If Label19.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label75.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label75.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label78.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label78.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label81.Caption = x24 Then
            Label24.BackColor = vbWhite
            Label81.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk25()
    If Label25.Caption <> "" Then
        x25 = Label25.Caption
        If Label19.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label18.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label46.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label73.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label73.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label76.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label76.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label79.Caption = x25 Then
            Label25.BackColor = vbWhite
            Label79.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk26()
    If Label26.Caption <> "" Then
        x26 = Label26.Caption
        If Label19.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label18.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label74.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label74.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label77.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label77.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label80.Caption = x26 Then
            Label26.BackColor = vbWhite
            Label80.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk27()
    If Label27.Caption <> "" Then
        x27 = Label27.Caption
        If Label19.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label18.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label75.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label75.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label78.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label78.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label81.Caption = x27 Then
            Label27.BackColor = vbWhite
            Label81.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk3()
    If Label3.Caption <> "" Then
        x3 = Label3.Caption
        If Label1.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = x3 Then
            Label2.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x3 Then
            Label19.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label36.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label36.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label33.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label33.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label30.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label30.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label63.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label63.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label60.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label60.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label57.Caption = x3 Then
            Label3.BackColor = vbWhite
            Label57.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub


Private Sub chk33()
    If Label33.Caption <> "" Then
            x33 = Label33.Caption
            If Label36.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
                    If Label3.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label3.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label6.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label9.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label9.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If

            If Label49.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x33 Then
                Label33.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk34()
    If Label34.Caption <> "" Then
            x34 = Label34.Caption
            If Label36.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label52.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label55.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label7.Caption = x34 Then
                Label34.BackColor = vbWhite
                Label7.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk35()
    If Label35.Caption <> "" Then
            x35 = Label35.Caption
            If Label36.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label52.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label56.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label8.Caption = x35 Then
                Label35.BackColor = vbWhite
                Label8.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk36()
    If Label36.Caption <> "" Then
            x36 = Label36.Caption
            If Label34.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
             
            If Label3.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label3.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label6.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label9.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label9.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = x36 Then
                Label36.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk4()
    If Label4.Caption <> "" Then
        X4 = Label4.Caption
        If Label1.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X4 Then
            Label2.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X4 Then
            Label3.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X4 Then
            Label14.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label28.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label28.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label31.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label31.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label34.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label34.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label61.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label61.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label55.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label55.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label58.Caption = X4 Then
            Label4.BackColor = vbWhite
            Label58.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk46()
    If Label46.Caption <> "" Then
        x46 = Label46.Caption
        If Label48.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label28.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label28.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label29.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label29.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label30.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label30.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label37.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label37.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label38.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label38.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label39.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label39.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label73.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label73.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label76.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label76.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label79.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label79.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x46 Then
            Label46.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk47()
    If Label47.Caption <> "" Then
        x47 = Label47.Caption
        If Label48.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label46.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label28.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label28.BackColor = vbWhite
            CheckError = "1"
        End If
        
        If Label29.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label29.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label30.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label30.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label37.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label37.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label38.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label38.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label39.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label39.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label74.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label74.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label77.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label77.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label80.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label80.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x47 Then
            Label47.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk48()
    If Label48.Caption <> "" Then
        x48 = Label48.Caption
        If Label46.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label75.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label75.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label78.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label78.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label81.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label81.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label28.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label28.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label29.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label29.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label30.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label30.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label37.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label37.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label38.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label38.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label39.Caption = x48 Then
            Label48.BackColor = vbWhite
            Label39.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk49()
    If Label49.Caption <> "" Then
        x49 = Label49.Caption
        If Label46.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label31.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label31.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label32.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label32.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label33.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label33.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label40.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label40.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label41.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label41.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label42.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label42.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label73.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label73.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label76.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label76.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label79.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label79.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x49 Then
            Label49.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk50()
    If Label50.Caption <> "" Then
        x50 = Label50.Caption
        If Label46.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label31.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label31.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label32.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label32.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label33.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label33.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label40.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label40.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label41.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label41.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label42.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label42.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label74.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label74.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label77.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label77.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label80.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label80.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x50 Then
            Label50.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If

    End If
End Sub

Private Sub chk51()
    If Label51.Caption <> "" Then
        x51 = Label51.Caption
        If Label46.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label75.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label75.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label78.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label78.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label81.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label81.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label31.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label31.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label32.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label32.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label33.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label33.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label40.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label40.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label41.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label41.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label42.Caption = x51 Then
            Label51.BackColor = vbWhite
            Label42.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk52()
    If Label52.Caption <> "" Then
        x52 = Label52.Caption
        If Label46.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label34.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label34.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label35.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label35.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label36.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label36.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label43.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label43.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label44.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label44.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label45.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label45.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label73.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label73.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label76.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label76.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label79.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label79.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = x52 Then
            Label52.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk53()
    If Label53.Caption <> "" Then
        x53 = Label53.Caption
        If Label46.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label54.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label54.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label34.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label34.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label35.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label35.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label36.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label36.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label43.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label43.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label44.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label44.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label45.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label45.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label74.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label74.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label77.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label77.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label80.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label80.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = x53 Then
            Label53.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk54()
    If Label54.Caption <> "" Then
        x54 = Label54.Caption
        If Label46.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label46.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label47.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label47.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label48.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label48.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label49.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label49.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label50.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label50.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label51.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label51.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label52.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label52.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label53.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label53.BackColor = vbWhite
            CheckError = "1"
        End If
            If Label34.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label34.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label35.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label35.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label36.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label36.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label43.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label43.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label44.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label44.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label45.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label45.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label75.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label75.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label78.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label78.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label81.Caption = x54 Then
            Label54.BackColor = vbWhite
            Label81.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk6()
    If Label6.Caption <> "" Then
        X6 = Label6.Caption
        If Label1.Caption = X6 Then
            Label1.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X6 Then
            Label4.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X6 Then
            Label5.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X6 Then
            Label14.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label36.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label36.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label33.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label33.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label30.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label30.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label63.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label63.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label60.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label60.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label57.Caption = X6 Then
            Label6.BackColor = vbWhite
            Label57.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub


Private Sub chk7()
    If Label7.Caption <> "" Then
        X7 = Label7.Caption
        If Label1.Caption = X7 Then
            Label1.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X7 Then
            Label2.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X7 Then
            Label5.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X7 Then
            Label6.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
            If Label26.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label18.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label28.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label28.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label31.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label31.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label34.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label34.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label61.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label61.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label55.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label55.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label58.Caption = X7 Then
            Label7.BackColor = vbWhite
            Label58.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk8()
    If Label8.Caption <> "" Then
        X8 = Label8.Caption
        If Label1.Caption = X8 Then
            Label1.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X8 Then
            Label4.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X8 Then
            Label5.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X8 Then
            Label6.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X8 Then
            Label7.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
                If Label29.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label29.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label32.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label32.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label35.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label35.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label62.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label62.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label56.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label56.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label59.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label59.BackColor = vbWhite
            CheckError = "1"
        End If
                If Label26.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label18.Caption = X8 Then
            Label8.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub

Private Sub chk9()
    If Label9.Caption <> "" Then
        X9 = Label9.Caption
        If Label1.Caption = X9 Then
            Label1.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X9 Then
            Label8.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X9 Then
            Label8.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X9 Then
            Label4.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X9 Then
            Label5.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X9 Then
            Label6.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X9 Then
            Label7.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X9 Then
            Label8.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
       If Label26.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label18.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        
        If Label36.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label36.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label33.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label33.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label30.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label30.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label63.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label63.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label60.Caption = X9 Then
            Label8.BackColor = vbWhite
            Label60.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label57.Caption = X9 Then
            Label9.BackColor = vbWhite
            Label57.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub


 
Private Sub chk5()
    If Label5.Caption <> "" Then
        X5 = Label5.Caption
        If Label1.Caption = X5 Then
            Label1.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label3.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X5 Then
            Label4.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label6.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label9.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X5 Then
            Label14.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label29.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label29.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label32.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label32.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label35.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label35.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label62.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label62.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label56.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label56.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label59.Caption = X5 Then
            Label5.BackColor = vbWhite
            Label59.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub


Private Sub INPUTE()
    If CheckNumber = 1 Then
        Label1.Caption = InputNumber
    End If
    If CheckNumber = 2 Then
        Label2.Caption = InputNumber
    End If
    If CheckNumber = 3 Then
        Label3.Caption = InputNumber
    End If
    If CheckNumber = 4 Then
        Label4.Caption = InputNumber
    End If
    If CheckNumber = 5 Then
        Label5.Caption = InputNumber
    End If
    If CheckNumber = 6 Then
        Label6.Caption = InputNumber
    End If
    If CheckNumber = 7 Then
        Label7.Caption = InputNumber
    End If
    If CheckNumber = 8 Then
        Label8.Caption = InputNumber
    End If
    If CheckNumber = 9 Then
        Label9.Caption = InputNumber
    End If
    If CheckNumber = 10 Then
        Label10.Caption = InputNumber
    End If
    If CheckNumber = 11 Then
        Label11.Caption = InputNumber
    End If
    If CheckNumber = 12 Then
        Label12.Caption = InputNumber
    End If
    If CheckNumber = 13 Then
        Label13.Caption = InputNumber
    End If
    If CheckNumber = 14 Then
        Label14.Caption = InputNumber
    End If
    If CheckNumber = 15 Then
        Label15.Caption = InputNumber
    End If
    If CheckNumber = 16 Then
        Label16.Caption = InputNumber
    End If
    If CheckNumber = 17 Then
        Label17.Caption = InputNumber
    End If
    If CheckNumber = 18 Then
        Label18.Caption = InputNumber
    End If
    If CheckNumber = 19 Then
        Label19.Caption = InputNumber
    End If
    If CheckNumber = 20 Then
        Label20.Caption = InputNumber
    End If
    If CheckNumber = 21 Then
        Label21.Caption = InputNumber
    End If
    If CheckNumber = 22 Then
        Label22.Caption = InputNumber
    End If
    If CheckNumber = 23 Then
        Label23.Caption = InputNumber
    End If
    If CheckNumber = 24 Then
        Label24.Caption = InputNumber
    End If
    If CheckNumber = 25 Then
        Label25.Caption = InputNumber
    End If
    If CheckNumber = 26 Then
        Label26.Caption = InputNumber
    End If
    If CheckNumber = 27 Then
        Label27.Caption = InputNumber
    End If
    If CheckNumber = 28 Then
        Label28.Caption = InputNumber
    End If
    If CheckNumber = 29 Then
        Label29.Caption = InputNumber
    End If
    If CheckNumber = 30 Then
        Label30.Caption = InputNumber
    End If
    If CheckNumber = 31 Then
        Label31.Caption = InputNumber
    End If
    If CheckNumber = 32 Then
        Label32.Caption = InputNumber
    End If
    If CheckNumber = 33 Then
        Label33.Caption = InputNumber
    End If
    If CheckNumber = 34 Then
        Label34.Caption = InputNumber
    End If
    If CheckNumber = 35 Then
        Label35.Caption = InputNumber
    End If
    If CheckNumber = 36 Then
        Label36.Caption = InputNumber
    End If
    If CheckNumber = 37 Then
        Label37.Caption = InputNumber
    End If
    If CheckNumber = 38 Then
        Label38.Caption = InputNumber
    End If
    If CheckNumber = 39 Then
        Label39.Caption = InputNumber
    End If
    If CheckNumber = 40 Then
        Label40.Caption = InputNumber
    End If
    If CheckNumber = 41 Then
        Label41.Caption = InputNumber
    End If
    If CheckNumber = 42 Then
        Label42.Caption = InputNumber
    End If
    If CheckNumber = 43 Then
        Label43.Caption = InputNumber
    End If
    If CheckNumber = 44 Then
        Label44.Caption = InputNumber
    End If
    If CheckNumber = 45 Then
        Label45.Caption = InputNumber
    End If
    If CheckNumber = 46 Then
        Label46.Caption = InputNumber
    End If
    If CheckNumber = 47 Then
        Label47.Caption = InputNumber
    End If
    If CheckNumber = 48 Then
        Label48.Caption = InputNumber
    End If
    If CheckNumber = 49 Then
        Label49.Caption = InputNumber
    End If
    If CheckNumber = 50 Then
        Label50.Caption = InputNumber
    End If
    If CheckNumber = 51 Then
        Label51.Caption = InputNumber
    End If
    If CheckNumber = 52 Then
        Label52.Caption = InputNumber
    End If
    If CheckNumber = 53 Then
        Label53.Caption = InputNumber
    End If
    If CheckNumber = 54 Then
        Label54.Caption = InputNumber
    End If
    If CheckNumber = 55 Then
        Label55.Caption = InputNumber
    End If
    If CheckNumber = 56 Then
        Label56.Caption = InputNumber
    End If
    If CheckNumber = 57 Then
        Label57.Caption = InputNumber
    End If
    If CheckNumber = 58 Then
        Label58.Caption = InputNumber
    End If
    If CheckNumber = 59 Then
        Label59.Caption = InputNumber
    End If
    If CheckNumber = 60 Then
        Label60.Caption = InputNumber
    End If
    If CheckNumber = 61 Then
        Label61.Caption = InputNumber
    End If
    If CheckNumber = 62 Then
        Label62.Caption = InputNumber
    End If
    If CheckNumber = 63 Then
        Label63.Caption = InputNumber
    End If
    If CheckNumber = 64 Then
        Label64.Caption = InputNumber
    End If
    If CheckNumber = 65 Then
        Label65.Caption = InputNumber
    End If
    If CheckNumber = 66 Then
        Label66.Caption = InputNumber
    End If
    If CheckNumber = 67 Then
        Label67.Caption = InputNumber
    End If
    If CheckNumber = 68 Then
        Label68.Caption = InputNumber
    End If
    If CheckNumber = 69 Then
        Label69.Caption = InputNumber
    End If
    If CheckNumber = 70 Then
        Label70.Caption = InputNumber
    End If
    If CheckNumber = 71 Then
        Label71.Caption = InputNumber
    End If
    If CheckNumber = 72 Then
        Label72.Caption = InputNumber
    End If
    If CheckNumber = 73 Then
        Label73.Caption = InputNumber
    End If
    If CheckNumber = 74 Then
        Label74.Caption = InputNumber
    End If
    If CheckNumber = 75 Then
        Label75.Caption = InputNumber
    End If
    If CheckNumber = 76 Then
        Label76.Caption = InputNumber
    End If
    If CheckNumber = 77 Then
        Label77.Caption = InputNumber
    End If
    If CheckNumber = 78 Then
        Label78.Caption = InputNumber
    End If
    If CheckNumber = 79 Then
        Label79.Caption = InputNumber
    End If
    If CheckNumber = 80 Then
        Label80.Caption = InputNumber
    End If
    If CheckNumber = 81 Then
        Label81.Caption = InputNumber
    End If
    
End Sub

Private Sub ResetGrid()
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = ""
    Label16.Caption = ""
    Label17.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    Label22.Caption = ""
    Label23.Caption = ""
    Label24.Caption = ""
    Label25.Caption = ""
    Label26.Caption = ""
    Label27.Caption = ""
    Label28.Caption = ""
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""
    Label32.Caption = ""
    Label33.Caption = ""
    Label34.Caption = ""
    Label35.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
    Label42.Caption = ""
    Label43.Caption = ""
    Label44.Caption = ""
    Label45.Caption = ""
    Label46.Caption = ""
    Label47.Caption = ""
    Label48.Caption = ""
    Label49.Caption = ""
    Label50.Caption = ""
    Label51.Caption = ""
    Label52.Caption = ""
    Label53.Caption = ""
    Label54.Caption = ""
    Label55.Caption = ""
    Label56.Caption = ""
    Label57.Caption = ""
    Label58.Caption = ""
    Label59.Caption = ""
    Label60.Caption = ""
    Label61.Caption = ""
    Label62.Caption = ""
    Label63.Caption = ""
    Label64.Caption = ""
    Label65.Caption = ""
    Label66.Caption = ""
    Label67.Caption = ""
    Label68.Caption = ""
    Label69.Caption = ""
    Label70.Caption = ""
    Label71.Caption = ""
    Label72.Caption = ""
    Label73.Caption = ""
    Label74.Caption = ""
    Label75.Caption = ""
    Label76.Caption = ""
    Label77.Caption = ""
    Label78.Caption = ""
    Label79.Caption = ""
    Label80.Caption = ""
    Label81.Caption = ""
    
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = True
    Label7.Enabled = True
    Label8.Enabled = True
    Label9.Enabled = True
    Label10.Enabled = True
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = True
    Label14.Enabled = True
    Label15.Enabled = True
    Label16.Enabled = True
    Label17.Enabled = True
    Label18.Enabled = True
    Label19.Enabled = True
    Label20.Enabled = True
    Label21.Enabled = True
    Label22.Enabled = True
    Label23.Enabled = True
    Label24.Enabled = True
    Label25.Enabled = True
    Label26.Enabled = True
    Label27.Enabled = True
    Label28.Enabled = True
    Label29.Enabled = True
    Label30.Enabled = True
    Label31.Enabled = True
    Label32.Enabled = True
    Label33.Enabled = True
    Label34.Enabled = True
    Label35.Enabled = True
    Label36.Enabled = True
    Label37.Enabled = True
    Label38.Enabled = True
    Label39.Enabled = True
    Label40.Enabled = True
    Label41.Enabled = True
    Label42.Enabled = True
    Label43.Enabled = True
    Label44.Enabled = True
    Label45.Enabled = True
    Label46.Enabled = True
    Label47.Enabled = True
    Label48.Enabled = True
    Label49.Enabled = True
    Label50.Enabled = True
    Label51.Enabled = True
    Label52.Enabled = True
    Label53.Enabled = True
    Label54.Enabled = True
    Label55.Enabled = True
    Label56.Enabled = True
    Label57.Enabled = True
    Label58.Enabled = True
    Label59.Enabled = True
    Label60.Enabled = True
    Label61.Enabled = True
    Label62.Enabled = True
    Label63.Enabled = True
    Label64.Enabled = True
    Label65.Enabled = True
    Label66.Enabled = True
    Label67.Enabled = True
    Label68.Enabled = True
    Label69.Enabled = True
    Label70.Enabled = True
    Label71.Enabled = True
    Label72.Enabled = True
    Label73.Enabled = True
    Label74.Enabled = True
    Label75.Enabled = True
    Label76.Enabled = True
    Label77.Enabled = True
    Label78.Enabled = True
    Label79.Enabled = True
    Label80.Enabled = True
    Label81.Enabled = True
    
        Label1.Tag = True
    Label2.Tag = True
    Label3.Tag = True
    Label4.Tag = True
    Label5.Tag = True
    Label6.Tag = True
    Label7.Tag = True
    Label8.Tag = True
    Label9.Tag = True
    Label10.Tag = True
    Label11.Tag = True
    Label12.Tag = True
    Label13.Tag = True
    Label14.Tag = True
    Label15.Tag = True
    Label16.Tag = True
    Label17.Tag = True
    Label18.Tag = True
    Label19.Tag = True
    Label20.Tag = True
    Label21.Tag = True
    Label22.Tag = True
    Label23.Tag = True
    Label24.Tag = True
    Label25.Tag = True
    Label26.Tag = True
    Label27.Tag = True
    Label28.Tag = True
    Label29.Tag = True
    Label30.Tag = True
    Label31.Tag = True
    Label32.Tag = True
    Label33.Tag = True
    Label34.Tag = True
    Label35.Tag = True
    Label36.Tag = True
    Label37.Tag = True
    Label38.Tag = True
    Label39.Tag = True
    Label40.Tag = True
    Label41.Tag = True
    Label42.Tag = True
    Label43.Tag = True
    Label44.Tag = True
    Label45.Tag = True
    Label46.Tag = True
    Label47.Tag = True
    Label48.Tag = True
    Label49.Tag = True
    Label50.Tag = True
    Label51.Tag = True
    Label52.Tag = True
    Label53.Tag = True
    Label54.Tag = True
    Label55.Tag = True
    Label56.Tag = True
    Label57.Tag = True
    Label58.Tag = True
    Label59.Tag = True
    Label60.Tag = True
    Label61.Tag = True
    Label62.Tag = True
    Label63.Tag = True
    Label64.Tag = True
    Label65.Tag = True
    Label66.Tag = True
    Label67.Tag = True
    Label68.Tag = True
    Label69.Tag = True
    Label70.Tag = True
    Label71.Tag = True
    Label72.Tag = True
    Label73.Tag = True
    Label74.Tag = True
    Label75.Tag = True
    Label76.Tag = True
    Label77.Tag = True
    Label78.Tag = True
    Label79.Tag = True
    Label80.Tag = True
    Label81.Tag = True
    LabelToolTipText = ""
    lblmode.Caption = "Game Mode"
    Call NORMAL
End Sub

Private Sub About_Click()
    MENU.Visible = False
    frmhlp.Visible = True
    txtgoogle.Visible = False
    frmhlp.Caption = "ABOUT"
    str = ""
    str = "This is totally a home made game.               GOBINDA NANDI, the programmer, designer, editor of this game."
    str = str + "He does not earn any money from this game,this game is not for sale."
    str = str + "He made this game just just to let you know a beginner can also pregramme a GAME."
    str = str + "Your Comment & Appriciation is awaited [Blog]......"
    txthlp.Left = 120
    txthlp.Top = 240
    frmhlp.Left = 120
    frmhlp.Top = 800
    CMDHLP.Left = 2160
    CMDHLP.Top = 3840
    txthlp.Text = ""
    txthlp.Text = str
    CMDHLP.SetFocus
    bg.Enabled = True
End Sub

Private Sub bg_Timer()
frmhlp.BackColor = Rnd * vbWhite
End Sub

Private Sub chkmusic_Click()
    If chkmusic.Value = 1 Then
        musicI = 0
        chkmusic.Caption = "&Music On"
        wmpmusicGobinda.URL = "C:\WINDOWS\Media\onestop.mid"
    ElseIf chkmusic.Value = 0 Then
        chkmusic.Caption = "&Music Off"
        wmpmusicGobinda.URL = ""
    End If
End Sub

Private Sub chkmusic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD1_Click()
    InputNumber = 0
    InputNumber = CMD1.Caption
    Call INPUTE
End Sub

Private Sub CMD1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD2_Click()
    InputNumber = 0
    InputNumber = CMD2.Caption
    Call INPUTE
End Sub

Private Sub CMD2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD3_Click()
    InputNumber = 0
    InputNumber = CMD3.Caption
    Call INPUTE
End Sub

Private Sub CMD3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD4_Click()
    InputNumber = 0
    InputNumber = CMD4.Caption
    Call INPUTE
End Sub

Private Sub CMD4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD5_Click()
    InputNumber = 0
    InputNumber = CMD5.Caption
    Call INPUTE
End Sub

Private Sub CMD5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD6_Click()
    InputNumber = 0
    InputNumber = CMD6.Caption
    Call INPUTE
End Sub

Private Sub CMD6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD7_Click()
    InputNumber = 0
    InputNumber = CMD7.Caption
    Call INPUTE
End Sub

Private Sub CMD7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD8_Click()
    InputNumber = 0
    InputNumber = CMD8.Caption
    Call INPUTE
End Sub

Private Sub CMD8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMD9_Click()
    InputNumber = 0
    InputNumber = CMD9.Caption
    Call INPUTE
End Sub

Private Sub CMDBACK_Click()
    Form_Load
End Sub

Private Sub CMD9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub

Private Sub cmdchk_Click()
 
    lblMsgBox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''' ''''strt 8th box  start'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CheckError = "0"
    NORMAL
    Call chk1
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk2
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk3
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk4
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk5
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk6
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk7
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk8
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk9
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''8th no box..end'''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''    '''''''strt 9th box''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk27
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk25
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk26
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk24
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk22
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk23
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk19
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk20
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk21
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''9th box..^^^^^^^^^^'''''''''END'''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''   ''''''''''''''7th box downnnn'''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk18
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk17
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk16
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk13
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk14
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk15
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk10
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk11
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk12
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''' '''''end of 7th box''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''6th box downnnn''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call chk54
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk52
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk53
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk51
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk49
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk50
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk48
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk46
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk47
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''   ''''''''''''''6th box up'''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '''''''''''''''''''''''''''''''''5th box START'''''''''''''''''''''
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk36
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk34
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk35
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk33
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk31
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk32
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk30
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk28
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk29
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''5th box END  '''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''4th bOx START'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk45
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk43
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk44
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk42
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk40
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk41
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk39
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk37
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk38
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''4th bOx END'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''3RD BOX START '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk81
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk79
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk80
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk78
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk76
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk77
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk75
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk73
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk74
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''3RD BOX END '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''2ND BOX START '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk63
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk61
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk62
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk60
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk58
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk59
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk57
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk55
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk56
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''2ND BOX END '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''1ST BOX START '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call chk72
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk70
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk71
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk69
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk67
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk68
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk66
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk64
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If
    Call chk65
    If CheckError = "1" Then
        Exit Sub
    Else: Call NORMAL
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''1ST BOX END '''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    FinishGame
End Sub

Private Sub chk13()
        If Label13.Caption <> "" Then
                X13 = Label13.Caption
            If Label18.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label12.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
             End If
             If Label22.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label22.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label23.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label23.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label24.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label24.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = X13 Then
                Label13.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub
Private Sub chk15()
    If Label15.Caption <> "" Then
        X15 = Label15.Caption
        If Label18.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label39.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label39.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label42.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label42.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label45.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label45.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label66.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label66.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label69.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label69.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label72.Caption = X15 Then
            Label15.BackColor = vbWhite
            Label72.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub
        
Private Sub chk16()
If Label16.Caption <> "" Then
        X16 = Label16.Caption
        If Label18.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label26.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label26.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label25.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label25.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label27.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label27.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label7.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label7.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label8.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label8.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label9.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label37.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label37.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label40.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label40.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label43.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label43.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label64.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label64.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label67.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label67.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label70.Caption = X16 Then
            Label16.BackColor = vbWhite
            Label70.BackColor = vbWhite
            CheckError = "1"
        End If
End If
End Sub

Private Sub cmdchk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub

Private Sub CmdClear_Click()
If MsgBox("You Are About To Clear All Boxes.", vbYesNo + vbInformation, "W A R N I N G") = vbYes Then

    If Label1.Tag = "" Then
        Label1.Tag = True
    End If
    If Label2.Tag = "" Then
        Label2.Tag = True
    End If
    If Label3.Tag = "" Then
        Label3.Tag = True
    End If
    If Label4.Tag = "" Then
        Label4.Tag = True
    End If
    If Label5.Tag = "" Then
        Label5.Tag = True
    End If
    If Label6.Tag = "" Then
        Label6.Tag = True
    End If
    If Label7.Tag = "" Then
        Label7.Tag = True
    End If
    If Label8.Tag = "" Then
        Label8.Tag = True
    End If
    If Label9.Tag = "" Then
        Label9.Tag = True
    End If
    If Label10.Tag = "" Then
        Label10.Tag = True
    End If
    If Label11.Tag = "" Then
        Label11.Tag = True
    End If
    If Label12.Tag = "" Then
        Label12.Tag = True
    End If
    If Label13.Tag = "" Then
        Label13.Tag = True
    End If
    If Label14.Tag = "" Then
        Label14.Tag = True
    End If
    If Label15.Tag = "" Then
        Label15.Tag = True
    End If
    If Label16.Tag = "" Then
        Label16.Tag = True
    End If
    If Label17.Tag = "" Then
        Label17.Tag = True
    End If
    If Label18.Tag = "" Then
        Label18.Tag = True
    End If
    If Label19.Tag = "" Then
        Label19.Tag = True
    End If
    If Label20.Tag = "" Then
        Label20.Tag = True
    End If
    If Label21.Tag = "" Then
        Label21.Tag = True
    End If
    If Label22.Tag = "" Then
        Label22.Tag = True
    End If
    If Label23.Tag = "" Then
        Label23.Tag = True
    End If
    If Label24.Tag = "" Then
        Label24.Tag = True
    End If
    If Label25.Tag = "" Then
        Label25.Tag = True
    End If
    If Label26.Tag = "" Then
        Label26.Tag = True
    End If
    If Label27.Tag = "" Then
        Label27.Tag = True
    End If
    If Label28.Tag = "" Then
        Label28.Tag = True
    End If
    If Label29.Tag = "" Then
        Label29.Tag = True
    End If
    If Label30.Tag = "" Then
        Label30.Tag = True
    End If
    If Label31.Tag = "" Then
        Label31.Tag = True
    End If
    If Label32.Tag = "" Then
        Label32.Tag = True
    End If
    If Label33.Tag = "" Then
        Label33.Tag = True
    End If
    If Label34.Tag = "" Then
        Label34.Tag = True
    End If
    If Label35.Tag = "" Then
        Label35.Tag = True
    End If
    If Label36.Tag = "" Then
        Label36.Tag = True
    End If
    If Label37.Tag = "" Then
        Label37.Tag = True
    End If
    If Label38.Tag = "" Then
        Label38.Tag = True
    End If
    If Label39.Tag = "" Then
        Label39.Tag = True
    End If
    If Label40.Tag = "" Then
        Label40.Tag = True
    End If
    If Label41.Tag = "" Then
        Label41.Tag = True
    End If
    If Label42.Tag = "" Then
        Label42.Tag = True
    End If
    If Label43.Tag = "" Then
        Label43.Tag = True
    End If
    If Label44.Tag = "" Then
        Label44.Tag = True
    End If
    If Label45.Tag = "" Then
        Label45.Tag = True
    End If
    If Label46.Tag = "" Then
        Label46.Tag = True
    End If
    If Label47.Tag = "" Then
        Label47.Tag = True
    End If
    If Label48.Tag = "" Then
        Label48.Tag = True
    End If
    If Label49.Tag = "" Then
        Label49.Tag = True
    End If
    If Label50.Tag = "" Then
        Label50.Tag = True
    End If
    If Label51.Tag = "" Then
        Label51.Tag = True
    End If
    If Label52.Tag = "" Then
        Label52.Tag = True
    End If
    If Label53.Tag = "" Then
        Label53.Tag = True
    End If
    If Label54.Tag = "" Then
        Label54.Tag = True
    End If
    If Label55.Tag = "" Then
        Label55.Tag = True
    End If
    If Label56.Tag = "" Then
        Label56.Tag = True
    End If
    If Label57.Tag = "" Then
        Label57.Tag = True
    End If
    If Label58.Tag = "" Then
        Label58.Tag = True
    End If
    If Label59.Tag = "" Then
        Label59.Tag = True
    End If
    If Label60.Tag = "" Then
        Label60.Tag = True
    End If
    If Label61.Tag = "" Then
        Label61.Tag = True
    End If
    If Label62.Tag = "" Then
        Label62.Tag = True
    End If
    If Label63.Tag = "" Then
        Label63.Tag = True
    End If
    If Label64.Tag = "" Then
        Label64.Tag = True
    End If
    If Label65.Tag = "" Then
        Label65.Tag = True
    End If
    If CheckNumber = 66 Then
        Label66.Tag = True
    End If
    If Label67.Tag = "" Then
        Label67.Tag = True
    End If
    If Label68.Tag = "" Then
        Label68.Tag = True
    End If
    If Label69.Tag = "" Then
        Label69.Tag = True
    End If
    If Label70.Tag = "" Then
        Label70.Tag = True
    End If
    If Label71.Tag = "" Then
        Label71.Tag = True
    End If
    If Label72.Tag = "" Then
        Label72.Tag = True
    End If
    If Label73.Tag = "" Then
        Label73.Tag = True
    End If
    If Label74.Tag = "" Then
        Label74.Tag = True
    End If
    If Label75.Tag = "" Then
        Label75.Tag = True
    End If
    If Label76.Tag = "" Then
        Label76.Tag = True
    End If
    If Label77.Tag = "" Then
        Label77.Tag = True
    End If
    If Label78.Tag = "" Then
        Label78.Tag = True
    End If
    If Label79.Tag = "" Then
        Label79.Tag = True
    End If
    If Label80.Tag = "" Then
        Label80.Tag = True
    End If
    If Label81.Tag = "" Then
        Label81.Tag = True
    End If


    If Label1.Tag = True Then
        Label1.Caption = ""
    End If
    If Label2.Tag = True Then
        Label2.Caption = ""
    End If
    If Label3.Tag = True Then
        Label3.Caption = ""
    End If
    If Label4.Tag = True Then
        Label4.Caption = ""
    End If
    If Label5.Tag = True Then
        Label5.Caption = ""
    End If
    If Label6.Tag = True Then
        Label6.Caption = ""
    End If
    If Label7.Tag = True Then
        Label7.Caption = ""
    End If
    If Label8.Tag = True Then
        Label8.Caption = ""
    End If
    If Label9.Tag = True Then
        Label9.Caption = ""
    End If
    If Label10.Tag = True Then
        Label10.Caption = ""
    End If
    If Label11.Tag = True Then
        Label11.Caption = ""
    End If
    If Label12.Tag = True Then
        Label12.Caption = ""
    End If
    If Label13.Tag = True Then
        Label13.Caption = ""
    End If
    If Label14.Tag = True Then
        Label14.Caption = ""
    End If
    If Label15.Tag = True Then
        Label15.Caption = ""
    End If
    If Label16.Tag = True Then
        Label16.Caption = ""
    End If
    If Label17.Tag = True Then
        Label17.Caption = ""
    End If
    If Label18.Tag = True Then
        Label18.Caption = ""
    End If
    If Label19.Tag = True Then
        Label19.Caption = ""
    End If
    If Label20.Tag = True Then
        Label20.Caption = ""
    End If
    If Label21.Tag = True Then
        Label21.Caption = ""
    End If
    If Label22.Tag = True Then
        Label22.Caption = ""
    End If
    If Label23.Tag = True Then
        Label23.Caption = ""
    End If
    If Label24.Tag = True Then
        Label24.Caption = ""
    End If
    If Label25.Tag = True Then
        Label25.Caption = ""
    End If
    If Label26.Tag = True Then
        Label26.Caption = ""
    End If
    If Label27.Tag = True Then
        Label27.Caption = ""
    End If
    If Label28.Tag = True Then
        Label28.Caption = ""
    End If
    If Label29.Tag = True Then
        Label29.Caption = ""
    End If
    If Label30.Tag = True Then
        Label30.Caption = ""
    End If
    If Label31.Tag = True Then
        Label31.Caption = ""
    End If
    If Label32.Tag = True Then
        Label32.Caption = ""
    End If
    If Label33.Tag = True Then
        Label33.Caption = ""
    End If
    If Label34.Tag = True Then
        Label34.Caption = ""
    End If
    If Label35.Tag = True Then
        Label35.Caption = ""
    End If
    If Label36.Tag = True Then
        Label36.Caption = ""
    End If
    If Label37.Tag = True Then
        Label37.Caption = ""
    End If
    If Label38.Tag = True Then
        Label38.Caption = ""
    End If
    If Label39.Tag = True Then
        Label39.Caption = ""
    End If
    If Label40.Tag = True Then
        Label40.Caption = ""
    End If
    If Label41.Tag = True Then
        Label41.Caption = ""
    End If
    If Label42.Tag = True Then
        Label42.Caption = ""
    End If
    If Label43.Tag = True Then
        Label43.Caption = ""
    End If
    If Label44.Tag = True Then
        Label44.Caption = ""
    End If
    If Label45.Tag = True Then
        Label45.Caption = ""
    End If
    If Label46.Tag = True Then
        Label46.Caption = ""
    End If
    If Label47.Tag = True Then
        Label47.Caption = ""
    End If
    If Label48.Tag = True Then
        Label48.Caption = ""
    End If
    If Label49.Tag = True Then
        Label49.Caption = ""
    End If
    If Label50.Tag = True Then
        Label50.Caption = ""
    End If
    If Label51.Tag = True Then
        Label51.Caption = ""
    End If
    If Label52.Tag = True Then
        Label52.Caption = ""
    End If
    If Label53.Tag = True Then
        Label53.Caption = ""
    End If
    If Label54.Tag = True Then
        Label54.Caption = ""
    End If
    If Label55.Tag = True Then
        Label55.Caption = ""
    End If
    If Label56.Tag = True Then
        Label56.Caption = ""
    End If
    If Label57.Tag = True Then
        Label57.Caption = ""
    End If
    If Label58.Tag = True Then
        Label58.Caption = ""
    End If
    If Label59.Tag = True Then
        Label59.Caption = ""
    End If
    If Label60.Tag = True Then
        Label60.Caption = ""
    End If
    If Label61.Tag = True Then
        Label61.Caption = ""
    End If
    If Label62.Tag = True Then
        Label62.Caption = ""
    End If
    If Label63.Tag = True Then
        Label63.Caption = ""
    End If
    If Label64.Tag = True Then
        Label64.Caption = ""
    End If
    If Label65.Tag = True Then
        Label65.Caption = ""
    End If
    If Label66.Tag = True Then
        Label66.Caption = ""
    End If
    If Label67.Tag = True Then
        Label67.Caption = ""
    End If
    If Label68.Tag = True Then
        Label68.Caption = ""
    End If
    If Label69.Tag = True Then
        Label69.Caption = ""
    End If
    If Label70.Tag = True Then
        Label70.Caption = ""
    End If
    If Label71.Tag = True Then
        Label71.Caption = ""
    End If
    If Label72.Tag = True Then
        Label72.Caption = ""
    End If
    If Label73.Tag = True Then
        Label73.Caption = ""
    End If
    If Label74.Tag = True Then
        Label74.Caption = ""
    End If
    If Label75.Tag = True Then
        Label75.Caption = ""
    End If
    If Label76.Tag = True Then
        Label76.Caption = ""
    End If
    If Label77.Tag = True Then
        Label77.Caption = ""
    End If
    If Label78.Tag = True Then
        Label78.Caption = ""
    End If
    If Label79.Tag = True Then
        Label79.Caption = ""
    End If
    If Label80.Tag = True Then
        Label80.Caption = ""
    End If
    If Label81.Tag = True Then
        Label81.Caption = ""
    End If

   End If
End Sub

Private Sub CmdClear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdclHlp_Click()
    MENU.Visible = False
    ResetGrid
End Sub

Private Sub cmdclHlp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdcustome_Click()
    lbltimer1.Caption = 0
    lbltimer2.Caption = 0
    ResetGrid
    MENU.Visible = False
    NORMAL
    cmdlock.Visible = True
    lblmode.Caption = "Custome Mode"
 End Sub

Private Sub cmdeasy_Click()
    PlayMode = "Easy"
    lbltimer1.Caption = 0
    lbltimer2.Caption = 0
    finishI = 0
    FinishB = False
    ResetGrid
    nuGameClear                     ''''only clear all boxes
    If Right(Format(Time, "ss"), 1) = 1 Then
        easy1
    ElseIf Right(Format(Time, "ss"), 1) = 2 Then
        easy2
    ElseIf Right(Format(Time, "ss"), 1) = 3 Then
        easy3
    ElseIf Right(Format(Time, "ss"), 1) = 4 Then
        easy4
    ElseIf Right(Format(Time, "ss"), 1) = 5 Then
        easy5
    ElseIf Right(Format(Time, "ss"), 1) = 6 Then
        easy6
    ElseIf Right(Format(Time, "ss"), 1) = 7 Then
        easy7
    ElseIf Right(Format(Time, "ss"), 1) = 8 Then
        easy8
    ElseIf Right(Format(Time, "ss"), 1) = 9 Then
        easy9
    ElseIf Right(Format(Time, "ss"), 1) = 0 Then
        easy10
    Else
        easyF
    End If
    lblmode.Caption = PlayMode & " Mode"
     MENU.Visible = False
End Sub

Private Sub cmdeasy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMDEND_Click()
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub

Private Sub CMDEND_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub CMDHLP_Click()
    txthlp.Left = 3960
    txthlp.Top = 3840
    frmhlp.Left = 3960
    frmhlp.Top = 3840
    CMDHLP.Left = 3960
    CMDHLP.Top = 3840
    bg.Enabled = False
End Sub

'Private Sub Command1_Click()
'    lbltimer.Enabled = False
'    Call NORMAL
'    Label82.Visible = False
'    Call enbl
'    Command1.Visible = False
''    CMDBACK.Visible = True
'    CMDEND.Visible = True
'    lbltime.Visible = True
'    lbllog.Visible = True
'    lbltime.Caption = Now
'    lbltimer1.Visible = True
'    lbltimer2.Visible = True
'    lbldiv.Visible = True
'End Sub
 
Private Sub enbl()
    CMD1.Enabled = True
    CMD2.Enabled = True
    CMD3.Enabled = True
    CMD4.Enabled = True
    CMD5.Enabled = True
    CMD6.Enabled = True
    CMD7.Enabled = True
    CMD8.Enabled = True
    CMD9.Enabled = True
    CmdClear.Enabled = True
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub CMDHLP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        frmhlp.Visible = False
    End If
End Sub

Private Sub cmdhrd_Click()
    PlayMode = "Hard"
    lbltimer1.Caption = 0
    lbltimer2.Caption = 0
    finishI = 0
    FinishB = False
    ResetGrid
    nuGameClear
    If Right(Format(Time, "ss"), 1) = 1 Then
        Hard1
    ElseIf Right(Format(Time, "ss"), 1) = 2 Then
        Hard2
    ElseIf Right(Format(Time, "ss"), 1) = 3 Then
        Hard3
    ElseIf Right(Format(Time, "ss"), 1) = 4 Then
        Hard4
    ElseIf Right(Format(Time, "ss"), 1) = 5 Then
        Hard5
    ElseIf Right(Format(Time, "ss"), 1) = 6 Then
        Hard6
    ElseIf Right(Format(Time, "ss"), 1) = 7 Then
        Hard7
    ElseIf Right(Format(Time, "ss"), 1) = 8 Then
        Hard8
    ElseIf Right(Format(Time, "ss"), 1) = 9 Then
        Hard9
    ElseIf Right(Format(Time, "ss"), 1) = 0 Then
        Hard10
    Else
        HardDefault
    End If
    lblmode.Caption = PlayMode & " Mode"
    MENU.Visible = False
End Sub

Private Sub cmdhrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdimp_Click()
    PlayMode = "Impossible"
    lbltimer1.Caption = 0
    lbltimer2.Caption = 0
    finishI = 0
    FinishB = False
    ResetGrid
    nuGameClear
    If Right(Format(Time, "ss"), 1) = 1 Then
        Impossible1
    ElseIf Right(Format(Time, "ss"), 1) = 2 Then
        Impossible2
    ElseIf Right(Format(Time, "ss"), 1) = 3 Then
        Impossible3
    ElseIf Right(Format(Time, "ss"), 1) = 4 Then
        Impossible4
    ElseIf Right(Format(Time, "ss"), 1) = 5 Then
        Impossible5
    ElseIf Right(Format(Time, "ss"), 1) = 6 Then
        Impossible6
    ElseIf Right(Format(Time, "ss"), 1) = 7 Then
        Impossible7
    ElseIf Right(Format(Time, "ss"), 1) = 8 Then
        Impossible8
    ElseIf Right(Format(Time, "ss"), 1) = 9 Then
        Impossible9
    ElseIf Right(Format(Time, "ss"), 1) = 0 Then
        Impossible10
    Else
        ImpossibleDefault
    End If
    lblmode.Caption = PlayMode & " Mode"
    MENU.Visible = False
End Sub

Private Sub cmdimp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdlock_Click()
    Call cmdchk_Click
    If CheckError = "1" Then
        MsgBox "oooopppsss....!!!! Your Data Values Are Not Sync, Make Your Grid Carefully.", vbCritical
        Exit Sub
    End If
'    lockBar.Value = 1
'    lockBar.Visible = True
        If lblLock(Label1) = True Then
        End If
        If lblLock(Label2) = True Then
        End If
        If lblLock(Label3) = True Then
        End If
        If lblLock(Label4) = True Then
        End If
        If lblLock(Label5) = True Then
        End If
        If lblLock(Label6) = True Then
        End If
        If lblLock(Label7) = True Then
        End If
        If lblLock(Label8) = True Then
        End If
        If lblLock(Label9) = True Then
        End If
        If lblLock(Label10) = True Then
        End If
        If lblLock(Label11) = True Then
        End If
        If lblLock(Label12) = True Then
        End If
        If lblLock(Label13) = True Then
        End If
        If lblLock(Label14) = True Then
        End If
        If lblLock(Label15) = True Then
        End If
        If lblLock(Label16) = True Then
        End If
        If lblLock(Label17) = True Then
        End If
        If lblLock(Label18) = True Then
        End If
        If lblLock(Label19) = True Then
        End If
        If lblLock(Label20) = True Then
        End If
        If lblLock(Label21) = True Then
        End If
        If lblLock(Label22) = True Then
        End If
        If lblLock(Label23) = True Then
        End If
        If lblLock(Label24) = True Then
        End If
        If lblLock(Label25) = True Then
        End If
        If lblLock(Label26) = True Then
        End If
        If lblLock(Label27) = True Then
        End If
        If lblLock(Label28) = True Then
        End If
        If lblLock(Label29) = True Then
        End If
        If lblLock(Label30) = True Then
        End If
        If lblLock(Label31) = True Then
        End If
        If lblLock(Label32) = True Then
        End If
        If lblLock(Label33) = True Then
        End If
        If lblLock(Label34) = True Then
        End If
        If lblLock(Label35) = True Then
        End If
        If lblLock(Label36) = True Then
        End If
        If lblLock(Label37) = True Then
        End If
        If lblLock(Label38) = True Then
        End If
        If lblLock(Label39) = True Then
        End If
        If lblLock(Label40) = True Then
        End If
        If lblLock(Label41) = True Then
        End If
        If lblLock(Label42) = True Then
        End If
        If lblLock(Label43) = True Then
        End If
        If lblLock(Label44) = True Then
        End If
        If lblLock(Label45) = True Then
        End If
        If lblLock(Label46) = True Then
        End If
        If lblLock(Label47) = True Then
        End If
        If lblLock(Label48) = True Then
        End If
        If lblLock(Label49) = True Then
        End If
        If lblLock(Label50) = True Then
        End If
        If lblLock(Label51) = True Then
        End If
        If lblLock(Label52) = True Then
        End If
        If lblLock(Label53) = True Then
        End If
        If lblLock(Label54) = True Then
        End If
        If lblLock(Label55) = True Then
        End If
        If lblLock(Label56) = True Then
        End If
        If lblLock(Label57) = True Then
        End If
        If lblLock(Label58) = True Then
        End If
        If lblLock(Label59) = True Then
        End If
        If lblLock(Label60) = True Then
        End If
        If lblLock(Label61) = True Then
        End If
        If lblLock(Label62) = True Then
        End If
        If lblLock(Label63) = True Then
        End If
        If lblLock(Label64) = True Then
        End If
        If lblLock(Label65) = True Then
        End If
        If lblLock(Label66) = True Then
        End If
        If lblLock(Label67) = True Then
        End If
        If lblLock(Label68) = True Then
        End If
        If lblLock(Label69) = True Then
        End If
        If lblLock(Label70) = True Then
        End If
        If lblLock(Label71) = True Then
        End If
        If lblLock(Label72) = True Then
        End If
        If lblLock(Label73) = True Then
        End If
        If lblLock(Label74) = True Then
        End If
        If lblLock(Label75) = True Then
        End If
        If lblLock(Label76) = True Then
        End If
        If lblLock(Label77) = True Then
        End If
        If lblLock(Label78) = True Then
        End If
        If lblLock(Label79) = True Then
        End If
        If lblLock(Label80) = True Then
        End If
        If lblLock(Label81) = True Then
        End If
'        FinishB = True
'        finishI = 100
'        timermarquee_Timer
'        NORMAL
'        FinishB = False
'        finishI = 0
        lbltimer1.Caption = 0
        lbltimer2.Caption = 0
        cmdlock.Visible = False
'        lockBar.Visible = False

End Sub

Private Sub cmdlock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdmd_Click()
    PlayMode = "Medium"
    lbltimer1.Caption = 0
    lbltimer2.Caption = 0
    finishI = 0
    FinishB = False
    ResetGrid
    nuGameClear
    If Right(Format(Time, "ss"), 1) = 1 Then
        Medium1
    ElseIf Right(Format(Time, "ss"), 1) = 2 Then
        Medium2
    ElseIf Right(Format(Time, "ss"), 1) = 3 Then
        Medium3
    ElseIf Right(Format(Time, "ss"), 1) = 4 Then
        Medium4
    ElseIf Right(Format(Time, "ss"), 1) = 5 Then
        Medium5
    ElseIf Right(Format(Time, "ss"), 1) = 6 Then
        Medium6
    ElseIf Right(Format(Time, "ss"), 1) = 7 Then
        Medium7
    ElseIf Right(Format(Time, "ss"), 1) = 8 Then
        Medium8
    ElseIf Right(Format(Time, "ss"), 1) = 9 Then
        Medium9
    ElseIf Right(Format(Time, "ss"), 1) = 0 Then
        Medium10
    Else
        MediumDefault
    End If
    lblmode.Caption = PlayMode & " Mode"
    MENU.Visible = False
End Sub

Private Sub cmdmd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdprof_Click()
    PlayMode = "Professional"
    lbltimer1.Caption = 0
    lbltimer2.Caption = 0
    finishI = 0
    FinishB = False
    ResetGrid
    nuGameClear
    If Right(Format(Time, "ss"), 1) = 1 Then
        Professional1
    ElseIf Right(Format(Time, "ss"), 1) = 2 Then
        Professional2
    ElseIf Right(Format(Time, "ss"), 1) = 3 Then
        Professional3
    ElseIf Right(Format(Time, "ss"), 1) = 4 Then
        Professional4
    ElseIf Right(Format(Time, "ss"), 1) = 5 Then
        Professional5
    ElseIf Right(Format(Time, "ss"), 1) = 6 Then
        Professional6
    ElseIf Right(Format(Time, "ss"), 1) = 7 Then
        Professional7
    ElseIf Right(Format(Time, "ss"), 1) = 8 Then
        Professional8
    ElseIf Right(Format(Time, "ss"), 1) = 9 Then
        Professional9
    ElseIf Right(Format(Time, "ss"), 1) = 0 Then
        Professional10
    Else
        ProfessionalDefault
    End If
    lblmode.Caption = PlayMode & " Mode"
    MENU.Visible = False
End Sub

Private Sub cmdprof_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdstart_Click()
    startClick = True
    lblUserMsg.Visible = True
    cmdstartString = 1
    lbltimer.Enabled = False
    Call NORMAL
    Label82.Visible = False
    Call enbl
    cmdstart.Visible = False
'    CMDBACK.Visible = True
    CMDEND.Visible = True
    lbltime.Visible = True
    lbllog.Visible = True
'    lbltime.Caption = Now
    lbltimer1.Visible = True
    lbltimer2.Visible = True
    lbldiv.Visible = True
'    txtgoogle.Visible = True
    chkmusic.Visible = True
    chkmusic.Value = 1
    chkmusic_Click
'    chkmusic.Caption = "Music On"
'    wmpmusicGobinda.URL = "C:\WINDOWS\Media\onestop.mid"
    lblmode.Visible = True
    lbmistakes.Visible = True
    ResetGrid
End Sub

Private Sub cmdstart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub cmdtip_Click()
    txthp = 0
    txthpi = 0
    txthp.Text = ""
    txttip.Interval = 200
    txthp.Visible = True
    txthp.Top = 149
    txthp.Left = 120
    txthp.SetFocus
    MENU.Visible = False
End Sub

Private Sub cmdtip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub Command2_Click()
 
    If CheckNumber = 1 Then
        Label1.Caption = ""
    End If
    If CheckNumber = 2 Then
        Label2.Caption = ""
    End If
    If CheckNumber = 3 Then
        Label3.Caption = ""
    End If
    If CheckNumber = 4 Then
        Label4.Caption = ""
    End If
    If CheckNumber = 5 Then
        Label5.Caption = ""
    End If
    If CheckNumber = 6 Then
        Label6.Caption = ""
    End If
    If CheckNumber = 7 Then
        Label7.Caption = ""
    End If
    If CheckNumber = 8 Then
        Label8.Caption = ""
    End If
    If CheckNumber = 9 Then
        Label9.Caption = ""
    End If
    If CheckNumber = 10 Then
        Label10.Caption = ""
    End If
    If CheckNumber = 11 Then
        Label11.Caption = ""
    End If
    If CheckNumber = 12 Then
        Label12.Caption = ""
    End If
    If CheckNumber = 13 Then
        Label13.Caption = ""
    End If
    If CheckNumber = 14 Then
        Label14.Caption = ""
    End If
    If CheckNumber = 15 Then
        Label15.Caption = ""
    End If
    If CheckNumber = 16 Then
        Label16.Caption = ""
    End If
    If CheckNumber = 17 Then
        Label17.Caption = ""
    End If
    If CheckNumber = 18 Then
        Label18.Caption = ""
    End If
    If CheckNumber = 19 Then
        Label19.Caption = ""
    End If
    If CheckNumber = 20 Then
        Label20.Caption = ""
    End If
    If CheckNumber = 21 Then
        Label21.Caption = ""
    End If
    If CheckNumber = 22 Then
        Label22.Caption = ""
    End If
    If CheckNumber = 23 Then
        Label23.Caption = ""
    End If
    If CheckNumber = 24 Then
        Label24.Caption = ""
    End If
    If CheckNumber = 25 Then
        Label25.Caption = ""
    End If
    If CheckNumber = 26 Then
        Label26.Caption = ""
    End If
    If CheckNumber = 27 Then
        Label27.Caption = ""
    End If
    If CheckNumber = 28 Then
        Label28.Caption = ""
    End If
    If CheckNumber = 29 Then
        Label29.Caption = ""
    End If
    If CheckNumber = 30 Then
        Label30.Caption = ""
    End If
    If CheckNumber = 31 Then
        Label31.Caption = ""
    End If
    If CheckNumber = 32 Then
        Label32.Caption = ""
    End If
    If CheckNumber = 33 Then
        Label33.Caption = ""
    End If
    If CheckNumber = 34 Then
        Label34.Caption = ""
    End If
    If CheckNumber = 35 Then
        Label35.Caption = ""
    End If
    If CheckNumber = 36 Then
        Label36.Caption = ""
    End If
    If CheckNumber = 37 Then
        Label37.Caption = ""
    End If
    If CheckNumber = 38 Then
        Label38.Caption = ""
    End If
    If CheckNumber = 39 Then
        Label39.Caption = ""
    End If
    If CheckNumber = 40 Then
        Label40.Caption = ""
    End If
    If CheckNumber = 41 Then
        Label41.Caption = ""
    End If
    If CheckNumber = 42 Then
        Label42.Caption = ""
    End If
    If CheckNumber = 43 Then
        Label43.Caption = ""
    End If
    If CheckNumber = 44 Then
        Label44.Caption = ""
    End If
    If CheckNumber = 45 Then
        Label45.Caption = ""
    End If
    If CheckNumber = 46 Then
        Label46.Caption = ""
    End If
    If CheckNumber = 47 Then
        Label47.Caption = ""
    End If
    If CheckNumber = 48 Then
        Label48.Caption = ""
    End If
    If CheckNumber = 49 Then
        Label49.Caption = ""
    End If
    If CheckNumber = 50 Then
        Label50.Caption = ""
    End If
    If CheckNumber = 51 Then
        Label51.Caption = ""
    End If
    If CheckNumber = 52 Then
        Label52.Caption = ""
    End If
    If CheckNumber = 53 Then
        Label53.Caption = ""
    End If
    If CheckNumber = 54 Then
        Label54.Caption = ""
    End If
    If CheckNumber = 55 Then
        Label55.Caption = ""
    End If
    If CheckNumber = 56 Then
        Label56.Caption = ""
    End If
    If CheckNumber = 57 Then
        Label57.Caption = ""
    End If
    If CheckNumber = 58 Then
        Label58.Caption = ""
    End If
    If CheckNumber = 59 Then
        Label59.Caption = ""
    End If
    If CheckNumber = 60 Then
        Label60.Caption = ""
    End If
    If CheckNumber = 61 Then
        Label61.Caption = ""
    End If
    If CheckNumber = 62 Then
        Label62.Caption = ""
    End If
    If CheckNumber = 63 Then
        Label63.Caption = ""
    End If
    If CheckNumber = 64 Then
        Label64.Caption = ""
    End If
    If CheckNumber = 65 Then
        Label65.Caption = ""
    End If
    If CheckNumber = 66 Then
        Label66.Caption = ""
    End If
    If CheckNumber = 67 Then
        Label67.Caption = ""
    End If
    If CheckNumber = 68 Then
        Label68.Caption = ""
    End If
    If CheckNumber = 69 Then
        Label69.Caption = ""
    End If
    If CheckNumber = 70 Then
        Label70.Caption = ""
    End If
    If CheckNumber = 71 Then
        Label71.Caption = ""
    End If
    If CheckNumber = 72 Then
        Label72.Caption = ""
    End If
    If CheckNumber = 73 Then
        Label73.Caption = ""
    End If
    If CheckNumber = 74 Then
        Label74.Caption = ""
    End If
    If CheckNumber = 75 Then
        Label75.Caption = ""
    End If
    If CheckNumber = 76 Then
        Label76.Caption = ""
    End If
    If CheckNumber = 77 Then
        Label77.Caption = ""
    End If
    If CheckNumber = 78 Then
        Label78.Caption = ""
    End If
    If CheckNumber = 79 Then
        Label79.Caption = ""
    End If
    If CheckNumber = 80 Then
        Label80.Caption = ""
    End If
    If CheckNumber = 81 Then
        Label81.Caption = ""
    End If
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

End Sub


Private Sub Exit_Click()
    If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "G O B I N D A   N A N D I") = vbYes Then
        End
    End If
End Sub

Private Sub FinishTimerI()
    If finishI = 100 Then
        Label65.BackColor = vbRed
    ElseIf finishI < 100 Then
        finishI = 100
        Exit Sub
    End If
    
    If finishI = 200 Then
        Label64.BackColor = vbRed
    ElseIf finishI < 200 Then
        finishI = 200
        Exit Sub
    End If
    
    If finishI = 300 Then
        Label66.BackColor = vbRed
    ElseIf finishI < 300 Then
        finishI = 300
        Exit Sub
    End If
    
    If finishI = 400 Then
        Label56.BackColor = vbRed
    ElseIf finishI < 400 Then
        finishI = 400
        Exit Sub
    End If
    
    If finishI = 500 Then
        Label55.BackColor = vbRed
    ElseIf finishI < 500 Then
        finishI = 500
        Exit Sub
    End If
    
    If finishI = 600 Then
        Label57.BackColor = vbRed
    ElseIf finishI < 600 Then
        finishI = 600
        Exit Sub
    End If
    
    If finishI = 700 Then
        Label74.BackColor = vbRed
    ElseIf finishI < 700 Then
        finishI = 700
        Exit Sub
    End If
    
    If finishI = 800 Then
        Label73.BackColor = vbRed
    ElseIf finishI < 800 Then
        finishI = 800
        Exit Sub
    End If
    
    If finishI = 900 Then
        Label75.BackColor = vbRed
    ElseIf finishI < 900 Then
        finishI = 900
        Exit Sub
    End If
    
    If finishI = 1000 Then
        Label68.BackColor = vbRed
    ElseIf finishI < 1000 Then
        finishI = 1000
        Exit Sub
    End If
    
    If finishI = 1100 Then
        Label67.BackColor = vbRed
    ElseIf finishI < 1100 Then
        finishI = 1100
        Exit Sub
    End If
    
    If finishI = 1200 Then
        Label69.BackColor = vbRed
    ElseIf finishI < 1200 Then
        finishI = 1200
        Exit Sub
    End If
    
    If finishI = 1300 Then
        Label59.BackColor = vbRed
    ElseIf finishI < 1300 Then
        finishI = 1300
        Exit Sub
    End If
    
    If finishI = 1400 Then
        Label58.BackColor = vbRed
    ElseIf finishI < 1400 Then
        finishI = 1400
        Exit Sub
    End If
    
    If finishI = 1500 Then
        Label60.BackColor = vbRed
    ElseIf finishI < 1500 Then
        finishI = 1500
        Exit Sub
    End If
    
    If finishI = 1600 Then
        Label77.BackColor = vbRed
    ElseIf finishI < 1600 Then
        finishI = 1600
        Exit Sub
    End If
    
    If finishI = 1700 Then
        Label76.BackColor = vbRed
    ElseIf finishI < 1700 Then
        finishI = 1700
        Exit Sub
    End If
    
    If finishI = 1800 Then
        Label78.BackColor = vbRed
    ElseIf finishI < 1800 Then
        finishI = 1800
        Exit Sub
    End If
    
    If finishI = 1900 Then
        Label71.BackColor = vbRed
    ElseIf finishI < 1900 Then
        finishI = 1900
        Exit Sub
    End If
    
    If finishI = 2000 Then
        Label70.BackColor = vbRed
    ElseIf finishI < 2000 Then
        finishI = 2000
        Exit Sub
    End If
    
    If finishI = 2100 Then
        Label72.BackColor = vbRed
    ElseIf finishI < 2100 Then
        finishI = 2100
        Exit Sub
    End If
    
    If finishI = 2200 Then
        Label62.BackColor = vbRed
    ElseIf finishI < 2200 Then
        finishI = 2200
        Exit Sub
    End If
    
    If finishI = 2300 Then
        Label61.BackColor = vbRed
    ElseIf finishI < 2300 Then
        finishI = 2300
        Exit Sub
    End If
    
    If finishI = 2400 Then
        Label63.BackColor = vbRed
    ElseIf finishI < 2400 Then
        finishI = 2400
        Exit Sub
    End If
    
    If finishI = 2500 Then
        Label80.BackColor = vbRed
    ElseIf finishI < 2500 Then
        finishI = 2500
        Exit Sub
    End If
    
    If finishI = 2600 Then
        Label79.BackColor = vbRed
    ElseIf finishI < 2600 Then
        finishI = 2600
        Exit Sub
    End If
    
    If finishI = 2700 Then
        Label81.BackColor = vbRed
    ElseIf finishI < 2700 Then
        finishI = 2700
        Exit Sub
    End If
    
    If finishI = 2800 Then
        Label38.BackColor = vbRed
    ElseIf finishI < 2800 Then
        finishI = 2800
        Exit Sub
    End If
    
    If finishI = 2900 Then
        Label37.BackColor = vbRed
    ElseIf finishI < 2900 Then
        finishI = 2900
        Exit Sub
    End If
    
    If finishI = 3000 Then
        Label39.BackColor = vbRed
    ElseIf finishI < 3000 Then
        finishI = 3000
        Exit Sub
    End If
    
    If finishI = 3100 Then
        Label29.BackColor = vbRed
    ElseIf finishI < 3100 Then
        finishI = 3100
        Exit Sub
    End If
    
    If finishI = 3200 Then
        Label28.BackColor = vbRed
    ElseIf finishI < 3200 Then
        finishI = 3200
        Exit Sub
    End If
    
    If finishI = 3300 Then
        Label30.BackColor = vbRed
    ElseIf finishI < 3300 Then
        finishI = 3300
        Exit Sub
    End If
    
    If finishI = 3400 Then
        Label47.BackColor = vbRed
    ElseIf finishI < 3400 Then
        finishI = 3400
        Exit Sub
    End If
    
    If finishI = 3500 Then
        Label46.BackColor = vbRed
    ElseIf finishI < 3500 Then
        finishI = 3500
        Exit Sub
    End If
    
    If finishI = 3600 Then
        Label48.BackColor = vbRed
    ElseIf finishI < 3600 Then
        finishI = 3600
        Exit Sub
    End If
    
    If finishI = 3700 Then
        Label41.BackColor = vbRed
    ElseIf finishI < 3700 Then
        finishI = 3700
        Exit Sub
    End If
    
    If finishI = 3800 Then
        Label40.BackColor = vbRed
    ElseIf finishI < 3800 Then
        finishI = 3800
        Exit Sub
    End If
    
    If finishI = 3900 Then
        Label42.BackColor = vbRed
    ElseIf finishI < 3900 Then
        finishI = 3900
        Exit Sub
    End If
    
    If finishI = 4000 Then
        Label32.BackColor = vbRed
    ElseIf finishI < 4000 Then
        finishI = 4000
        Exit Sub
    End If
    
    If finishI = 4100 Then
        Label31.BackColor = vbRed
    ElseIf finishI < 4100 Then
        finishI = 4100
        Exit Sub
    End If
    
    If finishI = 4200 Then
        Label33.BackColor = vbRed
    ElseIf finishI < 4200 Then
        finishI = 4200
        Exit Sub
    End If
    
    If finishI = 4300 Then
        Label50.BackColor = vbRed
    ElseIf finishI < 4300 Then
        finishI = 4300
        Exit Sub
    End If
    
    If finishI = 4400 Then
        Label49.BackColor = vbRed
    ElseIf finishI < 4400 Then
        finishI = 4400
        Exit Sub
    End If
    
    If finishI = 4500 Then
        Label51.BackColor = vbRed
    ElseIf finishI < 4500 Then
        finishI = 4500
        Exit Sub
    End If
    
    If finishI = 4600 Then
        Label44.BackColor = vbRed
    ElseIf finishI < 4600 Then
        finishI = 4600
        Exit Sub
    End If
    
    If finishI = 4700 Then
        Label43.BackColor = vbRed
    ElseIf finishI < 4700 Then
        finishI = 4700
        Exit Sub
    End If
    
    If finishI = 4800 Then
        Label45.BackColor = vbRed
    ElseIf finishI < 4800 Then
        finishI = 4800
        Exit Sub
    End If
    
    If finishI = 4900 Then
        Label35.BackColor = vbRed
    ElseIf finishI < 4900 Then
        finishI = 4900
        Exit Sub
    End If
    
    If finishI = 5000 Then
        Label34.BackColor = vbRed
    ElseIf finishI < 5000 Then
        finishI = 5000
        Exit Sub
    End If
    
    If finishI = 5100 Then
        Label36.BackColor = vbRed
    ElseIf finishI < 5100 Then
        finishI = 5100
        Exit Sub
    End If
    
    If finishI = 5200 Then
        Label53.BackColor = vbRed
    ElseIf finishI < 5200 Then
        finishI = 5200
        Exit Sub
    End If
    
    If finishI = 5300 Then
        Label52.BackColor = vbRed
    ElseIf finishI < 5300 Then
        finishI = 5300
        Exit Sub
    End If
    
    If finishI = 5400 Then
        Label54.BackColor = vbRed
    ElseIf finishI < 5400 Then
        finishI = 5400
        Exit Sub
    End If
    
    If finishI = 5500 Then
        Label11.BackColor = vbRed
    ElseIf finishI < 5500 Then
        finishI = 5500
        Exit Sub
    End If
    
    If finishI = 5600 Then
        Label10.BackColor = vbRed
    ElseIf finishI < 5600 Then
        finishI = 5600
        Exit Sub
    End If
    
    If finishI = 5700 Then
        Label12.BackColor = vbRed
    ElseIf finishI < 5700 Then
        finishI = 5700
        Exit Sub
    End If
    
    If finishI = 5800 Then
        Label2.BackColor = vbRed
    ElseIf finishI < 5800 Then
        finishI = 5800
        Exit Sub
    End If
    
    If finishI = 5900 Then
        Label1.BackColor = vbRed
    ElseIf finishI < 5900 Then
        finishI = 5900
        Exit Sub
    End If
    
    If finishI = 6000 Then
        Label3.BackColor = vbRed
    ElseIf finishI < 6000 Then
        finishI = 6000
        Exit Sub
    End If
    
    If finishI = 6100 Then
        Label20.BackColor = vbRed
    ElseIf finishI < 6100 Then
        finishI = 6100
        Exit Sub
    End If
    
    If finishI = 6200 Then
        Label19.BackColor = vbRed
    ElseIf finishI < 6200 Then
        finishI = 6200
        Exit Sub
    End If
    
    If finishI = 6300 Then
        Label21.BackColor = vbRed
    ElseIf finishI < 6300 Then
        finishI = 6300
        Exit Sub
    End If
    
    If finishI = 6400 Then
        Label14.BackColor = vbRed
    ElseIf finishI < 6400 Then
        finishI = 6400
        Exit Sub
    End If
    
    If finishI = 6500 Then
        Label13.BackColor = vbRed
    ElseIf finishI < 6500 Then
        finishI = 6500
        Exit Sub
    End If
    If finishI = 6600 Then
        Label15.BackColor = vbRed
    ElseIf finishI < 6600 Then
        finishI = 6600
        Exit Sub
    End If
    
    If finishI = 6700 Then
        Label5.BackColor = vbRed
    ElseIf finishI < 6700 Then
        finishI = 6700
        Exit Sub
    End If
    
    If finishI = 6800 Then
        Label4.BackColor = vbRed
    ElseIf finishI < 6800 Then
        finishI = 6800
        Exit Sub
    End If
    
    If finishI = 6900 Then
        Label6.BackColor = vbRed
    ElseIf finishI < 6900 Then
        finishI = 6900
        Exit Sub
    End If
    
    If finishI = 7000 Then
        Label23.BackColor = vbRed
    ElseIf finishI < 7000 Then
        finishI = 7000
        Exit Sub
    End If
    
    If finishI = 7100 Then
        Label22.BackColor = vbRed
    ElseIf finishI < 7100 Then
        finishI = 7100
        Exit Sub
    End If
    
    If finishI = 7200 Then
        Label24.BackColor = vbRed
    ElseIf finishI < 7200 Then
        finishI = 7200
        Exit Sub
    End If
    
    If finishI = 7300 Then
        Label17.BackColor = vbRed
    ElseIf finishI < 7300 Then
        finishI = 7300
        Exit Sub
    End If
    
    If finishI = 7400 Then
        Label16.BackColor = vbRed
    ElseIf finishI < 7400 Then
        finishI = 7400
        Exit Sub
    End If
    
    If finishI = 7500 Then
        Label18.BackColor = vbRed
    ElseIf finishI < 7500 Then
        finishI = 7500
        Exit Sub
    End If
    
    If finishI = 7600 Then
        Label8.BackColor = vbRed
    ElseIf finishI < 7600 Then
        finishI = 7600
        Exit Sub
    End If
    
    If finishI = 7700 Then
        Label7.BackColor = vbRed
    ElseIf finishI < 7700 Then
        finishI = 7700
        Exit Sub
    End If
    
    If finishI = 7800 Then
        Label9.BackColor = vbRed
    ElseIf finishI < 7800 Then
        finishI = 7800
        Exit Sub
    End If
    
    If finishI = 7900 Then
        Label26.BackColor = vbRed
    ElseIf finishI < 7900 Then
        finishI = 7900
        Exit Sub
    End If
    
    If finishI = 8000 Then
        Label25.BackColor = vbRed
    ElseIf finishI < 8000 Then
        finishI = 8000
        Exit Sub
    End If
    
    If finishI = 8100 Then
        Label27.BackColor = vbRed
    ElseIf finishI < 8100 Then
        finishI = 8100
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    ComputerName
    disableB
    lblUserMsg.Visible = False
    FillBoxCount = 0
    txtgoogle.Visible = False
    InputNumber = 0
    CheckNumber = 0
    MsgBox "Welcome To Sudoku,Wish You Love It.", vbOKOnly + vbInformation, "Gobinda Nandi"
    lbltimer_Timer
    cmdstartString = 0
    txtgoogle.Visible = False
    txthelp = "Look at the size of your diagram. Each row, column and square on your diagram must have each number once but no more. For example, in a diagram measuring 9 by 9, each row, column and square must have the digits 1 through 9 once without duplication of any number."
    txthelp = txthelp + " Analyze the location of squares that already are filled in. If a box has a 5 and 9, you know that the box needs the seven other digits to be complete. Look at the diagram in terms of three sections, the top, middle and bottom boxes in rows."
    txthelp = txthelp + " Use a pencil to fill in possible numbers so you can erase them if you're wrong. It's easier to write 5 or 6 in a box than to remember that the number in that box needs to be one or the other. Use a pen to finalize your answer if you know a number is correct."
    txthelp = txthelp + "Continue filling in boxes until each row, column and box has all digits once and only once. Check your work along the way to verify you don't have duplicated numbers."
'    FinishTimer.Enabled = False
    bg.Enabled = False
    FinishB = False
    chkmusic.Visible = False
    wmpmusicGobinda.URL = ""
'    wmpmusicGobinda.settings.mute = True
' wmpmusicGobinda.Enabled = False
    ResetGrid
    startClick = False
    cmdlock.Visible = False
'    lockBar.Visible = False
    lblmode.Visible = False
    lbmistakes.Visible = False
End Sub

Private Sub Label1_Change()
Call cmdchk_Click
    If lblmusicSE(Label1) = True Then
    End If

End Sub

Private Sub Label1_Click()
    SHP.Left = 1560
    SHP.Top = 3000
    CheckNumber = 1
    PressNumber.SetFocus
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label1.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label10_Change()
Call cmdchk_Click
    If lblmusicSE(Label10) = True Then
    End If

End Sub

Private Sub chk10()
        If Label10.Caption <> "" Then
            X10 = Label10.Caption
        If Label18.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label1.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label3.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label37.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label37.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label40.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label40.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label43.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label43.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label64.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label64.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label67.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label67.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label70.Caption = X10 Then
            Label10.BackColor = vbWhite
            Label70.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub
Private Sub Label10_Click()
    SHP.Left = 480
    SHP.Top = 3000
    CheckNumber = 10
    PressNumber.SetFocus
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label10.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label11_Change()
Call cmdchk_Click
    If lblmusicSE(Label11) = True Then
    End If

End Sub

Private Sub chk11()
        If Label11.Caption <> "" Then
            X11 = Label11.Caption
        If Label18.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label10.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label14.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label14.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label19.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label19.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label20.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label20.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label21.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label21.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label1.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label1.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label2.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label2.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label38.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label38.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label41.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label41.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label44.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label44.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label65.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label65.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label68.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label68.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label71.Caption = X11 Then
            Label11.BackColor = vbWhite
            Label71.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub
Private Sub Label11_Click()
    SHP.Left = 120
    SHP.Top = 3000
    CheckNumber = 11
    PressNumber.SetFocus
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label11.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label12_Change()
Call cmdchk_Click
    If lblmusicSE(Label12) = True Then
    End If

End Sub

Private Sub chk12()
        If Label12.Caption <> "" Then
            X12 = Label12.Caption
            If Label18.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label19.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label19.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label20.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label20.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label21.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label21.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = X12 Then
                Label12.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub
Private Sub Label12_Click()
    SHP.Left = 840
    SHP.Top = 3000
    CheckNumber = 12
    PressNumber.SetFocus
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label12.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label13_Change()
Call cmdchk_Click
    If lblmusicSE(Label13) = True Then
    End If

End Sub

Private Sub Label13_Click()
    SHP.Left = 480
    SHP.Top = 3360
    CheckNumber = 13
    PressNumber.SetFocus
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label13.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label14_Change()
Call cmdchk_Click
    If lblmusicSE(Label14) = True Then
    End If

End Sub

Private Sub chk14()
        If Label14.Caption <> "" Then
        X14 = Label14.Caption
        If Label18.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label18.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label16.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label15.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label15.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label13.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label13.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label17.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label17.BackColor = vbWhite
            CheckError = "1"
        End If
         If Label10.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label10.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label11.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label11.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label12.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label12.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label22.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label22.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label23.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label23.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label24.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label24.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label4.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label4.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label5.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label5.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label6.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label16.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label38.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label38.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label41.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label41.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label44.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label44.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label65.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label65.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label68.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label68.BackColor = vbWhite
            CheckError = "1"
        End If
        If Label71.Caption = X14 Then
            Label14.BackColor = vbWhite
            Label71.BackColor = vbWhite
            CheckError = "1"
        End If
    End If
End Sub
Private Sub Label14_Click()
    SHP.Left = 120
    SHP.Top = 3360
    CheckNumber = 14
    PressNumber.SetFocus
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label14.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label15_Change()
Call cmdchk_Click
    If lblmusicSE(Label15) = True Then
    End If

End Sub

Private Sub Label15_Click()
    SHP.Left = 840
    SHP.Top = 3360
    CheckNumber = 15
    PressNumber.SetFocus
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label15.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label16_Change()
Call cmdchk_Click
    If lblmusicSE(Label16) = True Then
    End If

End Sub

Private Sub Label16_Click()
    SHP.Left = 480
    SHP.Top = 3720
    CheckNumber = 16
    PressNumber.SetFocus
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label16.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label17_Change()
Call cmdchk_Click
    If lblmusicSE(Label17) = True Then
    End If

End Sub

Private Sub Label17_Click()
    SHP.Left = 120
    SHP.Top = 3720
    CheckNumber = 17
    PressNumber.SetFocus
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label17.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label18_Change()
Call cmdchk_Click
    If lblmusicSE(Label18) = True Then
    End If

End Sub

Private Sub Label18_Click()
    SHP.Left = 840
    SHP.Top = 3720
    CheckNumber = 18
    PressNumber.SetFocus
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label18.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label19_Change()
Call cmdchk_Click
    If lblmusicSE(Label19) = True Then
    End If

End Sub

Private Sub Label19_Click()
    SHP.Left = 2640
    SHP.Top = 3000
    CheckNumber = 19
    PressNumber.SetFocus
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label19.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label2_Change()
Call cmdchk_Click
    If lblmusicSE(Label2) = True Then
    End If

End Sub

Private Sub Label2_Click()
    SHP.Left = 1200
    SHP.Top = 3000
    CheckNumber = 2
    PressNumber.SetFocus
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label2.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label20_Change()
Call cmdchk_Click
    If lblmusicSE(Label20) = True Then
    End If

End Sub

Private Sub Label20_Click()
    SHP.Left = 2280
    SHP.Top = 3000
    CheckNumber = 20
    PressNumber.SetFocus
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label20.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label21_Change()
Call cmdchk_Click
    If lblmusicSE(Label21) = True Then
    End If

End Sub

Private Sub Label21_Click()
    CheckNumber = 21
    SHP.Left = 3000
    SHP.Top = 3000
    PressNumber.SetFocus
'    Call cmdchk_Click
End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label21.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label22_Change()
Call cmdchk_Click
    If lblmusicSE(Label22) = True Then
    End If

End Sub

Private Sub Label22_Click()
    SHP.Left = 2640
    SHP.Top = 3360
    CheckNumber = 22
    PressNumber.SetFocus
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label22.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label23_Change()
Call cmdchk_Click
    If lblmusicSE(Label23) = True Then
    End If

End Sub

Private Sub Label23_Click()
    SHP.Left = 2280
    SHP.Top = 3360
    CheckNumber = 23
    PressNumber.SetFocus
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label23.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label24_Change()
Call cmdchk_Click
    If lblmusicSE(Label24) = True Then
    End If

End Sub

Private Sub Label24_Click()
    SHP.Left = 3000
    SHP.Top = 3360
    CheckNumber = 24
    PressNumber.SetFocus
End Sub


Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label24.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label25_Change()
Call cmdchk_Click
    If lblmusicSE(Label25) = True Then
    End If

End Sub

Private Sub Label25_Click()
    SHP.Left = 2640
    SHP.Top = 3720
    CheckNumber = 25
    PressNumber.SetFocus
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label25.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label26_Change()
Call cmdchk_Click
    If lblmusicSE(Label26) = True Then
    End If

End Sub

Private Sub Label26_Click()
    SHP.Left = 2280
    SHP.Top = 3720
    CheckNumber = 26
    PressNumber.SetFocus
End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label26.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label27_Change()
Call cmdchk_Click
    If lblmusicSE(Label27) = True Then
    End If

End Sub

Private Sub Label27_Click()
    SHP.Left = 3000
    SHP.Top = 3720
    CheckNumber = 27
    PressNumber.SetFocus
End Sub


Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label27.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label28_Change()
Call cmdchk_Click
    If lblmusicSE(Label28) = True Then
    End If

End Sub

Private Sub Label28_Click()
    SHP.Left = 1560
    SHP.Top = 1920
    CheckNumber = 28
    PressNumber.SetFocus
End Sub

Private Sub Label28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label28.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label29_Change()
Call cmdchk_Click
    If lblmusicSE(Label29) = True Then
    End If

End Sub

Private Sub Label29_Click()
    SHP.Left = 1200
    SHP.Top = 1920
    CheckNumber = 29
    PressNumber.SetFocus
End Sub

Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label29.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label3_Change()
Call cmdchk_Click
    If lblmusicSE(Label3) = True Then
    End If

End Sub

Private Sub Label3_Click()
    SHP.Left = 1920
    SHP.Top = 3000
    CheckNumber = 3
    PressNumber.SetFocus
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label3.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label30_Change()
Call cmdchk_Click
    If lblmusicSE(Label30) = True Then
    End If

End Sub

Private Sub Label30_Click()
    SHP.Left = 1920
    SHP.Top = 1920
    CheckNumber = 30
    PressNumber.SetFocus
End Sub

Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label30.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label31_Change()
Call cmdchk_Click
    If lblmusicSE(Label31) = True Then
    End If

End Sub

Private Sub Label31_Click()
    SHP.Left = 1560
    SHP.Top = 2280
    CheckNumber = 31
    PressNumber.SetFocus
End Sub

Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label31.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label32_Change()
Call cmdchk_Click
    If lblmusicSE(Label32) = True Then
    End If

End Sub

Private Sub Label32_Click()
    SHP.Left = 1200
    SHP.Top = 2280
    CheckNumber = 32
    PressNumber.SetFocus
End Sub

Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label32.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label33_Change()
Call cmdchk_Click
    If lblmusicSE(Label33) = True Then
    End If

End Sub

Private Sub Label33_Click()
    SHP.Left = 1920
    SHP.Top = 2280
    CheckNumber = 33
    PressNumber.SetFocus
End Sub


Private Sub Label33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label33.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label34_Change()
Call cmdchk_Click
    If lblmusicSE(Label34) = True Then
    End If

End Sub

Private Sub Label34_Click()
    SHP.Left = 1560
    SHP.Top = 2640
    CheckNumber = 34
    PressNumber.SetFocus
End Sub

Private Sub Label34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label34.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label35_Change()
Call cmdchk_Click
    If lblmusicSE(Label35) = True Then
    End If

End Sub

Private Sub Label35_Click()
    SHP.Left = 1200
    SHP.Top = 2640
    CheckNumber = 35
    PressNumber.SetFocus
End Sub

Private Sub Label35_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label35.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label36_Change()
Call cmdchk_Click
    If lblmusicSE(Label36) = True Then
    End If

End Sub

Private Sub Label36_Click()
    SHP.Left = 1920
    SHP.Top = 2640
    CheckNumber = 36
    PressNumber.SetFocus
End Sub


Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label36.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label37_Change()
Call cmdchk_Click
    If lblmusicSE(Label37) = True Then
    End If

End Sub

Private Sub Label37_Click()
    SHP.Left = 480
    SHP.Top = 1920
    CheckNumber = 37
    PressNumber.SetFocus
End Sub

Private Sub Label37_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label37.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label38_Change()
Call cmdchk_Click
    If lblmusicSE(Label38) = True Then
    End If

End Sub

Private Sub Label38_Click()
    SHP.Left = 120
    SHP.Top = 1920
    CheckNumber = 38
    PressNumber.SetFocus
End Sub

Private Sub Label38_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label38.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label39_Change()
Call cmdchk_Click
    If lblmusicSE(Label39) = True Then
    End If

End Sub

Private Sub Label39_Click()
    SHP.Left = 840
    SHP.Top = 1920
    CheckNumber = 39
    PressNumber.SetFocus
End Sub

Private Sub Label39_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label39.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label4_Change()
Call cmdchk_Click
    If lblmusicSE(Label4) = True Then
    End If

End Sub

Private Sub Label4_Click()
    SHP.Left = 1560
    SHP.Top = 3360
    CheckNumber = 4
    PressNumber.SetFocus
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label4.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label40_Change()
Call cmdchk_Click
    If lblmusicSE(Label40) = True Then
    End If

End Sub

Private Sub Label40_Click()
    SHP.Left = 480
    SHP.Top = 2280
    CheckNumber = 40
    PressNumber.SetFocus
End Sub

Private Sub Label40_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label40.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label41_Change()
Call cmdchk_Click
    If lblmusicSE(Label41) = True Then
    End If

End Sub

Private Sub Label41_Click()
    SHP.Left = 120
    SHP.Top = 2280
    CheckNumber = 41
    PressNumber.SetFocus
End Sub

Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label41.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label42_Change()
Call cmdchk_Click
    If lblmusicSE(Label42) = True Then
    End If

End Sub

Private Sub Label42_Click()
    SHP.Left = 840
    SHP.Top = 2280
    CheckNumber = 42
    PressNumber.SetFocus
End Sub

Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label42.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label43_Change()
Call cmdchk_Click
    If lblmusicSE(Label43) = True Then
    End If

End Sub

Private Sub Label43_Click()
    SHP.Left = 480
    SHP.Top = 2640
    CheckNumber = 43
    PressNumber.SetFocus
End Sub

Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label43.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label44_Change()
Call cmdchk_Click
    If lblmusicSE(Label44) = True Then
    End If

End Sub

Private Sub Label44_Click()
    SHP.Left = 120
    SHP.Top = 2640
    CheckNumber = 44
    PressNumber.SetFocus
End Sub

Private Sub Label44_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label44.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label45_Change()
Call cmdchk_Click
    If lblmusicSE(Label45) = True Then
    End If

End Sub

Private Sub Label45_Click()
    SHP.Left = 840
    SHP.Top = 2640
    CheckNumber = 45
    PressNumber.SetFocus
End Sub

Private Sub Label45_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label45.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label46_Change()
Call cmdchk_Click
    If lblmusicSE(Label46) = True Then
    End If

End Sub

Private Sub Label46_Click()
    SHP.Left = 2640
    SHP.Top = 1920
    CheckNumber = 46
    PressNumber.SetFocus
End Sub

Private Sub Label46_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label46.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label47_Change()
Call cmdchk_Click
    If lblmusicSE(Label47) = True Then
    End If

End Sub

Private Sub Label47_Click()
    SHP.Left = 2280
    SHP.Top = 1920
    CheckNumber = 47
    PressNumber.SetFocus
End Sub

Private Sub Label47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label47.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label48_Change()
Call cmdchk_Click
    If lblmusicSE(Label48) = True Then
    End If

End Sub

Private Sub Label48_Click()
    SHP.Left = 3000
    SHP.Top = 1920
    CheckNumber = 48
    PressNumber.SetFocus
End Sub

Private Sub Label48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label48.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label49_Change()
Call cmdchk_Click
    If lblmusicSE(Label49) = True Then
    End If

End Sub

Private Sub Label49_Click()
    SHP.Left = 2640
    SHP.Top = 2280
    CheckNumber = 49
    PressNumber.SetFocus
End Sub

Private Sub Label49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label49.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label5_Change()
Call cmdchk_Click
    If lblmusicSE(Label5) = True Then
    End If

End Sub

Private Sub Label5_Click()
    SHP.Left = 1200
    SHP.Top = 3360
    CheckNumber = 5
    PressNumber.SetFocus
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label5.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label50_Change()
Call cmdchk_Click
    If lblmusicSE(Label50) = True Then
    End If

End Sub

Private Sub Label50_Click()
    SHP.Left = 2280
    SHP.Top = 2280
    CheckNumber = 50
    PressNumber.SetFocus
End Sub

Private Sub Label50_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label50.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label51_Change()
Call cmdchk_Click
    If lblmusicSE(Label51) = True Then
    End If

End Sub

Private Sub Label51_Click()
    SHP.Left = 3000
    SHP.Top = 2280
    CheckNumber = 51
    PressNumber.SetFocus
End Sub


Private Sub Label51_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label51.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label52_Change()
Call cmdchk_Click
    If lblmusicSE(Label52) = True Then
    End If

End Sub

Private Sub Label52_Click()
    SHP.Left = 2640
    SHP.Top = 2640
    CheckNumber = 52
    PressNumber.SetFocus
End Sub

Private Sub Label52_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label52.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label53_Change()
Call cmdchk_Click
    If lblmusicSE(Label53) = True Then
    End If

End Sub

Private Sub Label53_Click()
    SHP.Left = 2280
    SHP.Top = 2640
    CheckNumber = 53
    PressNumber.SetFocus
End Sub

Private Sub Label53_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label53.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label54_Change()
Call cmdchk_Click
    If lblmusicSE(Label54) = True Then
    End If

End Sub

Private Sub Label54_Click()
    SHP.Left = 3000
    SHP.Top = 2640
    CheckNumber = 54
    PressNumber.SetFocus
End Sub


Private Sub Label54_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label54.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label55_Change()
Call cmdchk_Click
    If lblmusicSE(Label55) = True Then
    End If

'lblMsgBox
End Sub

Private Sub Label55_Click()
    SHP.Left = 1560
    SHP.Top = 840
    CheckNumber = 55
    PressNumber.SetFocus
End Sub

Private Sub Label55_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label55.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label56_Change()
Call cmdchk_Click
    If lblmusicSE(Label56) = True Then
    End If

'lblMsgBox
End Sub

Private Sub Label56_Click()
    SHP.Left = 1200
    SHP.Top = 840
    CheckNumber = 56
    PressNumber.SetFocus
End Sub

Private Sub Label56_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label56.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label57_Change()
    Call cmdchk_Click
    If lblmusicSE(Label57) = True Then
    End If
    
    'lblMsgBox
End Sub

Private Sub Label57_Click()
    SHP.Left = 1920
    SHP.Top = 840
    CheckNumber = 57
    PressNumber.SetFocus
End Sub

Private Sub Label57_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label57.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label58_Change()
Call cmdchk_Click
    If lblmusicSE(Label58) = True Then
    End If

End Sub

Private Sub Label58_Click()
    SHP.Left = 1560
    SHP.Top = 1200
    CheckNumber = 58
    PressNumber.SetFocus
End Sub

Private Sub Label58_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label58.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label59_Change()
Call cmdchk_Click
    If lblmusicSE(Label59) = True Then
    End If

End Sub

Private Sub Label59_Click()
    SHP.Left = 1200
    SHP.Top = 1200
    CheckNumber = 59
    PressNumber.SetFocus
End Sub

Private Sub Label59_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label59.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label6_Change()
Call cmdchk_Click
    If lblmusicSE(Label6) = True Then
    End If

End Sub

Private Sub Label6_Click()
    SHP.Left = 1920
    SHP.Top = 3360
    CheckNumber = 6
    PressNumber.SetFocus
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label6.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label60_Change()
Call cmdchk_Click
    If lblmusicSE(Label60) = True Then
    End If

End Sub

Private Sub Label60_Click()
    SHP.Left = 1920
    SHP.Top = 1200
    CheckNumber = 60
    PressNumber.SetFocus
End Sub

Private Sub Label60_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label60.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label61_Change()
Call cmdchk_Click
    If lblmusicSE(Label61) = True Then
    End If

End Sub

Private Sub Label61_Click()
    SHP.Left = 1560
    SHP.Top = 1560
    CheckNumber = 61
    PressNumber.SetFocus
End Sub

Private Sub Label61_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label61.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label62_Change()
Call cmdchk_Click
    If lblmusicSE(Label62) = True Then
    End If

End Sub

Private Sub Label62_Click()
    SHP.Left = 1200
    SHP.Top = 1560
    CheckNumber = 62
    PressNumber.SetFocus
End Sub

Private Sub Label62_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label62.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label63_Change()
Call cmdchk_Click
    If lblmusicSE(Label63) = True Then
    End If

End Sub

Private Sub Label63_Click()
    SHP.Left = 1920
    SHP.Top = 1560
    CheckNumber = 63
    PressNumber.SetFocus
End Sub


Private Sub Label63_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label63.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label64_Change()
Call cmdchk_Click
    If lblmusicSE(Label64) = True Then
    End If

'lblMsgBox
End Sub

Private Sub Label64_Click()
    SHP.Left = 480
    SHP.Top = 840
    CheckNumber = 64
    PressNumber.SetFocus
End Sub

Private Sub Label64_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label64.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label65_Change()
Call cmdchk_Click
    If lblmusicSE(Label65) = True Then
    End If

'lblMsgBox
End Sub

Private Sub Label65_Click()
    SHP.Left = 120
    SHP.Top = 840
    CheckNumber = 65
    PressNumber.SetFocus
End Sub

Private Sub Label65_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label65.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label66_Change()
Call cmdchk_Click
    If lblmusicSE(Label66) = True Then
    End If
    
'lblMsgBox
End Sub

Private Sub Label66_Click()
    SHP.Left = 840
    SHP.Top = 840
    CheckNumber = 66
    PressNumber.SetFocus
End Sub

Private Sub Label66_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label66.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label67_Change()
Call cmdchk_Click
    If lblmusicSE(Label67) = True Then
    End If

End Sub

Private Sub Label67_Click()
    SHP.Left = 480
    SHP.Top = 1200
    CheckNumber = 67
    PressNumber.SetFocus
End Sub

Private Sub Label67_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label67.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label68_Change()
Call cmdchk_Click
    If lblmusicSE(Label68) = True Then
    End If

End Sub

Private Sub Label68_Click()
    SHP.Left = 120
    SHP.Top = 1200
    CheckNumber = 68
    PressNumber.SetFocus
End Sub

Private Sub Label68_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label68.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label69_Change()
Call cmdchk_Click
    If lblmusicSE(Label69) = True Then
    End If

End Sub

Private Sub Label69_Click()
    SHP.Left = 840
    SHP.Top = 1200
    CheckNumber = 69
    PressNumber.SetFocus
End Sub

Private Sub Label69_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label69.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label7_Change()
Call cmdchk_Click
    If lblmusicSE(Label7) = True Then
    End If

End Sub

Private Sub Label7_Click()
    SHP.Left = 1560
    SHP.Top = 3720
    CheckNumber = 7
    PressNumber.SetFocus
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label7.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label70_Change()
Call cmdchk_Click
    If lblmusicSE(Label70) = True Then
    End If

End Sub

Private Sub Label70_Click()
    SHP.Left = 480
    SHP.Top = 1560
    CheckNumber = 70
    PressNumber.SetFocus
End Sub

Private Sub Label70_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label70.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label71_Change()
Call cmdchk_Click
    If lblmusicSE(Label71) = True Then
    End If

End Sub

Private Sub Label71_Click()
    SHP.Left = 120
    SHP.Top = 1560
    CheckNumber = 71
    PressNumber.SetFocus
End Sub

Private Sub Label71_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label71.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label72_Change()
Call cmdchk_Click
    If lblmusicSE(Label72) = True Then
    End If

End Sub

Private Sub Label72_Click()
    SHP.Left = 840
    SHP.Top = 1560
    CheckNumber = 72
    PressNumber.SetFocus
End Sub

Private Sub Label72_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label72.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label73_Change()
Call cmdchk_Click
    If lblmusicSE(Label73) = True Then
    End If

'lblMsgBox
End Sub

Private Sub Label73_Click()
    SHP.Left = 2640
    SHP.Top = 840
    CheckNumber = 73
    PressNumber.SetFocus
End Sub

Private Sub Label73_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label73.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label74_Change()
    Call cmdchk_Click
    If lblmusicSE(Label74) = True Then
    End If
    
    'lblMsgBox
End Sub

Private Sub Label74_Click()
    SHP.Left = 2280
    SHP.Top = 840
    CheckNumber = 74
    PressNumber.SetFocus
End Sub

Private Sub Label74_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label74.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label75_Change()
Call cmdchk_Click
    If lblmusicSE(Label75) = True Then
    End If
End Sub

Private Sub Label75_Click()
    SHP.Left = 3000
    SHP.Top = 840
    CheckNumber = 75
    PressNumber.SetFocus
End Sub

Private Sub Label75_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label75.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label76_Change()
Call cmdchk_Click
    If lblmusicSE(Label76) = True Then
    End If

End Sub

Private Sub Label76_Click()
    SHP.Left = 2640
    SHP.Top = 1200
    CheckNumber = 76
    PressNumber.SetFocus
End Sub

Private Sub Label76_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label76.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label77_Change()
Call cmdchk_Click
    If lblmusicSE(Label77) = True Then
    End If

End Sub

Private Sub Label77_Click()
    SHP.Left = 2280
    SHP.Top = 1200
    CheckNumber = 77
    PressNumber.SetFocus
End Sub

Private Sub Label77_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label77.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label78_Change()
Call cmdchk_Click
    If lblmusicSE(Label78) = True Then
    End If

End Sub

Private Sub Label78_Click()
    SHP.Left = 3000
    SHP.Top = 1200
    CheckNumber = 78
    PressNumber.SetFocus
End Sub


Private Sub Label78_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label78.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label79_Change()
Call cmdchk_Click
    If lblmusicSE(Label79) = True Then
    End If

End Sub

Private Sub Label79_Click()
    SHP.Left = 2640
    SHP.Top = 1560
    CheckNumber = 79
    PressNumber.SetFocus
End Sub

Private Sub Label79_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label79.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label8_Change()
Call cmdchk_Click
    If lblmusicSE(Label8) = True Then
    End If

End Sub

Private Sub Label8_Click()
    SHP.Left = 1200
    SHP.Top = 3720
    CheckNumber = 8
    PressNumber.SetFocus
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label8.BackColor = &HAEDEF4
    End If
End Sub

Private Sub Label80_Change()
Call cmdchk_Click
    If lblmusicSE(Label80) = True Then
    End If

End Sub

Private Sub Label80_Click()
    SHP.Left = 2280
    SHP.Top = 1560
    CheckNumber = 80
    PressNumber.SetFocus
End Sub

Private Sub Label80_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label80.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label81_Change()
Call cmdchk_Click
    If lblmusicSE(Label81) = True Then
    End If

End Sub

Private Sub Label81_Click()
    SHP.Left = 3000
    SHP.Top = 1560
    CheckNumber = 81
    PressNumber.SetFocus
End Sub

Private Sub Label81_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label81.BackColor = &H9FFF9D
    End If
End Sub

Private Sub Label83_Click()
Shell "explorer.exe http://en.wikipedia.org/wiki/User:Gobinda_Nandi"
End Sub

Private Sub Label9_Change()
Call cmdchk_Click
    If lblmusicSE(Label9) = True Then
    End If

End Sub

Private Sub Label9_Click()
    SHP.Left = 1920
    SHP.Top = 3720
    CheckNumber = 9
    PressNumber.SetFocus
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If startClick = True And CheckError <> "1" Then
        NORMAL
        Label9.BackColor = &HAEDEF4
    End If
End Sub

Private Sub lblDateTimer_Timer()
       lbltime.Caption = Now
End Sub

Private Sub lbltimer_Timer()
'    SHP.Left = Rnd * 1089
'    SHP.Top = Rnd * 1089
    'Label82.Left = Rnd * 1000
'    Label82.Top = Rnd * 6211

''    Label64.Caption = Right(Rnd, 1)        'first box
''    Label65.Caption = Right(Rnd, 1)
''    Label66.Caption = Right(Rnd, 1)
''    Label67.Caption = Right(Rnd, 1)
''    Label68.Caption = Right(Rnd, 1)
''    Label69.Caption = Right(Rnd, 1)
''    Label70.Caption = Right(Rnd, 1)
''    Label71.Caption = Right(Rnd, 1)
''    Label72.Caption = Right(Rnd, 1)
'
'    Label55.Caption = Right(Rnd, 1)         'second box
'    Label56.Caption = Right(Rnd, 1)
'    Label57.Caption = Right(Rnd, 1)
'    Label58.Caption = Right(Rnd, 1)
'    Label59.Caption = Right(Rnd, 1)
'    Label60.Caption = Right(Rnd, 1)
'    Label61.Caption = Right(Rnd, 1)
'    Label62.Caption = Right(Rnd, 1)
'    Label63.Caption = Right(Rnd, 1)
'
''    Label73.Caption = Right(Rnd, 1)        'third box
''    Label74.Caption = Right(Rnd, 1)
''    Label75.Caption = Right(Rnd, 1)
''    Label76.Caption = Right(Rnd, 1)
''    Label77.Caption = Right(Rnd, 1)
''    Label78.Caption = Right(Rnd, 1)
''    Label79.Caption = Right(Rnd, 1)
''    Label80.Caption = Right(Rnd, 1)
''    Label81.Caption = Right(Rnd, 1)
'
'    Label37.Caption = Right(Rnd, 1)            'fourth box
'    Label38.Caption = Right(Rnd, 1)
'    Label39.Caption = Right(Rnd, 1)
'    Label40.Caption = Right(Rnd, 1)
'    Label41.Caption = Right(Rnd, 1)
'    Label42.Caption = Right(Rnd, 1)
'    Label43.Caption = Right(Rnd, 1)
'    Label44.Caption = Right(Rnd, 1)
'    Label45.Caption = Right(Rnd, 1)
'
    Label28.Caption = Right(Rnd, 1)        'fifth box
    Label29.Caption = Right(Rnd, 1)
    Label30.Caption = Right(Rnd, 1)
    Label31.Caption = Right(Rnd, 1)
    Label32.Caption = Right(Rnd, 1)
    Label33.Caption = Right(Rnd, 1)
    Label34.Caption = Right(Rnd, 1)
    Label35.Caption = Right(Rnd, 1)
    Label36.Caption = Right(Rnd, 1)
'
'    Label46.Caption = Right(Rnd, 1)           'sixth box
'    Label47.Caption = Right(Rnd, 1)
'    Label48.Caption = Right(Rnd, 1)
'    Label49.Caption = Right(Rnd, 1)
'    Label50.Caption = Right(Rnd, 1)
'    Label51.Caption = Right(Rnd, 1)
'    Label52.Caption = Right(Rnd, 1)
'    Label53.Caption = Right(Rnd, 1)
'    Label54.Caption = Right(Rnd, 1)
'
''    Label10.Caption = Right(Rnd, 1)        'seventh box
''    Label11.Caption = Right(Rnd, 1)
''    Label12.Caption = Right(Rnd, 1)
''    Label13.Caption = Right(Rnd, 1)
''    Label14.Caption = Right(Rnd, 1)
''    Label15.Caption = Right(Rnd, 1)
''    Label16.Caption = Right(Rnd, 1)
''    Label17.Caption = Right(Rnd, 1)
''    Label18.Caption = Right(Rnd, 1)
'
'    Label1.Caption = Right(Rnd, 1)            'eighth box
'    Label2.Caption = Right(Rnd, 1)
'    Label3.Caption = Right(Rnd, 1)
'    Label4.Caption = Right(Rnd, 1)
'    Label5.Caption = Right(Rnd, 1)
'    Label6.Caption = Right(Rnd, 1)
'    Label7.Caption = Right(Rnd, 1)
'    Label8.Caption = Right(Rnd, 1)
'    Label9.Caption = Right(Rnd, 1)
'
''    Label19.Caption = Right(Rnd, 1)             'ninth box
''    Label20.Caption = Right(Rnd, 1)
''    Label21.Caption = Right(Rnd, 1)
''    Label22.Caption = Right(Rnd, 1)
''    Label23.Caption = Right(Rnd, 1)
''    Label24.Caption = Right(Rnd, 1)
''    Label25.Caption = Right(Rnd, 1)
''    Label26.Caption = Right(Rnd, 1)
''    Label27.Caption = Right(Rnd, 1)

'****************************************************************************

    Label64.BackColor = Rnd * vbGreen          'first box
    Label65.BackColor = Rnd * vbWhite
    Label66.BackColor = Rnd * vbRed
    Label67.BackColor = Rnd * vbYellow
    Label68.BackColor = Rnd * vbBlack
    Label69.BackColor = Rnd * vbRed
    Label70.BackColor = Rnd * vbBlue
    Label71.BackColor = Rnd * vbYellow
    Label72.BackColor = Rnd * vbGreen
    
    Label55.BackColor = Rnd * vbYellow             'second box
    Label56.BackColor = Rnd * vbYellow
    Label57.BackColor = Rnd * vbWhite
    Label58.BackColor = Rnd * vbYellow
    Label59.BackColor = Rnd * vbBlue
    Label60.BackColor = Rnd * vbYellow
    Label61.BackColor = Rnd * vbYellow
    Label62.BackColor = Rnd * vbRed
    Label63.BackColor = Rnd * vbBlue
    
    Label73.BackColor = Rnd * vbGreen          'third box
    Label74.BackColor = Rnd * vbGreen
    Label75.BackColor = Rnd * vbWhite
    Label76.BackColor = Rnd * vbGreen
    Label77.BackColor = Rnd * vbBlue
    Label78.BackColor = Rnd * vbBlack
    Label79.BackColor = Rnd * vbRed
    Label80.BackColor = Rnd * vbBlue
    Label81.BackColor = Rnd * vbGreen
            
    Label37.BackColor = Rnd * vbRed             'fourth box
    Label38.BackColor = Rnd * vbYellow
    Label39.BackColor = Rnd * vbBlue
    Label40.BackColor = Rnd * vbYellow
    Label41.BackColor = Rnd * vbRed
    Label42.BackColor = Rnd * vbYellow
    Label43.BackColor = Rnd * vbWhite
    Label44.BackColor = Rnd * vbYellow
    Label45.BackColor = Rnd * vbBlue
      
'    Label28.BackColor = Rnd * vbGreen          'fifth box
'    Label29.BackColor = Rnd * vbGreen
'    Label30.BackColor = Rnd * vbBlue
'    Label31.BackColor = Rnd * vbBlack
'    Label32.BackColor = Rnd * vbWhite
'    Label33.BackColor = Rnd * vbBlue
'    Label34.BackColor = Rnd * vbGreen
'    Label35.BackColor = Rnd * vbRed
'    Label36.BackColor = Rnd * vbGreen
    
    Label46.BackColor = Rnd * vbYellow             'sixth box
    Label47.BackColor = Rnd * vbBlue
    Label48.BackColor = Rnd * vbYellow
    Label49.BackColor = Rnd * vbWhite
    Label50.BackColor = Rnd * vbYellow
    Label51.BackColor = Rnd * vbRed
    Label52.BackColor = Rnd * vbWhite
    Label53.BackColor = Rnd * vbBlue
    Label54.BackColor = Rnd * vbYellow
    
    Label10.BackColor = Rnd * vbBlue          'seventh box
    Label11.BackColor = Rnd * vbBlack
    Label12.BackColor = Rnd * vbGreen
    Label13.BackColor = Rnd * vbBlue
    Label14.BackColor = Rnd * vbBlack
    Label15.BackColor = Rnd * vbGreen
    Label16.BackColor = Rnd * vbRed
    Label17.BackColor = Rnd * vbWhite
    Label18.BackColor = Rnd * vbGreen
    
    Label1.BackColor = Rnd * vbWhite             'eighth box
    Label2.BackColor = Rnd * vbYellow
    Label3.BackColor = Rnd * vbRed
    Label4.BackColor = Rnd * vbBlue
    Label5.BackColor = Rnd * vbWhite
    Label6.BackColor = Rnd * vbYellow
    Label7.BackColor = Rnd * vbBlack
    Label8.BackColor = Rnd * vbBlue
    Label9.BackColor = Rnd * vbRed
    
    Label19.BackColor = Rnd * vbGreen             'ninth box
    Label20.BackColor = Rnd * vbGreen
    Label21.BackColor = Rnd * vbGreen
    Label22.BackColor = Rnd * vbBlue
    Label23.BackColor = Rnd * vbGreen
    Label24.BackColor = Rnd * vbBlack
    Label25.BackColor = Rnd * vbWhite
    Label26.BackColor = Rnd * vbGreen
    Label27.BackColor = Rnd * vbBlue
End Sub
Private Sub NORMAL()
    Label64.BackColor = vbGreen         'first box
    Label65.BackColor = vbGreen
    Label66.BackColor = vbGreen
    Label67.BackColor = vbGreen
    Label68.BackColor = vbGreen
    Label69.BackColor = vbGreen
    Label70.BackColor = vbGreen
    Label71.BackColor = vbGreen
    Label72.BackColor = vbGreen

    
    Label73.BackColor = vbGreen         'third box
    Label74.BackColor = vbGreen
    Label75.BackColor = vbGreen
    Label76.BackColor = vbGreen
    Label77.BackColor = vbGreen
    Label78.BackColor = vbGreen
    Label79.BackColor = vbGreen
    Label80.BackColor = vbGreen
    Label81.BackColor = vbGreen
            
  
      
    Label28.BackColor = vbGreen         'fifth box
    Label29.BackColor = vbGreen
    Label30.BackColor = vbGreen
    Label31.BackColor = vbGreen
    Label32.BackColor = vbGreen
    Label33.BackColor = vbGreen
    Label34.BackColor = vbGreen
    Label35.BackColor = vbGreen
    Label36.BackColor = vbGreen
 
    
    Label10.BackColor = vbGreen         'seventh box
    Label11.BackColor = vbGreen
    Label12.BackColor = vbGreen
    Label13.BackColor = vbGreen
    Label14.BackColor = vbGreen
    Label15.BackColor = vbGreen
    Label16.BackColor = vbGreen
    Label17.BackColor = vbGreen
    Label18.BackColor = vbGreen
 
    
    Label19.BackColor = vbGreen            'ninth box
    Label20.BackColor = vbGreen
    Label21.BackColor = vbGreen
    Label22.BackColor = vbGreen
    Label23.BackColor = vbGreen
    Label24.BackColor = vbGreen
    Label25.BackColor = vbGreen
    Label26.BackColor = vbGreen
    Label27.BackColor = vbGreen
    
    
    
    Label55.BackColor = vbYellow             'second box
    Label56.BackColor = vbYellow
    Label57.BackColor = vbYellow
    Label58.BackColor = vbYellow
    Label59.BackColor = vbYellow
    Label60.BackColor = vbYellow
    Label61.BackColor = vbYellow
    Label62.BackColor = vbYellow
    Label63.BackColor = vbYellow
 
    Label37.BackColor = vbYellow             'fourth box
    Label38.BackColor = vbYellow
    Label39.BackColor = vbYellow
    Label40.BackColor = vbYellow
    Label41.BackColor = vbYellow
    Label42.BackColor = vbYellow
    Label43.BackColor = vbYellow
    Label44.BackColor = vbYellow
    Label45.BackColor = vbYellow
       
    Label46.BackColor = vbYellow             'sixth box
    Label47.BackColor = vbYellow
    Label48.BackColor = vbYellow
    Label49.BackColor = vbYellow
    Label50.BackColor = vbYellow
    Label51.BackColor = vbYellow
    Label52.BackColor = vbYellow
    Label53.BackColor = vbYellow
    Label54.BackColor = vbYellow
 
    
    Label1.BackColor = vbYellow             'eighth box
    Label2.BackColor = vbYellow
    Label3.BackColor = vbYellow
    Label4.BackColor = vbYellow
    Label5.BackColor = vbYellow
    Label6.BackColor = vbYellow
    Label7.BackColor = vbYellow
    Label8.BackColor = vbYellow
    Label9.BackColor = vbYellow
   
End Sub

Private Sub mnExit_Click()
End
End Sub

Private Sub sx_Click()

End Sub


Private Sub method_Click()
    MENU.Visible = False
    txtgoogle.Visible = False
    frmhlp.Visible = True
    frmhlp.Caption = "HELP"
    str = ""
    str = str + "The objective of the "
    str = str + "game is to fill all the"
    str = str + "blank squares in a game"
    str = str + "with the correct numbers. "
    str = str + "There are three very"
    str = str + "simple constraints to"
    str = str + "follow. In a 9 by 9 square"
    str = str + "Sudoku game:"
    str = str + "   (1). Every row of 9"
    str = str + "numbers"
    str = str + "must include all digits 1"
    str = str + "through 9 in any order."
    str = str + "                              (2). Every column of 9"
    str = str + "numbers must include all"
    str = str + "digits 1 through 9 in any"
    str = str + "Order."
    str = str + "                        (3). Every 3 by 3"
    str = str + "subsection"
    str = str + "of the 9 by 9 square must"
    str = str + "include all digits 1"
    str = str + "through 9."
    txthlp.Left = 120
    txthlp.Top = 240
    frmhlp.Left = 120
    frmhlp.Top = 800
    CMDHLP.Left = 2160
    CMDHLP.Top = 3840
    txthlp.Text = ""
    txthlp.Text = str
    CMDHLP.SetFocus
    bg.Enabled = True
End Sub

Private Sub mnublog_Click()
Shell "explorer.exe http://www.facebook.com/pages/Sudoku/283069425083172"
End Sub

Private Sub mnuG_Click()
    MENU.Visible = False
    frmhlp.Visible = False
    txtgoogle.Visible = True
'    txtgoogle.SetFocus
    txtgoogle.Text = ""
    txtgoogle_GotFocus
    
End Sub

Private Sub mnunew1_Click()
    MENU.Visible = True
   txtgoogle.Visible = False
   frmhlp.Visible = False
   MENU.Left = 370
   MENU.Top = 480
   cmdeasy.SetFocus
   cmdlock.Visible = False
'   lblUserMsg.Visible = False
End Sub

Private Sub mnuopt_Click()
    MENU.Visible = False
    txtgoogle.Visible = False
    frmhlp.Visible = False
    frmcolor.Show 1
End Sub

Private Sub mnuUpdate_Click()
Shell "explorer.exe http://www.4shared.com/folder/ACGrtK5c/_online.html"

End Sub

Private Sub PressNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        InputNumber = 0
        If InputNumber = 0 Then
            If CheckNumber = 1 Then
                Label1.Caption = ""
            End If
            If CheckNumber = 2 Then
                Label2.Caption = ""
            End If
            If CheckNumber = 3 Then
                Label3.Caption = ""
            End If
            If CheckNumber = 4 Then
                Label4.Caption = ""
            End If
            If CheckNumber = 5 Then
                Label5.Caption = ""
            End If
            If CheckNumber = 6 Then
                Label6.Caption = ""
            End If
            If CheckNumber = 7 Then
                Label7.Caption = ""
            End If
            If CheckNumber = 8 Then
                Label8.Caption = ""
            End If
            If CheckNumber = 9 Then
                Label9.Caption = ""
            End If
            If CheckNumber = 10 Then
                Label10.Caption = ""
            End If
            If CheckNumber = 11 Then
                Label11.Caption = ""
            End If
            If CheckNumber = 12 Then
                Label12.Caption = ""
            End If
            If CheckNumber = 13 Then
                Label13.Caption = ""
            End If
            If CheckNumber = 14 Then
                Label14.Caption = ""
            End If
            If CheckNumber = 15 Then
                Label15.Caption = ""
            End If
            If CheckNumber = 16 Then
                Label16.Caption = ""
            End If
            If CheckNumber = 17 Then
                Label17.Caption = ""
            End If
            If CheckNumber = 18 Then
                Label18.Caption = ""
            End If
            If CheckNumber = 19 Then
                Label19.Caption = ""
            End If
            If CheckNumber = 20 Then
                Label20.Caption = ""
            End If
            If CheckNumber = 21 Then
                Label21.Caption = ""
            End If
            If CheckNumber = 22 Then
                Label22.Caption = ""
            End If
            If CheckNumber = 23 Then
                Label23.Caption = ""
            End If
            If CheckNumber = 24 Then
                Label24.Caption = ""
            End If
            If CheckNumber = 25 Then
                Label25.Caption = ""
            End If
            If CheckNumber = 26 Then
                Label26.Caption = ""
            End If
            If CheckNumber = 27 Then
                Label27.Caption = ""
            End If
            If CheckNumber = 28 Then
                Label28.Caption = ""
            End If
            If CheckNumber = 29 Then
                Label29.Caption = ""
            End If
            If CheckNumber = 30 Then
                Label30.Caption = ""
            End If
            If CheckNumber = 31 Then
                Label31.Caption = ""
            End If
            If CheckNumber = 32 Then
                Label32.Caption = ""
            End If
            If CheckNumber = 33 Then
                Label33.Caption = ""
            End If
            If CheckNumber = 34 Then
                Label34.Caption = ""
            End If
            If CheckNumber = 35 Then
                Label35.Caption = ""
            End If
            If CheckNumber = 36 Then
                Label36.Caption = ""
            End If
            If CheckNumber = 37 Then
                Label37.Caption = ""
            End If
            If CheckNumber = 38 Then
                Label38.Caption = ""
            End If
            If CheckNumber = 39 Then
                Label39.Caption = ""
            End If
            If CheckNumber = 40 Then
                Label40.Caption = ""
            End If
            If CheckNumber = 41 Then
                Label41.Caption = ""
            End If
            If CheckNumber = 42 Then
                Label42.Caption = ""
            End If
            If CheckNumber = 43 Then
                Label43.Caption = ""
            End If
            If CheckNumber = 44 Then
                Label44.Caption = ""
            End If
            If CheckNumber = 45 Then
                Label45.Caption = ""
            End If
            If CheckNumber = 46 Then
                Label46.Caption = ""
            End If
            If CheckNumber = 47 Then
                Label47.Caption = ""
            End If
            If CheckNumber = 48 Then
                Label48.Caption = ""
            End If
            If CheckNumber = 49 Then
                Label49.Caption = ""
            End If
            If CheckNumber = 50 Then
                Label50.Caption = ""
            End If
            If CheckNumber = 51 Then
                Label51.Caption = ""
            End If
            If CheckNumber = 52 Then
                Label52.Caption = ""
            End If
            If CheckNumber = 53 Then
                Label53.Caption = ""
            End If
            If CheckNumber = 54 Then
                Label54.Caption = ""
            End If
            If CheckNumber = 55 Then
                Label55.Caption = ""
            End If
            If CheckNumber = 56 Then
                Label56.Caption = ""
            End If
            If CheckNumber = 57 Then
                Label57.Caption = ""
            End If
            If CheckNumber = 58 Then
                Label58.Caption = ""
            End If
            If CheckNumber = 59 Then
                Label59.Caption = ""
            End If
            If CheckNumber = 60 Then
                Label60.Caption = ""
            End If
            If CheckNumber = 61 Then
                Label61.Caption = ""
            End If
            If CheckNumber = 62 Then
                Label62.Caption = ""
            End If
            If CheckNumber = 63 Then
                Label63.Caption = ""
            End If
            If CheckNumber = 64 Then
                Label64.Caption = ""
            End If
            If CheckNumber = 65 Then
                Label65.Caption = ""
            End If
            If CheckNumber = 66 Then
                Label66.Caption = ""
            End If
            If CheckNumber = 67 Then
                Label67.Caption = ""
            End If
            If CheckNumber = 68 Then
                Label68.Caption = ""
            End If
            If CheckNumber = 69 Then
                Label69.Caption = ""
            End If
            If CheckNumber = 70 Then
                Label70.Caption = ""
            End If
            If CheckNumber = 71 Then
                Label71.Caption = ""
            End If
            If CheckNumber = 72 Then
                Label72.Caption = ""
            End If
            If CheckNumber = 73 Then
                Label73.Caption = ""
            End If
            If CheckNumber = 74 Then
                Label74.Caption = ""
            End If
            If CheckNumber = 75 Then
                Label75.Caption = ""
            End If
            If CheckNumber = 76 Then
                Label76.Caption = ""
            End If
            If CheckNumber = 77 Then
                Label77.Caption = ""
            End If
            If CheckNumber = 78 Then
                Label78.Caption = ""
            End If
            If CheckNumber = 79 Then
                Label79.Caption = ""
            End If
            If CheckNumber = 80 Then
                Label80.Caption = ""
            End If
            If CheckNumber = 81 Then
                Label81.Caption = ""
            End If
        End If
    End If
End Sub

Private Sub PressNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If MsgBox("Are You Sure Want To Close This Application ?", vbQuestion + vbYesNo, "GOBINDA") = vbYes Then
            Unload Me
        End If
    End If

    If cmdstartString <> 1 Then
        Exit Sub
    End If
    If KeyAscii < 49 Or KeyAscii > 57 Then
        Exit Sub
    End If
    If CheckNumber = 1 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
     ElseIf CheckNumber = 2 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    ElseIf CheckNumber = 3 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    ElseIf CheckNumber = 4 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    ElseIf CheckNumber = 5 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    ElseIf CheckNumber = 6 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    ElseIf CheckNumber = 7 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    ElseIf CheckNumber = 8 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 9 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 10 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 11 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 12 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 13 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 14 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 15 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 16 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 17 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 18 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 19 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 20 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 21 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 22 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 23 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 24 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 25 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 26 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 27 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 28 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 29 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 30 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 31 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 32 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 33 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 34 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 35 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 36 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 37 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 38 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 39 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 40 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 41 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 42 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 43 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 44 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 45 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 46 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 47 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 48 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 49 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 50 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 51 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 52 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 53 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 54 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 55 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 56 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 57 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 58 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 59 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 60 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 61 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 62 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 63 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 64 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 65 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 66 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 67 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 68 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 69 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 70 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 71 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 72 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 73 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 74 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 75 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 76 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 77 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 78 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 79 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 80 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    ElseIf CheckNumber = 81 Then
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
    
    End If
    Call INPUTE
End Sub


Private Sub Timer1_Timer()
lbltimer1.Caption = lbltimer1.Caption + 1
End Sub

Private Sub Timer2_Timer()
  If lbltimer2.Caption = 60 Then
        lbltimer2.Caption = 0
    End If
    lbltimer2.Caption = lbltimer2.Caption + 1

End Sub


Private Sub Timer3_Timer()
lbldiv.ForeColor = Rnd * 1098
End Sub



Public Sub chk31()
    If Label31.Caption <> "" Then
            x31 = Label31.Caption
            If Label36.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label55.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label7.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label7.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label49.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x31 Then
                Label31.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk32()
    If Label32.Caption <> "" Then
            x32 = Label32.Caption
            If Label36.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label56.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label8.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label8.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label49.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x32 Then
                Label32.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk30()
    If Label30.Caption <> "" Then
            x30 = Label30.Caption
            If Label36.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
                    If Label3.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label3.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label6.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label9.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label9.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label46.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x30 Then
                Label30.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If


    End If
End Sub

Public Sub chk28()
    If Label28.Caption <> "" Then
            x28 = Label28.Caption
            If Label36.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label55.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label7.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label7.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label46.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x28 Then
                Label28.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub chk29()
    If Label29.Caption <> "" Then
            x29 = Label29.Caption
            If Label36.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label56.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label8.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label8.BackColor = vbWhite
                CheckError = "1"
            End If
            
            If Label46.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x29 Then
                Label29.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk45()
    If Label45.Caption <> "" Then
            x45 = Label45.Caption
            If Label43.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label34.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label36.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label12.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label18.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x45 Then
                Label45.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk43()
    If Label43.Caption <> "" Then
            x43 = Label43.Caption
            If Label45.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label34.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label36.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x43 Then
                Label43.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk44()
    If Label44.Caption <> "" Then
            x44 = Label44.Caption
            If Label45.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label36.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x44 Then
                Label44.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk42()
    If Label42.Caption <> "" Then
            x42 = Label42.Caption
            If Label45.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label31.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label49.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label12.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label18.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x42 Then
                Label42.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk40()
    If Label40.Caption <> "" Then
            x40 = Label40.Caption
            If Label45.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label31.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label49.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x40 Then
                Label40.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk41()
     If Label41.Caption <> "" Then
            x41 = Label41.Caption
            If Label45.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label31.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label49.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x41 Then
                Label41.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk39()
     If Label39.Caption <> "" Then
            x39 = Label39.Caption
            If Label45.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label28.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label46.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label12.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label18.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x39 Then
                Label39.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk37()
     If Label37.Caption <> "" Then
            x37 = Label37.Caption
            If Label45.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label46.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x37 Then
                Label37.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk38()
     If Label38.Caption <> "" Then
            x38 = Label38.Caption
            If Label45.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
                        If Label28.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label46.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x38 Then
                Label38.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk81()
     If Label81.Caption <> "" Then
            x81 = Label81.Caption
            If Label73.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label21.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label21.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label24.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label24.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label27.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label27.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x81 Then
                Label81.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk79()
     If Label79.Caption <> "" Then
            x79 = Label79.Caption
            If Label73.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label46.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label49.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label19.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label19.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label22.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label22.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label25.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label25.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x79 Then
                Label79.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk80()
     If Label80.Caption <> "" Then
            x80 = Label80.Caption
            If Label73.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label20.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label20.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label23.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label23.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label26.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label26.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x80 Then
                Label80.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk78()
     If Label78.Caption <> "" Then
            x78 = Label78.Caption
            If Label73.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label21.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label21.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label24.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label24.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label27.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label27.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x78 Then
                Label78.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk76()
     If Label76.Caption <> "" Then
            x76 = Label76.Caption
            If Label73.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label46.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label49.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label19.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label19.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label22.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label22.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label25.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label25.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x76 Then
                Label76.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk77()
     If Label77.Caption <> "" Then
            x77 = Label77.Caption
            If Label73.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label20.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label20.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label23.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label23.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label26.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label26.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x77 Then
                Label77.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk75()
     If Label75.Caption <> "" Then
            x75 = Label75.Caption
            If Label73.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label48.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label48.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label51.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label51.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label54.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label54.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label21.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label21.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label24.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label24.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label27.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label27.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x75 Then
                Label75.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk73()
     If Label73.Caption <> "" Then
            x73 = Label73.Caption
            If Label75.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label46.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label46.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label49.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label49.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label52.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label52.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label19.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label19.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label22.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label22.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label25.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label25.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x73 Then
                Label73.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk74()
     If Label74.Caption <> "" Then
            x74 = Label74.Caption
            If Label75.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label47.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label47.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label50.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label50.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label53.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label53.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label20.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label20.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label23.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label23.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label26.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label26.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x74 Then
                Label74.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk63()
     If Label63.Caption <> "" Then
            x63 = Label63.Caption
            If Label55.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label36.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label3.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label3.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label6.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label9.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label9.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x63 Then
                Label63.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk61()
     If Label61.Caption <> "" Then
            x61 = Label61.Caption
            If Label55.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label7.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label7.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x61 Then
                Label61.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
 
End Sub

Public Sub chk62()
     If Label62.Caption <> "" Then
            x62 = Label62.Caption
            If Label55.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label8.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label8.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x62 Then
                Label62.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk60()
     If Label60.Caption <> "" Then
            x60 = Label60.Caption
            If Label55.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label36.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label3.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label3.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label6.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label9.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label9.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x60 Then
                Label60.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk58()
     If Label58.Caption <> "" Then
            x58 = Label58.Caption
            If Label55.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label7.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label7.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x58 Then
                Label58.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk59()
     If Label59.Caption <> "" Then
            x59 = Label59.Caption
            If Label55.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label8.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label8.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x59 Then
                Label59.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk57()
     If Label57.Caption <> "" Then
            x57 = Label57.Caption
            If Label55.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label30.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label30.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label33.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label33.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label36.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label36.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label3.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label3.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label6.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label6.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label9.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label9.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x57 Then
                Label57.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk55()
     If Label55.Caption <> "" Then
            x55 = Label55.Caption
            If Label57.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label28.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label28.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label31.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label31.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label34.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label34.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label1.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label1.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label4.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label4.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label7.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label7.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x55 Then
                Label55.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk56()
     If Label56.Caption <> "" Then
            x56 = Label56.Caption
            If Label57.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label29.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label29.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label32.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label32.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label35.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label35.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label2.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label2.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label5.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label5.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label8.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label8.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x56 Then
                Label56.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk72()
     If Label72.Caption <> "" Then
            x72 = Label72.Caption
            If Label64.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label12.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label18.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x72 Then
                Label72.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk70()
     If Label70.Caption <> "" Then
            x70 = Label70.Caption
            If Label64.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x70 Then
                Label70.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk71()
     If Label71.Caption <> "" Then
            x71 = Label71.Caption
            If Label64.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label61.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label61.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label62.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label62.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label63.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label63.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label79.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label79.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label80.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label80.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label81.Caption = x71 Then
                Label71.BackColor = vbWhite
                Label81.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk69()
     If Label69.Caption <> "" Then
            x69 = Label69.Caption
            If Label64.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label12.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label18.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x69 Then
                Label69.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk67()
     If Label67.Caption <> "" Then
            x67 = Label67.Caption
            If Label64.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x67 Then
                Label67.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk68()
     If Label68.Caption <> "" Then
            x68 = Label68.Caption
            If Label64.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label66.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label58.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label58.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label59.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label59.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label60.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label60.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label76.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label76.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label77.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label77.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label78.Caption = x68 Then
                Label68.BackColor = vbWhite
                Label78.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk66()
     If Label66.Caption <> "" Then
            x66 = Label66.Caption
            If Label64.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label39.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label39.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label42.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label42.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label45.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label45.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label12.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label12.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label15.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label15.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label18.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label18.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x66 Then
                Label66.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk64()
     If Label64.Caption <> "" Then
            x64 = Label64.Caption
            If Label66.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label65.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label65.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label37.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label37.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label40.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label40.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label43.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label43.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label10.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label10.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label13.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label13.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label16.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label16.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x64 Then
                Label64.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Public Sub chk65()
     If Label65.Caption <> "" Then
            x65 = Label65.Caption
            If Label66.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label66.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label64.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label64.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label68.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label68.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label69.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label69.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label67.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label67.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label71.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label71.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label72.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label72.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label70.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label70.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label38.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label38.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label41.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label41.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label44.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label44.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label11.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label11.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label14.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label14.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label17.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label17.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label55.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label55.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label56.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label56.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label57.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label57.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label73.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label73.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label74.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label74.BackColor = vbWhite
                CheckError = "1"
            End If
            If Label75.Caption = x65 Then
                Label65.BackColor = vbWhite
                Label75.BackColor = vbWhite
                CheckError = "1"
            End If
    End If
End Sub

Private Sub timermarquee_Timer()
    If lblmarquee.Left = 3570 Then
        lblmarquee.Left = -1080
    End If
    lblmarquee.Left = 50 + lblmarquee.Left
    If FinishB = True Then
        FinishTimerI
    End If
    If CheckError = "1" Then
        lbmistakes.ForeColor = vbRed
        lbmistakes.Caption = "** Your Mistakes Are Marked."
    Else
        lbmistakes.Caption = "** No Mistakes Yet."
        lbmistakes.ForeColor = vbWhite
    End If
End Sub

Private Sub title_Timer()
    Label82.Left = Rnd * 1000
End Sub

Private Sub KeyyAs()
        If KeyAscii = 49 Then
            InputNumber = 1
        ElseIf KeyAscii = 50 Then
            InputNumber = 2
        ElseIf KeyAscii = 51 Then
            InputNumber = 3
        ElseIf KeyAscii = 52 Then
            InputNumber = 4
        ElseIf KeyAscii = 53 Then
            InputNumber = 5
        ElseIf KeyAscii = 54 Then
            InputNumber = 6
        ElseIf KeyAscii = 55 Then
            InputNumber = 7
        ElseIf KeyAscii = 56 Then
            InputNumber = 8
        ElseIf KeyAscii = 57 Then
            InputNumber = 9
        End If
End Sub

Private Sub txtgoogle_Change()
'    txtgoogle.Text = UCase(txtgoogle.Text)
    If txtgoogle.Text = "Google Search" Then
        txtgoogle.ForeColor = &H80000010
    ElseIf txtgoogle.Text = "" Then
        txtgoogle.ForeColor = &H80000010
    Else
        txtgoogle.ForeColor = &HC00000
    End If

End Sub

Private Sub txtgoogle_GotFocus()
     If txtgoogle.Text = "Google Search" Then
           txtgoogle.Text = ""
     ElseIf txtgoogle.Text = "" Then
           txtgoogle.Text = "Google Search"
     End If

End Sub


Private Sub txtgoogle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmGoogle.webg.Navigate "http://www.google.co.in/search?sclient=psy-ab&hl=en&site=&source=hp&q=" & txtgoogle.Text & "&btnK=Google+Search"
        frmGoogle.Show 1
        txtgoogle.Text = ""
        txtgoogle.Visible = False
    ElseIf KeyAscii = 27 Then
        txtgoogle.Visible = False
    End If
End Sub

Private Sub txtgoogle_LostFocus()
    If txtgoogle.Text = "" Then
           txtgoogle.Text = "Google Search"
     End If
End Sub

Private Sub txthp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txthp.Visible = False
    End If
End Sub

Private Sub txttip_Timer()
    txthpi = txthpi + 1
    txthp.Text = UCase(Left(txthelp, txthpi))
End Sub

Public Sub easy1()
    Label1.Caption = 9
    Label2.Caption = 5
    Label3.Caption = 7
    Label4.Caption = 4
    Label5.Caption = 8
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = 1
    Label9.Caption = 2
    Label10.Caption = 3
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = 9
    Label14.Caption = 5
    Label15.Caption = ""
    Label16.Caption = 7
    Label17.Caption = 8
    Label18.Caption = ""
    Label19.Caption = ""
    Label20.Caption = 8
    Label21.Caption = 4
    Label22.Caption = ""
    Label23.Caption = 6
    Label24.Caption = 1
    Label25.Caption = ""
    Label26.Caption = 5
    Label27.Caption = 9
    Label28.Caption = 5
    Label29.Caption = 2
    Label30.Caption = 8
    Label31.Caption = ""
    Label32.Caption = ""
    Label33.Caption = ""
    Label34.Caption = 3
    Label35.Caption = 6
    Label36.Caption = 4
    Label37.Caption = 4
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = 8
    Label41.Caption = 2
    Label42.Caption = 6
    Label43.Caption = 5
    Label44.Caption = ""
    Label45.Caption = 9
    Label46.Caption = 9
    Label47.Caption = 7
    Label48.Caption = ""
    Label49.Caption = 4
    Label50.Caption = 3
    Label51.Caption = 5
    Label52.Caption = 1
    Label53.Caption = ""
    Label54.Caption = ""
    Label55.Caption = ""
    Label56.Caption = 7
    Label57.Caption = 6
    Label58.Caption = 8
    Label59.Caption = ""
    Label60.Caption = 5
    Label61.Caption = 2
    Label62.Caption = 4
    Label63.Caption = 9
    Label64.Caption = ""
    Label65.Caption = 9
    Label66.Caption = 5
    Label67.Caption = ""
    Label68.Caption = 4
    Label69.Caption = 7
    Label70.Caption = ""
    Label71.Caption = 3
    Label72.Caption = 8
    Label73.Caption = 8
    Label74.Caption = ""
    Label75.Caption = 3
    Label76.Caption = 6
    Label77.Caption = ""
    Label78.Caption = 2
    Label79.Caption = 5
    Label80.Caption = ""
    Label81.Caption = ""

'    Label1.Caption = ""
'    Label2.Caption = ""
'    Label3.Caption = ""
'    Label4.Caption = ""
'    Label5.Caption = ""
'    Label6.Caption = ""
'    Label7.Caption = ""
'    Label8.Caption = ""
'    Label9.Caption = ""
'    Label10.Caption = ""
'    Label11.Caption = ""
'    Label12.Caption = ""
'    Label13.Caption = ""
'    Label14.Caption = ""
'    Label15.Caption = ""
'    Label16.Caption = ""
'    Label17.Caption = ""
'    Label18.Caption = ""
'    Label19.Caption = ""
'    Label20.Caption = ""
'    Label21.Caption = ""
'    Label22.Caption = ""
'    Label23.Caption = ""
'    Label24.Caption = ""
'    Label25.Caption = ""
'    Label26.Caption = ""
'    Label27.Caption = ""
'    Label28.Caption = ""
'    Label29.Caption = ""
'    Label30.Caption = ""
'    Label31.Caption = ""
'    Label32.Caption = ""
'    Label33.Caption = ""
'    Label34.Caption = ""
'    Label35.Caption = ""
'    Label36.Caption = ""
'    Label37.Caption = ""
'    Label38.Caption = ""
'    Label39.Caption = ""
'    Label40.Caption = ""
'    Label41.Caption = ""
'    Label42.Caption = ""
'    Label43.Caption = ""
'    Label44.Caption = ""
'    Label45.Caption = ""
'    Label46.Caption = ""
'    Label47.Caption = ""
'    Label48.Caption = ""
'    Label49.Caption = ""
'    Label50.Caption = ""
'    Label51.Caption = ""
'    Label52.Caption = ""
'    Label53.Caption = ""
'    Label54.Caption = ""
'    Label55.Caption = ""
'    Label56.Caption = ""
'    Label57.Caption = ""
'    Label58.Caption = ""
'    Label59.Caption = ""
'    Label60.Caption = ""
'    Label61.Caption = ""
'    Label62.Caption = ""
'    Label63.Caption = ""
'    Label64.Caption = ""
'    Label65.Caption = ""
'    Label66.Caption = ""
'    Label67.Caption = ""
'    Label68.Caption = ""
'    Label69.Caption = ""
'    Label70.Caption = ""
'    Label71.Caption = ""
'    Label72.Caption = ""
'    Label73.Caption = ""
'    Label74.Caption = ""
'    Label75.Caption = ""
'    Label76.Caption = ""
'    Label77.Caption = ""
'    Label78.Caption = ""
'    Label79.Caption = ""
'    Label80.Caption = ""
'    Label81.Caption = ""
    
    
    '*********************************************
    Label1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = True
    Label7.Enabled = True
    Label8.Enabled = False
    Label9.Enabled = False
    Label10.Enabled = False
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = False
    Label14.Enabled = False
    Label15.Enabled = True
    Label16.Enabled = False
    Label17.Enabled = False
    Label18.Enabled = True
    Label19.Enabled = True
    Label20.Enabled = False
    Label21.Enabled = False
    Label22.Enabled = True
    Label23.Enabled = False
    Label24.Enabled = False
    Label25.Enabled = True
    Label26.Enabled = False
    Label27.Enabled = False
    Label28.Enabled = False
    Label29.Enabled = False
    Label30.Enabled = False
    Label31.Enabled = True
    Label32.Enabled = True
    Label33.Enabled = True
    Label34.Enabled = False
    Label35.Enabled = False
    Label36.Enabled = False
    Label37.Enabled = False
    Label38.Enabled = True
    Label39.Enabled = True
    Label40.Enabled = False
    Label41.Enabled = False
    Label42.Enabled = False
    Label43.Enabled = False
    Label44.Enabled = True
    Label45.Enabled = False
    Label46.Enabled = False
    Label47.Enabled = False
    Label48.Enabled = True
    Label49.Enabled = False
    Label50.Enabled = False
    Label51.Enabled = False
    Label52.Enabled = False
    Label53.Enabled = True
    Label54.Enabled = True
    Label55.Enabled = True
    Label56.Enabled = False
    Label57.Enabled = False
    Label58.Enabled = False
    Label59.Enabled = True
    Label60.Enabled = False
    Label61.Enabled = False
    Label62.Enabled = False
    Label63.Enabled = False
    Label64.Enabled = True
    Label65.Enabled = False
    Label66.Enabled = False
    Label67.Enabled = True
    Label68.Enabled = False
    Label69.Enabled = False
    Label70.Enabled = True
    Label71.Enabled = False
    Label72.Enabled = False
    Label73.Enabled = False
    Label74.Enabled = True
    Label75.Enabled = False
    Label76.Enabled = False
    Label77.Enabled = True
    Label78.Enabled = False
    Label79.Enabled = False
    Label80.Enabled = True
    Label81.Enabled = True

'    Label1.Enabled = False
'    Label2.Enabled = False
'    Label3.Enabled = False
'    Label4.Enabled = False
'    Label5.Enabled = False
'    Label6.Enabled = False
'    Label7.Enabled = False
'    Label8.Enabled = False
'    Label9.Enabled = False
'    Label10.Enabled = False
'    Label11.Enabled = False
'    Label12.Enabled = False
'    Label13.Enabled = False
'    Label14.Enabled = False
'    Label15.Enabled = False
'    Label16.Enabled = False
'    Label17.Enabled = False
'    Label18.Enabled = False
'    Label19.Enabled = False
'    Label20.Enabled = False
'    Label21.Enabled = False
'    Label22.Enabled = False
'    Label23.Enabled = False
'    Label24.Enabled = False
'    Label25.Enabled = False
'    Label26.Enabled = False
'    Label27.Enabled = False
'    Label28.Enabled = False
'    Label29.Enabled = False
'    Label30.Enabled = False
'    Label31.Enabled = False
'    Label32.Enabled = False
'    Label33.Enabled = False
'    Label34.Enabled = False
'    Label35.Enabled = False
'    Label36.Enabled = False
'    Label37.Enabled = False
'    Label38.Enabled = False
'    Label39.Enabled = False
'    Label40.Enabled = False
'    Label41.Enabled = False
'    Label42.Enabled = False
'    Label43.Enabled = False
'    Label44.Enabled = False
'    Label45.Enabled = False
'    Label46.Enabled = False
'    Label47.Enabled = False
'    Label48.Enabled = False
'    Label49.Enabled = False
'    Label50.Enabled = False
'    Label51.Enabled = False
'    Label52.Enabled = False
'    Label53.Enabled = False
'    Label54.Enabled = False
'    Label55.Enabled = False
'    Label56.Enabled = False
'    Label57.Enabled = False
'    Label58.Enabled = False
'    Label59.Enabled = False
'    Label60.Enabled = False
'    Label61.Enabled = False
'    Label62.Enabled = False
'    Label63.Enabled = False
'    Label64.Enabled = False
'    Label65.Enabled = False
'    Label66.Enabled = False
'    Label67.Enabled = False
'    Label68.Enabled = False
'    Label69.Enabled = False
'    Label70.Enabled = False
'    Label71.Enabled = False
'    Label72.Enabled = False
'    Label73.Enabled = False
'    Label74.Enabled = False
'    Label75.Enabled = False
'    Label76.Enabled = False
'    Label77.Enabled = False
'    Label78.Enabled = False
'    Label79.Enabled = False
'    Label80.Enabled = False
'    Label81.Enabled = False
    '************************************************
        Label1.Tag = False
    Label2.Tag = False
    Label3.Tag = False
    Label4.Tag = False
    Label5.Tag = False
    Label6.Tag = True
    Label7.Tag = True
    Label8.Tag = False
    Label9.Tag = False
    Label10.Tag = False
    Label11.Tag = True
    Label12.Tag = True
    Label13.Tag = False
    Label14.Tag = False
    Label15.Tag = True
    Label16.Tag = False
    Label17.Tag = False
    Label18.Tag = True
    Label19.Tag = True
    Label20.Tag = False
    Label21.Tag = False
    Label22.Tag = True
    Label23.Tag = False
    Label24.Tag = False
    Label25.Tag = True
    Label26.Tag = False
    Label27.Tag = False
    Label28.Tag = False
    Label29.Tag = False
    Label30.Tag = False
    Label31.Tag = True
    Label32.Tag = True
    Label33.Tag = True
    Label34.Tag = False
    Label35.Tag = False
    Label36.Tag = False
    Label37.Tag = False
    Label38.Tag = True
    Label39.Tag = True
    Label40.Tag = False
    Label41.Tag = False
    Label42.Tag = False
    Label43.Tag = False
    Label44.Tag = True
    Label45.Tag = False
    Label46.Tag = False
    Label47.Tag = False
    Label48.Tag = True
    Label49.Tag = False
    Label50.Tag = False
    Label51.Tag = False
    Label52.Tag = False
    Label53.Tag = True
    Label54.Tag = True
    Label55.Tag = True
    Label56.Tag = False
    Label57.Tag = False
    Label58.Tag = False
    Label59.Tag = True
    Label60.Tag = False
    Label61.Tag = False
    Label62.Tag = False
    Label63.Tag = False
    Label64.Tag = True
    Label65.Tag = False
    Label66.Tag = False
    Label67.Tag = True
    Label68.Tag = False
    Label69.Tag = False
    Label70.Tag = True
    Label71.Tag = False
    Label72.Tag = False
    Label73.Tag = False
    Label74.Tag = True
    Label75.Tag = False
    Label76.Tag = False
    Label77.Tag = True
    Label78.Tag = False
    Label79.Tag = False
    Label80.Tag = True
    Label81.Tag = True

'    Label1.Tag = False
'    Label2.Tag = False
'    Label3.Tag = False
'    Label4.Tag = False
'    Label5.Tag = False
'    Label6.Tag = False
'    Label7.Tag = False
'    Label8.Tag = False
'    Label9.Tag = False
'    Label10.Tag = False
'    Label11.Tag = False
'    Label12.Tag = False
'    Label13.Tag = False
'    Label14.Tag = False
'    Label15.Tag = False
'    Label16.Tag = False
'    Label17.Tag = False
'    Label18.Tag = False
'    Label19.Tag = False
'    Label20.Tag = False
'    Label21.Tag = False
'    Label22.Tag = False
'    Label23.Tag = False
'    Label24.Tag = False
'    Label25.Tag = False
'    Label26.Tag = False
'    Label27.Tag = False
'    Label28.Tag = False
'    Label29.Tag = False
'    Label30.Tag = False
'    Label31.Tag = False
'    Label32.Tag = False
'    Label33.Tag = False
'    Label34.Tag = False
'    Label35.Tag = False
'    Label36.Tag = False
'    Label37.Tag = False
'    Label38.Tag = False
'    Label39.Tag = False
'    Label40.Tag = False
'    Label41.Tag = False
'    Label42.Tag = False
'    Label43.Tag = False
'    Label44.Tag = False
'    Label45.Tag = False
'    Label46.Tag = False
'    Label47.Tag = False
'    Label48.Tag = False
'    Label49.Tag = False
'    Label50.Tag = False
'    Label51.Tag = False
'    Label52.Tag = False
'    Label53.Tag = False
'    Label54.Tag = False
'    Label55.Tag = False
'    Label56.Tag = False
'    Label57.Tag = False
'    Label58.Tag = False
'    Label59.Tag = False
'    Label60.Tag = False
'    Label61.Tag = False
'    Label62.Tag = False
'    Label63.Tag = False
'    Label64.Tag = False
'    Label65.Tag = False
'    Label66.Tag = False
'    Label67.Tag = False
'    Label68.Tag = False
'    Label69.Tag = False
'    Label70.Tag = False
'    Label71.Tag = False
'    Label72.Tag = False
'    Label73.Tag = False
'    Label74.Tag = False
'    Label75.Tag = False
'    Label76.Tag = False
'    Label77.Tag = False
'    Label78.Tag = False
'    Label79.Tag = False
'    Label80.Tag = False
'    Label81.Tag = False
End Sub

Public Sub easy2()
    Label1.Caption = 6
    Label2.Caption = 1
    Label3.Caption = ""
    Label4.Caption = 5
    Label5.Caption = 7
    Label6.Caption = ""
    Label7.Caption = 2
    Label8.Caption = 8
    Label9.Caption = ""
    Label10.Caption = 8
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = 3
    Label14.Caption = 9
    Label15.Caption = 6
    Label16.Caption = 4
    Label17.Caption = ""
    Label18.Caption = 1
    Label19.Caption = ""
    Label20.Caption = 9
    Label21.Caption = 4
    Label22.Caption = ""
    Label23.Caption = 1
    Label24.Caption = ""
    Label25.Caption = 3
    Label26.Caption = 7
    Label27.Caption = 6
    Label28.Caption = ""
    Label29.Caption = 3
    Label30.Caption = 2
    Label31.Caption = ""
    Label32.Caption = 5
    Label33.Caption = 8
    Label34.Caption = ""
    Label35.Caption = 9
    Label36.Caption = 6
    Label37.Caption = ""
    Label38.Caption = 6
    Label39.Caption = ""
    Label40.Caption = 9
    Label41.Caption = 2
    Label42.Caption = ""
    Label43.Caption = 5
    Label44.Caption = 4
    Label45.Caption = 7
    Label46.Caption = 9
    Label47.Caption = 5
    Label48.Caption = 7
    Label49.Caption = 4
    Label50.Caption = ""
    Label51.Caption = 1
    Label52.Caption = ""
    Label53.Caption = ""
    Label54.Caption = 8
    Label55.Caption = 3
    Label56.Caption = ""
    Label57.Caption = 1
    Label58.Caption = 8
    Label59.Caption = ""
    Label60.Caption = 5
    Label61.Caption = 9
    Label62.Caption = ""
    Label63.Caption = 7
    Label64.Caption = 6
    Label65.Caption = 8
    Label66.Caption = 5
    Label67.Caption = ""
    Label68.Caption = ""
    Label69.Caption = 9
    Label70.Caption = ""
    Label71.Caption = 3
    Label72.Caption = 4
    Label73.Caption = 7
    Label74.Caption = 2
    Label75.Caption = ""
    Label76.Caption = ""
    Label77.Caption = 4
    Label78.Caption = 3
    Label79.Caption = 1
    Label80.Caption = ""
    Label81.Caption = ""
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Label1.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = True
    Label4.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = True
    Label7.Enabled = False
    Label8.Enabled = False
    Label9.Enabled = True
    Label10.Enabled = False
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = False
    Label14.Enabled = False
    Label15.Enabled = False
    Label16.Enabled = False
    Label17.Enabled = True
    Label18.Enabled = False
    Label19.Enabled = True
    Label20.Enabled = False
    Label21.Enabled = False
    Label22.Enabled = True
    Label23.Enabled = False
    Label24.Enabled = True
    Label25.Enabled = False
    Label26.Enabled = False
    Label27.Enabled = False
    Label28.Enabled = True
    Label29.Enabled = False
    Label30.Enabled = False
    Label31.Enabled = True
    Label32.Enabled = False
    Label33.Enabled = False
    Label34.Enabled = True
    Label35.Enabled = False
    Label36.Enabled = False
    Label37.Enabled = True
    Label38.Enabled = False
    Label39.Enabled = True
    Label40.Enabled = False
    Label41.Enabled = False
    Label42.Enabled = True
    Label43.Enabled = False
    Label44.Enabled = False
    Label45.Enabled = False
    Label46.Enabled = False
    Label47.Enabled = False
    Label48.Enabled = False
    Label49.Enabled = False
    Label50.Enabled = True
    Label51.Enabled = False
    Label52.Enabled = True
    Label53.Enabled = True
    Label54.Enabled = False
    Label55.Enabled = False
    Label56.Enabled = True
    Label57.Enabled = False
    Label58.Enabled = False
    Label59.Enabled = True
    Label60.Enabled = False
    Label61.Enabled = False
    Label62.Enabled = True
    Label63.Enabled = False
    Label64.Enabled = False
    Label65.Enabled = False
    Label66.Enabled = False
    Label67.Enabled = True
    Label68.Enabled = True
    Label69.Enabled = False
    Label70.Enabled = True
    Label71.Enabled = False
    Label72.Enabled = False
    Label73.Enabled = False
    Label74.Enabled = False
    Label75.Enabled = True
    Label76.Enabled = True
    Label77.Enabled = False
    Label78.Enabled = False
    Label79.Enabled = False
    Label80.Enabled = True
    Label81.Enabled = True
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Label1.Tag = False
    Label2.Tag = False
    Label3.Tag = True
    Label4.Tag = False
    Label5.Tag = False
    Label6.Tag = True
    Label7.Tag = False
    Label8.Tag = False
    Label9.Tag = True
    Label10.Tag = False
    Label11.Tag = True
    Label12.Tag = True
    Label13.Tag = False
    Label14.Tag = False
    Label15.Tag = False
    Label16.Tag = False
    Label17.Tag = True
    Label18.Tag = False
    Label19.Tag = True
    Label20.Tag = False
    Label21.Tag = False
    Label22.Tag = True
    Label23.Tag = False
    Label24.Tag = True
    Label25.Tag = False
    Label26.Tag = False
    Label27.Tag = False
    Label28.Tag = True
    Label29.Tag = False
    Label30.Tag = False
    Label31.Tag = True
    Label32.Tag = False
    Label33.Tag = False
    Label34.Tag = True
    Label35.Tag = False
    Label36.Tag = False
    Label37.Tag = True
    Label38.Tag = False
    Label39.Tag = True
    Label40.Tag = False
    Label41.Tag = False
    Label42.Tag = True
    Label43.Tag = False
    Label44.Tag = False
    Label45.Tag = False
    Label46.Tag = False
    Label47.Tag = False
    Label48.Tag = False
    Label49.Tag = False
    Label50.Tag = True
    Label51.Tag = False
    Label52.Tag = True
    Label53.Tag = True
    Label54.Tag = False
    Label55.Tag = False
    Label56.Tag = True
    Label57.Tag = False
    Label58.Tag = False
    Label59.Tag = True
    Label60.Tag = False
    Label61.Tag = False
    Label62.Tag = True
    Label63.Tag = False
    Label64.Tag = False
    Label65.Tag = False
    Label66.Tag = False
    Label67.Tag = True
    Label68.Tag = True
    Label69.Tag = False
    Label70.Tag = True
    Label71.Tag = False
    Label72.Tag = False
    Label73.Tag = False
    Label74.Tag = False
    Label75.Tag = True
    Label76.Tag = True
    Label77.Tag = False
    Label78.Tag = False
    Label79.Tag = False
    Label80.Tag = True
    Label81.Tag = True

End Sub

Public Sub easy3()
    
    Label65.Caption = ""
    Label64.Caption = 1
    Label66.Caption = 5
    Label56.Caption = 3
    Label55.Caption = 4
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 8
    Label75.Caption = 6
    Label68.Caption = 6
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 1
    Label58.Caption = 9
    Label60.Caption = ""
    Label77.Caption = 2
    Label76.Caption = ""
    Label78.Caption = 5
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 9
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 7
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 4
    Label39.Caption = ""
    Label29.Caption = 7
    Label28.Caption = ""
    Label30.Caption = 5
    Label47.Caption = 3
    Label46.Caption = ""
    Label48.Caption = 8
    Label41.Caption = ""
    Label40.Caption = 7
    Label42.Caption = 8
    Label32.Caption = ""
    Label31.Caption = 3
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 1
    Label51.Caption = ""
    Label44.Caption = 3
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 6
    Label34.Caption = ""
    Label36.Caption = 9
    Label53.Caption = ""
    Label52.Caption = 4
    Label54.Caption = ""
    Label11.Caption = 4
    Label10.Caption = ""
    Label12.Caption = 7
    Label2.Caption = ""
    Label1.Caption = 2
    Label3.Caption = ""
    Label20.Caption = 8
    Label19.Caption = 6
    Label21.Caption = ""
    Label14.Caption = 8
    Label13.Caption = ""
    Label15.Caption = 3
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 1
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 2
    Label17.Caption = ""
    Label16.Caption = 5
    Label18.Caption = ""
    Label8.Caption = 4
    Label7.Caption = ""
    Label9.Caption = 8
    Label26.Caption = ""
    Label25.Caption = 9
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = False
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = False
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub easy4()
    
    Label65.Caption = 9
    Label64.Caption = ""
    Label66.Caption = 7
    Label56.Caption = ""
    Label55.Caption = 6
    Label57.Caption = 2
    Label74.Caption = 1
    Label73.Caption = 8
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 8
    Label69.Caption = ""
    Label59.Caption = 7
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 4
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = 5
    Label70.Caption = ""
    Label72.Caption = 3
    Label62.Caption = ""
    Label61.Caption = 4
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 6
    Label38.Caption = 2
    Label37.Caption = ""
    Label39.Caption = 9
    Label29.Caption = 1
    Label28.Caption = ""
    Label30.Caption = 7
    Label47.Caption = 8
    Label46.Caption = 6
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = 9
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 2
    Label44.Caption = 8
    Label43.Caption = ""
    Label45.Caption = 4
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 6
    Label53.Caption = 9
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 4
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 6
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 7
    Label14.Caption = 7
    Label13.Caption = ""
    Label15.Caption = 1
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 3
    Label23.Caption = ""
    Label22.Caption = 4
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 5
    Label8.Caption = 4
    Label7.Caption = 7
    Label9.Caption = ""
    Label26.Caption = 6
    Label25.Caption = ""
    Label27.Caption = 8
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = False
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = False
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = False
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = False
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub easy5()

Label65.Caption = 9
Label64.Caption = ""
Label66.Caption = 6
Label56.Caption = 1
Label55.Caption = 2
Label57.Caption = ""
Label74.Caption = 3
Label73.Caption = ""
Label75.Caption = 7
Label68.Caption = 7
Label67.Caption = ""
Label69.Caption = ""
Label59.Caption = ""
Label58.Caption = ""
Label60.Caption = 8
Label77.Caption = ""
Label76.Caption = ""
Label78.Caption = 6
Label71.Caption = ""
Label70.Caption = 5
Label72.Caption = ""
Label62.Caption = ""
Label61.Caption = 7
Label63.Caption = ""
Label80.Caption = 2
Label79.Caption = ""
Label81.Caption = ""
Label38.Caption = ""
Label37.Caption = ""
Label39.Caption = ""
Label29.Caption = 7
Label28.Caption = ""
Label30.Caption = 2
Label47.Caption = 5
Label46.Caption = 8
Label48.Caption = ""
Label41.Caption = ""
Label40.Caption = 6
Label42.Caption = ""
Label32.Caption = ""
Label31.Caption = 4
Label33.Caption = ""
Label50.Caption = ""
Label49.Caption = 1
Label51.Caption = ""
Label44.Caption = ""
Label43.Caption = 9
Label45.Caption = 7
Label35.Caption = 8
Label34.Caption = ""
Label36.Caption = 3
Label53.Caption = ""
Label52.Caption = ""
Label54.Caption = ""
Label11.Caption = 6
Label10.Caption = ""
Label12.Caption = 9
Label2.Caption = ""
Label1.Caption = 5
Label3.Caption = 7
Label20.Caption = ""
Label19.Caption = ""
Label21.Caption = 1
Label14.Caption = 8
Label13.Caption = ""
Label15.Caption = ""
Label5.Caption = ""
Label4.Caption = 3
Label6.Caption = ""
Label23.Caption = 9
Label22.Caption = ""
Label24.Caption = 4
Label17.Caption = ""
Label16.Caption = 1
Label18.Caption = ""
Label8.Caption = ""
Label7.Caption = ""
Label9.Caption = 4
Label26.Caption = ""
Label25.Caption = 7
Label27.Caption = ""


Label65.Tag = False
Label64.Tag = True
Label66.Tag = False
Label56.Tag = False
Label55.Tag = False
Label57.Tag = True
Label74.Tag = False
Label73.Tag = True
Label75.Tag = False
Label68.Tag = False
Label67.Tag = True
Label69.Tag = True
Label59.Tag = True
Label58.Tag = True
Label60.Tag = False
Label77.Tag = True
Label76.Tag = True
Label78.Tag = False
Label71.Tag = True
Label70.Tag = False
Label72.Tag = True
Label62.Tag = True
Label61.Tag = False
Label63.Tag = True
Label80.Tag = False
Label79.Tag = True
Label81.Tag = True
Label38.Tag = True
Label37.Tag = True
Label39.Tag = True
Label29.Tag = False
Label28.Tag = True
Label30.Tag = False
Label47.Tag = False
Label46.Tag = False
Label48.Tag = True
Label41.Tag = True
Label40.Tag = False
Label42.Tag = True
Label32.Tag = True
Label31.Tag = False
Label33.Tag = True
Label50.Tag = True
Label49.Tag = False
Label51.Tag = True
Label44.Tag = True
Label43.Tag = False
Label45.Tag = False
Label35.Tag = False
Label34.Tag = True
Label36.Tag = False
Label53.Tag = True
Label52.Tag = True
Label54.Tag = True
Label11.Tag = False
Label10.Tag = True
Label12.Tag = False
Label2.Tag = True
Label1.Tag = False
Label3.Tag = False
Label20.Tag = True
Label19.Tag = True
Label21.Tag = False
Label14.Tag = False
Label13.Tag = True
Label15.Tag = True
Label5.Tag = True
Label4.Tag = False
Label6.Tag = True
Label23.Tag = False
Label22.Tag = True
Label24.Tag = False
Label17.Tag = True
Label16.Tag = False
Label18.Tag = True
Label8.Tag = True
Label7.Tag = True
Label9.Tag = False
Label26.Tag = True
Label25.Tag = False
Label27.Tag = True


Label65.Enabled = False
Label64.Enabled = True
Label66.Enabled = False
Label56.Enabled = False
Label55.Enabled = False
Label57.Enabled = True
Label74.Enabled = False
Label73.Enabled = True
Label75.Enabled = False
Label68.Enabled = False
Label67.Enabled = True
Label69.Enabled = True
Label59.Enabled = True
Label58.Enabled = True
Label60.Enabled = False
Label77.Enabled = True
Label76.Enabled = True
Label78.Enabled = False
Label71.Enabled = True
Label70.Enabled = False
Label72.Enabled = True
Label62.Enabled = True
Label61.Enabled = False
Label63.Enabled = True
Label80.Enabled = False
Label79.Enabled = True
Label81.Enabled = True
Label38.Enabled = True
Label37.Enabled = True
Label39.Enabled = True
Label29.Enabled = False
Label28.Enabled = True
Label30.Enabled = False
Label47.Enabled = False
Label46.Enabled = False
Label48.Enabled = True
Label41.Enabled = True
Label40.Enabled = False
Label42.Enabled = True
Label32.Enabled = True
Label31.Enabled = False
Label33.Enabled = True
Label50.Enabled = True
Label49.Enabled = False
Label51.Enabled = True
Label44.Enabled = True
Label43.Enabled = False
Label45.Enabled = False
Label35.Enabled = False
Label34.Enabled = True
Label36.Enabled = False
Label53.Enabled = True
Label52.Enabled = True
Label54.Enabled = True
Label11.Enabled = False
Label10.Enabled = True
Label12.Enabled = False
Label2.Enabled = True
Label1.Enabled = False
Label3.Enabled = False
Label20.Enabled = True
Label19.Enabled = True
Label21.Enabled = False
Label14.Enabled = False
Label13.Enabled = True
Label15.Enabled = True
Label5.Enabled = True
Label4.Enabled = False
Label6.Enabled = True
Label23.Enabled = False
Label22.Enabled = True
Label24.Enabled = False
Label17.Enabled = True
Label16.Enabled = False
Label18.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label9.Enabled = False
Label26.Enabled = True
Label25.Enabled = False
Label27.Enabled = True
End Sub

Public Sub easy6()

    Label65.Caption = 6
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 7
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = 2
    Label68.Caption = ""
    Label67.Caption = 9
    Label69.Caption = 3
    Label59.Caption = 2
    Label58.Caption = ""
    Label60.Caption = 5
    Label77.Caption = 6
    Label76.Caption = ""
    Label78.Caption = 8
    Label71.Caption = 4
    Label70.Caption = ""
    Label72.Caption = 2
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 3
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 3
    Label37.Caption = 8
    Label39.Caption = ""
    Label29.Caption = 9
    Label28.Caption = 5
    Label30.Caption = ""
    Label47.Caption = 2
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = 5
    Label40.Caption = ""
    Label42.Caption = 7
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = 3
    Label50.Caption = 9
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 1
    Label35.Caption = 8
    Label34.Caption = ""
    Label36.Caption = 2
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 5
    Label11.Caption = 2
    Label10.Caption = ""
    Label12.Caption = 9
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 8
    Label20.Caption = 1
    Label19.Caption = ""
    Label21.Caption = 4
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = 4
    Label5.Caption = 7
    Label4.Caption = ""
    Label6.Caption = 1
    Label23.Caption = ""
    Label22.Caption = 6
    Label24.Caption = ""
    Label17.Caption = 1
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 9
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 3
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = False
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = False
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub easy7()
    
    Label65.Caption = ""
    Label64.Caption = 4
    Label66.Caption = 3
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 7
    Label74.Caption = ""
    Label73.Caption = 1
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 9
    Label69.Caption = 1
    Label59.Caption = ""
    Label58.Caption = 3
    Label60.Caption = 2
    Label77.Caption = ""
    Label76.Caption = 6
    Label78.Caption = 8
    Label71.Caption = ""
    Label70.Caption = 2
    Label72.Caption = ""
    Label62.Caption = 1
    Label61.Caption = 9
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 4
    Label81.Caption = ""
    Label38.Caption = 6
    Label37.Caption = ""
    Label39.Caption = 8
    Label29.Caption = 5
    Label28.Caption = ""
    Label30.Caption = 9
    Label47.Caption = ""
    Label46.Caption = 2
    Label48.Caption = ""
    Label41.Caption = 4
    Label40.Caption = 7
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = 2
    Label33.Caption = 6
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 5
    Label44.Caption = ""
    Label43.Caption = 1
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = 4
    Label36.Caption = ""
    Label53.Caption = 6
    Label52.Caption = 9
    Label54.Caption = 3
    Label11.Caption = ""
    Label10.Caption = 6
    Label12.Caption = ""
    Label2.Caption = 8
    Label1.Caption = ""
    Label3.Caption = 3
    Label20.Caption = ""
    Label19.Caption = 5
    Label21.Caption = ""
    Label14.Caption = 1
    Label13.Caption = 5
    Label15.Caption = ""
    Label5.Caption = 9
    Label4.Caption = 6
    Label6.Caption = ""
    Label23.Caption = 8
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = 8
    Label18.Caption = ""
    Label8.Caption = 2
    Label7.Caption = ""
    Label9.Caption = 1
    Label26.Caption = 9
    Label25.Caption = 7
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = False
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = False
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = False
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = False
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = False
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = False
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = False
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = False
    Label25.Enabled = False
    Label27.Enabled = True
End Sub



Public Sub easy8()

    Label65.Caption = 3
    Label64.Caption = 6
    Label66.Caption = ""
    Label56.Caption = 1
    Label55.Caption = ""
    Label57.Caption = 7
    Label74.Caption = ""
    Label73.Caption = 8
    Label75.Caption = 9
    Label68.Caption = 4
    Label67.Caption = ""
    Label69.Caption = 1
    Label59.Caption = ""
    Label58.Caption = 2
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 7
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 9
    Label62.Caption = 3
    Label61.Caption = ""
    Label63.Caption = 6
    Label80.Caption = 4
    Label79.Caption = ""
    Label81.Caption = 2
    Label38.Caption = ""
    Label37.Caption = 7
    Label39.Caption = 5
    Label29.Caption = 8
    Label28.Caption = ""
    Label30.Caption = 9
    Label47.Caption = 2
    Label46.Caption = ""
    Label48.Caption = 4
    Label41.Caption = 9
    Label40.Caption = ""
    Label42.Caption = 3
    Label32.Caption = ""
    Label31.Caption = 1
    Label33.Caption = ""
    Label50.Caption = 8
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = 8
    Label43.Caption = ""
    Label45.Caption = 6
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 2
    Label53.Caption = ""
    Label52.Caption = 5
    Label54.Caption = ""
    Label11.Caption = 5
    Label10.Caption = ""
    Label12.Caption = 7
    Label2.Caption = 2
    Label1.Caption = ""
    Label3.Caption = 1
    Label20.Caption = 6
    Label19.Caption = ""
    Label21.Caption = 3
    Label14.Caption = 2
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 7
    Label6.Caption = ""
    Label23.Caption = 1
    Label22.Caption = ""
    Label24.Caption = 8
    Label17.Caption = 6
    Label16.Caption = 1
    Label18.Caption = 8
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 3
    Label26.Caption = ""
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = False
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = False
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = False
    Label16.Tag = False
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = False
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = False
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = False
    Label16.Enabled = False
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub easy9()

    Label65.Caption = ""
    Label64.Caption = 7
    Label66.Caption = ""
    Label56.Caption = 9
    Label55.Caption = 1
    Label57.Caption = ""
    Label74.Caption = 6
    Label73.Caption = ""
    Label75.Caption = 8
    Label68.Caption = 4
    Label67.Caption = ""
    Label69.Caption = 6
    Label59.Caption = 3
    Label58.Caption = 7
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 1
    Label78.Caption = 9
    Label71.Caption = 9
    Label70.Caption = ""
    Label72.Caption = 3
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 8
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 7
    Label38.Caption = ""
    Label37.Caption = 4
    Label39.Caption = 7
    Label29.Caption = ""
    Label28.Caption = 9
    Label30.Caption = 6
    Label47.Caption = 1
    Label46.Caption = ""
    Label48.Caption = 3
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 9
    Label32.Caption = ""
    Label31.Caption = 4
    Label33.Caption = ""
    Label50.Caption = 5
    Label49.Caption = ""
    Label51.Caption = 6
    Label44.Caption = 1
    Label43.Caption = ""
    Label45.Caption = 5
    Label35.Caption = ""
    Label34.Caption = 3
    Label36.Caption = ""
    Label53.Caption = 9
    Label52.Caption = 7
    Label54.Caption = ""
    Label11.Caption = 7
    Label10.Caption = 5
    Label12.Caption = ""
    Label2.Caption = 1
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = 3
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 2
    Label13.Caption = 9
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 5
    Label6.Caption = 3
    Label23.Caption = 7
    Label22.Caption = ""
    Label24.Caption = 1
    Label17.Caption = ""
    Label16.Caption = 3
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 7
    Label26.Caption = ""
    Label25.Caption = 9
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = False
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = False
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = False
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = False
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = False
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = False
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = False
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = False
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub easy10()
    Label65.Caption = ""
    Label64.Caption = 7
    Label66.Caption = 5
    Label56.Caption = 1
    Label55.Caption = ""
    Label57.Caption = 6
    Label74.Caption = ""
    Label73.Caption = 3
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 2
    Label69.Caption = ""
    Label59.Caption = 8
    Label58.Caption = ""
    Label60.Caption = 4
    Label77.Caption = 6
    Label76.Caption = ""
    Label78.Caption = 9
    Label71.Caption = 9
    Label70.Caption = ""
    Label72.Caption = 8
    Label62.Caption = 2
    Label61.Caption = ""
    Label63.Caption = 5
    Label80.Caption = 4
    Label79.Caption = ""
    Label81.Caption = 7
    Label38.Caption = ""
    Label37.Caption = 4
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = 5
    Label30.Caption = ""
    Label47.Caption = 9
    Label46.Caption = 6
    Label48.Caption = ""
    Label41.Caption = 6
    Label40.Caption = 3
    Label42.Caption = 7
    Label32.Caption = ""
    Label31.Caption = 4
    Label33.Caption = ""
    Label50.Caption = 5
    Label49.Caption = ""
    Label51.Caption = 1
    Label44.Caption = ""
    Label43.Caption = 8
    Label45.Caption = 9
    Label35.Caption = 6
    Label34.Caption = 2
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 7
    Label54.Caption = ""
    Label11.Caption = 7
    Label10.Caption = ""
    Label12.Caption = 6
    Label2.Caption = 3
    Label1.Caption = ""
    Label3.Caption = 2
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 5
    Label14.Caption = 8
    Label13.Caption = ""
    Label15.Caption = 4
    Label5.Caption = 5
    Label4.Caption = ""
    Label6.Caption = 7
    Label23.Caption = 2
    Label22.Caption = 9
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = 5
    Label18.Caption = ""
    Label8.Caption = 4
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 7
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = False
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = False
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = False
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = False
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Medium1()

    Label65.Caption = 7
    Label64.Caption = 2
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 9
    Label57.Caption = ""
    Label74.Caption = 4
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 3
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 4
    Label58.Caption = ""
    Label60.Caption = 8
    Label77.Caption = 9
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = 4
    Label70.Caption = ""
    Label72.Caption = 9
    Label62.Caption = ""
    Label61.Caption = 6
    Label63.Caption = ""
    Label80.Caption = 5
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 9
    Label37.Caption = ""
    Label39.Caption = 4
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 6
    Label46.Caption = 2
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 3
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 1
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = 5
    Label45.Caption = 6
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 8
    Label52.Caption = ""
    Label54.Caption = 9
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 7
    Label2.Caption = ""
    Label1.Caption = 3
    Label3.Caption = ""
    Label20.Caption = 1
    Label19.Caption = ""
    Label21.Caption = 6
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = 8
    Label5.Caption = 9
    Label4.Caption = ""
    Label6.Caption = 2
    Label23.Caption = ""
    Label22.Caption = 5
    Label24.Caption = 4
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 3
    Label8.Caption = ""
    Label7.Caption = 1
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 9
    Label27.Caption = 7
    
    
    Label65.Tag = False
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = False
End Sub

Public Sub Medium2()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = 4
    Label55.Caption = 2
    Label57.Caption = ""
    Label74.Caption = 7
    Label73.Caption = ""
    Label75.Caption = 6
    Label68.Caption = ""
    Label67.Caption = 2
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 8
    Label78.Caption = ""
    Label71.Caption = 9
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 5
    Label80.Caption = 4
    Label79.Caption = ""
    Label81.Caption = 2
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 4
    Label29.Caption = 7
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 2
    Label46.Caption = 6
    Label48.Caption = ""
    Label41.Caption = 6
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 1
    Label44.Caption = ""
    Label43.Caption = 3
    Label45.Caption = 5
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 6
    Label53.Caption = 9
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 3
    Label10.Caption = 9
    Label12.Caption = 6
    Label2.Caption = 5
    Label1.Caption = ""
    Label3.Caption = 1
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 4
    Label14.Caption = ""
    Label13.Caption = 4
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 3
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 5
    Label24.Caption = ""
    Label17.Caption = 7
    Label16.Caption = ""
    Label18.Caption = 8
    Label8.Caption = ""
    Label7.Caption = 6
    Label9.Caption = 4
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = False
    Label12.Tag = False
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Medium3()

    Label65.Caption = ""
    Label64.Caption = 4
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 2
    Label74.Caption = ""
    Label73.Caption = 3
    Label75.Caption = 6
    Label68.Caption = 7
    Label67.Caption = 2
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 3
    Label60.Caption = 1
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = 8
    Label71.Caption = ""
    Label70.Caption = 3
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 6
    Label63.Caption = ""
    Label80.Caption = 2
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 2
    Label37.Caption = ""
    Label39.Caption = 5
    Label29.Caption = ""
    Label28.Caption = 4
    Label30.Caption = ""
    Label47.Caption = 7
    Label46.Caption = ""
    Label48.Caption = 9
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = 6
    Label43.Caption = ""
    Label45.Caption = 4
    Label35.Caption = ""
    Label34.Caption = 8
    Label36.Caption = ""
    Label53.Caption = 3
    Label52.Caption = ""
    Label54.Caption = 1
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 7
    Label2.Caption = ""
    Label1.Caption = 1
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 5
    Label21.Caption = 3
    Label14.Caption = 8
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = 3
    Label4.Caption = 9
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 1
    Label24.Caption = 7
    Label17.Caption = 3
    Label16.Caption = 1
    Label18.Caption = ""
    Label8.Caption = 6
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 4
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = False
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = False
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Medium4()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = 4
    Label57.Caption = ""
    Label74.Caption = 3
    Label73.Caption = 7
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 2
    Label69.Caption = 9
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 5
    Label76.Caption = 6
    Label78.Caption = ""
    Label71.Caption = 7
    Label70.Caption = ""
    Label72.Caption = 5
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 6
    Label80.Caption = ""
    Label79.Caption = 4
    Label81.Caption = 1
    Label38.Caption = ""
    Label37.Caption = 4
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 5
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = 3
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 3
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 8
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = 6
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 1
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 2
    Label54.Caption = ""
    Label11.Caption = 2
    Label10.Caption = 7
    Label12.Caption = ""
    Label2.Caption = 3
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = 4
    Label19.Caption = ""
    Label21.Caption = 5
    Label14.Caption = ""
    Label13.Caption = 8
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = 1
    Label22.Caption = ""
    Label24.Caption = 6
    Label17.Caption = ""
    Label16.Caption = 9
    Label18.Caption = 6
    Label8.Caption = ""
    Label7.Caption = 8
    Label9.Caption = ""
    Label26.Caption = 7
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = False
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = False
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Medium5()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = 4
    Label55.Caption = 9
    Label57.Caption = ""
    Label74.Caption = 8
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 6
    Label69.Caption = ""
    Label59.Caption = 1
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 5
    Label76.Caption = 9
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 9
    Label62.Caption = 7
    Label61.Caption = 6
    Label63.Caption = ""
    Label80.Caption = 3
    Label79.Caption = ""
    Label81.Caption = 2
    Label38.Caption = ""
    Label37.Caption = 1
    Label39.Caption = ""
    Label29.Caption = 5
    Label28.Caption = ""
    Label30.Caption = 4
    Label47.Caption = 9
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 4
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 6
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 7
    Label35.Caption = 6
    Label34.Caption = ""
    Label36.Caption = 9
    Label53.Caption = ""
    Label52.Caption = 3
    Label54.Caption = ""
    Label11.Caption = 2
    Label10.Caption = ""
    Label12.Caption = 8
    Label2.Caption = ""
    Label1.Caption = 4
    Label3.Caption = 6
    Label20.Caption = 1
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 6
    Label13.Caption = 9
    Label15.Caption = 3
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 1
    Label23.Caption = ""
    Label22.Caption = 4
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 4
    Label8.Caption = ""
    Label7.Caption = 7
    Label9.Caption = 3
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = False
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = False
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = False
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = False
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Medium6()

    Label65.Caption = ""
    Label64.Caption = 4
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 1
    Label57.Caption = 8
    Label74.Caption = ""
    Label73.Caption = 3
    Label75.Caption = ""
    Label68.Caption = 3
    Label67.Caption = ""
    Label69.Caption = 6
    Label59.Caption = ""
    Label58.Caption = 9
    Label60.Caption = ""
    Label77.Caption = 2
    Label76.Caption = 5
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = 7
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 3
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 6
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 1
    Label29.Caption = 3
    Label28.Caption = ""
    Label30.Caption = 4
    Label47.Caption = 6
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = 6
    Label40.Caption = 8
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 9
    Label51.Caption = 5
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 5
    Label35.Caption = 9
    Label34.Caption = ""
    Label36.Caption = 6
    Label53.Caption = 7
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 5
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = 6
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 4
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 6
    Label15.Caption = 4
    Label5.Caption = ""
    Label4.Caption = 2
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 1
    Label17.Caption = ""
    Label16.Caption = 3
    Label18.Caption = ""
    Label8.Caption = 7
    Label7.Caption = 4
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Medium7()

    Label65.Caption = ""
    Label64.Caption = 9
    Label66.Caption = 7
    Label56.Caption = ""
    Label55.Caption = 4
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 3
    Label75.Caption = ""
    Label68.Caption = 5
    Label67.Caption = 3
    Label69.Caption = ""
    Label59.Caption = 8
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 6
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = 8
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 5
    Label79.Caption = ""
    Label81.Caption = 2
    Label38.Caption = 8
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = 7
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 4
    Label46.Caption = ""
    Label48.Caption = 3
    Label41.Caption = ""
    Label40.Caption = 5
    Label42.Caption = ""
    Label32.Caption = 3
    Label31.Caption = ""
    Label33.Caption = 4
    Label50.Caption = ""
    Label49.Caption = 1
    Label51.Caption = ""
    Label44.Caption = 4
    Label43.Caption = ""
    Label45.Caption = 3
    Label35.Caption = 6
    Label34.Caption = ""
    Label36.Caption = 9
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 5
    Label11.Caption = 1
    Label10.Caption = ""
    Label12.Caption = 2
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 5
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = 5
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 3
    Label23.Caption = ""
    Label22.Caption = 4
    Label24.Caption = 8
    Label17.Caption = ""
    Label16.Caption = 4
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 7
    Label9.Caption = ""
    Label26.Caption = 9
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Medium8()
    
    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 1
    Label74.Caption = ""
    Label73.Caption = 7
    Label75.Caption = ""
    Label68.Caption = 7
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 6
    Label58.Caption = 9
    Label60.Caption = 4
    Label77.Caption = 8
    Label76.Caption = 5
    Label78.Caption = ""
    Label71.Caption = 5
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 3
    Label80.Caption = 2
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 5
    Label39.Caption = ""
    Label29.Caption = 3
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 1
    Label46.Caption = 9
    Label48.Caption = ""
    Label41.Caption = 2
    Label40.Caption = 3
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 6
    Label51.Caption = 8
    Label44.Caption = ""
    Label43.Caption = 7
    Label45.Caption = 4
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 8
    Label53.Caption = ""
    Label52.Caption = 3
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 7
    Label2.Caption = 2
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 5
    Label14.Caption = ""
    Label13.Caption = 1
    Label15.Caption = 5
    Label5.Caption = 4
    Label4.Caption = 8
    Label6.Caption = 6
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 3
    Label17.Caption = ""
    Label16.Caption = 6
    Label18.Caption = ""
    Label8.Caption = 7
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 9
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = False
    Label60.Tag = False
    Label77.Tag = False
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = False
    Label5.Tag = False
    Label4.Tag = False
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = False
    Label60.Enabled = False
    Label77.Enabled = False
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = False
    Label5.Enabled = False
    Label4.Enabled = False
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Medium9()

    Label65.Caption = 1
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = 9
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 7
    Label75.Caption = 3
    Label68.Caption = ""
    Label67.Caption = 2
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 1
    Label60.Caption = ""
    Label77.Caption = 6
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 9
    Label62.Caption = ""
    Label61.Caption = 6
    Label63.Caption = 8
    Label80.Caption = ""
    Label79.Caption = 1
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 1
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = 2
    Label30.Caption = 9
    Label47.Caption = 7
    Label46.Caption = 4
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 8
    Label42.Caption = 7
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 5
    Label49.Caption = 9
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = 6
    Label45.Caption = 2
    Label35.Caption = 4
    Label34.Caption = 7
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 8
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 9
    Label12.Caption = ""
    Label2.Caption = 6
    Label1.Caption = 4
    Label3.Caption = ""
    Label20.Caption = 1
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = 4
    Label5.Caption = ""
    Label4.Caption = 8
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 2
    Label24.Caption = ""
    Label17.Caption = 5
    Label16.Caption = 7
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 3
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 8
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Medium10()

    Label65.Caption = 7
    Label64.Caption = ""
    Label66.Caption = 4
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 1
    Label74.Caption = ""
    Label73.Caption = 8
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = 6
    Label59.Caption = ""
    Label58.Caption = 5
    Label60.Caption = ""
    Label77.Caption = 4
    Label76.Caption = 3
    Label78.Caption = ""
    Label71.Caption = 2
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 4
    Label61.Caption = ""
    Label63.Caption = 9
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 5
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 3
    Label29.Caption = ""
    Label28.Caption = 7
    Label30.Caption = ""
    Label47.Caption = 8
    Label46.Caption = ""
    Label48.Caption = 4
    Label41.Caption = ""
    Label40.Caption = 6
    Label42.Caption = 9
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 2
    Label49.Caption = 5
    Label51.Caption = ""
    Label44.Caption = 8
    Label43.Caption = ""
    Label45.Caption = 1
    Label35.Caption = ""
    Label34.Caption = 4
    Label36.Caption = ""
    Label53.Caption = 6
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 3
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 6
    Label1.Caption = ""
    Label3.Caption = 4
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 9
    Label14.Caption = ""
    Label13.Caption = 4
    Label15.Caption = 5
    Label5.Caption = ""
    Label4.Caption = 3
    Label6.Caption = ""
    Label23.Caption = 7
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = 8
    Label18.Caption = ""
    Label8.Caption = 5
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 3
    Label25.Caption = ""
    Label27.Caption = 1
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Hard1()
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = 7
    Label7.Caption = 6
    Label8.Caption = 5
    Label9.Caption = ""
    Label10.Caption = 8
    Label11.Caption = 2
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = 1
    Label16.Caption = ""
    Label17.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    Label22.Caption = ""
    Label23.Caption = ""
    Label24.Caption = 3
    Label25.Caption = 1
    Label26.Caption = ""
    Label27.Caption = ""
    Label28.Caption = 3
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""
    Label32.Caption = ""
    Label33.Caption = ""
    Label34.Caption = 1
    Label35.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = 4
    Label40.Caption = ""
    Label41.Caption = ""
    Label42.Caption = 9
    Label43.Caption = ""
    Label44.Caption = ""
    Label45.Caption = ""
    Label46.Caption = ""
    Label47.Caption = ""
    Label48.Caption = ""
    Label49.Caption = ""
    Label50.Caption = 5
    Label51.Caption = ""
    Label52.Caption = ""
    Label53.Caption = 6
    Label54.Caption = ""
    Label55.Caption = 9
    Label56.Caption = ""
    Label57.Caption = 2
    Label58.Caption = ""
    Label59.Caption = 8
    Label60.Caption = ""
    Label61.Caption = ""
    Label62.Caption = ""
    Label63.Caption = ""
    Label64.Caption = 7
    Label65.Caption = ""
    Label66.Caption = ""
    Label67.Caption = ""
    Label68.Caption = 5
    Label69.Caption = ""
    Label70.Caption = ""
    Label71.Caption = ""
    Label72.Caption = ""
    Label73.Caption = ""
    Label74.Caption = ""
    Label75.Caption = ""
    Label76.Caption = ""
    Label77.Caption = 3
    Label78.Caption = ""
    Label79.Caption = 9
    Label80.Caption = ""
    Label81.Caption = 4


    Label1.Tag = True
    Label2.Tag = True
    Label3.Tag = True
    Label4.Tag = True
    Label5.Tag = True
    Label6.Tag = False
    Label7.Tag = False
    Label8.Tag = False
    Label9.Tag = True
    Label10.Tag = False
    Label11.Tag = False
    Label12.Tag = True
    Label13.Tag = True
    Label14.Tag = True
    Label15.Tag = False
    Label16.Tag = True
    Label17.Tag = True
    Label18.Tag = True
    Label19.Tag = True
    Label20.Tag = True
    Label21.Tag = True
    Label22.Tag = True
    Label23.Tag = True
    Label24.Tag = False
    Label25.Tag = False
    Label26.Tag = True
    Label27.Tag = True
    Label28.Tag = False
    Label29.Tag = True
    Label30.Tag = True
    Label31.Tag = True
    Label32.Tag = True
    Label33.Tag = True
    Label34.Tag = False
    Label35.Tag = True
    Label36.Tag = True
    Label37.Tag = True
    Label38.Tag = True
    Label39.Tag = False
    Label40.Tag = True
    Label41.Tag = True
    Label42.Tag = False
    Label43.Tag = True
    Label44.Tag = True
    Label45.Tag = True
    Label46.Tag = True
    Label47.Tag = True
    Label48.Tag = True
    Label49.Tag = True
    Label50.Tag = False
    Label51.Tag = True
    Label52.Tag = True
    Label53.Tag = False
    Label54.Tag = True
    Label55.Tag = False
    Label56.Tag = True
    Label57.Tag = False
    Label58.Tag = True
    Label59.Tag = False
    Label60.Tag = True
    Label61.Tag = True
    Label62.Tag = True
    Label63.Tag = True
    Label64.Tag = False
    Label65.Tag = True
    Label66.Tag = True
    Label67.Tag = True
    Label68.Tag = False
    Label69.Tag = True
    Label70.Tag = True
    Label71.Tag = True
    Label72.Tag = True
    Label73.Tag = True
    Label74.Tag = True
    Label75.Tag = True
    Label76.Tag = True
    Label77.Tag = False
    Label78.Tag = True
    Label79.Tag = False
    Label80.Tag = True
    Label81.Tag = False

    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = False
    Label7.Enabled = False
    Label8.Enabled = False
    Label9.Enabled = True
    Label10.Enabled = False
    Label11.Enabled = False
    Label12.Enabled = True
    Label13.Enabled = True
    Label14.Enabled = True
    Label15.Enabled = False
    Label16.Enabled = True
    Label17.Enabled = True
    Label18.Enabled = True
    Label19.Enabled = True
    Label20.Enabled = True
    Label21.Enabled = True
    Label22.Enabled = True
    Label23.Enabled = True
    Label24.Enabled = False
    Label25.Enabled = False
    Label26.Enabled = True
    Label27.Enabled = True
    Label28.Enabled = False
    Label29.Enabled = True
    Label30.Enabled = True
    Label31.Enabled = True
    Label32.Enabled = True
    Label33.Enabled = True
    Label34.Enabled = False
    Label35.Enabled = True
    Label36.Enabled = True
    Label37.Enabled = True
    Label38.Enabled = True
    Label39.Enabled = False
    Label40.Enabled = True
    Label41.Enabled = True
    Label42.Enabled = False
    Label43.Enabled = True
    Label44.Enabled = True
    Label45.Enabled = True
    Label46.Enabled = True
    Label47.Enabled = True
    Label48.Enabled = True
    Label49.Enabled = True
    Label50.Enabled = False
    Label51.Enabled = True
    Label52.Enabled = True
    Label53.Enabled = False
    Label54.Enabled = True
    Label55.Enabled = False
    Label56.Enabled = True
    Label57.Enabled = False
    Label58.Enabled = True
    Label59.Enabled = False
    Label60.Enabled = True
    Label61.Enabled = True
    Label62.Enabled = True
    Label63.Enabled = True
    Label64.Enabled = False
    Label65.Enabled = True
    Label66.Enabled = True
    Label67.Enabled = True
    Label68.Enabled = False
    Label69.Enabled = True
    Label70.Enabled = True
    Label71.Enabled = True
    Label72.Enabled = True
    Label73.Enabled = True
    Label74.Enabled = True
    Label75.Enabled = True
    Label76.Enabled = True
    Label77.Enabled = False
    Label78.Enabled = True
    Label79.Enabled = False
    Label80.Enabled = True
    Label81.Enabled = False

End Sub

Public Sub Hard2()
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = 1
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = 5
    Label7.Caption = ""
    Label8.Caption = 4
    Label9.Caption = ""
    Label10.Caption = 2
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = ""
    Label16.Caption = 8
    Label17.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    Label20.Caption = ""
    Label21.Caption = 6
    Label22.Caption = ""
    Label23.Caption = ""
    Label24.Caption = 7
    Label25.Caption = ""
    Label26.Caption = ""
    Label27.Caption = ""
    Label28.Caption = 2
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""
    Label32.Caption = ""
    Label33.Caption = ""
    Label34.Caption = 6
    Label35.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = 9
    Label40.Caption = ""
    Label41.Caption = ""
    Label42.Caption = 4
    Label43.Caption = ""
    Label44.Caption = ""
    Label45.Caption = 7
    Label46.Caption = ""
    Label47.Caption = 8
    Label48.Caption = ""
    Label49.Caption = ""
    Label50.Caption = 1
    Label51.Caption = ""
    Label52.Caption = ""
    Label53.Caption = 3
    Label54.Caption = ""
    Label55.Caption = ""
    Label56.Caption = ""
    Label57.Caption = 3
    Label58.Caption = ""
    Label59.Caption = 9
    Label60.Caption = ""
    Label61.Caption = ""
    Label62.Caption = 8
    Label63.Caption = ""
    Label64.Caption = ""
    Label65.Caption = ""
    Label66.Caption = ""
    Label67.Caption = ""
    Label68.Caption = 3
    Label69.Caption = ""
    Label70.Caption = ""
    Label71.Caption = 6
    Label72.Caption = ""
    Label73.Caption = 5
    Label74.Caption = ""
    Label75.Caption = ""
    Label76.Caption = ""
    Label77.Caption = ""
    Label78.Caption = ""
    Label79.Caption = 4
    Label80.Caption = ""
    Label81.Caption = ""
 
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = False
    Label4.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = False
    Label7.Enabled = True
    Label8.Enabled = False
    Label9.Enabled = True
    Label10.Enabled = False
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = True
    Label14.Enabled = True
    Label15.Enabled = True
    Label16.Enabled = False
    Label17.Enabled = True
    Label18.Enabled = True
    Label19.Enabled = True
    Label20.Enabled = True
    Label21.Enabled = False
    Label22.Enabled = True
    Label23.Enabled = True
    Label24.Enabled = False
    Label25.Enabled = True
    Label26.Enabled = True
    Label27.Enabled = True
    Label28.Enabled = False
    Label29.Enabled = True
    Label30.Enabled = True
    Label31.Enabled = True
    Label32.Enabled = True
    Label33.Enabled = True
    Label34.Enabled = False
    Label35.Enabled = True
    Label36.Enabled = True
    Label37.Enabled = True
    Label38.Enabled = True
    Label39.Enabled = False
    Label40.Enabled = True
    Label41.Enabled = True
    Label42.Enabled = False
    Label43.Enabled = True
    Label44.Enabled = True
    Label45.Enabled = False
    Label46.Enabled = True
    Label47.Enabled = False
    Label48.Enabled = True
    Label49.Enabled = True
    Label50.Enabled = False
    Label51.Enabled = True
    Label52.Enabled = True
    Label53.Enabled = False
    Label54.Enabled = True
    Label55.Enabled = True
    Label56.Enabled = True
    Label57.Enabled = False
    Label58.Enabled = True
    Label59.Enabled = False
    Label60.Enabled = True
    Label61.Enabled = True
    Label62.Enabled = False
    Label63.Enabled = True
    Label64.Enabled = True
    Label65.Enabled = True
    Label66.Enabled = True
    Label67.Enabled = True
    Label68.Enabled = False
    Label69.Enabled = True
    Label70.Enabled = True
    Label71.Enabled = False
    Label72.Enabled = True
    Label73.Enabled = False
    Label74.Enabled = True
    Label75.Enabled = True
    Label76.Enabled = True
    Label77.Enabled = True
    Label78.Enabled = True
    Label79.Enabled = False
    Label80.Enabled = True
    Label81.Enabled = True
 
    Label1.Tag = True
    Label2.Tag = True
    Label3.Tag = False
    Label4.Tag = True
    Label5.Tag = True
    Label6.Tag = False
    Label7.Tag = True
    Label8.Tag = False
    Label9.Tag = True
    Label10.Tag = False
    Label11.Tag = True
    Label12.Tag = True
    Label13.Tag = True
    Label14.Tag = True
    Label15.Tag = True
    Label16.Tag = False
    Label17.Tag = True
    Label18.Tag = True
    Label19.Tag = True
    Label20.Tag = True
    Label21.Tag = False
    Label22.Tag = True
    Label23.Tag = True
    Label24.Tag = False
    Label25.Tag = True
    Label26.Tag = True
    Label27.Tag = True
    Label28.Tag = False
    Label29.Tag = True
    Label30.Tag = True
    Label31.Tag = True
    Label32.Tag = True
    Label33.Tag = True
    Label34.Tag = False
    Label35.Tag = True
    Label36.Tag = True
    Label37.Tag = True
    Label38.Tag = True
    Label39.Tag = False
    Label40.Tag = True
    Label41.Tag = True
    Label42.Tag = False
    Label43.Tag = True
    Label44.Tag = True
    Label45.Tag = False
    Label46.Tag = True
    Label47.Tag = False
    Label48.Tag = True
    Label49.Tag = True
    Label50.Tag = False
    Label51.Tag = True
    Label52.Tag = True
    Label53.Tag = False
    Label54.Tag = True
    Label55.Tag = True
    Label56.Tag = True
    Label57.Tag = False
    Label58.Tag = True
    Label59.Tag = False
    Label60.Tag = True
    Label61.Tag = True
    Label62.Tag = False
    Label63.Tag = True
    Label64.Tag = True
    Label65.Tag = True
    Label66.Tag = True
    Label67.Tag = True
    Label68.Tag = False
    Label69.Tag = True
    Label70.Tag = True
    Label71.Tag = False
    Label72.Tag = True
    Label73.Tag = False
    Label74.Tag = True
    Label75.Tag = True
    Label76.Tag = True
    Label77.Tag = True
    Label78.Tag = True
    Label79.Tag = False
    Label80.Tag = True
    Label81.Tag = True
 
End Sub

Public Sub Hard3()
    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 9
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 8
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 2
    Label78.Caption = 1
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 4
    Label61.Caption = 3
    Label63.Caption = 5
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 7
    Label46.Caption = 8
    Label48.Caption = 6
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = 1
    Label31.Caption = ""
    Label33.Caption = 4
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = 6
    Label43.Caption = 2
    Label45.Caption = 9
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 6
    Label1.Caption = 7
    Label3.Caption = 8
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 4
    Label13.Caption = 1
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 5
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 3
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = False
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = False
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = False
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = False
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Hard4()
    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 2
    Label67.Caption = 6
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 5
    Label78.Caption = 1
    Label71.Caption = 9
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 1
    Label61.Caption = 4
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 7
    Label39.Caption = 3
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = 6
    Label31.Caption = ""
    Label33.Caption = 2
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 4
    Label52.Caption = 9
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = 3
    Label3.Caption = 9
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 8
    Label14.Caption = 5
    Label13.Caption = 1
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 3
    Label24.Caption = 6
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 7
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = False
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = False
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Hard5()

    Label65.Caption = ""
    Label64.Caption = 9
    Label66.Caption = ""
    Label56.Caption = 1
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 8
    Label73.Caption = 7
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 3
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 5
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = 2
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 8
    Label62.Caption = ""
    Label61.Caption = 2
    Label63.Caption = ""
    Label80.Caption = 1
    Label79.Caption = ""
    Label81.Caption = 4
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 5
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 4
    Label46.Caption = ""
    Label48.Caption = 7
    Label41.Caption = 3
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = 4
    Label31.Caption = ""
    Label33.Caption = 8
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 9
    Label44.Caption = 8
    Label43.Caption = ""
    Label45.Caption = 2
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 6
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 6
    Label10.Caption = ""
    Label12.Caption = 4
    Label2.Caption = ""
    Label1.Caption = 7
    Label3.Caption = ""
    Label20.Caption = 5
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 1
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 8
    Label6.Caption = 4
    Label23.Caption = ""
    Label22.Caption = 2
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = 8
    Label18.Caption = 3
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 9
    Label26.Caption = ""
    Label25.Caption = 4
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Hard6()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 3
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 9
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = 6
    Label68.Caption = 6
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 7
    Label60.Caption = ""
    Label77.Caption = 5
    Label76.Caption = ""
    Label78.Caption = 4
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 7
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 2
    Label80.Caption = 8
    Label79.Caption = ""
    Label81.Caption = 1
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = 9
    Label28.Caption = 1
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = 8
    Label48.Caption = 2
    Label41.Caption = ""
    Label40.Caption = 4
    Label42.Caption = 9
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 1
    Label49.Caption = 6
    Label51.Caption = ""
    Label44.Caption = 2
    Label43.Caption = 8
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = 3
    Label36.Caption = 6
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 7
    Label12.Caption = 6
    Label2.Caption = 5
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = 4
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 8
    Label13.Caption = ""
    Label15.Caption = 2
    Label5.Caption = ""
    Label4.Caption = 4
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 3
    Label17.Caption = 5
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = 6
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 7
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = False
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Hard7()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 9
    Label56.Caption = ""
    Label55.Caption = 8
    Label57.Caption = 6
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 5
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 9
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 1
    Label78.Caption = 8
    Label71.Caption = ""
    Label70.Caption = 3
    Label72.Caption = ""
    Label62.Caption = 5
    Label61.Caption = ""
    Label63.Caption = 7
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 6
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = 8
    Label28.Caption = ""
    Label30.Caption = 3
    Label47.Caption = 1
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = 3
    Label40.Caption = ""
    Label42.Caption = 7
    Label32.Caption = 4
    Label31.Caption = ""
    Label33.Caption = 1
    Label50.Caption = 5
    Label49.Caption = ""
    Label51.Caption = 9
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 1
    Label35.Caption = 2
    Label34.Caption = ""
    Label36.Caption = 5
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 3
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 1
    Label1.Caption = 3
    Label3.Caption = 8
    Label20.Caption = ""
    Label19.Caption = 7
    Label21.Caption = ""
    Label14.Caption = 8
    Label13.Caption = 1
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 9
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 2
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = 6
    Label7.Caption = 5
    Label9.Caption = ""
    Label26.Caption = 4
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = False
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = False
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Hard8()

    Label65.Caption = 8
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 6
    Label57.Caption = ""
    Label74.Caption = 7
    Label73.Caption = ""
    Label75.Caption = 5
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 9
    Label58.Caption = ""
    Label60.Caption = 4
    Label77.Caption = ""
    Label76.Caption = 3
    Label78.Caption = 1
    Label71.Caption = 1
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 3
    Label80.Caption = ""
    Label79.Caption = 2
    Label81.Caption = ""
    Label38.Caption = 3
    Label37.Caption = 6
    Label39.Caption = ""
    Label29.Caption = 4
    Label28.Caption = ""
    Label30.Caption = 2
    Label47.Caption = ""
    Label46.Caption = 9
    Label48.Caption = ""
    Label41.Caption = 2
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = 7
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 3
    Label44.Caption = ""
    Label43.Caption = 4
    Label45.Caption = ""
    Label35.Caption = 6
    Label34.Caption = ""
    Label36.Caption = 8
    Label53.Caption = ""
    Label52.Caption = 1
    Label54.Caption = 7
    Label11.Caption = ""
    Label10.Caption = 5
    Label12.Caption = ""
    Label2.Caption = 1
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 6
    Label14.Caption = 7
    Label13.Caption = 1
    Label15.Caption = ""
    Label5.Caption = 3
    Label4.Caption = ""
    Label6.Caption = 6
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = 9
    Label16.Caption = ""
    Label18.Caption = 6
    Label8.Caption = ""
    Label7.Caption = 2
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 4
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = False
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = False
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Hard9()

    Label65.Caption = ""
    Label64.Caption = 3
    Label66.Caption = ""
    Label56.Caption = 8
    Label55.Caption = 2
    Label57.Caption = 6
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = 9
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 4
    Label60.Caption = ""
    Label77.Caption = 8
    Label76.Caption = ""
    Label78.Caption = 6
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 5
    Label62.Caption = ""
    Label61.Caption = 3
    Label63.Caption = ""
    Label80.Caption = 2
    Label79.Caption = 1
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 9
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 3
    Label47.Caption = 1
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = 2
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 5
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 4
    Label35.Caption = 6
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 8
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 1
    Label12.Caption = 2
    Label2.Caption = ""
    Label1.Caption = 6
    Label3.Caption = ""
    Label20.Caption = 4
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 5
    Label13.Caption = ""
    Label15.Caption = 8
    Label5.Caption = ""
    Label4.Caption = 7
    Label6.Caption = 9
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = 6
    Label18.Caption = ""
    Label8.Caption = 5
    Label7.Caption = ""
    Label9.Caption = 4
    Label26.Caption = ""
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Hard10()
 
    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 1
    Label56.Caption = ""
    Label55.Caption = 9
    Label57.Caption = 6
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 6
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = 7
    Label77.Caption = 9
    Label76.Caption = ""
    Label78.Caption = 5
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 4
    Label62.Caption = ""
    Label61.Caption = 1
    Label63.Caption = ""
    Label80.Caption = 8
    Label79.Caption = 6
    Label81.Caption = 2
    Label38.Caption = ""
    Label37.Caption = 7
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = 5
    Label30.Caption = ""
    Label47.Caption = 2
    Label46.Caption = ""
    Label48.Caption = 8
    Label41.Caption = ""
    Label40.Caption = 6
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 9
    Label51.Caption = ""
    Label44.Caption = 1
    Label43.Caption = ""
    Label45.Caption = 8
    Label35.Caption = ""
    Label34.Caption = 4
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 3
    Label54.Caption = ""
    Label11.Caption = 8
    Label10.Caption = 1
    Label12.Caption = 6
    Label2.Caption = ""
    Label1.Caption = 7
    Label3.Caption = ""
    Label20.Caption = 4
    Label19.Caption = 5
    Label21.Caption = ""
    Label14.Caption = 9
    Label13.Caption = ""
    Label15.Caption = 2
    Label5.Caption = 5
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 7
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = 1
    Label7.Caption = 3
    Label9.Caption = ""
    Label26.Caption = 6
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = False
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = False
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = False
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional1()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 2
    Label56.Caption = 5
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 8
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 1
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 3
    Label60.Caption = ""
    Label77.Caption = 6
    Label76.Caption = ""
    Label78.Caption = 7
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 8
    Label62.Caption = ""
    Label61.Caption = 4
    Label63.Caption = 6
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 2
    Label38.Caption = ""
    Label37.Caption = 4
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = 1
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 3
    Label44.Caption = ""
    Label43.Caption = 9
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 2
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 6
    Label12.Caption = ""
    Label2.Caption = 2
    Label1.Caption = 8
    Label3.Caption = ""
    Label20.Caption = 4
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 4
    Label13.Caption = ""
    Label15.Caption = 5
    Label5.Caption = ""
    Label4.Caption = 9
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 6
    Label24.Caption = 1
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 9
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 1
    Label26.Caption = 8
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional2()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 5
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 1
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 5
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 2
    Label62.Caption = 4
    Label61.Caption = ""
    Label63.Caption = 1
    Label80.Caption = ""
    Label79.Caption = 6
    Label81.Caption = 7
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 5
    Label29.Caption = 1
    Label28.Caption = ""
    Label30.Caption = 3
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 4
    Label42.Caption = 7
    Label32.Caption = 8
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 1
    Label49.Caption = 9
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 2
    Label53.Caption = 6
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 2
    Label12.Caption = ""
    Label2.Caption = 6
    Label1.Caption = ""
    Label3.Caption = 8
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 5
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 1
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 3
    Label24.Caption = 6
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 3
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 2
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional3()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 1
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 8
    Label75.Caption = ""
    Label68.Caption = 8
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 2
    Label58.Caption = ""
    Label60.Caption = 7
    Label77.Caption = ""
    Label76.Caption = 9
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 6
    Label62.Caption = ""
    Label61.Caption = 5
    Label63.Caption = ""
    Label80.Caption = 7
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 4
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = 8
    Label30.Caption = ""
    Label47.Caption = 6
    Label46.Caption = 1
    Label48.Caption = ""
    Label41.Caption = 9
    Label40.Caption = 6
    Label42.Caption = ""
    Label32.Caption = 3
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 4
    Label51.Caption = 5
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 8
    Label35.Caption = ""
    Label34.Caption = 6
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 7
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 2
    Label2.Caption = ""
    Label1.Caption = 9
    Label3.Caption = ""
    Label20.Caption = 4
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 8
    Label15.Caption = ""
    Label5.Caption = 4
    Label4.Caption = ""
    Label6.Caption = 3
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 1
    Label17.Caption = ""
    Label16.Caption = 7
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 3
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional4()

    Label65.Caption = ""
    Label64.Caption = 6
    Label66.Caption = ""
    Label56.Caption = 2
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 9
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 5
    Label69.Caption = ""
    Label59.Caption = 9
    Label58.Caption = 6
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 3
    Label61.Caption = ""
    Label63.Caption = 8
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 6
    Label38.Caption = 4
    Label37.Caption = 2
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 9
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 3
    Label42.Caption = ""
    Label32.Caption = 7
    Label31.Caption = ""
    Label33.Caption = 2
    Label50.Caption = ""
    Label49.Caption = 4
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 6
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 5
    Label54.Caption = 3
    Label11.Caption = 1
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 8
    Label1.Caption = ""
    Label3.Caption = 3
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 4
    Label14.Caption = ""
    Label13.Caption = 8
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 9
    Label6.Caption = 7
    Label23.Caption = ""
    Label22.Caption = 3
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = 9
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 6
    Label26.Caption = ""
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = False
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = False
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = False
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = False
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Professional5()

    Label65.Caption = 8
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = 2
    Label55.Caption = 4
    Label57.Caption = 6
    Label74.Caption = 3
    Label73.Caption = 1
    Label75.Caption = ""
    Label68.Caption = 2
    Label67.Caption = ""
    Label69.Caption = 3
    Label59.Caption = 5
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = 4
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 2
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 9
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 3
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = 6
    Label40.Caption = ""
    Label42.Caption = 4
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 9
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 8
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 4
    Label52.Caption = ""
    Label54.Caption = 7
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = 5
    Label19.Caption = ""
    Label21.Caption = 6
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = 9
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 3
    Label17.Caption = ""
    Label16.Caption = 4
    Label18.Caption = 8
    Label8.Caption = 6
    Label7.Caption = 7
    Label9.Caption = 5
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 2
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = False
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = False
    Label18.Tag = False
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = False
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = False
    Label18.Enabled = False
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Professional6()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 2
    Label56.Caption = ""
    Label55.Caption = 1
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 3
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 4
    Label58.Caption = ""
    Label60.Caption = 6
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = 8
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 2
    Label61.Caption = ""
    Label63.Caption = 7
    Label80.Caption = 3
    Label79.Caption = ""
    Label81.Caption = 1
    Label38.Caption = 4
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = 8
    Label28.Caption = ""
    Label30.Caption = 1
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = 9
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = 6
    Label31.Caption = ""
    Label33.Caption = 2
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 3
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 3
    Label34.Caption = ""
    Label36.Caption = 5
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 8
    Label11.Caption = 5
    Label10.Caption = ""
    Label12.Caption = 6
    Label2.Caption = 1
    Label1.Caption = ""
    Label3.Caption = 9
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 4
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = 7
    Label4.Caption = ""
    Label6.Caption = 8
    Label23.Caption = 1
    Label22.Caption = ""
    Label24.Caption = 9
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 3
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional7()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 1
    Label56.Caption = ""
    Label55.Caption = 7
    Label57.Caption = 8
    Label74.Caption = 2
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 2
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 1
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = 8
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 2
    Label63.Caption = ""
    Label80.Caption = 7
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 6
    Label39.Caption = ""
    Label29.Caption = 9
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 3
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 3
    Label42.Caption = 4
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 9
    Label49.Caption = 7
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 8
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 2
    Label53.Caption = ""
    Label52.Caption = 5
    Label54.Caption = ""
    Label11.Caption = 8
    Label10.Caption = ""
    Label12.Caption = 2
    Label2.Caption = ""
    Label1.Caption = 6
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 9
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = 5
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 9
    Label23.Caption = ""
    Label22.Caption = 3
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 9
    Label8.Caption = 1
    Label7.Caption = 8
    Label9.Caption = ""
    Label26.Caption = 4
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional8()

    Label65.Caption = 4
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 5
    Label57.Caption = ""
    Label74.Caption = 8
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 6
    Label69.Caption = ""
    Label59.Caption = 3
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 9
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = 3
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 1
    Label79.Caption = 5
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 8
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = 4
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = 9
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 1
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 7
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = 2
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = 8
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 6
    Label54.Caption = 5
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 8
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 2
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 7
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 6
    Label23.Caption = ""
    Label22.Caption = 4
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 6
    Label8.Caption = ""
    Label7.Caption = 1
    Label9.Caption = ""
    Label26.Caption = 5
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Professional9()

    Label65.Caption = 6
    Label64.Caption = ""
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = 9
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 7
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 1
    Label60.Caption = ""
    Label77.Caption = 4
    Label76.Caption = ""
    Label78.Caption = 2
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 4
    Label62.Caption = 5
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 7
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = 3
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 4
    Label47.Caption = ""
    Label46.Caption = 2
    Label48.Caption = ""
    Label41.Caption = 8
    Label40.Caption = ""
    Label42.Caption = 1
    Label32.Caption = 2
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 5
    Label44.Caption = ""
    Label43.Caption = 7
    Label45.Caption = ""
    Label35.Caption = 1
    Label34.Caption = ""
    Label36.Caption = 6
    Label53.Caption = ""
    Label52.Caption = 4
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 6
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 5
    Label20.Caption = 2
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 2
    Label13.Caption = ""
    Label15.Caption = 9
    Label5.Caption = ""
    Label4.Caption = 6
    Label6.Caption = ""
    Label23.Caption = 7
    Label22.Caption = ""
    Label24.Caption = 4
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 4
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 8
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Professional10()

    Label65.Caption = 9
    Label64.Caption = 2
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 6
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 5
    Label69.Caption = ""
    Label59.Caption = 4
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 6
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 1
    Label63.Caption = 3
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 9
    Label38.Caption = ""
    Label37.Caption = 9
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 2
    Label47.Caption = ""
    Label46.Caption = 4
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 1
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 8
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = 3
    Label43.Caption = 7
    Label45.Caption = ""
    Label35.Caption = 5
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 9
    Label54.Caption = ""
    Label11.Caption = 5
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 3
    Label1.Caption = ""
    Label3.Caption = 4
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 1
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = 6
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 5
    Label23.Caption = ""
    Label22.Caption = 8
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = 1
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 4
    Label25.Caption = 5
    Label27.Caption = ""
    
    
    Label65.Tag = False
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = False
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Impossible1()

    Label65.Caption = ""
    Label64.Caption = 7
    Label66.Caption = ""
    Label56.Caption = 9
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 8
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = 3
    Label71.Caption = ""
    Label70.Caption = 2
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 8
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 9
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 3
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 7
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 8
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 5
    Label53.Caption = ""
    Label52.Caption = 6
    Label54.Caption = ""
    Label11.Caption = 9
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 3
    Label1.Caption = 1
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 5
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 6
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 8
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Impossible2()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 8
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 3
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = 1
    Label68.Caption = ""
    Label67.Caption = 1
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 8
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 7
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 5
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 9
    Label81.Caption = ""
    Label38.Caption = 6
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 4
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 2
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 8
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = 9
    Label45.Caption = ""
    Label35.Caption = 8
    Label34.Caption = ""
    Label36.Caption = 1
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 3
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = 2
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 5
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 9
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 6
    Label24.Caption = ""
    Label17.Caption = 8
    Label16.Caption = 6
    Label18.Caption = ""
    Label8.Caption = 7
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 4
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Impossible3()

    Label65.Caption = ""
    Label64.Caption = 1
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 7
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = 5
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 4
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 3
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 8
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = 8
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 3
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 2
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 9
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = 4
    Label45.Caption = 5
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 7
    Label10.Caption = 8
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 1
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 6
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 9
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 1
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 5
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Impossible4()

    Label65.Caption = ""
    Label64.Caption = 1
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 7
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 4
    Label76.Caption = ""
    Label78.Caption = 6
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 3
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 8
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = 8
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = 3
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 2
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 9
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = 4
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 6
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = 8
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 1
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = 6
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 9
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 1
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 5
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 2
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Impossible5()

    Label65.Caption = ""
    Label64.Caption = 3
    Label66.Caption = 2
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 9
    Label74.Caption = ""
    Label73.Caption = 1
    Label75.Caption = 5
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 6
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 6
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = 7
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 4
    Label38.Caption = ""
    Label37.Caption = 9
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 5
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 8
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 3
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 4
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 2
    Label54.Caption = ""
    Label11.Caption = 3
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 2
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = 1
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 2
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = 5
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = 1
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 3
    Label25.Caption = 9
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = False
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = False
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = False
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = False
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Public Sub Impossible6()

    Label65.Caption = 3
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = 4
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 9
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 7
    Label60.Caption = 3
    Label77.Caption = ""
    Label76.Caption = 8
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = 4
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 1
    Label37.Caption = ""
    Label39.Caption = 6
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 8
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 5
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 4
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 2
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 3
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = 3
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 6
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 2
    Label15.Caption = ""
    Label5.Caption = 5
    Label4.Caption = 6
    Label6.Caption = ""
    Label23.Caption = 1
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 7
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 3
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = False
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = False
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Impossible7()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 1
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 5
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = 3
    Label69.Caption = ""
    Label59.Caption = 1
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 4
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 2
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 9
    Label38.Caption = ""
    Label37.Caption = 9
    Label39.Caption = 8
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 5
    Label47.Caption = 7
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 7
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 9
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 6
    Label35.Caption = 4
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = 1
    Label54.Caption = ""
    Label11.Caption = 9
    Label10.Caption = ""
    Label12.Caption = 5
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 7
    Label20.Caption = ""
    Label19.Caption = 3
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 8
    Label23.Caption = ""
    Label22.Caption = 9
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 3
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 2
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = False
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = False
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = False
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = False
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Impossible8()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = 7
    Label55.Caption = ""
    Label57.Caption = 3
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = 1
    Label68.Caption = 7
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 5
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = 1
    Label62.Caption = ""
    Label61.Caption = 9
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 2
    Label81.Caption = ""
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 4
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 6
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = 7
    Label41.Caption = ""
    Label40.Caption = 2
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 4
    Label51.Caption = ""
    Label44.Caption = 5
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = 9
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 8
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 6
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 8
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = 1
    Label6.Caption = ""
    Label23.Caption = 3
    Label22.Caption = ""
    Label24.Caption = 9
    Label17.Caption = 6
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = 5
    Label7.Caption = ""
    Label9.Caption = 7
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub Impossible9()

    Label65.Caption = 8
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 5
    Label74.Caption = 2
    Label73.Caption = ""
    Label75.Caption = 7
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = 6
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = ""
    Label71.Caption = 9
    Label70.Caption = ""
    Label72.Caption = 7
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = ""
    Label81.Caption = 4
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = 6
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 7
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = 8
    Label32.Caption = 6
    Label31.Caption = ""
    Label33.Caption = 2
    Label50.Caption = 1
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = 6
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 7
    Label10.Caption = ""
    Label12.Caption = 3
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 9
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = 3
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 5
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 5
    Label8.Caption = 2
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = 6
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = False
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = False
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = False
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = False
End Sub

Public Sub Impossible10()
     
    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 5
    Label56.Caption = ""
    Label55.Caption = ""
    Label57.Caption = 4
    Label74.Caption = ""
    Label73.Caption = 9
    Label75.Caption = 1
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = 6
    Label59.Caption = ""
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 8
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = 9
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 2
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 8
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = 2
    Label30.Caption = ""
    Label47.Caption = 1
    Label46.Caption = ""
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 7
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = ""
    Label44.Caption = 1
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = 5
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 3
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 2
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 6
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 8
    Label15.Caption = ""
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 3
    Label23.Caption = 9
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = 3
    Label16.Caption = 4
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = ""
    Label26.Caption = 7
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = True
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = False
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = True
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = False
    Label46.Tag = True
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = False
    Label16.Tag = False
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = True
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = False
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = True
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = False
    Label46.Enabled = True
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = False
    Label16.Enabled = False
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Public Sub nuGameClear()
Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = ""
    Label16.Caption = ""
    Label17.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""
    Label22.Caption = ""
    Label23.Caption = ""
    Label24.Caption = ""
    Label25.Caption = ""
    Label26.Caption = ""
    Label27.Caption = ""
    Label28.Caption = ""
    Label29.Caption = ""
    Label30.Caption = ""
    Label31.Caption = ""
    Label32.Caption = ""
    Label33.Caption = ""
    Label34.Caption = ""
    Label35.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
    Label42.Caption = ""
    Label43.Caption = ""
    Label44.Caption = ""
    Label45.Caption = ""
    Label46.Caption = ""
    Label47.Caption = ""
    Label48.Caption = ""
    Label49.Caption = ""
    Label50.Caption = ""
    Label51.Caption = ""
    Label52.Caption = ""
    Label53.Caption = ""
    Label54.Caption = ""
    Label55.Caption = ""
    Label56.Caption = ""
    Label57.Caption = ""
    Label58.Caption = ""
    Label59.Caption = ""
    Label60.Caption = ""
    Label61.Caption = ""
    Label62.Caption = ""
    Label63.Caption = ""
    Label64.Caption = ""
    Label65.Caption = ""
    Label66.Caption = ""
    Label67.Caption = ""
    Label68.Caption = ""
    Label69.Caption = ""
    Label70.Caption = ""
    Label71.Caption = ""
    Label72.Caption = ""
    Label73.Caption = ""
    Label74.Caption = ""
    Label75.Caption = ""
    Label76.Caption = ""
    Label77.Caption = ""
    Label78.Caption = ""
    Label79.Caption = ""
    Label80.Caption = ""
    Label81.Caption = ""
End Sub

Public Sub lblMsgBox()
    FillBoxCount = 0
    If Label1.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label2.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label3.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label4.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label5.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label6.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label7.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label8.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label9.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label10.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label11.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label12.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label13.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label14.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label15.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label16.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label17.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label18.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label19.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label20.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label21.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label22.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label23.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label24.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label25.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label26.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label27.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label28.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label29.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label30.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label31.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label32.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label33.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label34.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label35.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label36.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label37.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label38.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label39.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label40.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label41.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label42.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label43.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label44.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label45.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label46.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label47.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label48.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label49.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label50.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label51.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label52.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label53.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label54.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label55.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label56.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label57.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label58.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label59.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label60.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label61.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label62.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label63.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label64.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label65.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label66.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label67.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label68.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label69.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label70.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label71.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label72.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label73.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label74.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label75.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label76.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label77.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label78.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label79.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label80.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    If Label81.Caption = "" Then
            FillBoxCount = FillBoxCount + 1
    End If
    lblUserMsg.Caption = "You Have " & FillBoxCount & " More Empty Boxes"
End Sub

Private Sub FinishGame()
    I = 0
    j = 0
    FinishTimeT = 0
    finishI = 0
    FinishB = False
    If Label1.Caption <> "" Then
    If Label2.Caption <> "" Then
    If Label3.Caption <> "" Then
    If Label4.Caption <> "" Then
    If Label5.Caption <> "" Then
    If Label6.Caption <> "" Then
    If Label7.Caption <> "" Then
    If Label8.Caption <> "" Then
    If Label9.Caption <> "" Then
    If Label10.Caption <> "" Then
    If Label11.Caption <> "" Then
    If Label12.Caption <> "" Then
    If Label13.Caption <> "" Then
    If Label14.Caption <> "" Then
    If Label15.Caption <> "" Then
    If Label16.Caption <> "" Then
    If Label17.Caption <> "" Then
    If Label18.Caption <> "" Then
    If Label19.Caption <> "" Then
    If Label20.Caption <> "" Then
    If Label21.Caption <> "" Then
    If Label22.Caption <> "" Then
    If Label23.Caption <> "" Then
    If Label24.Caption <> "" Then
    If Label25.Caption <> "" Then
    If Label26.Caption <> "" Then
    If Label27.Caption <> "" Then
    If Label28.Caption <> "" Then
    If Label29.Caption <> "" Then
    If Label30.Caption <> "" Then
    If Label31.Caption <> "" Then
    If Label32.Caption <> "" Then
    If Label33.Caption <> "" Then
    If Label34.Caption <> "" Then
    If Label35.Caption <> "" Then
    If Label36.Caption <> "" Then
    If Label37.Caption <> "" Then
    If Label38.Caption <> "" Then
    If Label39.Caption <> "" Then
    If Label40.Caption <> "" Then
    If Label41.Caption <> "" Then
    If Label42.Caption <> "" Then
    If Label43.Caption <> "" Then
    If Label44.Caption <> "" Then
    If Label45.Caption <> "" Then
    If Label46.Caption <> "" Then
    If Label47.Caption <> "" Then
    If Label48.Caption <> "" Then
    If Label49.Caption <> "" Then
    If Label50.Caption <> "" Then
    If Label51.Caption <> "" Then
    If Label52.Caption <> "" Then
    If Label53.Caption <> "" Then
    If Label54.Caption <> "" Then
    If Label55.Caption <> "" Then
    If Label56.Caption <> "" Then
    If Label57.Caption <> "" Then
    If Label58.Caption <> "" Then
    If Label59.Caption <> "" Then
    If Label60.Caption <> "" Then
    If Label61.Caption <> "" Then
    If Label62.Caption <> "" Then
    If Label63.Caption <> "" Then
    If Label64.Caption <> "" Then
    If Label65.Caption <> "" Then
    If Label66.Caption <> "" Then
    If Label67.Caption <> "" Then
    If Label68.Caption <> "" Then
    If Label69.Caption <> "" Then
    If Label70.Caption <> "" Then
    If Label71.Caption <> "" Then
    If Label72.Caption <> "" Then
    If Label73.Caption <> "" Then
    If Label74.Caption <> "" Then
    If Label75.Caption <> "" Then
    If Label76.Caption <> "" Then
    If Label77.Caption <> "" Then
    If Label78.Caption <> "" Then
    If Label79.Caption <> "" Then
    If Label80.Caption <> "" Then
    If Label81.Caption <> "" Then
        I = CInt(lbltimer1.Caption)
        j = CInt(lbltimer2.Caption)
        finishI = 100
        FinishB = True
        FinishTimeT = I * 60 + j
        If FinishTimeT > 0 And FinishTimeT <= 600 Then
            FinishScore = Round((90 + (1 - (FinishTimeT / 600)) * 10), 2)
        ElseIf FinishTimeT > 600 And FinishTimeT <= 1200 Then
            FinishScore = Round((80 + (1 - (FinishTimeT / 1200)) * 10), 2)
        ElseIf FinishTimeT > 1200 And FinishTimeT <= 1800 Then
            FinishScore = Round((70 + (1 - (FinishTimeT / 1800)) * 10), 2)
        ElseIf FinishTimeT > 1800 And FinishTimeT <= 2400 Then
            FinishScore = Round((60 + (1 - (FinishTimeT / 2400)) * 10), 2)
        ElseIf FinishTimeT > 2400 And FinishTimeT <= 3000 Then
            FinishScore = Round((50 + (1 - (FinishTimeT / 3000)) * 10), 2)
        ElseIf FinishTimeT > 3000 And FinishTimeT <= 3600 Then
            FinishScore = Round((40 + (1 - (FinishTimeT / 3600)) * 10), 2)
        ElseIf FinishTimeT > 3600 And FinishTimeT <= 4200 Then
            FinishScore = Round((30 + (1 - (FinishTimeT / 4200)) * 10), 2)
        ElseIf FinishTimeT > 4200 And FinishTimeT <= 4800 Then
            FinishScore = Round((20 + (1 - (FinishTimeT / 4800)) * 10), 2)
        ElseIf FinishTimeT > 4800 And FinishTimeT <= 5400 Then
            FinishScore = Round((10 + (1 - (FinishTimeT / 5400)) * 10), 2)
        ElseIf FinishTimeT > 5400 And FinishTimeT <= 6000 Then
            FinishScore = Round((3 + (1 - (FinishTimeT / 6000)) * 10), 2)
        Else
            FinishScore = 0
        End If
'        timermarquee_Timer
        MsgBox "You Took Total " & FinishTimeT & " sec in " & UCase(PlayMode) & " Mode; Hence You Scored  " & FinishScore & " % !", vbInformation, "Game Over"
        timermarquee_Timer
        ResetGrid
'        finishI = 0
'        FinishB = False
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If

End Sub

 
Private Sub wmptimer_Timer()
    If chkmusic.Value = 1 Then
    musicI = musicI + 1
        If musicI = 248 Then
            wmpmusicGobinda.URL = "C:\WINDOWS\Media\onestop.mid"
            musicI = 0
        End If
    End If
End Sub

Public Function lblLock(LBLNAME As Label) As Boolean
    If LBLNAME.Caption = "" Then
        LBLNAME.Enabled = True
        LBLNAME.Tag = True
'        lockBar.Value = lockBar.Value + 1
    Else
        LBLNAME.Enabled = False
        LBLNAME.Tag = False
'        lockBar.Value = lockBar.Value + 1
    End If
End Function

Public Function lblmusicSE(LBLNAME As Label) As Boolean
    LabelToolTipText = ""
    If LBLNAME.Caption = "1" Then
        LabelToolTipText = "One"
    ElseIf LBLNAME.Caption = "2" Then
        LabelToolTipText = "Two"
    ElseIf LBLNAME.Caption = "3" Then
        LabelToolTipText = "Three"
    ElseIf LBLNAME.Caption = "4" Then
        LabelToolTipText = "Four"
    ElseIf LBLNAME.Caption = "5" Then
        LabelToolTipText = "Five"
    ElseIf LBLNAME.Caption = "6" Then
        LabelToolTipText = "Six"
    ElseIf LBLNAME.Caption = "7" Then
        LabelToolTipText = "Seven"
    ElseIf LBLNAME.Caption = "8" Then
        LabelToolTipText = "Eight"
    ElseIf LBLNAME.Caption = "9" Then
        LabelToolTipText = "Nine"
    ElseIf LBLNAME.Caption = "" Then
        LabelToolTipText = ""
    Else
        LabelToolTipText = ""
    End If
    LBLNAME.ToolTipText = LabelToolTipText
    If LBLNAME.Caption = "" Then
        lblmusic.URL = "C:\WINDOWS\Media\Windows Recycle.wav"
    Else
        lblmusic.URL = "C:\WINDOWS\Media\Windows Pop-up Blocked.wav"
    End If
End Function

Private Sub disableB()
    CMD1.Enabled = False
    CMD2.Enabled = False
    CMD3.Enabled = False
    CMD4.Enabled = False
    CMD5.Enabled = False
    CMD6.Enabled = False
    CMD7.Enabled = False
    CMD8.Enabled = False
    CMD9.Enabled = False
    CmdClear.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub HardDefault()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 9
    Label56.Caption = 5
    Label55.Caption = 1
    Label57.Caption = ""
    Label74.Caption = 2
    Label73.Caption = ""
    Label75.Caption = 7
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = 2
    Label59.Caption = 3
    Label58.Caption = ""
    Label60.Caption = 7
    Label77.Caption = ""
    Label76.Caption = 9
    Label78.Caption = ""
    Label71.Caption = ""
    Label70.Caption = 4
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 8
    Label63.Caption = ""
    Label80.Caption = ""
    Label79.Caption = 3
    Label81.Caption = ""
    Label38.Caption = 5
    Label37.Caption = ""
    Label39.Caption = 3
    Label29.Caption = 1
    Label28.Caption = 7
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = 4
    Label41.Caption = ""
    Label40.Caption = 8
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 2
    Label51.Caption = ""
    Label44.Caption = 7
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = 9
    Label36.Caption = 6
    Label53.Caption = 3
    Label52.Caption = ""
    Label54.Caption = 1
    Label11.Caption = ""
    Label10.Caption = 9
    Label12.Caption = ""
    Label2.Caption = ""
    Label1.Caption = 4
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 7
    Label21.Caption = ""
    Label14.Caption = ""
    Label13.Caption = 1
    Label15.Caption = ""
    Label5.Caption = 7
    Label4.Caption = ""
    Label6.Caption = ""
    Label23.Caption = 6
    Label22.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 6
    Label8.Caption = ""
    Label7.Caption = 5
    Label9.Caption = 3
    Label26.Caption = 8
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = False
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = False
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = True
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = False
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = False
    Label36.Tag = False
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = True
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = True
    Label13.Tag = False
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = True
    Label23.Tag = False
    Label22.Tag = True
    Label24.Tag = True
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = False
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = False
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = False
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = True
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = False
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = False
    Label36.Enabled = False
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = True
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = True
    Label13.Enabled = False
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = True
    Label23.Enabled = False
    Label22.Enabled = True
    Label24.Enabled = True
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = False
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Private Sub easyF()

    Label65.Caption = ""
    Label64.Caption = 6
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 8
    Label57.Caption = 9
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = 7
    Label68.Caption = 8
    Label67.Caption = 4
    Label69.Caption = ""
    Label59.Caption = 7
    Label58.Caption = 2
    Label60.Caption = ""
    Label77.Caption = ""
    Label76.Caption = 1
    Label78.Caption = ""
    Label71.Caption = 3
    Label70.Caption = ""
    Label72.Caption = 9
    Label62.Caption = 6
    Label61.Caption = 1
    Label63.Caption = ""
    Label80.Caption = 5
    Label79.Caption = 2
    Label81.Caption = 8
    Label38.Caption = 5
    Label37.Caption = ""
    Label39.Caption = 7
    Label29.Caption = 2
    Label28.Caption = ""
    Label30.Caption = 1
    Label47.Caption = 8
    Label46.Caption = 6
    Label48.Caption = ""
    Label41.Caption = ""
    Label40.Caption = 8
    Label42.Caption = 6
    Label32.Caption = 9
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = 2
    Label49.Caption = 4
    Label51.Caption = ""
    Label44.Caption = ""
    Label43.Caption = ""
    Label45.Caption = 4
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 6
    Label53.Caption = 7
    Label52.Caption = ""
    Label54.Caption = 1
    Label11.Caption = 4
    Label10.Caption = 9
    Label12.Caption = 2
    Label2.Caption = ""
    Label1.Caption = 6
    Label3.Caption = 7
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 5
    Label14.Caption = ""
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = 4
    Label4.Caption = 5
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = 9
    Label24.Caption = 2
    Label17.Caption = 6
    Label16.Caption = ""
    Label18.Caption = 8
    Label8.Caption = 3
    Label7.Caption = 9
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = 7
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = False
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = False
    Label68.Tag = False
    Label67.Tag = False
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = False
    Label60.Tag = True
    Label77.Tag = True
    Label76.Tag = False
    Label78.Tag = True
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = False
    Label62.Tag = False
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = False
    Label81.Tag = False
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = False
    Label46.Tag = False
    Label48.Tag = True
    Label41.Tag = True
    Label40.Tag = False
    Label42.Tag = False
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = False
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = True
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = False
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = False
    Label10.Tag = False
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = False
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = True
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = False
    Label24.Tag = False
    Label17.Tag = False
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = False
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = False
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = False
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = False
    Label68.Enabled = False
    Label67.Enabled = False
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = False
    Label60.Enabled = True
    Label77.Enabled = True
    Label76.Enabled = False
    Label78.Enabled = True
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = False
    Label62.Enabled = False
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = False
    Label81.Enabled = False
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = False
    Label46.Enabled = False
    Label48.Enabled = True
    Label41.Enabled = True
    Label40.Enabled = False
    Label42.Enabled = False
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = False
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = True
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = False
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = False
    Label10.Enabled = False
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = False
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = True
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = False
    Label24.Enabled = False
    Label17.Enabled = False
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = False
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = False
    Label27.Enabled = True
End Sub

Private Sub ImpossibleDefault()

    Label65.Caption = ""
    Label64.Caption = 5
    Label66.Caption = 9
    Label56.Caption = ""
    Label55.Caption = 7
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = ""
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 4
    Label58.Caption = ""
    Label60.Caption = ""
    Label77.Caption = 1
    Label76.Caption = ""
    Label78.Caption = 8
    Label71.Caption = 6
    Label70.Caption = ""
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 2
    Label79.Caption = ""
    Label81.Caption = ""
    Label38.Caption = 9
    Label37.Caption = 3
    Label39.Caption = ""
    Label29.Caption = 6
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = ""
    Label48.Caption = 4
    Label41.Caption = ""
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = 3
    Label51.Caption = ""
    Label44.Caption = 4
    Label43.Caption = ""
    Label45.Caption = ""
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = 1
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 5
    Label11.Caption = ""
    Label10.Caption = ""
    Label12.Caption = 3
    Label2.Caption = ""
    Label1.Caption = 5
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = ""
    Label21.Caption = 7
    Label14.Caption = 7
    Label13.Caption = ""
    Label15.Caption = 8
    Label5.Caption = ""
    Label4.Caption = ""
    Label6.Caption = 2
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 6
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 9
    Label9.Caption = ""
    Label26.Caption = 3
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = False
    Label66.Tag = False
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = True
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = True
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = False
    Label70.Tag = True
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = False
    Label39.Tag = True
    Label29.Tag = False
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = True
    Label48.Tag = False
    Label41.Tag = True
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = False
    Label51.Tag = True
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = True
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = False
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = True
    Label21.Tag = False
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = True
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = False
    Label66.Enabled = False
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = True
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = True
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = False
    Label70.Enabled = True
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = False
    Label39.Enabled = True
    Label29.Enabled = False
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = True
    Label48.Enabled = False
    Label41.Enabled = True
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = False
    Label51.Enabled = True
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = True
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = False
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = True
    Label21.Enabled = False
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = True
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Private Sub MediumDefault()

    Label65.Caption = ""
    Label64.Caption = ""
    Label66.Caption = 6
    Label56.Caption = 8
    Label55.Caption = ""
    Label57.Caption = ""
    Label74.Caption = 4
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 8
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = ""
    Label58.Caption = 5
    Label60.Caption = 7
    Label77.Caption = 9
    Label76.Caption = ""
    Label78.Caption = 6
    Label71.Caption = ""
    Label70.Caption = 2
    Label72.Caption = ""
    Label62.Caption = ""
    Label61.Caption = 1
    Label63.Caption = ""
    Label80.Caption = 7
    Label79.Caption = ""
    Label81.Caption = 8
    Label38.Caption = ""
    Label37.Caption = ""
    Label39.Caption = ""
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = 5
    Label47.Caption = ""
    Label46.Caption = 6
    Label48.Caption = 3
    Label41.Caption = 6
    Label40.Caption = ""
    Label42.Caption = 2
    Label32.Caption = 9
    Label31.Caption = ""
    Label33.Caption = 8
    Label50.Caption = 1
    Label49.Caption = ""
    Label51.Caption = 5
    Label44.Caption = 4
    Label43.Caption = 5
    Label45.Caption = ""
    Label35.Caption = 2
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = ""
    Label11.Caption = 3
    Label10.Caption = ""
    Label12.Caption = 7
    Label2.Caption = ""
    Label1.Caption = 2
    Label3.Caption = ""
    Label20.Caption = ""
    Label19.Caption = 1
    Label21.Caption = ""
    Label14.Caption = 5
    Label13.Caption = ""
    Label15.Caption = 4
    Label5.Caption = 6
    Label4.Caption = 8
    Label6.Caption = ""
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 2
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = 9
    Label8.Caption = ""
    Label7.Caption = ""
    Label9.Caption = 3
    Label26.Caption = 5
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = True
    Label64.Tag = True
    Label66.Tag = False
    Label56.Tag = False
    Label55.Tag = True
    Label57.Tag = True
    Label74.Tag = False
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = True
    Label58.Tag = False
    Label60.Tag = False
    Label77.Tag = False
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = True
    Label61.Tag = False
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = True
    Label81.Tag = False
    Label38.Tag = True
    Label37.Tag = True
    Label39.Tag = True
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = False
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = False
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = False
    Label32.Tag = False
    Label31.Tag = True
    Label33.Tag = False
    Label50.Tag = False
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = False
    Label43.Tag = False
    Label45.Tag = True
    Label35.Tag = False
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = True
    Label11.Tag = False
    Label10.Tag = True
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = False
    Label3.Tag = True
    Label20.Tag = True
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = False
    Label5.Tag = False
    Label4.Tag = False
    Label6.Tag = True
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = False
    Label8.Tag = True
    Label7.Tag = True
    Label9.Tag = False
    Label26.Tag = False
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = True
    Label64.Enabled = True
    Label66.Enabled = False
    Label56.Enabled = False
    Label55.Enabled = True
    Label57.Enabled = True
    Label74.Enabled = False
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = True
    Label58.Enabled = False
    Label60.Enabled = False
    Label77.Enabled = False
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = True
    Label61.Enabled = False
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = True
    Label81.Enabled = False
    Label38.Enabled = True
    Label37.Enabled = True
    Label39.Enabled = True
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = False
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = False
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = False
    Label32.Enabled = False
    Label31.Enabled = True
    Label33.Enabled = False
    Label50.Enabled = False
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = False
    Label43.Enabled = False
    Label45.Enabled = True
    Label35.Enabled = False
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = True
    Label11.Enabled = False
    Label10.Enabled = True
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = False
    Label3.Enabled = True
    Label20.Enabled = True
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = False
    Label5.Enabled = False
    Label4.Enabled = False
    Label6.Enabled = True
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = False
    Label8.Enabled = True
    Label7.Enabled = True
    Label9.Enabled = False
    Label26.Enabled = False
    Label25.Enabled = True
    Label27.Enabled = True
End Sub

Private Sub ProfessionalDefault()
    
    Label65.Caption = 6
    Label64.Caption = ""
    Label66.Caption = ""
    Label56.Caption = ""
    Label55.Caption = 2
    Label57.Caption = ""
    Label74.Caption = ""
    Label73.Caption = ""
    Label75.Caption = ""
    Label68.Caption = 7
    Label67.Caption = ""
    Label69.Caption = ""
    Label59.Caption = 3
    Label58.Caption = ""
    Label60.Caption = 8
    Label77.Caption = ""
    Label76.Caption = ""
    Label78.Caption = 4
    Label71.Caption = ""
    Label70.Caption = 1
    Label72.Caption = ""
    Label62.Caption = 6
    Label61.Caption = ""
    Label63.Caption = ""
    Label80.Caption = 9
    Label79.Caption = 8
    Label81.Caption = ""
    Label38.Caption = 2
    Label37.Caption = ""
    Label39.Caption = 7
    Label29.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label47.Caption = ""
    Label46.Caption = 9
    Label48.Caption = 8
    Label41.Caption = 9
    Label40.Caption = ""
    Label42.Caption = ""
    Label32.Caption = ""
    Label31.Caption = ""
    Label33.Caption = ""
    Label50.Caption = ""
    Label49.Caption = ""
    Label51.Caption = 7
    Label44.Caption = 4
    Label43.Caption = ""
    Label45.Caption = 6
    Label35.Caption = ""
    Label34.Caption = ""
    Label36.Caption = ""
    Label53.Caption = ""
    Label52.Caption = ""
    Label54.Caption = 1
    Label11.Caption = ""
    Label10.Caption = 6
    Label12.Caption = 9
    Label2.Caption = ""
    Label1.Caption = ""
    Label3.Caption = 4
    Label20.Caption = 2
    Label19.Caption = 3
    Label21.Caption = ""
    Label14.Caption = 8
    Label13.Caption = ""
    Label15.Caption = ""
    Label5.Caption = 5
    Label4.Caption = ""
    Label6.Caption = 1
    Label23.Caption = ""
    Label22.Caption = ""
    Label24.Caption = 9
    Label17.Caption = ""
    Label16.Caption = ""
    Label18.Caption = ""
    Label8.Caption = ""
    Label7.Caption = 9
    Label9.Caption = ""
    Label26.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    
    
    Label65.Tag = False
    Label64.Tag = True
    Label66.Tag = True
    Label56.Tag = True
    Label55.Tag = False
    Label57.Tag = True
    Label74.Tag = True
    Label73.Tag = True
    Label75.Tag = True
    Label68.Tag = False
    Label67.Tag = True
    Label69.Tag = True
    Label59.Tag = False
    Label58.Tag = True
    Label60.Tag = False
    Label77.Tag = True
    Label76.Tag = True
    Label78.Tag = False
    Label71.Tag = True
    Label70.Tag = False
    Label72.Tag = True
    Label62.Tag = False
    Label61.Tag = True
    Label63.Tag = True
    Label80.Tag = False
    Label79.Tag = False
    Label81.Tag = True
    Label38.Tag = False
    Label37.Tag = True
    Label39.Tag = False
    Label29.Tag = True
    Label28.Tag = True
    Label30.Tag = True
    Label47.Tag = True
    Label46.Tag = False
    Label48.Tag = False
    Label41.Tag = False
    Label40.Tag = True
    Label42.Tag = True
    Label32.Tag = True
    Label31.Tag = True
    Label33.Tag = True
    Label50.Tag = True
    Label49.Tag = True
    Label51.Tag = False
    Label44.Tag = False
    Label43.Tag = True
    Label45.Tag = False
    Label35.Tag = True
    Label34.Tag = True
    Label36.Tag = True
    Label53.Tag = True
    Label52.Tag = True
    Label54.Tag = False
    Label11.Tag = True
    Label10.Tag = False
    Label12.Tag = False
    Label2.Tag = True
    Label1.Tag = True
    Label3.Tag = False
    Label20.Tag = False
    Label19.Tag = False
    Label21.Tag = True
    Label14.Tag = False
    Label13.Tag = True
    Label15.Tag = True
    Label5.Tag = False
    Label4.Tag = True
    Label6.Tag = False
    Label23.Tag = True
    Label22.Tag = True
    Label24.Tag = False
    Label17.Tag = True
    Label16.Tag = True
    Label18.Tag = True
    Label8.Tag = True
    Label7.Tag = False
    Label9.Tag = True
    Label26.Tag = True
    Label25.Tag = True
    Label27.Tag = True
    
    
    Label65.Enabled = False
    Label64.Enabled = True
    Label66.Enabled = True
    Label56.Enabled = True
    Label55.Enabled = False
    Label57.Enabled = True
    Label74.Enabled = True
    Label73.Enabled = True
    Label75.Enabled = True
    Label68.Enabled = False
    Label67.Enabled = True
    Label69.Enabled = True
    Label59.Enabled = False
    Label58.Enabled = True
    Label60.Enabled = False
    Label77.Enabled = True
    Label76.Enabled = True
    Label78.Enabled = False
    Label71.Enabled = True
    Label70.Enabled = False
    Label72.Enabled = True
    Label62.Enabled = False
    Label61.Enabled = True
    Label63.Enabled = True
    Label80.Enabled = False
    Label79.Enabled = False
    Label81.Enabled = True
    Label38.Enabled = False
    Label37.Enabled = True
    Label39.Enabled = False
    Label29.Enabled = True
    Label28.Enabled = True
    Label30.Enabled = True
    Label47.Enabled = True
    Label46.Enabled = False
    Label48.Enabled = False
    Label41.Enabled = False
    Label40.Enabled = True
    Label42.Enabled = True
    Label32.Enabled = True
    Label31.Enabled = True
    Label33.Enabled = True
    Label50.Enabled = True
    Label49.Enabled = True
    Label51.Enabled = False
    Label44.Enabled = False
    Label43.Enabled = True
    Label45.Enabled = False
    Label35.Enabled = True
    Label34.Enabled = True
    Label36.Enabled = True
    Label53.Enabled = True
    Label52.Enabled = True
    Label54.Enabled = False
    Label11.Enabled = True
    Label10.Enabled = False
    Label12.Enabled = False
    Label2.Enabled = True
    Label1.Enabled = True
    Label3.Enabled = False
    Label20.Enabled = False
    Label19.Enabled = False
    Label21.Enabled = True
    Label14.Enabled = False
    Label13.Enabled = True
    Label15.Enabled = True
    Label5.Enabled = False
    Label4.Enabled = True
    Label6.Enabled = False
    Label23.Enabled = True
    Label22.Enabled = True
    Label24.Enabled = False
    Label17.Enabled = True
    Label16.Enabled = True
    Label18.Enabled = True
    Label8.Enabled = True
    Label7.Enabled = False
    Label9.Enabled = True
    Label26.Enabled = True
    Label25.Enabled = True
    Label27.Enabled = True
End Sub
