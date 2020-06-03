VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtgridname 
      Height          =   375
      Left            =   360
      TabIndex        =   82
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "save"
      Height          =   495
      Left            =   4440
      TabIndex        =   81
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text81 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   80
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text80 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   79
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text79 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   78
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text78 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   77
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text77 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   76
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text76 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   75
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text75 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   74
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text74 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   73
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text73 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   72
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text72 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   71
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text71 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   70
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text70 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   69
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text69 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   68
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text68 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   67
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text67 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   66
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text66 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   65
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text65 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   64
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text64 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   63
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text63 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   62
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text62 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   61
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text61 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   60
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text60 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   59
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text59 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   58
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text58 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   57
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text57 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   56
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text56 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   55
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text55 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   54
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text54 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   53
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text53 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   52
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text52 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   51
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text51 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   50
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text50 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   49
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text49 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   48
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text48 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   47
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text47 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   46
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text46 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   45
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text45 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   44
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text44 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   43
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text43 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   42
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text42 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   41
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text41 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   40
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text40 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   39
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text39 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   38
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text38 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   37
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text37 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   36
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text36 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   35
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text35 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   34
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text34 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   33
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text33 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   32
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text32 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   31
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text31 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   30
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text30 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   29
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text29 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   28
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text28 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   27
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   26
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   25
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   23
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   22
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   21
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   16
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   13
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   1
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Dim str As String
Public Sub Connection()
    Set con = New ADODB.Connection
    con.Open "gridc", "", ""

End Sub

Private Sub cmdsave_Click()
    str = ""
    str = str + "   CREATE TABLE " & txtgridname.Text & ""
    str = str + "   ("
'    str = str + "   Gridname VARCHAR(50) ,"
    str = str + "   caption VARCHAR(50)  ,"
    str = str + "   tag VARCHAR(50),"
    str = str + "   Enabled VarChar(50)"
    str = str + "   );"
    If rs.State = 1 Then rs.Close
     rs.Open str, con, adOpenDynamic, adLockOptimistic
    If rs.State = 1 Then rs.Close
     rs.Open "select * from " & txtgridname.Text & "", con, adOpenDynamic, adLockOptimistic
    For i = 1 To 81
        If i = 1 Then
            If Text1.Text <> "" Then
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label65.caption=" & Text1.Text & ""
                rs!Tag = "Label65.tag= false"
                rs!Enabled = "Label65.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label65.caption=0"
                rs!Tag = "Label65.tag= true"
                rs!Enabled = "Label65.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################
        
                If i = 2 Then
            If Text3.Text <> "" Then
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label64.caption=" & Text3.Text & ""
                rs!Tag = "Label64.tag= false"
                rs!Enabled = "Label64.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label64.caption=0"
                rs!Tag = "Label64.tag= true"
                rs!Enabled = "Label64.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################
        If i = 3 Then
            If Text2.Text <> "" Then
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label66.caption=" & Text2.Text & ""
                rs!Tag = "Label66.tag= false"
                rs!Enabled = "Label66.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label66.caption=0"
                rs!Tag = "Label66.tag= true"
                rs!Enabled = "Label66.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################
        If i = 4 Then
            If Text4.Text <> "" Then
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label56.caption=" & Text4.Text & ""
                rs!Tag = "Label56.tag= false"
                rs!Enabled = "Label56.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label56.caption=0"
                rs!Tag = "Label56.tag= true"
                rs!Enabled = "Label56.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 5 Then
            If Text6.Text <> "" Then
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label55.caption=" & Text6.Text & ""
                rs!Tag = "Label55.tag= false"
                rs!Enabled = "Label55.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                ''rs!Gridname = txtgridname.Text
                rs!Caption = "Label55.caption=0"
                rs!Tag = "Label55.tag= true"
                rs!Enabled = "Label55.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################
        If i = 6 Then
            If Text5.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label57.caption=" & Text5.Text & ""
                rs!Tag = "Label57.tag= false"
                rs!Enabled = "Label57.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label57.caption=0"
                rs!Tag = "Label57.tag= true"
                rs!Enabled = "Label57.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 7 Then
            If Text7.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label74.caption=" & Text7.Text & ""
                rs!Tag = "Label74.tag= false"
                rs!Enabled = "Label74.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label74.caption=0"
                rs!Tag = "Label74.tag= true"
                rs!Enabled = "Label74.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 8 Then
            If Text9.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label73.caption=" & Text9.Text & ""
                rs!Tag = "Label73.tag= false"
                rs!Enabled = "Label73.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label73.caption=0"
                rs!Tag = "Label73.tag= true"
                rs!Enabled = "Label73.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 9 Then
            If Text8.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label75.caption=" & Text8.Text & ""
                rs!Tag = "Label75.tag= false"
                rs!Enabled = "Label75.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label75.caption=0"
                rs!Tag = "Label75.tag= true"
                rs!Enabled = "Label75.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################
        If i = 10 Then
            If Text10.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label68.caption=" & Text10.Text & ""
                rs!Tag = "Label68.tag= false"
                rs!Enabled = "Label68.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label68.caption=0"
                rs!Tag = "Label68.tag= true"
                rs!Enabled = "Label68.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 11 Then
            If Text12.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label67.caption=" & Text12.Text & ""
                rs!Tag = "Label67.tag= false"
                rs!Enabled = "Label67.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label67.caption=0"
                rs!Tag = "Label67.tag= true"
                rs!Enabled = "Label67.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 12 Then
            If Text11.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label69.caption=" & Text11.Text & ""
                rs!Tag = "Label69.tag= false"
                rs!Enabled = "Label69.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label69.caption=0"
                rs!Tag = "Label69.tag= true"
                rs!Enabled = "Label69.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 13 Then
            If Text13.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label59.caption=" & Text13.Text & ""
                rs!Tag = "Label59.tag= false"
                rs!Enabled = "Label59.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label59.caption=0"
                rs!Tag = "Label59.tag= true"
                rs!Enabled = "Label59.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 14 Then
            If Text15.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label58.caption=" & Text15.Text & ""
                rs!Tag = "Label58.tag= false"
                rs!Enabled = "Label58.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label58.caption=0"
                rs!Tag = "Label58.tag= true"
                rs!Enabled = "Label58.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 15 Then
            If Text14.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label60.caption=" & Text14.Text & ""
                rs!Tag = "Label60.tag= false"
                rs!Enabled = "Label60.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label60.caption=0"
                rs!Tag = "Label60.tag= true"
                rs!Enabled = "Label60.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 16 Then
            If Text16.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label77.caption=" & Text16.Text & ""
                rs!Tag = "Label77.tag= false"
                rs!Enabled = "Label77.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label77.caption=0"
                rs!Tag = "Label77.tag= true"
                rs!Enabled = "Label77.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 17 Then
            If Text18.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label76.caption=" & Text18.Text & ""
                rs!Tag = "Label76.tag= false"
                rs!Enabled = "Label76.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label76.caption=0"
                rs!Tag = "Label76.tag= true"
                rs!Enabled = "Label76.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 18 Then
            If Text17.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label78.caption=" & Text17.Text & ""
                rs!Tag = "Label78.tag= false"
                rs!Enabled = "Label78.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label78.caption=0"
                rs!Tag = "Label78.tag= true"
                rs!Enabled = "Label78.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 19 Then
            If Text19.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label71.caption=" & Text19.Text & ""
                rs!Tag = "Label71.tag= false"
                rs!Enabled = "Label71.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label71.caption=0"
                rs!Tag = "Label71.tag= true"
                rs!Enabled = "Label71.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 20 Then
            If Text21.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label70.caption=" & Text21.Text & ""
                rs!Tag = "Label70.tag= false"
                rs!Enabled = "Label70.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label70.caption=0"
                rs!Tag = "Label70.tag= true"
                rs!Enabled = "Label70.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################
        If i = 21 Then
            If Text20.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label72.caption=" & Text20.Text & ""
                rs!Tag = "Label72.tag= false"
                rs!Enabled = "Label72.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label72.caption=0"
                rs!Tag = "Label72.tag= true"
                rs!Enabled = "Label72.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 22 Then
            If Text22.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label62.caption=" & Text22.Text & ""
                rs!Tag = "Label62.tag= false"
                rs!Enabled = "Label62.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label62.caption=0"
                rs!Tag = "Label62.tag= true"
                rs!Enabled = "Label62.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 23 Then
            If Text24.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label61.caption=" & Text24.Text & ""
                rs!Tag = "Label61.tag= false"
                rs!Enabled = "Label61.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label61.caption=0"
                rs!Tag = "Label61.tag= true"
                rs!Enabled = "Label61.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 24 Then
            If Text23.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label63.caption=" & Text23.Text & ""
                rs!Tag = "Label63.tag= false"
                rs!Enabled = "Label63.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label63.caption=0"
                rs!Tag = "Label63.tag= true"
                rs!Enabled = "Label63.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 25 Then
            If Text25.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label80.caption=" & Text25.Text & ""
                rs!Tag = "Label80.tag= false"
                rs!Enabled = "Label80.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label80.caption=0"
                rs!Tag = "Label80.tag= true"
                rs!Enabled = "Label80.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 26 Then
            If Text27.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label79.caption=" & Text27.Text & ""
                rs!Tag = "Label79.tag= false"
                rs!Enabled = "Label79.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label79.caption=0"
                rs!Tag = "Label79.tag= true"
                rs!Enabled = "Label79.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 27 Then
            If Text26.Text <> "" Then
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label81.caption=" & Text26.Text & ""
                rs!Tag = "Label81.tag= false"
                rs!Enabled = "Label81.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                'rs!Gridname = txtgridname.Text
                rs!Caption = "Label81.caption=0"
                rs!Tag = "Label81.tag= true"
                rs!Enabled = "Label81.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 28 Then
            If Text28.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label38.caption=" & Text28.Text & ""
                rs!Tag = "Label38.tag= false"
                rs!Enabled = "Label38.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label38.caption=0"
                rs!Tag = "Label38.tag= true"
                rs!Enabled = "Label38.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 29 Then
            If Text30.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label37.caption=" & Text30.Text & ""
                rs!Tag = "Label37.tag= false"
                rs!Enabled = "Label37.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label37.caption=0"
                rs!Tag = "Label37.tag= true"
                rs!Enabled = "Label37.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 30 Then
            If Text29.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label39.caption=" & Text29.Text & ""
                rs!Tag = "Label39.tag= false"
                rs!Enabled = "Label39.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label39.caption=0"
                rs!Tag = "Label39.tag= true"
                rs!Enabled = "Label39.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 31 Then
            If Text31.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label29.caption=" & Text31.Text & ""
                rs!Tag = "Label29.tag= false"
                rs!Enabled = "Label29.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label29.caption=0"
                rs!Tag = "Label29.tag= true"
                rs!Enabled = "Label29.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 32 Then
            If Text33.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label28.caption=" & Text33.Text & ""
                rs!Tag = "Label28.tag= false"
                rs!Enabled = "Label28.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label28.caption=0"
                rs!Tag = "Label28.tag= true"
                rs!Enabled = "Label28.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 33 Then
            If Text32.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label30.caption=" & Text32.Text & ""
                rs!Tag = "Label30.tag= false"
                rs!Enabled = "Label30.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label30.caption=0"
                rs!Tag = "Label30.tag= true"
                rs!Enabled = "Label30.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 34 Then
            If Text34.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label47.caption=" & Text34.Text & ""
                rs!Tag = "Label47.tag= false"
                rs!Enabled = "Label47.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label47.caption=0"
                rs!Tag = "Label47.tag= true"
                rs!Enabled = "Label47.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 35 Then
            If Text36.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label46.caption=" & Text36.Text & ""
                rs!Tag = "Label46.tag= false"
                rs!Enabled = "Label46.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label46.caption=0"
                rs!Tag = "Label46.tag= true"
                rs!Enabled = "Label46.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 36 Then
            If Text35.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label48.caption=" & Text35.Text & ""
                rs!Tag = "Label48.tag= false"
                rs!Enabled = "Label48.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label48.caption=0"
                rs!Tag = "Label48.tag= true"
                rs!Enabled = "Label48.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 37 Then
            If Text37.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label41.caption=" & Text37.Text & ""
                rs!Tag = "Label41.tag= false"
                rs!Enabled = "Label41.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label41.caption=0"
                rs!Tag = "Label41.tag= true"
                rs!Enabled = "Label41.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 38 Then
            If Text39.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label40.caption=" & Text39.Text & ""
                rs!Tag = "Label40.tag= false"
                rs!Enabled = "Label40.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label40.caption=0"
                rs!Tag = "Label40.tag= true"
                rs!Enabled = "Label40.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 39 Then
            If Text38.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label42.caption=" & Text38.Text & ""
                rs!Tag = "Label42.tag= false"
                rs!Enabled = "Label42.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label42.caption=0"
                rs!Tag = "Label42.tag= true"
                rs!Enabled = "Label42.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 40 Then
            If Text40.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label32.caption=" & Text40.Text & ""
                rs!Tag = "Label32.tag= false"
                rs!Enabled = "Label32.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label32.caption=0"
                rs!Tag = "Label32.tag= true"
                rs!Enabled = "Label32.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 41 Then
            If Text42.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label31.caption=" & Text42.Text & ""
                rs!Tag = "Label31.tag= false"
                rs!Enabled = "Label31.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label31.caption=0"
                rs!Tag = "Label31.tag= true"
                rs!Enabled = "Label31.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 42 Then
            If Text41.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label33.caption=" & Text41.Text & ""
                rs!Tag = "Label33.tag= false"
                rs!Enabled = "Label33.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label33.caption=0"
                rs!Tag = "Label33.tag= true"
                rs!Enabled = "Label33.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 43 Then
            If Text43.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label50.caption=" & Text43.Text & ""
                rs!Tag = "Label50.tag= false"
                rs!Enabled = "Label50.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label50.caption=0"
                rs!Tag = "Label50.tag= true"
                rs!Enabled = "Label50.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 44 Then
            If Text45.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label49.caption=" & Text45.Text & ""
                rs!Tag = "Label49.tag= false"
                rs!Enabled = "Label49.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label49.caption=0"
                rs!Tag = "Label49.tag= true"
                rs!Enabled = "Label49.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 45 Then
            If Text44.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label51.caption=" & Text44.Text & ""
                rs!Tag = "Label51.tag= false"
                rs!Enabled = "Label51.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label51.caption=0"
                rs!Tag = "Label51.tag= true"
                rs!Enabled = "Label51.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 46 Then
            If Text46.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label44.caption=" & Text46.Text & ""
                rs!Tag = "Label44.tag= false"
                rs!Enabled = "Label44.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label44.caption=0"
                rs!Tag = "Label44.tag= true"
                rs!Enabled = "Label44.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 47 Then
            If Text48.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label43.caption=" & Text48.Text & ""
                rs!Tag = "Label43.tag= false"
                rs!Enabled = "Label43.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label43.caption=0"
                rs!Tag = "Label43.tag= true"
                rs!Enabled = "Label43.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 48 Then
            If Text47.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label45.caption=" & Text47.Text & ""
                rs!Tag = "Label45.tag= false"
                rs!Enabled = "Label45.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label45.caption=0"
                rs!Tag = "Label45.tag= true"
                rs!Enabled = "Label45.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 49 Then
            If Text49.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label35.caption=" & Text49.Text & ""
                rs!Tag = "Label35.tag= false"
                rs!Enabled = "Label35.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label35.caption=0"
                rs!Tag = "Label35.tag= true"
                rs!Enabled = "Label35.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 50 Then
            If Text51.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label34.caption=" & Text51.Text & ""
                rs!Tag = "Label34.tag= false"
                rs!Enabled = "Label34.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label34.caption=0"
                rs!Tag = "Label34.tag= true"
                rs!Enabled = "Label34.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 51 Then
            If Text50.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label36.caption=" & Text50.Text & ""
                rs!Tag = "Label36.tag= false"
                rs!Enabled = "Label36.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label36.caption=0"
                rs!Tag = "Label36.tag= true"
                rs!Enabled = "Label36.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 52 Then
            If Text52.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label53.caption=" & Text52.Text & ""
                rs!Tag = "Label53.tag= false"
                rs!Enabled = "Label53.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label53.caption=0"
                rs!Tag = "Label53.tag= true"
                rs!Enabled = "Label53.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 53 Then
            If Text54.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label52.caption=" & Text54.Text & ""
                rs!Tag = "Label52.tag= false"
                rs!Enabled = "Label52.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label52.caption=0"
                rs!Tag = "Label52.tag= true"
                rs!Enabled = "Label52.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 54 Then
            If Text53.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label54.caption=" & Text53.Text & ""
                rs!Tag = "Label54.tag= false"
                rs!Enabled = "Label54.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label54.caption=0"
                rs!Tag = "Label54.tag= true"
                rs!Enabled = "Label54.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 55 Then
            If Text55.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label11.caption=" & Text55.Text & ""
                rs!Tag = "Label11.tag= false"
                rs!Enabled = "Label11.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label11.caption=0"
                rs!Tag = "Label11.tag= true"
                rs!Enabled = "Label11.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 56 Then
            If Text57.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label10.caption=" & Text57.Text & ""
                rs!Tag = "Label10.tag= false"
                rs!Enabled = "Label10.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label10.caption=0"
                rs!Tag = "Label10.tag= true"
                rs!Enabled = "Label10.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 57 Then
            If Text56.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label12.caption=" & Text56.Text & ""
                rs!Tag = "Label12.tag= false"
                rs!Enabled = "Label12.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label12.caption=0"
                rs!Tag = "Label12.tag= true"
                rs!Enabled = "Label12.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 58 Then
            If Text58.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label2.caption=" & Text58.Text & ""
                rs!Tag = "Label2.tag= false"
                rs!Enabled = "Label2.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label2.caption=0"
                rs!Tag = "Label2.tag= true"
                rs!Enabled = "Label2.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 59 Then
            If Text60.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label1.caption=" & Text60.Text & ""
                rs!Tag = "Label1.tag= false"
                rs!Enabled = "Label1.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label1.caption=0"
                rs!Tag = "Label1.tag= true"
                rs!Enabled = "Label1.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 60 Then
            If Text59.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label3.caption=" & Text59.Text & ""
                rs!Tag = "Label3.tag= false"
                rs!Enabled = "Label3.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label3.caption=0"
                rs!Tag = "Label3.tag= true"
                rs!Enabled = "Label3.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 61 Then
            If Text61.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label20.caption=" & Text61.Text & ""
                rs!Tag = "Label20.tag= false"
                rs!Enabled = "Label20.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label20.caption=0"
                rs!Tag = "Label20.tag= true"
                rs!Enabled = "Label20.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 62 Then
            If Text63.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label19.caption=" & Text63.Text & ""
                rs!Tag = "Label19.tag= false"
                rs!Enabled = "Label19.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label19.caption=0"
                rs!Tag = "Label19.tag= true"
                rs!Enabled = "Label19.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 63 Then
            If Text62.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label21.caption=" & Text62.Text & ""
                rs!Tag = "Label21.tag= false"
                rs!Enabled = "Label21.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label21.caption=0"
                rs!Tag = "Label21.tag= true"
                rs!Enabled = "Label21.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 64 Then
            If Text64.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label14.caption=" & Text64.Text & ""
                rs!Tag = "Label14.tag= false"
                rs!Enabled = "Label14.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label14.caption=0"
                rs!Tag = "Label14.tag= true"
                rs!Enabled = "Label14.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################


        If i = 65 Then
            If Text66.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label13.caption=" & Text66.Text & ""
                rs!Tag = "Label13.tag= false"
                rs!Enabled = "Label13.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label13.caption=0"
                rs!Tag = "Label13.tag= true"
                rs!Enabled = "Label13.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 67 Then
            If Text67.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label5.caption=" & Text67.Text & ""
                rs!Tag = "Label5.tag= false"
                rs!Enabled = "Label5.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label5.caption=0"
                rs!Tag = "Label5.tag= true"
                rs!Enabled = "Label5.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 66 Then
            If Text65.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label15.caption=" & Text65.Text & ""
                rs!Tag = "Label15.tag= false"
                rs!Enabled = "Label15.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label15.caption=0"
                rs!Tag = "Label15.tag= true"
                rs!Enabled = "Label15.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 68 Then
            If Text69.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label4.caption=" & Text69.Text & ""
                rs!Tag = "Label4.tag= false"
                rs!Enabled = "Label4.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label4.caption=0"
                rs!Tag = "Label4.tag= true"
                rs!Enabled = "Label4.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 69 Then
            If Text68.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label6.caption=" & Text68.Text & ""
                rs!Tag = "Label6.tag= false"
                rs!Enabled = "Label6.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label6.caption=0"
                rs!Tag = "Label6.tag= true"
                rs!Enabled = "Label6.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 70 Then
            If Text70.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label23.caption=" & Text70.Text & ""
                rs!Tag = "Label23.tag= false"
                rs!Enabled = "Label23.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label23.caption=0"
                rs!Tag = "Label23.tag= true"
                rs!Enabled = "Label23.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 71 Then
            If Text72.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label22.caption=" & Text72.Text & ""
                rs!Tag = "Label22.tag= false"
                rs!Enabled = "Label22.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label22.caption=0"
                rs!Tag = "Label22.tag= true"
                rs!Enabled = "Label22.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 72 Then
            If Text71.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label24.caption=" & Text71.Text & ""
                rs!Tag = "Label24.tag= false"
                rs!Enabled = "Label24.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label24.caption=0"
                rs!Tag = "Label24.tag= true"
                rs!Enabled = "Label24.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 73 Then
            If Text73.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label17.caption=" & Text73.Text & ""
                rs!Tag = "Label17.tag= false"
                rs!Enabled = "Label17.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label17.caption=0"
                rs!Tag = "Label17.tag= true"
                rs!Enabled = "Label17.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################

        If i = 74 Then
            If Text75.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label16.caption=" & Text75.Text & ""
                rs!Tag = "Label16.tag= false"
                rs!Enabled = "Label16.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label16.caption=0"
                rs!Tag = "Label16.tag= true"
                rs!Enabled = "Label16.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 75 Then
            If Text74.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label18.caption=" & Text74.Text & ""
                rs!Tag = "Label18.tag= false"
                rs!Enabled = "Label18.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label18.caption=0"
                rs!Tag = "Label18.tag= true"
                rs!Enabled = "Label18.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 76 Then
            If Text76.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label8.caption=" & Text76.Text & ""
                rs!Tag = "Label8.tag= false"
                rs!Enabled = "Label8.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label8.caption=0"
                rs!Tag = "Label8.tag= true"
                rs!Enabled = "Label8.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 77 Then
            If Text78.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label7.caption=" & Text78.Text & ""
                rs!Tag = "Label7.tag= false"
                rs!Enabled = "Label7.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label7.caption=0"
                rs!Tag = "Label7.tag= true"
                rs!Enabled = "Label7.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 78 Then
            If Text77.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label9.caption=" & Text77.Text & ""
                rs!Tag = "Label9.tag= false"
                rs!Enabled = "Label9.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label9.caption=0"
                rs!Tag = "Label9.tag= true"
                rs!Enabled = "Label9.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 79 Then
            If Text79.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label26.caption=" & Text79.Text & ""
                rs!Tag = "Label26.tag= false"
                rs!Enabled = "Label26.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label26.caption=0"
                rs!Tag = "Label26.tag= true"
                rs!Enabled = "Label26.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 80 Then
            If Text81.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label25.caption=" & Text81.Text & ""
                rs!Tag = "Label25.tag= false"
                rs!Enabled = "Label25.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label25.caption=0"
                rs!Tag = "Label25.tag= true"
                rs!Enabled = "Label25.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################



        If i = 81 Then
            If Text80.Text <> "" Then
                rs.AddNew
                rs!Caption = "Label27.caption=" & Text80.Text & ""
                rs!Tag = "Label27.tag= false"
                rs!Enabled = "Label27.Enabled= false"
                rs.Update
            Else
                rs.AddNew
                rs!Caption = "Label27.caption=0"
                rs!Tag = "Label27.tag= true"
                rs!Enabled = "Label27.Enabled= true"
                rs.Update
            End If
        End If
        '#####################################










    Next i
    
MsgBox "done"
End
End Sub

Private Sub Form_Load()
Connection
End Sub
