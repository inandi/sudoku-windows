VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmGoogle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gobinda Search"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   18540
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser webg 
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18255
      ExtentX         =   32200
      ExtentY         =   17595
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmGoogle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
