VERSION 5.00
Begin VB.Form frmcolor 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   Caption         =   "ColorAdjustment"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcolorcl 
      BackColor       =   &H000000FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar hs3 
      Height          =   375
      Left            =   120
      Max             =   100
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.HScrollBar hs2 
      Height          =   375
      Left            =   120
      Max             =   100
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.HScrollBar hs1 
      Height          =   375
      Left            =   120
      Max             =   100
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Press Esc To Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmcolor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcolorcl_Click()
 Unload Me
End Sub

Private Sub hs1_Change()
    SOD.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltimer1.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbldiv.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltimer2.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltime.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbllog.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lblUserMsg.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lblmarquee.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
End Sub

Private Sub hs1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub hs2_Change()
    SOD.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltimer1.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbldiv.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltimer2.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltime.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbllog.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lblUserMsg.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lblmarquee.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
End Sub

Private Sub hs2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub hs3_Change()
    SOD.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltimer1.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbldiv.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltimer2.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbltime.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lbllog.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lblUserMsg.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
    SOD.lblmarquee.BackColor = RGB(hs1.Value, hs2.Value, hs3.Value)
End Sub

Private Sub hs3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       Unload Me
    End If

End Sub
