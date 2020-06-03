VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00AEDEF4&
      Caption         =   "Label1"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer lblmusic 
      CausesValidation=   0   'False
      Height          =   975
      Left            =   6480
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
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
      _cx             =   2355
      _cy             =   1720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim st As String
'''Dim i As Integer
'''Private Sub Command1_Click()
'''WindowsMediaPlayer1.URL = "C:\WINDOWS\Media\town.mid"
'''End Sub
'''
'''Private Sub Form_Load()
'''i = 0
'''End Sub
'''
'''Private Sub Timer1_Timer()
'''i = i + 1
'''If i = 78 Then
'''Command1_Click
'''i = 0
'''End If
'''End Sub
'''
'''Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
'''Debug.Print
'''End Sub
'''
''''Dim abc As Integer
''''
''''Private Sub Command1_Click()
''''pb.Max = 100
'''''  pb.Value = pb.Max / 6
'''' Timer1.Enabled = True
''''End Sub
''''
''''Private Sub Command2_KeyPress(KeyAscii As Integer)
''''If abc = 1 Then
''''        If KeyAscii = 49 Then
''''        Label1.Caption = 1
''''     ElseIf KeyAscii = 50 Then
''''        Label1.Caption = 2
''''         ElseIf KeyAscii = 51 Then
''''        Label1.Caption = 3
''''         ElseIf KeyAscii = 52 Then
''''        Label1.Caption = 4
''''         ElseIf KeyAscii = 53 Then
''''        Label1.Caption = 5
''''        End If
''''ElseIf abc = 2 Then
''''     If KeyAscii = 49 Then
''''        Label2.Caption = 1
''''    End If
''''End If
''''End Sub
''''
''''Private Sub Form_Load()
'''' Timer1.Enabled = False
''''' Animation1.
''''End Sub
''''
''''
''''Private Sub HScroll1_Change()
''''
''''End Sub
''''
''''Private Sub hs_Change()
''''r = hs.Value
''''Form1.BackColor = vbRed * r
''''End Sub
''''
''''
''''
''''Private Sub Label1_Click()
''''abc = 1
''''Command2.SetFocus
''''End Sub
''''
''''Private Sub Label2_Click()
''''abc = 2
''''Command2.SetFocus
''''End Sub
''''
''''Private Sub Timer1_Timer()
''''
'''''Label1.BackColor = vbWhite
''''
''''If Timer1.Interval = 500 Then
''''    Label1.BackColor = vbWhite
''''    pb.Value = pb.Value + pb.Max / 7
''''ElseIf Timer1.Interval < 500 Then
''''Timer1.Interval = 500
''''Exit Sub
''''End If
''''If Timer1.Interval = 800 Then
''''    Label2.BackColor = vbWhite
''''    pb.Value = pb.Value + pb.Max / 7
''''ElseIf Timer1.Interval < 800 Then
''''Timer1.Interval = 800
''''Exit Sub
''''End If
''''
''''
''''If Timer1.Interval = 1100 Then
''''    Label3.BackColor = vbWhite
''''    pb.Value = pb.Value + pb.Max / 7
''''ElseIf Timer1.Interval < 1100 Then
'''' Timer1.Interval = 1100
''''Exit Sub
''''End If
''''
''''If Timer1.Interval = 1400 Then
''''    Label4.BackColor = vbWhite
''''    pb.Value = pb.Value + pb.Max / 7
''''ElseIf Timer1.Interval < 1400 Then
'''' Timer1.Interval = 1400
''''Exit Sub
''''End If
''''
''''If Timer1.Interval = 1700 Then
''''    Label5.BackColor = vbWhite
''''    pb.Value = pb.Value + pb.Max / 7
''''ElseIf Timer1.Interval < 1700 Then
'''' Timer1.Interval = 1700
''''Exit Sub
''''End If
''''
''''If Timer1.Interval = 2000 Then
''''    Label6.BackColor = vbWhite
''''    If pb.Value < pb.Max Then
''''    pb.Value = pb.Value + pb.Max / 7
''''    End If
''''ElseIf Timer1.Interval < 2000 Then
'''' Timer1.Interval = 2000
''''Exit Sub
''''End If
''''
''''End Sub
''''
''''
''''
'''Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)
'''
'''End Sub
'''
'''Private Sub WindowsMediaPlayer1_PlayerReconnect()
'''WindowsMediaPlayer1.URL = "C:\WINDOWS\Media\town.mid"
'''End Sub


Private Function ttttttttt(Label75 As Label) As Boolean

    If Label75.Caption = 3 Then
        lblmusic.URL = "C:\WINDOWS\Media\Windows Recycle.wav"
    Else
        lblmusic.URL = "C:\WINDOWS\Media\Windows Print complete.wav"
    End If
End Function

Private Sub Command1_Click()
    Label1.Caption = 3
    st = "Label1"
    If ttttttttt(Label1) = True Then
    
    End If
End Sub

Private Sub Command2_Click()
    Label1.Caption = 0
        If ttttttttt(Label1) = True Then
    
    End If
End Sub

Private Sub Label1_Click()

End Sub
