VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "V I B E P L A Y E R"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "vibeplayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1980
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2535
      Left            =   0
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   6
      Top             =   0
      Width           =   6975
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   480
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   480
      Left            =   1920
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Play"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00808000&
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "c:\winnt\media\canyon.mid"
      Top             =   2520
      Width           =   3495
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2520
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   661
      _Version        =   393216
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      OLEDropMode     =   1
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2490
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If ListIndex = 0 Then
File1.Pattern = ("*.wav")
ElseIf ListIndex = 1 Then
File1.Pattern = ("*.mid")
ElseIf ListIndex = 2 Then
File1.Pattern = ("*.avi;*.mpg")
Else
File1.Pattern = ("*.*")
End If
End Sub

Private Sub Command1_Click()
MMControl1.FileName = Text1.Text
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
If Combo1.ListIndex = 0 Then
File1.Pattern = ("*.wav")
ElseIf Combo1.ListIndex = 1 Then
File1.Pattern = ("*.mid")
ElseIf Combo1.ListIndex = 2 Then
File1.Pattern = ("*.avi;*.mpg")
Else
File1.Pattern = ("*.*")
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If Combo1.ListIndex = 0 Then
File1.Pattern = ("*.wav")
ElseIf Combo1.ListIndex = 1 Then
File1.Pattern = ("*.mid")
ElseIf Combo1.ListIndex = 2 Then
File1.Pattern = ("*.avi;*.mpg")
Else
File1.Pattern = ("*.*")
End If
If Right(File1.Path, 1) <> "\" Then
filenam = File1.Path + "\" + File1.FileName
Else
filenam = File1.Path + File1.FileName
End If
Text1.Text = filenam
End Sub

Private Sub Form_Load()
MMControl1.hWndDisplay = Picture1.hWnd
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.Command = "Close"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
Combo1.Text = "*.wav"
Combo1.AddItem "*.wav"
Combo1.AddItem "*.mid"
Combo1.AddItem "*.avi;*.mpg"
Combo1.AddItem "All files"
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
'show
End Sub

Private Sub play_Click()
MMControl1.FileName = Text1.Text
MMControl1.Command = "Open"
MMControl1.Command = "Play"
MMControl1.hWndDisplay = Picture1.hWnd
End Sub

Private Sub stop_Click()
If MMControl1.Mode = 524 Then Exit Sub
If MMControl1.Mode <> 525 Then
MMControl1.Wait = True
MMControl1.Command = "Stop"
End If
MMControl1.Wait = True
MMControl1.Command = "Close"
End Sub

Private Sub Picture1_Click()
'show
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
MMControl1.Command = "Close"
End Sub
