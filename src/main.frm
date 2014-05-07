VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H80000003&
   Caption         =   "The Guess Game!"
   ClientHeight    =   5550
   ClientLeft      =   3885
   ClientTop       =   3315
   ClientWidth     =   9195
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9195
   Begin VB.Frame mainfrm 
      BackColor       =   &H80000001&
      Caption         =   "The Guess Game"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton play 
         BackColor       =   &H80000002&
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   10
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton Cup3cmd 
         Caption         =   "Cup#3"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   9
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton Cup2cmd 
         Caption         =   "Cup#2"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton Cup1cmd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cup#1"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   7
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton replay 
         BackColor       =   &H00004080&
         Caption         =   "Replay"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   5
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox Cup3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Choose This Cup!"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Cup2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Choose This Cup!"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Cup1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   2
         ToolTipText     =   "Choose This Cup!"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label inf 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " InfoBar "
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label infobar 
         BackColor       =   &H80000016&
         Caption         =   "         Choose a Cup To Start With.                Guess Where the Pie is?"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Info Bar"
         Top             =   1800
         Width           =   5415
      End
      Begin VB.Label Title 
         BackColor       =   &H00800000&
         Caption         =   " Are You Good At Guessing?     Play This and Findout             Yourself!"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rndNum, op As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMiliseconds As Long)

Private Sub Form_Load()
    Randomize
    rndNum = Int(Rnd() * 3) + 1
End Sub

Private Sub Cup1cmd_Click()
    op = 1
    infobar.Caption = "              Cup#1 Choosen.                        Press Play to Play :)"
    Cup1.BackColor = &H8000000D
    Cup2cmd.Enabled = False
    Cup3cmd.Enabled = False
End Sub

Private Sub Cup2cmd_Click()
    op = 2
    infobar.Caption = "              Cup#2 Choosen.                        Press Play to Play :)"
    Cup2.BackColor = &H8000000D
    Cup1cmd.Enabled = False
    Cup3cmd.Enabled = False
End Sub

Private Sub Cup3cmd_Click()
    op = 3
    infobar.Caption = "              Cup#3 Choosen.                        Press Play to Play :)"
    Cup3.BackColor = &H8000000D
    Cup1cmd.Enabled = False
    Cup2cmd.Enabled = False
End Sub

Private Sub play_Click()
    infobar.Caption = "         The Pie Is At...  Cup#" & rndNum
    'Sleep 500
    Select Case rndNum
    Case 1
        Cup1.Text = "     @@@"
        Cup1.BackColor = &H80FF80
    Case 2
        Cup2.Text = "     @@@"
        Cup2.BackColor = &H80FF80
    Case 3
        Cup3.Text = "     @@@"
        Cup3.BackColor = &H80FF80
    End Select
    
    If op = rndNum Then
        infobar.Caption = "            You Guessed Right!                          You Won! :)"
    Else
        infobar.Caption = "            You Guessed Wrong!                          You Lost! :("
    End If
    Cup1cmd.Visible = False
    Cup2cmd.Visible = False
    Cup3cmd.Visible = False
End Sub

Private Sub replay_Click()
    Cup1.Text = " "
    Cup2.Text = " "
    Cup3.Text = " "
    Cup1.BackColor = &HFFFFFF
    Cup2.BackColor = &HFFFFFF
    Cup3.BackColor = &HFFFFFF
    Cup1cmd.Visible = True
    Cup2cmd.Visible = True
    Cup3cmd.Visible = True
    Cup1cmd.Enabled = True
    Cup2cmd.Enabled = True
    Cup3cmd.Enabled = True
    infobar.Caption = "         Choose a Cup To Start With.                Guess Where the Pie is?"
    rndNum = Int(Rnd() * 3) + 1
End Sub

Private Sub Form_Terminate()
    MsgBox "ThanQ for Playing...      ~~Developed By ARVEE!!"
End Sub

'Coded by Arvind Kumar aka arvee
