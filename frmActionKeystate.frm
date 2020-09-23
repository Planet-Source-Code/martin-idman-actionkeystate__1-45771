VERSION 5.00
Begin VB.Form frmActionKeystate 
   Caption         =   "Action Keystate"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "frmActionKeystate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   3660
      Top             =   990
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "Off"
      Height          =   300
      Index           =   2
      Left            =   2520
      TabIndex        =   14
      Top             =   660
      Width           =   800
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "Off"
      Height          =   300
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   330
      Width           =   800
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "Off"
      Height          =   300
      Index           =   0
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "On"
      Height          =   300
      Index           =   2
      Left            =   1710
      TabIndex        =   11
      Top             =   660
      Width           =   800
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "On"
      Height          =   300
      Index           =   1
      Left            =   1710
      TabIndex        =   10
      Top             =   330
      Width           =   800
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "On"
      Height          =   300
      Index           =   0
      Left            =   1710
      TabIndex        =   9
      Top             =   0
      Width           =   800
   End
   Begin VB.TextBox txtState 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   870
      TabIndex        =   5
      Text            =   "OFF"
      Top             =   330
      Width           =   800
   End
   Begin VB.TextBox txtState 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   870
      TabIndex        =   4
      Text            =   "OFF"
      Top             =   0
      Width           =   800
   End
   Begin VB.TextBox txtState 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   870
      TabIndex        =   3
      Text            =   "OFF"
      Top             =   660
      Width           =   800
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle"
      Height          =   300
      Index           =   1
      Left            =   3330
      TabIndex        =   2
      Top             =   330
      Width           =   800
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle"
      Height          =   300
      Index           =   0
      Left            =   3330
      TabIndex        =   1
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle"
      Height          =   300
      Index           =   2
      Left            =   3330
      TabIndex        =   0
      Top             =   660
      Width           =   800
   End
   Begin VB.Label lblLabel 
      Caption         =   "Caps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   660
      Width           =   795
   End
   Begin VB.Label lblLabel 
      Caption         =   "Num"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   795
   End
   Begin VB.Label lblLabel 
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   45
      TabIndex        =   6
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "frmActionKeystate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sub to handle form load
Private Sub Form_Load()
   ' Update state
   UpdateState
End Sub

' Sub to handle on button
Private Sub cmdOn_Click(Index As Integer)
   ' Send on (1)
   GetSetKS False, IIf(Index = 0, 1, 0), IIf(Index = 1, 1, 0), IIf(Index = 2, 1, 0)
   ' Update state
   UpdateState
End Sub

' Sub to handle off button
Private Sub cmdOff_Click(Index As Integer)
   ' Send off (2)
   GetSetKS False, IIf(Index = 0, 2, 0), IIf(Index = 1, 2, 0), IIf(Index = 2, 2, 0)
   ' Update state
   UpdateState
End Sub

' Sub to handle toggle button
Private Sub cmdToggle_Click(Index As Integer)
   ' Send toggle (3)
   GetSetKS False, IIf(Index = 0, 3, 0), IIf(Index = 1, 3, 0), IIf(Index = 2, 3, 0)
   ' Update state
   UpdateState
End Sub

' Sub to handle update state
Private Sub UpdateState()
   ' Dimension temp variables
   Dim strTemp As String
   Dim intTemp As Integer
   ' Get state from function
   strTemp = GetSetKS(True)
   Debug.Print strTemp
   ' Fill in states
   For intTemp = 0 To 2
      txtState(intTemp).Text = IIf(Split(strTemp, ",")(intTemp) = "1", "On", "Off")
   Next intTemp
End Sub

' Sub to refresh status
Private Sub tmrTimer_Timer()
   ' Call sub to update every second
   UpdateState
End Sub
