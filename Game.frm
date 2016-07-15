VERSION 5.00
Begin VB.Form Game 
   AutoRedraw      =   -1  'True
   Caption         =   "Conway"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Myriad Arabic"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   30.25
   ScaleMode       =   0  'User
   ScaleWidth      =   75.875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmStats 
      Caption         =   "Stats"
      Height          =   1095
      Left            =   7200
      TabIndex        =   12
      Top             =   6000
      Width           =   1695
      Begin VB.Label lblDead 
         Caption         =   "Dead: 900"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblAlive 
         Caption         =   "Alive: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblGeneration 
         Caption         =   "Generation : 0"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimulate 
      Caption         =   "Start"
      Height          =   525
      Left            =   7200
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CheckBox chkEndless 
      Caption         =   "Neverending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7200
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMinGen 
      Caption         =   "-"
      Height          =   405
      Left            =   7200
      TabIndex        =   8
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton cmdPlusGen 
      Caption         =   "+"
      Height          =   405
      Left            =   8640
      TabIndex        =   7
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtGenerations 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   7560
      TabIndex        =   6
      Text            =   "25"
      Top             =   2640
      Width           =   975
   End
   Begin VB.HScrollBar hsbTime 
      Height          =   375
      Left            =   7200
      Max             =   300
      Min             =   1
      TabIndex        =   3
      Top             =   1680
      Value           =   50
      Width           =   1695
   End
   Begin VB.Timer timGame 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6240
      Top             =   240
   End
   Begin VB.HScrollBar Size 
      Height          =   375
      Left            =   7200
      Max             =   40
      Min             =   30
      TabIndex        =   1
      Top             =   720
      Value           =   30
      Width           =   1695
   End
   Begin VB.CommandButton cmdOneTick 
      Caption         =   "Tick Once"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblGenerations 
      Alignment       =   2  'Center
      Caption         =   "Generations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Time: 1x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   74
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   58
      X2              =   58
      Y1              =   0
      Y2              =   30.5
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Caption         =   "Size : 30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkEndless_Click()
If chkEndless.value = 1 Then
    txtGenerations.Enabled = False
    cmdPlusGen.Enabled = False
    cmdMinGen.Enabled = False
    endless = True
Else
    txtGenerations.Enabled = True
    cmdPlusGen.Enabled = True
    cmdMinGen.Enabled = True
    endless = False
End If


End Sub

Private Sub cmdMinGen_Click()
generations = generations - 1
txtGenerations.Text = generations
End Sub

Private Sub cmdOneTick_Click()
generation = generation + 1
Call CalculateChange
Call DisplayStats
End Sub


Private Sub cmdPlusGen_Click()
generations = generations + 1
txtGenerations.Text = generations
End Sub

Private Sub cmdReset_Click()
MsgBox ("Simulation Reset")
    timGame.Enabled = False
    generation = 0
    alive = 0
    dead = boardsize ^ 2
    Call DisplayStats
    cmdSimulate.Caption = "Start"
    chkEndless.Enabled = True
    
Call FullPopulate(False)
Call Render
End Sub

Private Sub cmdSimulate_Click()
If timGame.Enabled = False Then
    'begin sim
    cmdSimulate.Caption = "Stop"
    chkEndless.Enabled = False
    timGame.Enabled = True
Else
    'end sim
    cmdSimulate.Caption = "Start"
    chkEndless.Enabled = True
    timGame.Enabled = False
End If


End Sub

Private Sub Form_Load()
'default board size is 30, configure
boardsize = 30
generations = 25
endless = False

generation = 0
dead = 900
alive = 0

ReDim Board(0 To 30, 0 To 30)
ReDim NextBoard(0 To 30, 0 To 30)
Call FullPopulate(False) 'fill the board with true as a dummy value
Call Render

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = True
Call ClickBoard(X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mousedown = True Then
    Call ClickBoard(X, Y)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = False
lastX = 0
lastY = 0
End Sub

Private Sub hsbTime_Change()
timeRate = hsbTime.value / 50
lblTime.Caption = "Time: " + CStr(timeRate) + "x"
timGame.Interval = 500 / timeRate
End Sub

Private Sub Size_Change()
timGame.Enabled = False


boardsize = Size.value
dead = (Size.value) ^ 2
alive = 0
lblSize.Caption = "Size: " + CStr(Size.value)
ReDim Board(0 To Size.value, 0 To Size.value)
ReDim NextBoard(0 To Size.value, 0 To Size.value)
Call FullPopulate(False) 'fill the board with true as a dummy value
Call Render
End Sub

Private Sub timGame_Timer()
If generation < generations Or endless = True Then
    'continue simulation
    generation = generation + 1
    Call CalculateChange
    Call DisplayStats
Else
    MsgBox ("Simulation Complete!")
    timGame.Enabled = False
    generation = 0
    cmdSimulate.Caption = "Start"
    chkEndless.Enabled = True
End If


End Sub

Private Sub txtGenerations_Change()
generations = Val(txtGenerations.Text)
End Sub
