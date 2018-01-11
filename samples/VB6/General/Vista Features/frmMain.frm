VERSION 5.00
Object = "{52D76F35-4551-4C56-B53B-A343E42B0AF8}#2.6#0"; "ProgBarU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ProgressBar 2.6 - Vista Features Sample"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame2 
      Caption         =   "&Current Value"
      Height          =   975
      Left            =   4200
      TabIndex        =   6
      Top             =   600
      Width           =   3735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Kein
         Height          =   615
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   7
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox chkSmoothReverse 
            Caption         =   "S&mooth Reverse"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Value           =   1  'Aktiviert
            Width           =   1575
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Appl&y"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2520
            TabIndex        =   10
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox txtCurrentValue 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Text            =   "0"
            Top             =   210
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&State"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   615
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   249
         TabIndex        =   2
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton optState 
            Caption         =   "&Failed"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optState 
            Caption         =   "&Paused"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optState 
            Caption         =   "&Normal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
   End
   Begin ProgBarLibUCtl.ProgressBar ProgBarU 
      Height          =   255
      Left            =   120
      Top             =   180
      Width           =   5985
      _cx             =   10557
      _cy             =   450
      ActivateMarquee =   -1  'True
      Appearance      =   3
      BackColor       =   -1
      BarColor        =   -1
      BarStyle        =   0
      BorderStyle     =   0
      CurrentValue    =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   3
      DisplayText     =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      HAlignment      =   1
      HoverTime       =   -1
      MarqueeStepDuration=   50
      Maximum         =   100
      Minimum         =   0
      MousePointer    =   0
      Orientation     =   0
      ProgressState   =   1
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      RightToLeftLayout=   0   'False
      SmoothReverse   =   -1  'True
      StepWidth       =   10
      SupportOLEDragImages=   -1  'True
      TextShadowColor =   3158064
      TextShadowOffsetX=   1
      TextShadowOffsetY=   1
      UseSystemFont   =   -1  'True
      DisplayProgressInTaskBar=   0   'False
      Text            =   "frmMain.frx":0000
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Sub chkSmoothReverse_Click()
  ProgBarU.SmoothReverse = (chkSmoothReverse.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  ProgBarU.About
End Sub

Private Sub cmdApply_Click()
  On Error Resume Next
  ProgBarU.CurrentValue = CInt(txtCurrentValue.Text)
  If Err Then
    txtCurrentValue.Text = "0"
    ProgBarU.CurrentValue = 0
  End If
  cmdApply.Enabled = False
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  Const ES_NUMBER = &H2000&
  Const GWL_STYLE = -16
  Dim style As Long

  style = GetWindowLong(txtCurrentValue.hWnd, GWL_STYLE)
  style = style Or ES_NUMBER
  SetWindowLong txtCurrentValue.hWnd, GWL_STYLE, style
End Sub

Private Sub optState_Click(Index As Integer)
  If optState(0).Value Then
    ProgBarU.ProgressState = ProgressStateConstants.psNormal
  ElseIf optState(1).Value Then
    ProgBarU.ProgressState = ProgressStateConstants.psPaused
  ElseIf optState(2).Value Then
    ProgBarU.ProgressState = ProgressStateConstants.psFailed
  End If
End Sub

Private Sub txtCurrentValue_Change()
  cmdApply.Enabled = True
End Sub
