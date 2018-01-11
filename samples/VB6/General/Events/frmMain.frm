VERSION 5.00
Object = "{52D76F35-4551-4C56-B53B-A343E42B0AF8}#2.6#0"; "ProgBarU.ocx"
Object = "{74B9FB9B-58D2-4789-9052-ED4EF2ADA75F}#2.6#0"; "ProgBarA.ocx"
Begin VB.Form frmMain 
   Caption         =   "ProgressBar 2.6 - Events Sample"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
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
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   9120
      Top             =   2280
   End
   Begin ProgBarLibACtl.ProgressBar ProgBarA 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      Top             =   5265
      Width           =   10545
      _cx             =   2646
      _cy             =   529
      ActivateMarquee =   -1  'True
      Appearance      =   3
      BackColor       =   -1
      BarColor        =   -1
      BarStyle        =   0
      BorderStyle     =   0
      CurrentValue    =   10
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   0
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
   Begin ProgBarLibUCtl.ProgressBar ProgBarU 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      Top             =   5520
      Width           =   10545
      _cx             =   18600
      _cy             =   450
      ActivateMarquee =   -1  'True
      Appearance      =   3
      BackColor       =   -1
      BarColor        =   -1
      BarStyle        =   0
      BorderStyle     =   0
      CurrentValue    =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   0
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
      Text            =   "frmMain.frx":002A
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   8760
      TabIndex        =   0
      Top             =   600
      Width           =   975
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
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtLog 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private bLog As Boolean


  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  ProgBarU.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    On Error Resume Next
    cmdAbout.Left = Me.ScaleWidth - 5 - cmdAbout.Width
    chkLog.Left = cmdAbout.Left
    txtLog.Width = cmdAbout.Left - 5 - txtLog.Left
    txtLog.Height = ProgBarA.Top - 2 - txtLog.Top
  End If
End Sub


Private Sub ProgBarA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ProgBarA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ProgBarA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ProgBarA_DragDrop"
End Sub

Private Sub ProgBarA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "ProgBarA_DragOver"
End Sub

Private Sub ProgBarA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "ProgBarA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_OLEDragDrop(ByVal Data As ProgBarLibACtl.IOLEDataObject, effect As ProgBarLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarA_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ProgBarA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ProgBarA_OLEDragEnter(ByVal Data As ProgBarLibACtl.IOLEDataObject, effect As ProgBarLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarA_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ProgBarA_OLEDragLeave(ByVal Data As ProgBarLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarA_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ProgBarA_OLEDragMouseMove(ByVal Data As ProgBarLibACtl.IOLEDataObject, effect As ProgBarLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarA_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ProgBarA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ProgBarA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ProgBarA_ResizedControlWindow()
  AddLogEntry "ProgBarA_ResizedControlWindow"
  txtLog.Height = ProgBarA.Top - 2 - txtLog.Top
End Sub

Private Sub ProgBarA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ProgBarU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ProgBarU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ProgBarU_DragDrop"
End Sub

Private Sub ProgBarU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "ProgBarU_DragOver"
End Sub

Private Sub ProgBarU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "ProgBarU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_OLEDragDrop(ByVal Data As ProgBarLibUCtl.IOLEDataObject, effect As ProgBarLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarU_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ProgBarU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ProgBarU_OLEDragEnter(ByVal Data As ProgBarLibUCtl.IOLEDataObject, effect As ProgBarLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarU_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ProgBarU_OLEDragLeave(ByVal Data As ProgBarLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarU_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ProgBarU_OLEDragMouseMove(ByVal Data As ProgBarLibUCtl.IOLEDataObject, effect As ProgBarLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "ProgBarU_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub ProgBarU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ProgBarU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ProgBarU_ResizedControlWindow()
  AddLogEntry "ProgBarU_ResizedControlWindow"
End Sub

Private Sub ProgBarU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ProgBarU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ProgBarU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub Timer1_Timer()
  ProgBarU.ChangeCurrentValue
  Select Case ProgBarU.CurrentValue
    Case ProgBarU.Maximum
      ProgBarU.StepWidth = -10
    Case ProgBarU.Minimum
      ProgBarU.StepWidth = 10
  End Select
  ProgBarA.ChangeCurrentValue
  Select Case ProgBarA.CurrentValue
    Case ProgBarA.Maximum
      ProgBarA.StepWidth = -10
    Case ProgBarA.Minimum
      ProgBarA.StepWidth = 10
  End Select
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub
