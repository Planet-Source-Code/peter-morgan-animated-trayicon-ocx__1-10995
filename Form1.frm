VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\SYSTEM~1\projSysTray.vbp"
Begin VB.Form Form1 
   Caption         =   "System Tray ... test scenario"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin projSysTray.SysTray SysTray1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   1376
      _ExtentY        =   1376
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   3720
   End
   Begin VB.Frame Frame2 
      Caption         =   "Waiting time before result:"
      Height          =   1215
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
      Begin VB.TextBox txtDelay 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Seconds"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "By default, test will:"
      Height          =   1575
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "Fail"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Succeed"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Tray Menu"
      Begin VB.Menu mnuSend1 
         Caption         =   "Attempt test connection"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constants used with mouse click on the system tray icon
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONDBLCLK = &H203

Dim start As Boolean

Private Sub Form_Load()
' ... How can I get the adding of frames to be more like the
'adding of buttons to a toolbar control? (i.e. and save the
'results in the property bag)
    Dim stFrames As New colFrame

    'Set form default values
    txtDelay = 10

    'Load test graphics into imagelists (must be icons)
    ' ... these could come from a resource file in the exe
    ImageList1.ListImages.Add 1, , LoadPicture(App.Path & "\graphics\Chip0.ico")
    ImageList1.ListImages.Add 2, , LoadPicture(App.Path & "\graphics\Chip1.ico")
    ImageList1.ListImages.Add 3, , LoadPicture(App.Path & "\graphics\Chip2.ico")
    ImageList1.ListImages.Add 4, , LoadPicture(App.Path & "\graphics\Chip3.ico")
    ImageList1.ListImages.Add 5, , LoadPicture(App.Path & "\graphics\Chip4.ico")
    ImageList1.ListImages.Add 6, , LoadPicture(App.Path & "\graphics\Chip5.ico")
    ImageList1.ListImages.Add 7, , LoadPicture(App.Path & "\graphics\Chip6.ico")

    'Animation Sequence #1 - "connecting phase"
    'add individual frames to the frames collection
    'format = key, image_index_number, delay_in_seconds
    stFrames.Add 1, 1, 0.5
    stFrames.Add 2, 2, 0.5
    stFrames.Add 3, 3, 0.5
    stFrames.Add 4, 4, 0.5
    stFrames.Add 5, 5, 0.5
    
    stFrames.Add 6, 1, 0.5
    stFrames.Add 5, 5, 0.5
    stFrames.Add 4, 4, 0.5
    stFrames.Add 3, 3, 0.5
    stFrames.Add 2, 2, 0.5
    
    'assign these frames to sequence #1
    'specify that images will come from ImageList1
    SysTray1.stSequence.Add 1, "ImageList1", stFrames
    Set stFrames = Nothing
    
    'Sequence #2 - "connected"
    Set stFrames = New colFrame
    
    stFrames.Add 1, 1, 1
    stFrames.Add 2, 6, 0.5
    stFrames.Add 3, 1, 0.5
    stFrames.Add 4, 6, 0.5
    stFrames.Add 5, 1, 0.5
    stFrames.Add 6, 6, 1

    SysTray1.stSequence.Add 2, "ImageList1", stFrames
    Set stFrames = Nothing
    
    'Sequence #3 - "no connection"
    Set stFrames = New colFrame
    
    stFrames.Add 1, 1, 1
    stFrames.Add 2, 7, 0.5
    stFrames.Add 3, 1, 0.5
    stFrames.Add 4, 7, 0.5
    stFrames.Add 5, 1, 0.5
    stFrames.Add 6, 7, 0.5
    
    SysTray1.stSequence.Add 3, "ImageList1", stFrames
    Set stFrames = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Enable the popup menu for the system tray icon
' ... is there a better way to do this ???
    Static Message As Long
    Message = x / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_LBUTTONDBLCLK
            Call mnuRestore_Click
        Case WM_RBUTTONUP:
            Me.PopupMenu mnuPopUp
    End Select
End Sub

Private Sub mnuSend1_Click()
    start = True
    Timer1.Interval = 1000
    ToggleMenu True
    SysTray1.Animate = True
    SysTray1.InitialSequence = 1
    SysTray1.InitialImage = 1
    SysTray1.TipText = "Attempting Connection ..."
    SysTray1.SendToTray
    SysTray1.PlayAnimation 1, True
End Sub

Private Sub mnuRestore_Click()
    ToggleMenu False
    
    SysTray1.RestoreFromTray
End Sub

Private Sub ToggleMenu(ByVal blnSendToTray As Boolean)
    If blnSendToTray = True Then
        mnuSend1.Enabled = False
        mnuRestore.Enabled = True
    Else
        mnuSend1.Enabled = True
        mnuRestore.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    Static x As Integer

    If start = True Then
        x = x + 1
        If x >= txtDelay.Text Then
            Timer1.Interval = 0
            x = 0
            If Option1.Value = True Then
                'display success
                SysTray1.TipText = "Success"
                SysTray1.PlayAnimation 2, False
            End If
            If Option2.Value = True Then
                'display failure
                SysTray1.TipText = "Failure"
                SysTray1.PlayAnimation 3, False
            End If
        End If
    End If
End Sub
