VERSION 5.00
Begin VB.UserControl SysTray 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "SysTray.ctx":0000
   ScaleHeight     =   780
   ScaleWidth      =   2565
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   75
      Picture         =   "SysTray.ctx":0014
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   90
      Width           =   615
   End
End
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'AUTHOR:    Peter Morgan    <firefly_nz24@yahoo.com>
'DATE:      25 August 2000
'
'Please feel free to use this code how you see fit
'Any comments or enhancements to the existing control would
'be great. If someone would like to double-team on a project,
'please drop me a line
'
'
'CREDITS:
'
'   * Jonathan Morrison  <jonathanm@mindspring.com>
'   * Pascal van de Wijdeven
'
'Thanks guys
Option Explicit

'Property Variables:
Public stSequence As New colSeq         'this object holds all frame and sequencing information

Private m_Animate As Boolean            'enable animated sequence?
Private m_CurrentSequence As Integer    'current sequence number
Private m_InitialSequence As Integer    'starting sequence number
Private m_InitialImage As Integer       'starting image number
Private m_TipText As String             'text to be displayed on system tray icon
Private intCurrent As Currency          'timer variable used with animation frame change
Private intTotal As Integer             'number of seconds before next frame is displayed
Private intPos As Integer               'position in sequence
Private blnBegin As Boolean             'can we start yet? stops timer trigger in design mode
Private blnRepeat As Boolean            'should animation repeat or stop at last frame?

' Default values
Private Const def_InitialImage = 1
Private Const def_InitialSequence = 1
Private Const def_CurrentSequence = 1
Private Const def_Animate = False
Private Const def_TipText = ""

' PROPERTIES *************************************************
Public Property Get Animate() As Boolean
Attribute Animate.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Animate = m_Animate
End Property

Public Property Let Animate(ByVal New_Animate As Boolean)
    m_Animate = New_Animate
    PropertyChanged "Animate"
    

End Property

Public Property Get CurrentSequence() As Integer
Attribute CurrentSequence.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    CurrentSequence = m_CurrentSequence
End Property

Public Property Let CurrentSequence(ByVal New_CurrentSequence As Integer)
    m_CurrentSequence = New_CurrentSequence
    PropertyChanged "CurrentSequence"
End Property

Public Property Get InitialImage() As Integer
Attribute InitialImage.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InitialImage = m_InitialImage
End Property

Public Property Let InitialImage(ByVal New_InitialImage As Integer)
    m_InitialImage = New_InitialImage
    PropertyChanged "InitialImage"
End Property

Public Property Get InitialSequence() As Integer
Attribute InitialSequence.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    InitialSequence = m_InitialSequence
End Property

Public Property Let InitialSequence(ByVal New_InitialSequence As Integer)
    m_InitialSequence = New_InitialSequence
    PropertyChanged "InitialSequence"
End Property

Public Property Get TipText() As String
Attribute TipText.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    TipText = m_TipText
End Property

Public Property Let TipText(ByVal New_TipText As String)
    m_TipText = New_TipText
    PropertyChanged "TipText"
End Property

Private Sub UserControl_InitProperties()
'Initialize Properties for User Control
    m_InitialImage = def_InitialImage
    m_InitialSequence = def_InitialSequence
    m_CurrentSequence = def_CurrentSequence
    m_Animate = def_Animate
    m_TipText = def_TipText
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Load property values from storage
    Set frm = Parent
    
    m_InitialImage = PropBag.ReadProperty("InitialImage", def_InitialImage)
    m_InitialSequence = PropBag.ReadProperty("InitialSequence", def_InitialSequence)
    m_CurrentSequence = PropBag.ReadProperty("CurrentSequence", def_CurrentSequence)
    m_Animate = PropBag.ReadProperty("Animate", def_Animate)
    m_TipText = PropBag.ReadProperty("TipText", def_TipText)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Write property values to storage
    Call PropBag.WriteProperty("InitialImage", m_InitialImage, def_InitialImage)
    Call PropBag.WriteProperty("InitialSequence", m_InitialSequence, def_InitialSequence)
    Call PropBag.WriteProperty("CurrentSequence", m_CurrentSequence, def_CurrentSequence)
    Call PropBag.WriteProperty("Animate", m_Animate, def_Animate)
    Call PropBag.WriteProperty("TipText", m_TipText, def_TipText)
End Sub

Private Sub UserControl_Resize()
'Keep the control a specific height and width at design-time
    UserControl.Height = 780
    UserControl.Width = 780
End Sub

' USER CONTROL FUNCTIONS **************************************
Public Function SendToTray()
'Create effect where form minimizes into the tray
    Dim lngRetVal As Long

    ZoomForm ZOOM_TO_TRAY, frm.hwnd
    frm.Visible = False 'hide the form from view
    Picture2.Picture = frm.Icon 'store original icon from restoration on terminate
    m_CurrentSequence = m_InitialSequence   'init sequence
    
    'take the specified initial image
    frm.Icon = frm.Controls(stSequence(m_InitialSequence).ImageList).ListImages(m_InitialImage).Picture

    Set IconObject = frm.Icon
    'create the new icon on the system tray
    AddIcon frm, IconObject.Handle, IconObject, m_TipText
End Function

Public Function RestoreFromTray()
'Create the effect that original window is expanding from system tray
    delIcon IconObject.Handle   'remove icon from tray
    m_Animate = False           'stop animation
    frm.Icon = Picture2.Picture 'restore original icon
    ZoomForm ZOOM_FROM_TRAY, frm.hwnd
    frm.Visible = True          'make original form visible
End Function

Private Sub Timer1_Timer()
'Main animation cycle
    'The control must be allowed to animate
    If m_Animate = True Then
        'The icon needs to be on the tray before animation starts
        If blnBegin = True Then
            'change the animated frame if the duration has expired
            If intCurrent >= intTotal Then
                intCurrent = 0
                intPos = intPos + 1
                'loop to beginning of frames in this sequence if end is reached
                If intPos > stSequence(m_CurrentSequence).colFrame.Count Then
                    intPos = 1
                    If blnRepeat = False Then
                        'stop animation if this was a forward only animation
                        m_Animate = False
                        Exit Sub
                    End If
                End If
                Call AnimateIcon    'sub to repaint the icon
            End If
            'intCurrent is a currency variable to allow for a
            'floating point - dodgey, but it works
            intCurrent = intCurrent + (Timer1.Interval / 1000)
        End If
    End If
End Sub

Private Sub AnimateIcon()
'Repaint new icon image to system tray
    Dim intNextImage As Integer
    
    'get the next frame from the sequence
    intNextImage = stSequence(m_CurrentSequence).colFrame.Item(intPos).ImageNum
    intTotal = stSequence(m_CurrentSequence).colFrame.Item(intPos).Interval

    'set the form icon property to the next picture in the sequence
    frm.Icon = frm.Controls(stSequence(m_CurrentSequence).ImageList).ListImages(intNextImage).Picture
    intTotal = Me.stSequence(m_CurrentSequence).colFrame.Item(intPos).Interval
    'paint the forms icon to the system tray
    modIcon frm, IconObject.Handle, frm.Icon, m_TipText
End Sub

Public Sub PlayAnimation(ByVal Seq As Integer, ByVal Repeat As Boolean)
'(re)initialise variables - ready for animation play
    intPos = 1
    intCurrent = 0
    intTotal = 0
    m_CurrentSequence = Seq
    blnRepeat = Repeat
    blnBegin = True
    Timer1.Interval = 100   'frames can be changed every 1/10 second (I guess)
    
    Call AnimateIcon    'sub to repaint the icon
End Sub
