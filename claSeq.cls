VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colFrame"
'This is the structure for each animation 'Sequence'
'A sequence is a collection of related animation frames (see claFrame)
Option Explicit

Public Key As String
Public mvarImageList As String
Private mvarcolFrame As colFrame

' PROPERTIES *************************************************
Public Property Get colFrame() As colFrame
    If mvarcolFrame Is Nothing Then
        Set mvarcolFrame = New colFrame
    End If

    Set colFrame = mvarcolFrame
End Property

Public Property Set colFrame(vData As colFrame)
    Set mvarcolFrame = vData
End Property

Public Property Let ImageList(ByVal vData As String)
    mvarImageList = vData
End Property

Public Property Get ImageList() As String
    ImageList = mvarImageList
End Property

Private Sub Class_Terminate()
    Set mvarcolFrame = Nothing
End Sub



