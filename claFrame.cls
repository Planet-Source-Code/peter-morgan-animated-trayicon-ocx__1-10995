VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'Individual frames (for the animation) are defined in this structure
'Many frames make up an animation 'Sequence'
Option Explicit

Public Key As String

Private mvarInterval As Integer
Private mvarImageNum As Integer

' PROPERTIES *************************************************
Public Property Let ImageNum(ByVal vData As Integer)
    mvarImageNum = vData
End Property

Public Property Get ImageNum() As Integer
    ImageNum = mvarImageNum
End Property

Public Property Let Interval(ByVal vData As Integer)
    mvarInterval = vData
End Property

Public Property Get Interval() As Integer
    Interval = mvarInterval
End Property



