VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TeamData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTeamName As String
Private mTokuten As Integer
Private mTokushittensa As Integer
Private mRank As Integer
Private mSoutokuten As Integer
Private mShourisuu As Integer
Private mHaisensuu As Integer
Private mHikiwakesuu As Integer
Private mShiaisuu As Integer

Private Sub Class_Initialize()
    mShourisuu = 0
    mHaisensuu = 0
    mHikiwakesuu = 0
    mShiaisuu = 0
End Sub

Public Property Get Shiaisuu() As Integer
    Shiaisuu = mShiaisuu
End Property
Public Property Let Shiaisuu(ByVal suuji As Integer)
    mShiaisuu = suuji
End Property


Public Property Get Hikiwakesuu() As Integer
    Hikiwakesuu = mHikiwakesuu
End Property
Public Property Let Hikiwakesuu(ByVal suuji As Integer)
    mHikiwakesuu = suuji
End Property

Public Property Get Haisensuu() As Integer
    Haisensuu = mHaisensuu
End Property
Public Property Let Haisensuu(ByVal suuji As Integer)
    mHaisensuu = suuji
End Property

Public Property Get Shourisuu() As Integer
    Shourisuu = mShourisuu
End Property
Public Property Let Shourisuu(ByVal suuji As Integer)
    mShourisuu = suuji
End Property

Public Property Get Soutokuten() As Integer
    Soutokuten = mSoutokuten
End Property
Public Property Let Soutokuten(ByVal suuji As Integer)
    mSoutokuten = suuji
End Property

Public Property Get Rank() As Integer
    Rank = mRank
End Property
Public Property Let Rank(ByVal suuji As Integer)
    mRank = suuji
End Property

Public Property Get teamName() As String
    teamName = mTeamName
End Property
Public Property Let teamName(ByVal text As String)
    mTeamName = text
End Property

Public Property Get Tokuten() As Integer
    Tokuten = mTokuten
End Property
Public Property Let Tokuten(ByVal suuji As Integer)
    mTokuten = suuji
End Property

Public Property Get Tokushittensa() As Integer
    Tokushittensa = mTokushittensa
End Property
Public Property Let Tokushittensa(ByVal suuji As Integer)
    mTokushittensa = suuji
End Property

