VERSION 5.00
Begin VB.UserControl ZoomPic 
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   ScaleHeight     =   2505
   ScaleWidth      =   3300
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   45
      ScaleHeight     =   2490
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      Begin VB.Timer Timer1 
         Left            =   2925
         Top             =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   6165
      Top             =   4680
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "ZoomPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim AutoAniValue As Integer
Dim AutoAniInd As Integer


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    PropertyChanged "Picture"
    'Picture1.PaintPicture Image1.Picture, 1, 1, Picture1.Width, Picture1.Height
End Property



Sub ZoomByRatio(ZoomRatio As Integer)
Dim X As Integer, Y As Integer
X = (Image1.Width / 2) * ZoomRatio / 100
Y = (Image1.Height / 2) * ZoomRatio / 100
Picture1.PaintPicture Image1.Picture, 1, 1, Picture1.Width, Picture1.Height, X, Y, Image1.Width - X, Image1.Height - Y

End Sub


Private Sub Timer1_Timer()
If AutoAniValue + AutoAniInd > 150 Then
   AutoAniInd = AutoAniInd * -1
End If

If AutoAniValue + AutoAniInd < 1 Then
   AutoAniInd = AutoAniInd * -1
End If
AutoAniValue = AutoAniValue + AutoAniInd

Me.ZoomByRatio AutoAniValue

End Sub

Private Sub UserControl_Initialize()
AutoAniValue = 1
AutoAniInd = 1
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Timer1.Enabled = PropBag.ReadProperty("AutomaticAnimation", True)
    Timer1.Interval = PropBag.ReadProperty("Interval", 0)
End Sub

Private Sub UserControl_Resize()
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutomaticAnimation", Timer1.Enabled, True)
    Call PropBag.WriteProperty("Interval", Timer1.Interval, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Enabled
Public Property Get AutomaticAnimation() As Boolean
Attribute AutomaticAnimation.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    AutomaticAnimation = Timer1.Enabled
End Property

Public Property Let AutomaticAnimation(ByVal New_AutomaticAnimation As Boolean)
    Timer1.Enabled() = New_AutomaticAnimation
    PropertyChanged "AutomaticAnimation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    Timer1.Interval() = New_Interval
    PropertyChanged "Interval"
End Property

