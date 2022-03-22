VERSION 5.00
Begin VB.UserControl WooButton 
   BackStyle       =   0  '͸��
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   ScaleHeight     =   405
   ScaleWidth      =   1005
   ToolboxBitmap   =   "WooButton.ctx":0000
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   630
   End
   Begin VB.Shape ShapeBG 
      BorderColor     =   &H00EBAE1D&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   400
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1000
   End
End
Attribute VB_Name = "WooButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum ShapeType
    Rectangle������ = 0
    Square������ = 1
    Oval��Բ�� = 2
    CircleԲ�� = 3
    RoundedRectangleԲ�ǳ����� = 4
    RoundedSquareԲ�������� = 5
End Enum

Public Event click()
Attribute click.VB_Description = "��ť����¼�"

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShapeBG.BorderStyle = 0
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShapeBG.BorderStyle = 6
    RaiseEvent click
End Sub

Private Sub UserControl_Resize()
    ShapeBG.Width = UserControl.ScaleWidth
    ShapeBG.Height = UserControl.ScaleHeight
    lbl.Left = (UserControl.ScaleWidth - lbl.Width) / 2
    lbl.Top = (UserControl.ScaleHeight - lbl.Height) / 2
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "text", lbl.Caption
    PropBag.WriteProperty "borderWidth", ShapeBG.borderWidth
    PropBag.WriteProperty "shape", ShapeBG.shape
    PropBag.WriteProperty "borderColor", ShapeBG.borderColor
    PropBag.WriteProperty "bgColor", ShapeBG.FillColor
    PropBag.WriteProperty "fontColor", lbl.ForeColor
    PropBag.WriteProperty "font", lbl.font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lbl.Caption = PropBag.ReadProperty("text")
    ShapeBG.borderWidth = PropBag.ReadProperty("borderWidth")
    ShapeBG.shape = PropBag.ReadProperty("shape", 4)
    ShapeBG.borderColor = PropBag.ReadProperty("borderColor", &H808000)
    ShapeBG.FillColor = PropBag.ReadProperty("bgColor", &HC0FFC0)
    lbl.ForeColor = PropBag.ReadProperty("fontColor", &H80000008)
    Set lbl.font = PropBag.ReadProperty("font", UserControl.font)
End Sub


Public Property Let text(ByVal strText As String)
Attribute text.VB_Description = "��ť��ʾ����"
Attribute text.VB_ProcData.VB_Invoke_PropertyPut = ";����"
    lbl.Caption = strText
    UserControl_Resize
    PropertyChanged "text"
End Property
Public Property Get text() As String
    text = lbl.Caption
End Property

Public Property Let shape(ByVal btnShape As ShapeType)
Attribute shape.VB_Description = "��ť��״"
Attribute shape.VB_ProcData.VB_Invoke_PropertyPut = ";���"
    ShapeBG.shape = btnShape
    UserControl_Resize
    PropertyChanged "shape"
End Property
Public Property Get shape() As ShapeType
    shape = ShapeBG.shape
End Property

Public Property Let borderWidth(ByVal bd As Integer)
Attribute borderWidth.VB_Description = "��ť������ϸ�����أ�"
Attribute borderWidth.VB_ProcData.VB_Invoke_PropertyPut = ";���"
    ShapeBG.borderWidth = bd
    PropertyChanged "borderWidth"
End Property
Public Property Get borderWidth() As Integer
    borderWidth = ShapeBG.borderWidth
End Property

Public Property Let borderColor(ByVal clr As OLE_COLOR)
Attribute borderColor.VB_Description = "��ť�߿���ɫ"
Attribute borderColor.VB_ProcData.VB_Invoke_PropertyPut = ";���"
    ShapeBG.borderColor = clr
    PropertyChanged "borderColor"
End Property
Public Property Get borderColor() As OLE_COLOR
    borderColor = ShapeBG.borderColor
End Property

Public Property Let bgColor(ByVal clr As OLE_COLOR)
Attribute bgColor.VB_Description = "��ť������ɫ"
Attribute bgColor.VB_ProcData.VB_Invoke_PropertyPut = ";���"
    ShapeBG.FillColor = clr
    PropertyChanged "bgColor"
End Property
Public Property Get bgColor() As OLE_COLOR
    bgColor = ShapeBG.FillColor
End Property

Public Property Let fontColor(ByVal clr As OLE_COLOR)
    lbl.ForeColor = clr
    PropertyChanged "fontColor"
End Property
Public Property Get fontColor() As OLE_COLOR
    fontColor = lbl.ForeColor
End Property

Public Property Set font(ByVal newFont As StdFont)
    Set lbl.font = newFont
    UserControl_Resize
    PropertyChanged "font"
End Property
Public Property Get font() As StdFont
    Set font = lbl.font
End Property

