VERSION 5.00
Begin VB.UserControl WooTextBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00EBAE1D&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   ScaleHeight     =   375
   ScaleWidth      =   1995
   ToolboxBitmap   =   "WooTextBox.ctx":0000
   Begin VB.TextBox txt2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   50
      Visible         =   0   'False
      Width           =   1900
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   1900
   End
End
Attribute VB_Name = "WooTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private bdSize As Integer

Private Sub UserControl_Resize()
    Dim n As Long
    n = bdSize * Screen.TwipsPerPixelX
    txt.Width = UserControl.ScaleWidth - 2 * n
    txt.Height = UserControl.ScaleHeight - 2 * n
    txt.Left = n
    txt.Top = n
    txt2.Left = txt.Left
    txt2.Top = txt.Top
    txt2.Width = txt.Width
    txt2.Height = txt.Height
End Sub

Private Sub UserControl_InitProperties()
    bdSize = 1
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "text", txt.text
    PropBag.WriteProperty "passwordChar", txt.passwordChar
    PropBag.WriteProperty "borderSize", bdSize
    PropBag.WriteProperty "borderColor", UserControl.BackColor
    PropBag.WriteProperty "bgColor", txt.BackColor
    PropBag.WriteProperty "fontColor", txt.ForeColor
    PropBag.WriteProperty "font", txt.font
    PropBag.WriteProperty "multiLine", txt2.Visible
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txt.text = PropBag.ReadProperty("text")
    txt.passwordChar = PropBag.ReadProperty("passwordChar")
    bdSize = PropBag.ReadProperty("borderSize", 1)
    UserControl.BackColor = PropBag.ReadProperty("borderColor")
    txt.BackColor = PropBag.ReadProperty("bgColor")
    txt.ForeColor = PropBag.ReadProperty("fontColor")
    Set txt.font = PropBag.ReadProperty("font")
    txt2.Visible = PropBag.ReadProperty("multiLine", False)
    txt.Visible = (Not txt2.Visible)
    txt2.passwordChar = txt.passwordChar
    txt2.BackColor = txt.BackColor
    txt2.ForeColor = txt.ForeColor
    Set txt2.font = txt.font
End Sub

Private Sub txt_Change()
    If txt2.text <> txt.text Then txt2.text = txt.text
End Sub
Private Sub txt2_Change()
    If txt.text <> txt2.text Then txt.text = txt2.text
End Sub


Public Property Let text(ByVal strText As String)
Attribute text.VB_Description = "文本内容"
Attribute text.VB_ProcData.VB_Invoke_PropertyPut = ";数据"
    txt.text = strText
    PropertyChanged "text"
End Property
Public Property Get text() As String
    text = txt.text
End Property

Public Property Let passwordChar(ByVal char As String)
Attribute passwordChar.VB_Description = "作为密码框时显示的符号"
Attribute passwordChar.VB_ProcData.VB_Invoke_PropertyPut = ";外观"
    txt.passwordChar = char
    txt2.passwordChar = char
    PropertyChanged "passwordChar"
End Property
Public Property Get passwordChar() As String
    passwordChar = txt.passwordChar
End Property

Public Property Let borderSize(ByVal pixSize As Integer)
    bdSize = pixSize
    UserControl_Resize
    PropertyChanged "borderSize"
End Property
Public Property Get borderSize() As Integer
Attribute borderSize.VB_Description = "文本框的边框粗细（像素）"
Attribute borderSize.VB_ProcData.VB_Invoke_Property = ";外观"
    borderSize = bdSize
End Property

Public Property Let borderColor(ByVal clr As OLE_COLOR)
    UserControl.BackColor = clr
    PropertyChanged "borderColor"
End Property
Public Property Get borderColor() As OLE_COLOR
Attribute borderColor.VB_Description = "文本框的边框颜色"
Attribute borderColor.VB_ProcData.VB_Invoke_Property = ";外观"
    borderColor = UserControl.BackColor
End Property

Public Property Let bgColor(ByVal clr As OLE_COLOR)
    txt.BackColor = clr
    txt2.BackColor = clr
    PropertyChanged "bgColor"
End Property
Public Property Get bgColor() As OLE_COLOR
Attribute bgColor.VB_Description = "文本框背景色"
Attribute bgColor.VB_ProcData.VB_Invoke_Property = ";外观"
    bgColor = txt.BackColor
End Property

Public Property Let fontColor(ByVal clr As OLE_COLOR)
Attribute fontColor.VB_Description = "文本框字体颜色"
Attribute fontColor.VB_ProcData.VB_Invoke_PropertyPut = ";外观"
    txt.ForeColor = clr
    txt2.ForeColor = clr
    PropertyChanged "fontColor"
End Property
Public Property Get fontColor() As OLE_COLOR
    fontColor = txt.ForeColor
End Property

Public Property Set font(ByVal newFont As StdFont)
    Set txt.font = newFont
    Set txt2.font = newFont
    PropertyChanged "font"
End Property
Public Property Get font() As StdFont
Attribute font.VB_Description = "文本框字体属性"
Attribute font.VB_ProcData.VB_Invoke_Property = ";外观"
    Set font = txt.font
End Property

Public Property Let multiLine(ByVal multi As Boolean)
Attribute multiLine.VB_Description = "文本框是否允许多行输入（带滚动条）"
Attribute multiLine.VB_ProcData.VB_Invoke_PropertyPut = ";外观"
    txt.Visible = (Not multi)
    txt2.Visible = multi
    PropertyChanged "multiLine"
End Property
Public Property Get multiLine() As Boolean
    multiLine = txt2.Visible
End Property

