VERSION 5.00
Object = "{468AD253-514F-430C-B56B-46E81B982236}#68.0#0"; "WooSkin.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00F9EEE4&
   Caption         =   "WooSkin Demo"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4965
   StartUpPosition =   2  '屏幕中心
   Begin WooSkin.WooTextBox WooTextBoxB 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      text            =   "password"
      passwordChar    =   "*"
      borderSize      =   1
      borderColor     =   49152
      bgColor         =   -2147483643
      fontColor       =   -2147483640
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      multiLine       =   0   'False
   End
   Begin WooSkin.WooTextBox WooTextBoxA 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      text            =   "单行文本框"
      passwordChar    =   ""
      borderSize      =   1
      borderColor     =   15445533
      bgColor         =   -2147483643
      fontColor       =   -2147483640
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      multiLine       =   0   'False
   End
   Begin WooSkin.WooButton WooButton2 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      text            =   "无边框纯色按钮"
      borderWidth     =   3
      shape           =   4
      borderColor     =   15445533
      bgColor         =   15445533
      fontColor       =   16777215
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WooSkin.WooButton WooButton1 
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      text            =   "粗边框圆角按钮"
      borderWidth     =   3
      shape           =   4
      borderColor     =   12632064
      bgColor         =   12648447
      fontColor       =   16384
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WooSkin.WooTextBox WooTextBoxC 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2566
      text            =   ""
      passwordChar    =   ""
      borderSize      =   1
      borderColor     =   16576
      bgColor         =   -2147483624
      fontColor       =   33023
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      multiLine       =   -1  'True
   End
   Begin WooSkin.WooButton WooButton3 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      text            =   "椭圆按钮"
      borderWidth     =   2
      shape           =   2
      borderColor     =   15645627
      bgColor         =   16761024
      fontColor       =   16777215
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WooSkin.WooButton WooButton4 
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      text            =   "细边框矩形按钮"
      borderWidth     =   1
      shape           =   0
      borderColor     =   16711935
      bgColor         =   12640511
      fontColor       =   16512
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'扁平化彩色VB控件，首个版本，包含文本框和按钮控件。
'福建吴世昌原创
'主页: www.fj543.com
'
Private Sub Form_Load()
    WooTextBoxC.Text = "VB6.0彩色扁平化控件。" & vbCrLf & vbCrLf & "由fj543制作。"
End Sub

Private Sub WooButton1_click()
    MsgBox WooTextBoxC.Text
End Sub
 
 
