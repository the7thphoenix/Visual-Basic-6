VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Calculator"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleMode       =   0  'User
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnAC 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   100
      TabIndex        =   17
      Top             =   1900
      Width           =   2200
   End
   Begin VB.CommandButton Btn2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1250
      TabIndex        =   16
      Top             =   4905
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Vazir"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   100
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   100
      Width           =   4500
   End
   Begin VB.CommandButton BtnDivide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3550
      TabIndex        =   14
      Top             =   1900
      Width           =   1050
   End
   Begin VB.CommandButton Btnx 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2400
      TabIndex        =   13
      Top             =   1900
      Width           =   1050
   End
   Begin VB.CommandButton BtnNegative 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3550
      TabIndex        =   12
      Top             =   3900
      Width           =   1050
   End
   Begin VB.CommandButton BtnPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3550
      TabIndex        =   11
      Top             =   2900
      Width           =   1050
   End
   Begin VB.CommandButton BtnP 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2400
      TabIndex        =   10
      Top             =   5900
      Width           =   1050
   End
   Begin VB.CommandButton Btn0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   100
      TabIndex        =   9
      Top             =   5900
      Width           =   2200
   End
   Begin VB.CommandButton Btn1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   100
      TabIndex        =   8
      Top             =   4900
      Width           =   1050
   End
   Begin VB.CommandButton Btn3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2400
      TabIndex        =   7
      Top             =   4900
      Width           =   1050
   End
   Begin VB.CommandButton Btn4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   100
      TabIndex        =   6
      Top             =   3900
      Width           =   1050
   End
   Begin VB.CommandButton Btn5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1250
      TabIndex        =   5
      Top             =   3900
      Width           =   1050
   End
   Begin VB.CommandButton Btn6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2400
      TabIndex        =   4
      Top             =   3900
      Width           =   1050
   End
   Begin VB.CommandButton Btn9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2400
      TabIndex        =   3
      Top             =   2900
      Width           =   1050
   End
   Begin VB.CommandButton Btn8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1250
      TabIndex        =   2
      Top             =   2900
      Width           =   1050
   End
   Begin VB.CommandButton Btn7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   100
      TabIndex        =   1
      Top             =   2900
      Width           =   1050
   End
   Begin VB.CommandButton BtnANS 
      Caption         =   " = "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   3550
      TabIndex        =   0
      Top             =   4900
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Dim sign As String

Private Sub Btn0_Click()
Text1.Text = Text1.Text & 0
'Label1.Caption = Label1.Caption & 0
End Sub

Private Sub Btn1_Click()
Text1.Text = Text1.Text & 1
End Sub

Private Sub Btn2_Click()
Text1.Text = Text1.Text & 2
End Sub

Private Sub Btn3_Click()
Text1.Text = Text1.Text & 3
End Sub

Private Sub Btn4_Click()
Text1.Text = Text1.Text & 4
End Sub

Private Sub Btn5_Click()
Text1.Text = Text1.Text & 5
End Sub

Private Sub Btn6_Click()
Text1.Text = Text1.Text & 6
End Sub

Private Sub Btn7_Click()
Text1.Text = Text1.Text & 7
End Sub

Private Sub Btn8_Click()
Text1.Text = Text1.Text & 8
End Sub

Private Sub Btn9_Click()
Text1.Text = Text1.Text & 9
End Sub

Private Sub BtnAC_Click()
Text1.Text = ""
End Sub

Private Sub BtnANS_Click()
b = Text1.Text
If sign = "+" Then
Text1.Text = a + b
Else
If sign = "-" Then
Text1.Text = a - b
Else
If sign = "*" Then
Text1.Text = a * b
Else
If sign = "/" Then
Text1.Text = a / b
End If
End If
End If
End If
End Sub

Private Sub BtnDivide_Click()
a = Text1.Text
sign = "/"
Text1.Text = ""
End Sub

Private Sub BtnNegative_Click()
a = Text1.Text
sign = "-"
Text1.Text = ""
End Sub

Private Sub BtnP_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub BtnPlus_Click()
a = Text1.Text
sign = "+"
Text1.Text = ""
End Sub

Private Sub Btnx_Click()
a = Text1.Text
sign = "*"
Text1.Text = ""
End Sub

