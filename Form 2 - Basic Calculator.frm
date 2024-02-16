VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   250
      TabIndex        =   5
      Top             =   1695
      Width           =   4065
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   2500
      TabIndex        =   4
      Top             =   1000
      Width           =   1750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3075
      TabIndex        =   3
      Top             =   2100
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sub"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1650
      TabIndex        =   2
      Top             =   2100
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   250
      TabIndex        =   1
      Top             =   2100
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   2500
      TabIndex        =   0
      Top             =   300
      Width           =   1750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    Number 2    "
      BeginProperty Font 
         Name            =   "XB Titre"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   270
      TabIndex        =   7
      Top             =   1000
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    Number 1    "
      BeginProperty Font 
         Name            =   "XB Titre"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   270
      TabIndex        =   6
      Top             =   300
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = Str(Val(Text1.Text) + Val(Text2.Text))
End Sub

Private Sub Command2_Click()
Text3.Text = Str(Val(Text1.Text) - Val(Text2.Text))
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Form1.Caption = "Simple Calculator"
Text1.Text = ""
Text2.Text = ""

End Sub

