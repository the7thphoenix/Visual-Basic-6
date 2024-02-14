VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "XB Titre"
      Size            =   18
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Cbtn 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3500
      TabIndex        =   3
      Top             =   300
      Width           =   699
   End
   Begin VB.TextBox Abtn 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   2
      Top             =   300
      Width           =   699
   End
   Begin VB.TextBox Bbtn 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2000
      TabIndex        =   1
      Top             =   300
      Width           =   699
   End
   Begin VB.CommandButton Command 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "XB Titre"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1700
      TabIndex        =   0
      Top             =   1110
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1065
      TabIndex        =   11
      Top             =   3765
      Width           =   2850
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "X2="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   180
      TabIndex        =   10
      Top             =   3015
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "X1="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   165
      TabIndex        =   9
      Top             =   2250
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "c="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2925
      TabIndex        =   8
      Top             =   330
      Width           =   420
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "b="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   1410
      TabIndex        =   7
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "a="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   45
      TabIndex        =   6
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1000
      TabIndex        =   5
      Top             =   2940
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1000
      TabIndex        =   4
      Top             =   2085
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()

a = Val(Abtn.Text)
  b = Val(Bbtn.Text)
c = Val(Cbtn.Text)

delta = b ^ 2 - (4 * a * c)

If delta > 0 Then
 X1 = (-b + Sqr(delta)) / (2 * a)
 X2 = (-b - Sqr(delta)) / (2 * a)

If delta < 0 Then Label3.Caption = "No Root...": GoTo 10

If delta = 0 Then
  X1 = -b / (2 * a)
  X2 = X1: GoTo 5
End If
End If


5 Label1.Caption = X1: Label2.Caption = X2

10 Rem end

End Sub
