VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
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
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   6
      Top             =   6720
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   5
      Top             =   5760
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   5880
      TabIndex        =   4
      Top             =   4800
      Width           =   7815
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   5880
      TabIndex        =   3
      Top             =   3840
      Width           =   7815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Login to Your Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   2400
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a, username
Private Sub Command1_Click()
username = Text1
File = "D:\CS Project\usr&pass\" + Text1 + ".txt"
If Text1 <> "" And Text2 <> "" Then
Open File For Input As #1
Input #1, a
If Text2 = a And Dir(File) <> "" Then
username = Text1
Form1.Visible = False
Form3.Visible = True
Else
MsgBox "Incorrect Username or Password"
End If
Close #1
Else
MsgBox "Enter Username and Password"
End If
End Sub

Private Sub Command2_Click()
Form1.Visible = False
Form6.Visible = True
End Sub

Private Sub Form_Load()
Form1.Picture = LoadPicture("D:\CS Project\world-tourism-day.jpg")
Text2.PasswordChar = "*"
End Sub

