VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FF8080&
   Caption         =   "Form6"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Login"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   8040
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "I accept the terms and conditions"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   5760
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Retype-Password"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Password"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Name"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3000
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, File
Private Sub Command1_Click()
If Option1 = True Then
File = "D:\CS Project\usr&pass\" + Text1 + ".txt"
a = Text2
If Text2 = Text3 Then
Open File For Append As #1
Print #1, a
Close #1
MsgBox "Created user account"
Else
MsgBox "Both the Passwords do not match"
End If
Else
MsgBox "You must accept the terms and conditions"
End If
End Sub

Private Sub Command2_Click()
Form6.Visible = False
Form1.Visible = True
End Sub

Private Sub Form_Load()
Text2.PasswordChar = "*"
Text3.PasswordChar = "*"
End Sub
