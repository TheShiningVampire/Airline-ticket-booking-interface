VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF8080&
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   735
      Left            =   8520
      TabIndex        =   3
      Top             =   9480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   6075
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   17775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "This is your ticket which needs to be shown at the ticket counter at the airport"
      Height          =   975
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   9855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Success! You have successfully booked the tickets"
      Height          =   855
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   9015
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim File As Long
Dim strTheData As String


Private Sub Command1_Click()
End
End Sub


Private Sub Form_Load()
Form5.Picture = LoadPicture("D:\CS Project\6.jpg")
File = FreeFile
Open "D:\CS Project\bill.txt" For Input As #File
strTheData = StrConv(InputB(LOF(File), File), vbUnicode)
Close #iFile
Text1 = strTheData
End Sub
