VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FF8080&
   Caption         =   "Form 4"
   ClientHeight    =   3030
   ClientLeft      =   10650
   ClientTop       =   4200
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
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   555
      Left            =   10680
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "YYYY"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   555
      Left            =   8760
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "MM"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Return to Previous Page"
      Height          =   735
      Left            =   10920
      TabIndex        =   12
      Top             =   9360
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proceed"
      Height          =   735
      Left            =   5880
      TabIndex        =   11
      Top             =   9360
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Passenger"
      Height          =   735
      Left            =   5880
      TabIndex        =   10
      Top             =   7440
      Width           =   4695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "I know that if this information is proved to be false, strict action will be taken against me"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   8520
      Width           =   17895
   End
   Begin VB.TextBox Text4 
      Height          =   570
      Left            =   6960
      MaxLength       =   10
      TabIndex        =   7
      Top             =   6240
      Width           =   9255
   End
   Begin VB.TextBox Text3 
      Height          =   570
      Left            =   6960
      MaxLength       =   9
      TabIndex        =   6
      Top             =   5040
      Width           =   9255
   End
   Begin VB.TextBox Text2 
      Height          =   570
      Left            =   6960
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "DD"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   570
      Left            =   6960
      TabIndex        =   2
      Top             =   2640
      Width           =   9255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      TabIndex        =   15
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Mobile Number"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Passport Number"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Date of Birth"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Name (As on Passport)"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Passenger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public l, m, n, o, i, File, day, mon, year, seats, flight, username
Dim ls, con, nop, v, x, y

Private Sub Command1_Click()

l = Text1
m = Val(Text2)
n = Val(Text3)
o = Val(Text4)
x = Val(Text6)
y = Val(Text7)

If con = 0 Then
Open "D:\CS Project\bill.txt" For Output As #1
Print #1, "Origin:", Form3.origin
Print #1, "Departure:", Form3.depat
Print #1, "Date of Departure:", day, mon, year
Print #1, "Seats:", seats
con = 1
nop = seats
v = 1
End If

If seats > 0 Then

If IsNumeric(Text2) = True And IsNumeric(Text6) = True And IsNumeric(Text7) = True And IsNumeric(Text3) = True And IsNumeric(Text4) = True And Text1 <> "" And Text2 <> "" And Text3 <> "" And Text4 <> "" And Text6 <> "" And Text7 <> "" Then
Print #1, "Name:", l
Print #1, "Date of Birth:", m, x, y
Print #1, "Passport Number:", n
Print #1, "Moblie Number:", o

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text6 = ""
Text7 = ""

If v < nop And seats <> 0 Then
v = v + 1
Label6.Caption = v
End If
seats = seats - 1


Else
MsgBox "Data is invalid"
End If

If seats = 0 Then
Command1.Enabled = False
Command2.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Option1.Enabled = True
End If
End If

End Sub

Private Sub Command2_Click()
If Option1.Value = True Then
Print #1, "Total:", i
Close #1
Form4.Hide
Form5.Show
End If
End Sub

Private Sub Command3_Click()
Form4.Hide
Form3.Show
End Sub

Private Sub Form_Load()
Form4.Picture = LoadPicture("D:\CS Project\passport-in-shimoga.jpg")
Option1.Enabled = False
Command2.Enabled = False
Open "D:\CS Project\val.txt" For Input As #1
Input #1, seats, i, day, mon, year
Close #1
If seats = 1 Then
Option1.Enabled = True
Command2.Enabled = True
End If
Label6.Caption = "1"
End Sub

Private Sub Text2_Click()
Text2 = ""
End Sub

Private Sub Text6_Click()
Text6 = ""
End Sub

Private Sub Text7_Click()
Text7 = ""
End Sub
