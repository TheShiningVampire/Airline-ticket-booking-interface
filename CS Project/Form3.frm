VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF8080&
   Caption         =   "Form3"
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
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   10080
      MaxLength       =   2
      TabIndex        =   25
      Text            =   "MM"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   720
      Left            =   10920
      MaxLength       =   4
      TabIndex        =   24
      Text            =   "YYYY"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Proceed"
      Height          =   510
      Left            =   11280
      TabIndex        =   23
      Top             =   10200
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Height          =   630
      Left            =   5040
      TabIndex        =   22
      Top             =   9960
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   15480
      TabIndex        =   20
      Top             =   8880
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5520
      TabIndex        =   18
      Top             =   7800
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5040
      TabIndex        =   17
      Top             =   8880
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   12120
      ScaleHeight     =   4635
      ScaleWidth      =   7515
      TabIndex        =   15
      Top             =   3000
      Width           =   7575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   14
      Text            =   "0"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   13
      Text            =   "0"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   7200
      Max             =   15
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   7200
      Max             =   15
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   9240
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "DD"
      Top             =   2040
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   13080
      TabIndex        =   5
      Text            =   "Select Destination"
      Top             =   2160
      Width           =   6375
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Text            =   "Select Origin"
      Top             =   2160
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "SEARCH FOR AIRLINE AND COMPARE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   5520
      Width           =   5535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5040
      TabIndex        =   2
      Top             =   6960
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      Caption         =   "Your Payment :"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   10080
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      Caption         =   "Amount per head Child (Rs.)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   19
      Top             =   9000
      Width           =   4575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Amount per head Adult (Rs.)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   9000
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Number of Children"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Number of Adults"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Number of Passenger:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Departure Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Name of Airlines"
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
      Left            =   360
      TabIndex        =   1
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Journey Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a, b, c, d, e, f, g, h, i, k, adults, children, seats, depat, origin, ct, flight

Private Sub Combo1_Change()
Command2.BackColor = RGB(0, 250, 0)
End Sub

Private Sub Command1_Click()
If ct = 0 Then
Open "D:\CS Project\val.txt" For Output As #1
ct = ct + 1
End If
If Combo2.Text <> "Select Origin" And Combo3.Text <> "Select Destination" And seats <> 0 Then
Combo1.AddItem "Indigo Airlines"
Combo1.AddItem "Jet Airways"
Combo1.AddItem "Air India"
Combo1.AddItem "Spice Jet"
Combo1.AddItem "Go Air"
Combo1.AddItem "Etihad Airways"
Combo1.AddItem "Jetlite"
Combo1.AddItem "Air Asia"
Combo1.AddItem "Vistara"
Combo1.AddItem "Air Costa"
Combo1.AddItem "Air Pegasus"
If Combo3.Text = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Picture1.Picture = LoadPicture("D:\CS Project\1030.jpg")
End If
If Combo3.Text = "PUNE, Pune Airport" Then
Picture1.Picture = LoadPicture("D:\CS Project\shaniwar-wada-pune.jpg")
End If
If Combo3.Text = "NAGPUR, Dr. Babasaheb International Airport" Then
Picture1.Picture = LoadPicture("D:\CS Project\download.jpg")
End If
If Combo3.Text = "BENGALURU, Kempegowda International Airport" Then
Picture1.Picture = LoadPicture("D:\CS Project\image 3.jpg")
End If
Text4 = ""
Text5 = ""

d = Val(Text2)
e = Val(Text3)
k = d + e
h = Text1

Else
MsgBox "Invalid Data"
End If
End Sub

Private Sub Command2_Click()

a = Combo2.Text
b = Combo3.Text
c = Combo1.Text
flight = Combo1.Text

If c = "Indigo Airlines" Then

If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8740"
Text5 = "6740"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9680"
Text5 = "8880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12635"
Text5 = "10635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8740"
Text5 = "6740"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7690"
Text5 = "6880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13635"
Text5 = "12635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9635"
Text5 = "8635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8740"
Text5 = "6740"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9940"
Text5 = "8740"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "12635"
Text5 = "10635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10500"
Text5 = "8500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9690"
Text5 = "8880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Jet Airways" Then

If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8940"
Text5 = "6940"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9780"
Text5 = "8980"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "14435"
Text5 = "11635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8340"
Text5 = "6740"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7390"
Text5 = "6280"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12635"
Text5 = "11635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9935"
Text5 = "8935"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8940"
Text5 = "6940"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "11040"
Text5 = "8740"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "14635"
Text5 = "12635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "11500"
Text5 = "9500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9990"
Text5 = "8980"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Air India" Then

If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9580"
Text5 = "8580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "11635"
Text5 = "9635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7390"
Text5 = "6380"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12635"
Text5 = "11635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9435"
Text5 = "8435"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8340"
Text5 = "6340"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9840"
Text5 = "8640"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "11635"
Text5 = "9635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "11500"
Text5 = "8600"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9390"
Text5 = "8580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Spice Jet" Then
If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8640"
Text5 = "6640"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9580"
Text5 = "8780"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12535"
Text5 = "10635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8640"
Text5 = "6440"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7590"
Text5 = "6780"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13435"
Text5 = "12435"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9535"
Text5 = "8535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8640"
Text5 = "6640"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9740"
Text5 = "8640"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "12135"
Text5 = "10135"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10100"
Text5 = "8100"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9590"
Text5 = "8780"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Go Air" Then

If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8940"
Text5 = "6940"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9780"
Text5 = "8780"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12035"
Text5 = "10035"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7590"
Text5 = "6680"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13235"
Text5 = "12235"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "8235"
Text5 = "7635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "7740"
Text5 = "5740"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9140"
Text5 = "8140"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "11635"
Text5 = "9635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10500"
Text5 = "7500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "8690"
Text5 = "8280"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Etihad Airways" Then
If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "9040"
Text5 = "7040"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "10080"
Text5 = "9880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13635"
Text5 = "11635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8940"
Text5 = "6940"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7990"
Text5 = "6980"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "14635"
Text5 = "13935"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "10035"
Text5 = "9035"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "9540"
Text5 = "7540"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "1140"
Text5 = "9640"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "15635"
Text5 = "13635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "12500"
Text5 = "10500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "11690"
Text5 = "10880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Jetlite" Then
If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9580"
Text5 = "8580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12535"
Text5 = "10535"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7590"
Text5 = "6580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13535"
Text5 = "12535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9535"
Text5 = "8535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9540"
Text5 = "8540"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "12535"
Text5 = "10535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10500"
Text5 = "8500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9590"
Text5 = "8580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Air Asia" Then
If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9580"
Text5 = "8580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12535"
Text5 = "10535"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7590"
Text5 = "6580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13535"
Text5 = "12535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9535"
Text5 = "8535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8540"
Text5 = "6540"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9540"
Text5 = "8540"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "12535"
Text5 = "10535"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10500"
Text5 = "8500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9590"
Text5 = "8580"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Vistara" Then
If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "7740"
Text5 = "5740"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "8680"
Text5 = "7880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "11635"
Text5 = "9635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "7740"
Text5 = "5740"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "6690"
Text5 = "5880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12635"
Text5 = "11635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "8635"
Text5 = "7635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "7740"
Text5 = "5740"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8940"
Text5 = "7740"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "11635"
Text5 = "9635"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "9500"
Text5 = "7500"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "8690"
Text5 = "7880"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Air Costa" Then

If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8745"
Text5 = "6745"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9685"
Text5 = "8885"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12630"
Text5 = "10630"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8745"
Text5 = "6745"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7695"
Text5 = "6885"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13630"
Text5 = "12630"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9630"
Text5 = "8630"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8745"
Text5 = "6745"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9945"
Text5 = "8745"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "12630"
Text5 = "10630"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10505"
Text5 = "8505"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9695"
Text5 = "8885"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If

If c = "Air Pegasus" Then

If a = "MUMBAI, Chhatrapati Shivaji International Airport" Then
If b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8750"
Text5 = "6750"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9650"
Text5 = "8850"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "12655"
Text5 = "10635"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "NAGPUR, Dr. Babasaheb International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "8750"
Text5 = "6750"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "7650"
Text5 = "6850"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "13655"
Text5 = "12655"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "PUNE, Pune Airport" Then
If b = "BENGALURU, Kempegowda International Airport" Then
Text4 = "9655"
Text5 = "8655"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "8750"
Text5 = "6750"
ElseIf b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "9950"
Text5 = "8750"
ElseIf b = "PUNE, Pune Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

If a = "BENGALURU, Kempegowda International Airport" Then
If b = "MUMBAI, Chhatrapati Shivaji International Airport" Then
Text4 = "12655"
Text5 = "10655"
ElseIf b = "NAGPUR, Dr. Babasaheb International Airport" Then
Text4 = "10550"
Text5 = "8550"
ElseIf b = "PUNE, Pune Airport" Then
Text4 = "9650"
Text5 = "8850"
ElseIf b = "BENGALURU, Kempegowda International Airport" Then
MsgBox "Place of Origin and Destination should be different"
End If
End If

End If
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True

f = Val(Text4)
g = Val(Text5)
i = (d * f) + (e * g)
Text6 = i

origin = Combo2.Text
depat = Combo3.Text

Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False

End Sub

Private Sub Command3_Click()
Print #1, seats, i, Text1, Text7, Text8
Close #1
Form3.Hide
Form4.Show
End Sub

Private Sub Form_Load()
Form3.Picture = LoadPicture("D:\CS Project\northindia-state.jpg")
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
ct = 0
Combo2.AddItem "MUMBAI, Chhatrapati Shivaji International Airport"
Combo2.AddItem "NAGPUR, Dr. Babasaheb International Airport"
Combo2.AddItem "PUNE, Pune Airport"
Combo2.AddItem "BENGALURU, Kempegowda International Airport"
Combo3.AddItem "MUMBAI, Chhatrapati Shivaji International Airport"
Combo3.AddItem "NAGPUR, Dr. Babasaheb International Airport"
Combo3.AddItem "PUNE, Pune Airport"
Combo3.AddItem "BENGALURU, Kempegowda International Airport"
Command1.BackColor = RGB(0, 250, 0)
End Sub

Private Sub HScroll1_Change()
Text2 = HScroll1.Value
adults = Val(Text2)
seats = adults + children
End Sub

Private Sub HScroll2_Change()
Text3 = HScroll2.Value
children = Val(Text3)
seats = adults + children
End Sub

Private Sub Text1_Click()
Text1 = ""
End Sub

Private Sub Text7_Click()
Text7 = ""
End Sub

Private Sub Text8_Click()
Text8 = ""
End Sub
