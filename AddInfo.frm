VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add Information"
   ClientHeight    =   9030
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14775
   LinkTopic       =   "Form3"
   ScaleHeight     =   9030
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Sahil\vb projects\comp project\nikita\database\DATABASE1VB.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Add Information"
      Top             =   4680
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SAVE BIODATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5880
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H0080C0FF&
      DataField       =   "gender"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2160
      TabIndex        =   42
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H0080C0FF&
      DataField       =   "year"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7440
      TabIndex        =   41
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H0080C0FF&
      DataField       =   "month"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6600
      TabIndex        =   40
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080C0FF&
      DataField       =   "date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5760
      TabIndex        =   39
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Create Bio-Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   11760
      Width           =   5295
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFFF&
      DataField       =   "objective"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   7320
      Width           =   5295
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H0080C0FF&
      DataField       =   "experience"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   6480
      Width           =   5295
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFFF&
      DataField       =   "skills"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   5520
      Width           =   5295
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      DataField       =   "achievements"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   4680
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      DataField       =   "religion"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   12120
      TabIndex        =   29
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0080C0FF&
      DataField       =   "marital status"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6720
      TabIndex        =   27
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H0080C0FF&
      DataField       =   "final year percentage"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12120
      TabIndex        =   26
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H0080C0FF&
      DataField       =   "year of grad"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12120
      TabIndex        =   25
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H0080C0FF&
      DataField       =   "college"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12120
      TabIndex        =   24
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H0080C0FF&
      DataField       =   "former school"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12120
      TabIndex        =   23
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H0080C0FF&
      DataField       =   "language spoken"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12120
      TabIndex        =   18
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H0080C0FF&
      DataField       =   "nationality"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12120
      TabIndex        =   16
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0080C0FF&
      DataField       =   "address"
      DataSource      =   "Data1"
      Height          =   1575
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080C0FF&
      DataField       =   "email id"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080C0FF&
      DataField       =   "contact number"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080C0FF&
      DataField       =   "last name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      DataField       =   "first name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080C0FF&
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"AddInfo.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   46
      Top             =   8640
      Width           =   14775
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   44
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "*Please fill in all the details which have been highlighted otherwise your data will not be accepted."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   13200
      Width           =   8055
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Objective:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   35
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Experience:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Skills:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   31
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Achievements:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   28
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Final Year Percentage:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   22
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Year of Graduation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   21
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "College:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   20
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Former School:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Language Spoken:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   17
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nationality:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marital status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of Birth:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_lostfocus()
If Combo1.ListIndex < 0 Then
Combo1.SetFocus
MsgBox "Please Choose a Valid Title", vbCritical, "Title"
Combo1 = ""

End If
End Sub

Private Sub Combo2_lostfocus()
If Combo2.ListIndex < 0 Then
Combo2.SetFocus
MsgBox "Please Choose a Valid Gender", vbCritical, "Gender"
Combo2 = ""

End If
End Sub

Private Sub Combo3_lostfocus()
If Combo3.ListIndex < 0 Then
Combo3.SetFocus
MsgBox "Please Choose a marital status", vbCritical, "Marital Status"
Combo3 = ""
End If
End Sub


Private Sub Command3_Click()
Form3.Hide
Combo1 = ""
Combo2 = ""
Combo3 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text12 = ""
Text13 = ""
Text14 = ""
Text15 = ""
Text16 = ""
Text17 = ""
Text18 = ""
Text19 = ""


End Sub

Private Sub Text10_lostfocus()
For i = 1 To Len(Text10)
If IsNumeric(Mid(Text10, i, 1)) = True Or Text10 = "" Then
Text10.SetFocus
MsgBox "Please enter valid language spoken.", vbCritical, "Language Spoken"
Text10 = ""
End If
Next i
End Sub

Private Sub Text11_lostfocus()

If Text11 = "" Then
Text11.SetFocus
MsgBox "Please enter valid former school.", vbCritical, "Former School"
Text11 = ""
End If
End Sub

Private Sub Text12_lostfocus()
If Text12 = "" Then
Text12.SetFocus
MsgBox "Please enter valid college.", vbCritical, "College"
Text12 = ""
End If
End Sub

Private Sub Text13_lostfocus()

If Val(Text13) < 1990 Or Val(Text13) > 2014 Or IsNumeric(Text13) = False Or Text13 = "" Then
Text13.SetFocus
MsgBox " Please enter valid year of graduation.", vbCritical, "Year Of Graduation"
Text13 = ""
End If
End Sub

Private Sub Text14_lostfocus()
If Val(Text14) < 1 Or Val(Text14) > 100 Or IsNumeric(Text14) = False Or Text14 = "" Then
Text14.SetFocus
MsgBox "Please enter valid final year percentage.", vbCritical, "Final Year Percentage"
Text14 = ""

End If
End Sub

Private Sub Text16_lostfocus()
If Text16 = "" Then
Text16.SetFocus
MsgBox "Please mention your experience.", vbCritical, "Experience"
Text16 = ""

End If
End Sub

Private Sub Text18_lostfocus()

If Val(Text18) > 12 Or IsNumeric(Text18) = False Then
Text18.SetFocus
MsgBox "Please enter correct month for date of birth.", vbCritical, "Date Of Birth"
Text18 = ""
End If

If Val(Text4) > 31 Or IsNumeric(Text4) = False Then
If Val(Text18) = 1 Or Val(Text18) = 3 Or Val(Text18) = 5 Or Val(Text18) = 7 Or Val(Text18) = 8 Or Val(Text18) = 10 Or Val(Text18) = 12 Then
Text4.SetFocus
MsgBox "Please enter correct date for date of birth.", vbCritical, "Date Of Birth"
Text4 = ""
End If
End If

If Val(Text4) > 30 Or IsNumeric(Text4) = False Then
If Val(Text18) = 4 Or Val(Text18) = 6 Or Val(Text18) = 9 Or Val(Text18) = 11 Then
Text4.SetFocus
MsgBox "Please enter correct date for date of birth.", vbCritical, "Date Of Birth"
Text4 = ""
End If
End If

If Val(Text4) > 29 Or IsNumeric(Text4) = False Then
If Val(Text18) = 2 Then
Text4.SetFocus
MsgBox "Please enter correct date for date of birth.", vbCritical, "Date Of Birth"
Text4 = ""
End If
End If
End Sub

Private Sub Text19_lostfocus()
If Val(Text19) < 1970 Or Val(Text19) > 2014 Or IsNumeric(Text19) = False Then
Text19.SetFocus
MsgBox " Please enter valid year for date of birth.", vbCritical, "Year Of Birth"
Text19 = ""
End If
End Sub


Private Sub Text3_lostfocus()
For i = 1 To Len(Text3)
If IsNumeric(Mid(Text3, i, 1)) = True Or Text3 = "" Then
Text3.SetFocus
MsgBox "Please enter valid religion.", vbCritical, "Religion"
Text3 = ""
End If
Next i
End Sub

Private Sub Text4_lostfocus()
If Val(Text4) > 31 Or IsNumeric(Text4) = False Then
Text4.SetFocus
MsgBox "Please enter correct date for date of birth.", vbCritical, "Date Of Birth"
Text4 = ""
End If
End Sub

Private Sub Text5_lostfocus()
If IsNumeric(Text5) = False Or Len(Text5) <> 10 Then
Text5.SetFocus
MsgBox "Please enter correct mobile number.", vbCritical, "Mobile Number"
Text5 = ""
End If
End Sub

Private Sub Text6_lostfocus()
flag = 0
For i = 1 To Len(Text6)
If Mid(Text6, i, 1) = "@" Then
flag = flag + 1
End If
Next i
If flag = 0 Then
Text6.SetFocus
MsgBox "Enter Correct Email Address.", vbCritical, "Email Address"
Text6 = ""
End If
End Sub

Private Sub Text7_lostfocus()
If Text7 = "" Then
Text7.SetFocus
MsgBox "Please mention your address.", vbCritical, "Address"
Text7 = ""
End If

End Sub

Private Sub Text9_lostfocus()
For i = 1 To Len(Text9)
If IsNumeric(Mid(Text9, i, 1)) = True Or Text9 = "" Then
Text9.SetFocus
MsgBox "Please enter valid nationality.", vbCritical, "Nationality"
Text9 = ""
End If
Next i
End Sub
Private Sub Command2_Click()

If Combo1.ListIndex < 0 Then
MsgBox "Please Choose a Valid Title", vbCritical, "Title"
Combo1.SetFocus
Combo1 = ""
Else
k = k + 1
End If

If Combo2.ListIndex < 0 Then
MsgBox "Please Choose a Valid Gender", vbCritical, "Gender"
Combo2.SetFocus
Combo2 = ""
Else
k = k + 1
End If

If Combo3.ListIndex < 0 Then
MsgBox "Please Choose Valid Marital Status", vbCritical, "Marital Status"
Combo3 = ""
Combo3.SetFocus
Else
k = k + 1
End If

If Text1 = "" Or IsNumeric(Text1) = True Then
MsgBox "Please Enter a Valid First Name", vbCritical, "First Name"
Text1.SetFocus
Text1 = ""
Else
k = k + 1
End If

If Text2 = "" Or IsNumeric(Text2) = True Then
MsgBox "Please Enter a Valid Last Name", vbCritical, "Last Name"
Text2.SetFocus
Text2 = ""
Else
k = k + 1
End If

If Val(Text4) > 31 Or IsNumeric(Text4) = False Then
If Val(Text18) = 1 Or Val(Text18) = 3 Or Val(Text18) = 5 Or Val(Text18) = 7 Or Val(Text18) = 8 Or Val(Text18) = 10 Or Val(Text18) = 12 Then
MsgBox "Please Enter a Valid Date", vbCritical, "Date"
Text4 = ""
Text4.SetFocus
Else
k = k + 1
End If
End If
If Val(Text4) > 30 Or IsNumeric(Text4) = False Then
If Val(Text18) = 4 Or Val(Text18) = 6 Or Val(Text18) = 9 Or Val(Text18) = 11 Then
MsgBox "Please Enter a Correct Date", vbCritical, "Date"
Text4 = ""
Text4.SetFocus
Else
k = k + 1
End If
End If
If Val(Text4) > 29 Or IsNumeric(Text4) = False Then
If Val(Text18) = 2 Then
MsgBox "Please Enter a Correct Date", vbCritical, "Date"
Text4 = ""
Text4.SetFocus
Else
k = k + 1
End If
End If


If Val(Text18) > 12 Or IsNumeric(Text18) = False Then
MsgBox "Please Enter a Valid Month", vbCritical, "Month"
Text18.SetFocus
Text18 = ""
Else
k = k + 1
End If

If Val(Text19) < 1970 Or Val(Text19) > 2014 Or IsNumeric(Text19) = False Then
Text19.SetFocus
MsgBox "Please Enter a Valid Year of Birth", vbCritical, "Year of Birth"
Text19 = ""
Else
k = k + 1
End If

If IsNumeric(Text5) = False Or Len(Text5) <> 10 Then
MsgBox "Please Enter a Correct Mobile Number", vbCritical, "Contact Number"
Text5 = ""
Text5.SetFocus
Else
k = k + 1
End If

flag = 0
For i = 1 To Len(Text6)
If Mid(Text6, i, 1) = "@" Then
flag = flag + 1
End If
Next i
If flag = 0 Then
MsgBox "Please Enter a Valid Email ID", vbCritical, "E-mail ID"
Text6 = ""
Text6.SetFocus
Else
k = k + 1
End If


If Text7 = "" Then
MsgBox "Please mention your address.", vbCritical, "Address"
Text7 = ""
Text7.SetFocus
Else
k = k + 1
End If

If IsNumeric(Text9) = True Or Text9 = "" Then
MsgBox "Please Enter a Valid Nationality", vbCritical, "Nationality"
Text9 = ""
Text9.SetFocus
Else
k = k + 1
End If

If IsNumeric(Text3) = True Or Text3 = "" Then
MsgBox "Please Enter a Valid Religion", vbCritical, "Religion"
Text3 = ""
Text3.SetFocus
Else
k = k + 1
End If

If IsNumeric(Text10) = True Or Text10 = "" Then
MsgBox "Please Enter a Valid Language Spoken", vbCritical, "Language Spoken"
Text10 = ""
Text10.SetFocus
Else
k = k + 1
End If

If Text11 = "" Then
MsgBox "Please Enter a Valid Former School", vbCritical, "Former School"
Text11 = ""
Text11.SetFocus
Else
k = k + 1
End If

If Text12 = "" Then
MsgBox "Please Enter a Valid College", vbCritical, "College"
Text12 = ""
Text12.SetFocus
Else
k = k + 1
End If

If Val(Text13) < 1990 Or Val(Text13) > 2014 Or IsNumeric(Text13) = False Or Text13 = "" Then
Text13.SetFocus
MsgBox "Please Enter a Valid Year of Graduation", vbCritical, "Year of Graduation"
Text13 = ""
Else
k = k + 1
End If

If Val(Text14) < 1 Or Val(Text14) > 100 Or IsNumeric(Text14) = False Or Text14 = "" Then
MsgBox "Please Enter a Valid Final Year Percentage", vbCritical, "Final Year Percentage"
Text14 = ""
Text14.SetFocus
Else
k = k + 1
End If

If Text16 = "" Then
MsgBox "Please mention your experience.", vbCritical, "Experience"
Text16 = ""
Text16.SetFocus
Else
k = k + 1
End If


If k = 18 Then
Form5.Show
Data1.Recordset.Update
End If

End Sub

Private Sub Form_Load()
k = 0
Combo1.AddItem "Mr."
Combo1.AddItem "Mrs."
Combo1.AddItem "Dr."
Combo1.AddItem "Dr. Mrs."

Combo2.AddItem "Male"
Combo2.AddItem "Female"

Combo3.AddItem "Married"

Combo3.AddItem "Single"

Combo3.AddItem "Divorced"

Combo3.AddItem "Widowed"


End Sub

Private Sub Text1_change()
a = Len(Text1)
For i = 1 To a
If IsNumeric(Mid(Text1, i, 1)) Then
MsgBox "Please Enter Correct Name", vbCritical, "First Name"
Text1 = ""
Text1.SetFocus
End If
Next i
End Sub

Private Sub Text2_Change()
a = Len(Text2)
For i = 1 To a
If IsNumeric(Mid(Text2, i, 1)) Then
MsgBox "Please Enter Correct Name", vbCritical, "Last Name"
Text2 = ""
Text2.SetFocus
End If
Next i
End Sub

Private Sub Text1_lostfocus()
If Text1 = "" Then
Text1.SetFocus
MsgBox "Please Enter a Valid First Name", vbCritical, "First Name"
Text1 = ""
End If
End Sub


Private Sub Text2_lostfocus()
If Text2 = "" Then
Text2.SetFocus
MsgBox "Please Enter a Valid Last Name", vbCritical, "Last Name"
Text2 = ""
End If
End Sub
