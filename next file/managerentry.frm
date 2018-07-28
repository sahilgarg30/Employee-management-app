VERSION 5.00
Begin VB.Form form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Manager Entry"
   ClientHeight    =   9825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "Form4"
   ScaleHeight     =   9825
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8520
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "View as Paragraph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8520
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Go to Last Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Go to First Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5040
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080C0FF&
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      DataField       =   "first name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080C0FF&
      DataField       =   "last name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080C0FF&
      DataField       =   "contact number"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6840
      TabIndex        =   18
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080C0FF&
      DataField       =   "email id"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0080C0FF&
      DataField       =   "address"
      DataSource      =   "Data1"
      Height          =   1575
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H0080C0FF&
      DataField       =   "nationality"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12240
      TabIndex        =   15
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H0080C0FF&
      DataField       =   "language spoken"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12240
      TabIndex        =   14
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H0080C0FF&
      DataField       =   "former school"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12240
      TabIndex        =   13
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H0080C0FF&
      DataField       =   "college"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12240
      TabIndex        =   12
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H0080C0FF&
      DataField       =   "year of grad"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12240
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H0080C0FF&
      DataField       =   "final year percentage"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   12240
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0080C0FF&
      DataField       =   "marital status"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6840
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      DataField       =   "religion"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   12240
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      DataField       =   "achievements"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4680
      Width           =   5295
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFFF&
      DataField       =   "skills"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   5520
      Width           =   5295
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H0080C0FF&
      DataField       =   "experience"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6480
      Width           =   5295
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFFF&
      DataField       =   "objective"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   7320
      Width           =   5295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080C0FF&
      DataField       =   "date"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H0080C0FF&
      DataField       =   "month"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H0080C0FF&
      DataField       =   "year"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H0080C0FF&
      DataField       =   "gender"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\Computer Project\nikita\database\DATABASE1VB.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   9360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Add Information"
      Top             =   4680
      Visible         =   0   'False
      Width           =   4815
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
      Left            =   3120
      TabIndex        =   41
      Top             =   480
      Width           =   1695
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
      Left            =   5880
      TabIndex        =   40
      Top             =   480
      Width           =   1575
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
      Left            =   1200
      TabIndex        =   39
      Top             =   480
      Width           =   1095
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
      Left            =   4080
      TabIndex        =   38
      Top             =   1680
      Width           =   1695
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
      Left            =   4920
      TabIndex        =   37
      Top             =   2400
      Width           =   1815
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
      Left            =   480
      TabIndex        =   36
      Top             =   2400
      Width           =   1335
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
      Left            =   4920
      TabIndex        =   35
      Top             =   3600
      Width           =   1695
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
      Left            =   840
      TabIndex        =   34
      Top             =   1680
      Width           =   1215
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
      Left            =   9480
      TabIndex        =   33
      Top             =   840
      Width           =   1575
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
      Left            =   4920
      TabIndex        =   32
      Top             =   3000
      Width           =   1935
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
      Left            =   9480
      TabIndex        =   31
      Top             =   1800
      Width           =   2295
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
      Left            =   9480
      TabIndex        =   30
      Top             =   2280
      Width           =   1695
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
      Left            =   9480
      TabIndex        =   29
      Top             =   2760
      Width           =   1695
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
      Left            =   9480
      TabIndex        =   28
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Final Year Percentage:                                             %"
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
      Left            =   9480
      TabIndex        =   27
      Top             =   3720
      Width           =   5655
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
      Left            =   9480
      TabIndex        =   26
      Top             =   1320
      Width           =   1455
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
      Left            =   480
      TabIndex        =   25
      Top             =   4680
      Width           =   2655
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
      Left            =   480
      TabIndex        =   24
      Top             =   5640
      Width           =   2055
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
      Left            =   480
      TabIndex        =   23
      Top             =   6600
      Width           =   1815
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
      Left            =   480
      TabIndex        =   22
      Top             =   7440
      Width           =   1815
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command2_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command3_Click()
If Data1.Recordset.BOF Then
Data1.Recordset.MoveNext
Else
Data1.Recordset.MoveFirst
End If
End Sub

Private Sub Command4_Click()
If Data1.Recordset.EOF Then
Data1.Recordset.MovePrevious
Else
Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command5_Click()
Form7.Show
Form7.Label1.Caption = "Hello! I am " & form4.Combo1 & form4.Text1 & " " & form4.Text2 & ". I was born on " & form4.Text4 & "/" & form4.Text18 & "/" & form4.Text19 & ". I am " & form4.Combo3 & ". " & vbNewLine & "I am a " & form4.Text9 & " and am a " & form4.Text3 & ". I speak " & form4.Text10 & ". My former school's name is " & form4.Text11 & ". I passed out of " & form4.Text12 & " in the year " & form4.Text13 & " and secured " & form4.Text14 & "% in the final year. " & vbNewLine & "I have an experience of " & form4.Text16 & vbNewLine & "I live in " & form4.Text7 & ". To contact me you can call at " & form4.Text5 & " or send an e-mail at " & form4.Text6 & "."
End Sub

Private Sub Command6_Click()
Form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form2.Hide
Form2.Text1 = ""
Form2.Text2 = ""
Form2.Combo1 = ""
Form2.Frame1.Visible = False
End Sub
