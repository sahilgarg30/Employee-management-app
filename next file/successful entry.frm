VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Successful Entry"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   4500
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your details have been saved sucessfully."
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   9135
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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

Private Sub Command2_Click()
Form6.Show
Form6.Label1.Caption = "Hello! I am " & Form3.Combo1 & Form3.Text1 & " " & Form3.Text2 & ". I was born on " & Form3.Text4 & "/" & Form3.Text18 & "/" & Form3.Text19 & ". I am " & Form3.Combo3 & ". " & vbNewLine & "I am a " & Form3.Text9 & " and am a " & Form3.Text3 & ". I speak " & Form3.Text10 & ". My former school's name is " & Form3.Text11 & ". I passed out of " & Form3.Text12 & " in the year " & Form3.Text13 & " and secured " & Form3.Text14 & "% in the final year. " & vbNewLine & "I have an experience of " & Form3.Text16 & vbNewLine & "I live in " & Form3.Text7 & ". To contact me you can call at " & Form3.Text5 & " or send an e-mail at " & Form3.Text6 & "."
End Sub
