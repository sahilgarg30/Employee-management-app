VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Biodata For Manager"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   1215
      Left            =   8520
      TabIndex        =   2
      Top             =   7680
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "go back to records"
      Height          =   1215
      Left            =   1800
      TabIndex        =   1
      Top             =   7680
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
form4.Show
Form3.Hide
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
