VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Biodata For Employee"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   10920
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      Height          =   1095
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   8535
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   13215
   End
End
Attribute VB_Name = "Form6"
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
