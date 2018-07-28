VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Welcome"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Welcome.frx":0000
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   960
      TabIndex        =   2
      Top             =   4200
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visual Bio-Data Generator"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   8415
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form2.Show
End Sub
