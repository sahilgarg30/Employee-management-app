VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Log In"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10500
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   1080
      TabIndex        =   2
      Top             =   6000
      Width           =   5895
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1920
         TabIndex        =   8
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1560
         Width           =   3735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1920
         TabIndex        =   4
         Text            =   "Select Designation"
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Employee ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pass Key: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Designation:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Profile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "View Profile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "If you are a manager viewing profiles, please click 'View Profile' and enter the required key."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   11
      Top             =   1680
      Width           =   6975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "If you are a fresher, please click on 'Add Profile' to add your details to our database."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   10
      Top             =   600
      Width           =   6975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = True
Combo1.SetFocus
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
If MsgBox("Do You Want to Create a New Profile?", vbQuestion + vbYesNo, "New Profile") = vbYes Then
Form3.Show
If Form3.Data1.Recordset.EOF Then
Else
Form3.Data1.Recordset.MoveLast
Form3.Data1.Recordset.MoveNext
End If
Form3.Combo1.SetFocus
Else
MsgBox "Please Choose to View Profile.", vbInformation, "View profile"
End If
End Sub

Private Sub Command3_Click()
If Combo1 = "General Manager" Then
If Text1 = "gm" And Text2 = "11" Then
form4.Show
Else
form4.Hide
MsgBox "Enter Correct Details", vbCritical, "Wrong Details"
Text1 = ""
Text2 = ""
End If
End If

If Combo1 = "Human Resource Manager" Then
If Text1 = "hrm" And Text2 = "22" Then
form4.Show
Else
form4.Hide
MsgBox "Enter Correct Details", vbCritical, "Wrong Details"
Text1 = ""
Text2 = ""
End If
End If

If Combo1 = "Communication Manager" Then
If Text1 = "cm" And Text2 = "33" Then
form4.Show
Else
form4.Hide
MsgBox "Enter Correct Details", vbCritical, "Wrong Details"
Text1 = ""
Text2 = ""
End If
End If

If Combo1 = "Head of Department" Then
If Text1 = "hod" And Text2 = "44" Then
form4.Show
Else
form4.Hide
MsgBox "Enter Correct Details", vbCritical, "Wrong Details"
Text1 = ""
Text2 = ""
End If
End If

If Combo1 = "Head of Security" Then
If Text1 = "hos" And Text2 = "55" Then
form4.Show
Else
form4.Hide
MsgBox "Enter Correct Details", vbCritical, "Wrong Details"
Text1 = ""
Text2 = ""
End If
End If

If Combo1 = "Executive Manager" Then
If Text1 = "em" And Text2 = "66" Then
form4.Show
Else
form4.Hide
MsgBox "Enter Correct Details", vbCritical, "Wrong Details"
Text1 = ""
Text2 = ""
End If
End If


End Sub

Private Sub Command4_Click()
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

Private Sub Form_Load()
Text1 = ""
Text2 = ""
Combo1 = ""
Frame1.Visible = False
Combo1.AddItem "General Manager"
Combo1.AddItem "Human Resource Manager"
Combo1.AddItem "Communication Manager"
Combo1.AddItem "Head of Department"
Combo1.AddItem "Head of Security"
Combo1.AddItem "Executive Manager"

End Sub

