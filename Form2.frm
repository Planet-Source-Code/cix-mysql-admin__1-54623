VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Softbuddy - MYSQL LOGIN 1.0"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Profile"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "Form2.frx":030A
      ScaleHeight     =   1545
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6270
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MYSQL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009E6441&
         Height          =   435
         Index           =   2
         Left            =   3600
         TabIndex        =   10
         Top             =   120
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   7320
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   435
         Index           =   1
         Left            =   4920
         TabIndex        =   9
         Top             =   120
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Profiles"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Profile"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Host"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Profiles() As Profile
Private Sub cmdCancel_Click()
 End
End Sub

Private Sub cmdLogin_Click()
 
 modMYSQL.Server = txtHost
 modMYSQL.Username = txtUsername
 modMYSQL.Password = txtPassword
 
  If modMYSQL.MYSQL_Connect Then
    If modMYSQL.LoadEnv Then
      frmAdmin.Show
      Me.Hide
    Else
      MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
    End If
  Else
   MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
  End If

End Sub

Private Sub Combo1_Click()

 Dim Counter As Integer
  
  For Counter = 0 To UBound(Profiles())
   If Profiles(Counter).StrName = Combo1.Text Then
     With Profiles(Counter)
      txtUsername.Text = .Username
      txtPassword.Text = .Password
      txtHost = .Host
     End With
   End If
  Next Counter

End Sub

Private Sub Command1_Click()
 frmOptions.Show
End Sub

Private Sub Command2_Click()
 modProfiles.DeleteProfile Combo1.Text
 modProfiles.SaveProfiles App.Path & "\profiles.pro"
 Form_Load
End Sub

Private Sub Form_Load()
 
 modProfiles.LoadProfiles App.Path & "\profiles.pro"
 Profiles = modProfiles.CurrentProfiles
 
 Dim Counter As Integer
  Combo1.Clear
  For Counter = 0 To UBound(Profiles())
   If Profiles(Counter).StrName <> "" Then Combo1.AddItem Profiles(Counter).StrName
  Next Counter


End Sub

