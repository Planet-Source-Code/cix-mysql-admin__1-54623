VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAdmin 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Softbuddy - MYSQL ADMIN 1.0"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   4
      Left            =   -30000
      ScaleHeight     =   5025
      ScaleWidth      =   4935
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtSQL 
         Height          =   1215
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         ToolTipText     =   "PRESS F9 TO EXECUTE SQL"
         Top             =   240
         Width           =   5895
      End
      Begin MSFlexGridLib.MSFlexGrid P_SHEET3 
         Height          =   3255
         Left            =   0
         TabIndex        =   19
         Top             =   1920
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   14737632
         AllowUserResizing=   1
      End
      Begin VB.Label lblSQL 
         AutoSize        =   -1  'True
         Caption         =   "Select a Database"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   5025
      ScaleWidth      =   4935
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Index           =   2
      Left            =   3000
      ScaleHeight     =   5145
      ScaleWidth      =   5055
      TabIndex        =   8
      Top             =   -24000
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtScript 
         Height          =   4695
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   0
         Width           =   5895
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "&Save to file"
         Height          =   350
         Left            =   4680
         TabIndex        =   16
         Top             =   4720
         Width           =   1215
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5625
      Index           =   1
      Left            =   2880
      ScaleHeight     =   5625
      ScaleWidth      =   5175
      TabIndex        =   7
      Top             =   -20000
      Visible         =   0   'False
      Width           =   5175
      Begin MSFlexGridLib.MSFlexGrid P_Sheet1 
         Height          =   2415
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   14737632
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid P_Sheet2 
         Height          =   2055
         Left            =   0
         TabIndex        =   13
         Top             =   3120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   14737632
         AllowUserResizing=   1
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Select a Database"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Indexes"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   0
      Left            =   -30000
      ScaleHeight     =   5025
      ScaleWidth      =   4935
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   4935
      Begin MSFlexGridLib.MSFlexGrid Sheet 
         Height          =   5175
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   9128
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   14737632
         AllowUserResizing=   1
      End
   End
   Begin MSComctlLib.TabStrip RSTAB 
      Height          =   6615
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11668
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data"
            Key             =   "data"
            Object.Tag             =   "data"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Key             =   "properties"
            Object.Tag             =   "properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Script"
            Key             =   "script"
            Object.Tag             =   "script"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "export"
            Object.Tag             =   "export"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SQL"
            Key             =   "sql"
            Object.Tag             =   "sql"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picramme 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6840
      Left            =   0
      ScaleHeight     =   6840
      ScaleWidth      =   2775
      TabIndex        =   3
      Top             =   1455
      Width           =   2775
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2260
         Picture         =   "Form1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Refresh Databases and tables"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5775
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   10186
         _Version        =   393217
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1930
         Picture         =   "Form1.frx":03C7
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Change Profile"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   350
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1580
         Picture         =   "Form1.frx":04BA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Create New Window"
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   350
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
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
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   0
      Width           =   10485
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
         Left            =   7560
         TabIndex        =   2
         Top             =   0
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   3120
         X2              =   9960
         Y1              =   480
         Y2              =   480
      End
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
         Index           =   0
         Left            =   6240
         TabIndex        =   1
         Top             =   0
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":05AD
         Top             =   -120
         Width           =   6000
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2ADA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AE37
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AEF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AF92
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xNode As Node
Const AllLeft = 2940

Private Sub cmdSQL_Click()
If xNode.Parent Is Nothing Then
 If Not modMYSQL.PrefromSQL(txtSQL, xNode.Text, Me.P_SHEET3, "{null}") Then
     MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
 End If
Else
 If Not modMYSQL.PrefromSQL(txtSQL, xNode.Parent, Me.P_SHEET3, "{null}") Then
     MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
 End If
End If
End Sub

Private Sub Command1_Click()
 
Dim Strname As String

If Not xNode Is Nothing Then Strname = xNode.Text

 modMYSQL.LoadEnv
 Dim I As Integer
 
 For I = 1 To TreeView1.Nodes.Count
  If TreeView1.Nodes(I).Text = Strname Then Set xNode = TreeView1.Nodes(I)
  TreeView1.Nodes(I).Expanded = True
 Next I
 
 
End Sub

Private Sub Command2_Click()
 Load frmLogin
 frmLogin.Show
 Unload Me
End Sub

Private Sub Form_Load()
 picTab(0).ZOrder 0
 picTab(0).Visible = True
End Sub


Private Sub Form_Resize()
 On Error Resume Next
 
 
   TreeView1.Height = Picramme.ScaleHeight - TreeView1.Top
 
   RSTAB.Width = Me.ScaleWidth - RSTAB.Left - 60
   RSTAB.Height = Me.ScaleHeight - RSTAB.Top - 60
 
 Dim I As Integer
 Dim J As Integer
 
 J = picTab.UBound
 
 For I = 0 To J
  picTab(I).Left = AllLeft
  picTab(I).Top = RSTAB.Top + 450
  picTab(I).Height = RSTAB.Height - 520
  picTab(I).Width = RSTAB.Width - 120
 Next I
 
 
  Sheet.Height = picTab(0).ScaleHeight - (Sheet.Top + 60)
  Sheet.Width = picTab(0).ScaleWidth - (Sheet.Left + 60)
  
 Dim H As Integer
 
 H = Int((picTab(0).ScaleHeight - 500) / 2)
 
  
  P_Sheet1.Width = picTab(0).ScaleWidth - (P_Sheet1.Left + 60)
  P_Sheet1.Height = H
 
  
  P_Sheet2.Top = P_Sheet1.Top + P_Sheet1.Height + 300
  P_Sheet2.Width = picTab(0).ScaleWidth - (P_Sheet2.Left + 60)
  P_Sheet2.Height = H
  
  Label2.Top = P_Sheet2.Top - Label2.Height
  
  txtScript.Height = (picTab(0).ScaleHeight - 500)
  txtScript.Width = picTab(0).ScaleWidth
  cmdFile.Top = txtScript.Top + txtScript.Height + 60
  cmdFile.Left = (picTab(0).ScaleWidth) - (cmdFile.Width + 120)
  
  
  txtSQL.Height = H
  txtSQL.Width = picTab(0).ScaleWidth
  P_SHEET3.Top = txtSQL.Top + txtSQL.Height + 300
  P_SHEET3.Height = H
  P_SHEET3.Width = picTab(0).ScaleWidth
 
End Sub

Private Sub Form_Terminate()
 modMYSQL.CloseConnection
 Unload frmLogin
 Unload frmOptions
 Unload Me
 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
 modMYSQL.CloseConnection
 Unload frmLogin
 Unload frmOptions
 Unload Me
 End

End Sub

Private Sub P_Sheet1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = 2 Then Clipboard.SetText P_Sheet1.Text
End Sub

Private Sub P_Sheet2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 67 And Shift = 2 Then Clipboard.SetText P_Sheet2.Text
End Sub

Private Sub P_SHEET3_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 67 And Shift = 2 Then Clipboard.SetText P_SHEET3.Text
End Sub

Private Sub RSTAB_Click()
 Dim I As Integer
 For I = 0 To picTab.UBound
  picTab(I).Visible = False
 Next I
 
 picTab(RSTAB.SelectedItem.Index - 1).ZOrder 0
 picTab(RSTAB.SelectedItem.Index - 1).Visible = True
 
 
End Sub

Private Sub Sheet_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 67 And Shift = 2 Then Clipboard.SetText Sheet.Text
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 67 And Shift = 2 Then Clipboard.SetText xNode.Text
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Set xNode = Node
 On Error Resume Next
 
 If xNode.Parent Is Nothing Then
    
    Me.lblProperties = "Tables"
    Me.lblSQL = "Useing Database " & xNode.Text
    
    If Not modMYSQL.PrefromSQL("SHOW TABLES FROM " & xNode.Text, "", Me.P_Sheet1, "NULL") Then
      MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
    End If
    
    Me.P_Sheet2.Clear
    Me.Sheet.Clear
 
 Else
   If Not modMYSQL.PrefromSQL("SELECT * FROM " & xNode.Text, xNode.Parent, Me.Sheet) Then
      MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
   Else
    Me.lblSQL = "Useing Database " & xNode.Parent
    Me.lblProperties = "Properties"
    If Not modMYSQL.PrefromSQL("DESC " & xNode.Text, xNode.Parent, Me.P_Sheet1, "") Then
      MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
    End If
    
    If Not modMYSQL.PrefromSQL("SHOW INDEX FROM " & xNode.Parent & "." & xNode.Text, "", Me.P_Sheet2) Then
      MsgBox modMYSQL.StringError, vbInformation, "Softbuddy - MYSQL (ADMIN)"
    End If
    
    txtScript = modMYSQL.BuildScript(xNode.Text, P_Sheet1, P_Sheet2)
    
   End If
 
 End If

End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF9 Then
    cmdSQL_Click
 End If
    
End Sub
