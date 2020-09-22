Attribute VB_Name = "modMYSQL"
Option Explicit

Public Server As String
Public Username As String
Public Password As String
Public StringError As String

Const Row1 = &HFFFFFF
Const Row2 = &HF0F0F0

Private Mysql_Connection As New ADODB.Connection
Private rs As New ADODB.Recordset
Function CloseConnection()
 Mysql_Connection.Close
End Function

Function MYSQL_Connect() As Boolean
 MYSQL_Connect = True
 
 On Error GoTo Err
 
   If Mysql_Connection.State = adStateOpen Then Mysql_Connection.Close
      
   Mysql_Connection.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Server & ";UID=" & Username & ";PWD=" & Password & ";DATABASE="

  
 Exit Function
  
Err:
 MYSQL_Connect = False
 StringError = "ERROR : " & Err.Number & vbNewLine & Err.Description & vbNewLine
End Function

Function MYSQL_Connect_DB(DB As String) As Boolean
 MYSQL_Connect_DB = True
 
 On Error GoTo Err
 
   If Mysql_Connection.State = adStateOpen Then Mysql_Connection.Close
      
   Mysql_Connection.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Server & ";UID=" & Username & ";PWD=" & Password & ";DATABASE=" & DB & ""

  
 Exit Function
  
Err:
 MYSQL_Connect_DB = False
 StringError = "ERROR : " & Err.Number & vbNewLine & Err.Description & vbNewLine
End Function


Function LoadEnv() As Boolean
LoadEnv = True
 On Error GoTo Err
 
   If Mysql_Connection.State <> adStateOpen Then
    MYSQL_Connect
   End If
      
   With rs
     .Open "SHOW DATABASES", Mysql_Connection, adOpenStatic, adLockReadOnly
   End With
   
   With frmAdmin.TreeView1
    .Nodes.Clear
    Dim xNode As Node
      rs.MoveFirst
      While rs.EOF = False
       Set xNode = .Nodes.Add(, , rs.Fields(0).Value, rs.Fields(0).Value, 4, 4)
       rs.MoveNext
      Wend
      
    rs.Close
      
    Dim I As Integer
    Dim ICount As Integer
    ICount = .Nodes.Count
    For I = 1 To ICount
      rs.Open "SHOW TABLES FROM " & .Nodes(I).Text, Mysql_Connection, adOpenStatic, adLockReadOnly
      If rs.RecordCount <> 0 And Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
         Set xNode = .Nodes.Add(.Nodes(I).Text, tvwChild, rs.Fields(0).Value, rs.Fields(0).Value, 3, 3)
         rs.MoveNext
        Wend
      End If
       If rs.State <> adStateClosed Then
        On Error Resume Next
        rs.Close
       End If
    Next I
      
      
    rs.Close
   End With

  
 Exit Function
  
Err:
 LoadEnv = False
 StringError = "ERROR : " & Err.Number & vbNewLine & Err.Description & vbNewLine
 
End Function


Function PrefromSQL(SQL As String, DB As String, Sheet As MSFlexGrid, Optional strNull As String = "") As Boolean
 PrefromSQL = True

 On Error GoTo Err
 
  Sheet.Visible = False
 
  If DB <> "" Then
    Mysql_Connection.Close
    MYSQL_Connect_DB DB
  Else
   If Mysql_Connection.State <> adStateOpen Then
    MYSQL_Connect
   End If
  
  End If
      
   With rs
     .Open SQL, Mysql_Connection, adOpenStatic, adLockReadOnly
   End With
   
   Dim ColorRow As Integer
   Dim ColCount As Integer
   Dim I As Integer
   
   ColorRow = 0
   
   With Sheet
      .Clear
      ColCount = rs.Fields.Count - 1
      .Cols = ColCount + 1
      .Rows = 2
    
    If rs.State = adStateClosed Then
      MsgBox "Operation done and no result returned", vbInformation, "Softbuddy - MYSQL (ADMIN)"
      Sheet.Visible = True
     Exit Function
    End If
    
      
      For I = 0 To ColCount
       .Row = 0
       .Col = I
       .CellAlignment = 3
       .Text = rs.Fields.Item(I).Name
       .ColWidth(I) = frmAdmin.TextWidth(rs.Fields.Item(I).Name & "DDD")
      Next I
    
      If rs.RecordCount <> 0 And Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
          DoEvents
              
              
                 
          For I = 0 To ColCount
             .Col = I
             .Row = .Rows - 1
             If Not IsNull(rs.Fields(I)) Then
              .Text = rs.Fields(I)
              If frmAdmin.TextWidth(rs.Fields(I) & "DDD") > .ColWidth(I) Then
               .ColWidth(I) = frmAdmin.TextWidth(rs.Fields(I) & "DDD")
              End If
             Else
              .Text = strNull
             End If
             .CellAlignment = 1
             If ColorRow > 0 Then
               .CellBackColor = Row2
             Else
               .CellBackColor = Row1
             End If
    
          Next I
             
             If ColorRow > 0 Then ColorRow = -1
             ColorRow = ColorRow + 1
          rs.MoveNext
         .Rows = .Rows + 1
        Wend
      End If
       
      If rs.State <> adStateClosed Then
       On Error Resume Next
       rs.Close
      End If
      
    rs.Close
   End With

  Sheet.Visible = True
 Exit Function
  
Err:
 Sheet.Visible = True
 PrefromSQL = False
 StringError = "ERROR : " & Err.Number & vbNewLine & Err.Description & vbNewLine


End Function

Function BuildScript(Strname As String, Sheet As MSFlexGrid, Sheet2 As MSFlexGrid) As String
 
  Dim Rows As Integer
  Dim Cols As Integer
  Dim I As Integer, J As Integer
  Dim Buf As String
  Dim Buf1 As String
  Dim Buf2 As String
  Dim Fieldname As String
  
  Rows = Sheet.Rows - 1
  Cols = Sheet.Cols
  
'  Buf = "DROP TABLE IF EXISTS `" & StrName & "`;" & vbNewLine
  Buf = Buf & "CREATE TABLE `" & Strname & "` (" & vbNewLine
  
  For I = 1 To Rows - 1
    DoEvents
     With Sheet
      .Col = 0
      .Row = I
      Fieldname = .Text
      If Buf1 <> "" Then
       Buf1 = Buf1 & Space(3) & ",`" & .Text & "` "
      Else
       Buf1 = Buf1 & Space(3) & "`" & .Text & "` "
      End If
      .Col = 1
      .Row = I
      Buf1 = Buf1 & .Text & " "
      .Col = 2
      .Row = I
      If Trim(.Text) <> "YES" Then Buf1 = Buf1 & "NOT NULL "
      .Col = 3
      .Row = I
      
      If .Text = "PRI" Then
           Buf2 = Buf2 & Space(3) & ",PRIMARY KEY  (`" & Fieldname & "`)" & vbNewLine
      ElseIf .Text = "UNI" Then
           Buf2 = Buf2 & Space(3) & ",UNIQUE KEY `" & FindFieldKeyName(Fieldname, Sheet2, "0") & "` (`" & Fieldname & "`)" & vbNewLine
      End If
      
       
      .Col = 4
      .Row = I
      If .Text <> "" Then
       Buf1 = Buf1 & " default '" & .Text & "' "
      End If
      
      .Col = 5
      .Row = I
      If .Text <> "" Then
       Buf1 = Buf1 & .Text & vbNewLine
      Else
       Buf1 = Buf1 & vbNewLine
      End If
     
      Debug.Print Buf & Buf1
     
   End With
  Next I
  
 BuildScript = Buf & Buf1 & Buf2 & ")" & vbNewLine
 
End Function

Function FindFieldKeyName(Strname As String, Sheet As MSFlexGrid, str_UniQue As String) As String
  Dim Rows As Integer
  Dim Cols As Integer
  Dim I As Integer, J As Integer
  
  Dim Buf1 As String, Buf2 As String, Buf3 As String
  
  Rows = Sheet.Rows - 1
  
  For I = 1 To Rows
    DoEvents
     With Sheet
      .Col = 4
      .Row = I
      Buf1 = .Text
      .Col = 1
      .Row = I
      Buf2 = .Text
      .Col = 2
      .Row = I
      Buf3 = .Text
     End With
     
   If Strname = Buf1 And Buf2 = str_UniQue Then
     FindFieldKeyName = Buf3
     Exit Function
   End If
 Next I
End Function
