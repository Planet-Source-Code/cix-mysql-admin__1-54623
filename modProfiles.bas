Attribute VB_Name = "modProfiles"
Option Explicit

Type Profile
 StrName As String
 Username As String
 Password As String
 Host As String
End Type

Public CurrentProfiles() As Profile

Function LoadProfiles(Strfile As String)
 Close #1
 Open Strfile For Binary As #1
  
  Dim Profiles() As Profile
  Dim Counter As Integer
  
  While Not EOF(1)
   DoEvents
   Counter = Counter + 1
   ReDim Preserve Profiles(Counter) As Profile
   Get #1, , Profiles(Counter)
  Wend
  
  CurrentProfiles = Profiles
  
 Close #1
 
 
End Function

Function SaveProfiles(Strfile As String)
 
 
 Close #1
 Kill Strfile
 Open Strfile For Binary As #1
  Dim Counter As Integer
  
  For Counter = 0 To UBound(CurrentProfiles())
   If CurrentProfiles(Counter).Host = "" And CurrentProfiles(Counter).Password = "" And CurrentProfiles(Counter).StrName = "" And CurrentProfiles(Counter).Username = "" Then
   Else
    Put #1, , CurrentProfiles(Counter)
   End If
  Next Counter
  
 Close #1
End Function

Function DeleteProfile(StrName As String)
  Dim Counter As Integer
  
  For Counter = 0 To UBound(CurrentProfiles())
   If CurrentProfiles(Counter).StrName = StrName Then
     With CurrentProfiles(Counter)
      .Host = ""
      .Password = ""
      .StrName = ""
      .Username = ""
     End With
   End If
  Next Counter

End Function

