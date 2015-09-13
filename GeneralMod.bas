Attribute VB_Name = "GeneralMod"
Option Explicit

Public Const de_Field = "||"
Public Const de_Record = "&&&"
Public fso As New FileSystemObject
Public Log As New cLog


Public Function Second2Time(ByVal dblSecond As Double, Optional RemoveLastDigit As Boolean) As String
   On Error GoTo Error
   'edit seng 2005-08-19
  Dim intSecond As Double, lIntSecond As Single
  Dim iniMinute As Single
  Dim strTime As String

  intSecond = GetDecimal(dblSecond)
  lIntSecond = Fix(dblSecond)
  iniMinute = lIntSecond Mod 3600
  strTime = Format(lIntSecond \ 3600, "00") + ":" + Format(iniMinute \ 60, "00") + ":" + Format(iniMinute Mod 60, "00")
  If Not RemoveLastDigit Then strTime = strTime & Format(intSecond - lIntSecond, "#.00")
  Second2Time = strTime
   Exit Function
Error:
   'WriteEvent "Second2Time", Err.Description, Err.Source
End Function

Public Function GetDecimal(ByVal Value As Double) As Double
   On Error GoTo Error
   Dim myString As String
   Dim decimalPosition As Integer
   myString = CStr(Value)
   decimalPosition = InStrRev(myString, ".")
   If (decimalPosition > 0) And (Len(myString) - decimalPosition > 2) Then
      myString = Left(myString, decimalPosition + 2)
   End If
   GetDecimal = CDbl(myString)
   Exit Function
Error:
   'WriteEvent "GetDecimal", Err.Description, Err.Source
   Log.Writelog "GetDecimal", "ERROR", Err.Description, "Value = " & Value
End Function

'''Public Function WriteEvent(lpFunctionName As String, Value As String, Title As String)
'''   On Error Resume Next
'''   Debug.Print "Error =" & lpFunctionName, "Value=" & Value, "Title=" & Title
'''
'''End Function

