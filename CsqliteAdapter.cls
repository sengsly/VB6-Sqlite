VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CsqliteAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim sqlConnection As New cConnection
Const rsaKey = "kimseng"

Public Function OpenDB(lpFileName As String) As Boolean
   OpenDB = sqlConnection.OpenDB(lpFileName, rsaKey)
End Function

Public Function selectRecord() As String

End Function
