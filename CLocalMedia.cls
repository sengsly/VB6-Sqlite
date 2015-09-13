VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLocalMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================API Declaration=====================================
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

'================Variable Declaration================================
Private Const DB_Name As String = "usermedia.sqlite"
Private Const MediaExt As String = ";mp3;wma;vob;dat;ifo;*tp;trp;ts;wmv;asf;avi;mpg;mov;mp4;m2ts;evo;"
Private mConnection As New cConnection



Public Function Initialize() As Boolean
   If mConnection.OpenDB(App.Path & "\" & DB_Name) Then
      Initialize = True
   End If
End Function

Public Function addFolder(lpFolderPath As String) As Boolean
   Dim cRec As cRecordset
   If fso.FolderExists(lpFolderPath) Then
      Set cRec = mConnection.OpenRecordset("SELECT id from MEDIAFOLDER Where folders ='" & lpFolderPath & "'")
      If cRec.EOF Then
         If Me.ScanMedias(lpFolderPath) Then
            mConnection.Execute "INSERT INTO  MEDIAFOLDER(folders, language) VALUES('" & lpFolderPath & "','" & "English" & "');"
            addFolder = True
         End If
      End If
   End If
End Function
Public Function ScanMedias(lpFolder As String) As Boolean

   Dim ty As WIN32_FIND_DATA
   Dim iLen As Long
   Dim lpFileName As String
   Dim lpFullPathName As String
   Dim hHandle As Long
   Dim bResult As Long
   Dim lpScanFolder As String
   
   lpScanFolder = fso.BuildPath(lpFolder, "\*.*")
   hHandle = FindFirstFile(StrPtr(lpScanFolder), ty)
   If hHandle = INVALID_HANDLE_VALUE Then Exit Function
   Do
      lpFileName = StrConv(ty.cFileName, vbFromUnicode)
      iLen = lstrlen(StrPtr(lpFileName))
      If InStr(MediaExt, ";" & fso.GetExtensionName(lpFileName) & ";") > 0 Then
         lpFullPathName = fso.BuildPath(lpFolder, Left(lpFileName, iLen))
         mConnection.Execute "INSERT INTO  Medias(TrackPath, TrackName) VALUES('" & lpFullPathName & "','" & lpFileName & "');"
         Debug.Print fso.BuildPath(lpFolder, Left(lpFileName, iLen))
         
      End If
      bResult = FindNextFile(hHandle, ty)
   Loop While (bResult <> 0)
   ScanMedias = True
   
End Function

Public Function getMedias() As String           ' Return xml string
   Dim i As Long
   Dim cRecord As cRecordset
   Dim xmlElement As IXMLDOMElement
   Dim xmlDoc As New DOMDocument
   Dim xmlRoot As MSXML2.IXMLDOMElement
   
   Set xmlRoot = xmlDoc.createElement("Medias")
   Set cRecord = mConnection.OpenRecordset("SELECT id, trackname FROM Medias")
   xmlDoc.appendChild xmlRoot
   While Not cRecord.EOF
      Debug.Print cRecord.Fields("id"), cRecord.Fields("trackname")
      Set xmlElement = xmlDoc.createElement("Tracks")
      xmlElement.setAttribute "id", cRecord.Fields("id")
      xmlElement.setAttribute "trackname", cRecord.Fields("trackname")
      xmlRoot.appendChild xmlElement
      cRecord.MoveNext
   Wend
   getMedias = xmlDoc.XML

End Function
