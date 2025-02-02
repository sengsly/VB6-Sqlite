VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'=====================Local Enumeration
Public Enum AlbumEnum
   NotAvailable = 0
   Available = 1
   All = 3
End Enum
Public Enum CDTypeEnum
   Clean = 2
   Explicit = 3
   AllCD = 0
End Enum
Public Enum SortEnum
   stDate = 0
   stName = 1
   stNone = 3
End Enum


'=====================Local Constant
Const mRsaKey = "ymguide"

'=====================Local variable
Dim mConnection As cConnection
Dim mCountry As String
Dim mLanguage As String
Dim mCountryPath As String
Dim mOpenSuccess As Boolean

Dim var_Album() As Variant

'======================Browsing album
Public selectAlbumType As AlbumEnum



Private Sub Class_Initialize()
   Set mConnection = Nothing
End Sub


Public Function Initialize(lpFileName As String, Optional aDatabaseType As Integer = 0) As Boolean
   On Error GoTo Error:
   Set mConnection = New cConnection
   'If mConnection.OpenDB(lpFileName, mRsaKey) Then
   If mConnection.OpenDB(lpFileName) Then
      mOpenSuccess = True
   Else
      mOpenSuccess = False
   End If
   Initialize = mOpenSuccess
   Exit Function
Error:
   Log.Writelog "Initialize", "ERROR", Err.Description, "lpFileName = " & lpFileName & ", aDatabaseType=" & aDatabaseType
End Function



Public Function GetTracksByAlbum(ByVal albumID As Long, boolShowArtist As Boolean) As String
   'On Error GoTo Error
   Static bQuerying As Boolean
   Dim tmp As String
   Dim Data As Variant
   Dim stringShowArtist As String
   Dim showArtistUnicode As String
   Dim sTrackLength As String
   Dim iAvailable As Integer
   
   
'''   SELECT     dbo.sqlite_track.TrackId, dbo.sqlite_track.TrackPosition, dbo.sqlite_track.Name, dbo.sqlite_track.Unicode, dbo.sqlite_track.Rating, dbo.sqlite_track.TrackLength, dbo.sqlite_track.Language,
'''                      dbo.sqlite_track.SingType, dbo.sqlite_label.Name AS Expr1, dbo.sqlite_label.Unicode AS Expr2
'''FROM         dbo.sqlite_label INNER JOIN
'''                      dbo.sqlite_movie ON dbo.sqlite_label.LabelId = dbo.sqlite_movie.LabelId INNER JOIN
'''                      dbo.sqlite_title ON dbo.sqlite_movie.TitleID = dbo.sqlite_title.TitleID INNER JOIN
'''                      dbo.sqlite_track ON dbo.sqlite_title.TrackId = dbo.sqlite_track.TrackId
   
   Dim sqlTracks As cRecordset
   Dim sqlArtist As cRecordset
   Dim sql As String
'''   sql = "SELECT Track.ID, Track.Position, Track.Name, Track.Unicode, Track.Rating, Track.Length, Track.Language, " & _
'''         "Track.SingType, Label.Name AS LabelName, Label.Unicode as LabelUnicode, Actor.Name as ActorName, Actor.Unicode as ActorUnicode, " & _
'''         "Movie.Path As MoviePath, Movie.PathSmall as MoviePathSmall, Movie.PathMedium As MoviePathMedium " & _
'''         "FROM Label " & _
'''         "INNER JOIN Movie ON Label.ID = Movie.LabelID " & _
'''         "INNER JOIN Title ON Movie.MainTitle = Title.ID " & _
'''         "INNER JOIN Track ON Title.TrackID = Track.ID " & _
'''         "LEFT JOIN TrackArtist ON Track.ID = TrackArtist.TrackID " & _
'''         "INNER JOIN Actor ON TrackArtist.ArtistID= Actor.ID " & _
'''         "WHERE Movie.ID=" & albumID & " " & _
'''         "ORDER by Track.Position,Track.ID"
   
   sql = "SELECT Track.ID, Track.Position, Track.Name, Track.Unicode, Track.Rating, Track.Length, Track.Language, " & _
         "Track.SingType, Track.Available, Track.SeekOffset, Label.Name AS LabelName, Label.Unicode as LabelUnicode, " & _
         "Movie.Path As MoviePath, Movie.PathSmall as MoviePathSmall, Movie.PathMedium As MoviePathMedium " & _
         "FROM Label " & _
         "INNER JOIN Movie ON Label.ID = Movie.LabelID " & _
         "INNER JOIN Title ON Movie.MainTitle = Title.ID " & _
         "INNER JOIN Track ON Title.TrackID = Track.ID " & _
         "WHERE Movie.ID=" & albumID & " " & _
         "ORDER by Track.Position"
   
   
   If bQuerying = False Then
      bQuerying = True
      
      Set sqlTracks = mConnection.OpenRecordset(sql)
      While Not sqlTracks.EOF
         With sqlTracks
            stringShowArtist = vbNullString
            showArtistUnicode = vbNullString
            If boolShowArtist Then
               Set sqlArtist = mConnection.OpenRecordset("Select Actor.Name, Actor.Unicode from Actor INNER JOIN TrackArtist ON Actor.ID=TrackArtist.ArtistID Where TrackArtist.TrackID=" & .Fields("ID"))
               While Not sqlArtist.EOF
                  If stringShowArtist = "" Then stringShowArtist = sqlArtist.Fields("Name") Else stringShowArtist = stringShowArtist & "+" & sqlArtist.Fields("Name")
                  If showArtistUnicode = "" Then showArtistUnicode = sqlArtist.Fields("Unicode") Else showArtistUnicode = showArtistUnicode & "+" & sqlArtist.Fields("Unicode")
                  sqlArtist.MoveNext
               Wend
            Else
               stringShowArtist = .Fields("LabelName")
               showArtistUnicode = .Fields("LabelUnicode")
            End If
            sTrackLength = Second2Time(.Fields("Length"), True)
            iAvailable = (.Fields("MoviePath")) <> "" Or (.Fields("MoviePathSmall") <> "") Or (.Fields("MoviePathMedium") <> "") Or (.Fields("Available") = 1)
            Data = Array(.Fields("ID"), .Fields("Position"), .Fields("Name"), .Fields("Unicode"), "Mood", "None", "DanceStyle", "Track Artist", stringShowArtist, showArtistUnicode, .Fields("Rating"), sTrackLength, .Fields("Language"), iAvailable, 0, "", .Fields("SingType"), .Fields("SeekOffset"), .Fields("LabelName"), .Fields("LabelUnicode"))
            If tmp = "" Then tmp = Join(Data, de_Field) Else tmp = tmp & de_Record & Join(Data, de_Field)
            .MoveNext
         End With
      Wend
   End If
   
   Set sqlArtist = Nothing
   Set sqlTracks = Nothing
   bQuerying = False
   GetTracksByAlbum = tmp
Error:
   'WriteEvent "GetTracksByAlbum", Err.Description, Err.Source
End Function

Public Function GetGenre() As String
   On Error GoTo Error
   
   Dim sqlGenre As cRecordset
   Dim varData As Variant
   Dim strTmp As String
   Dim sql As String
  
   sql = "SELECT ID, Name, Unicode FROM Genre ORDER BY Name, Unicode"
   Set sqlGenre = mConnection.OpenRecordset(sql, True)
   
   While Not sqlGenre.EOF
      varData = Array(sqlGenre.Fields("ID"), sqlGenre.Fields("Name"), sqlGenre.Fields("Unicode"))
      If strTmp = "" Then strTmp = Join(varData, de_Field) Else strTmp = strTmp & de_Record & Join(varData, de_Field)
      sqlGenre.MoveNext
   Wend
   
   GetGenre = strTmp
   Exit Function
Error:
   Log.Writelog "GeGenre", "ERROR", Err.Description, "NULL"
End Function

Public Function GetLabel() As String

   On Error GoTo Error
   Dim sqlLabel As cRecordset
   Dim varData As Variant
   Dim strTmp As String
   Dim sql As String
  
   sql = "SELECT ID, Name, Unicode FROM Label ORDER BY Name, Unicode"
   Set sqlLabel = mConnection.OpenRecordset(sql, True)
   
   While Not sqlLabel.EOF
      varData = Array(sqlLabel.Fields("ID"), sqlLabel.Fields("Name"), sqlLabel.Fields("Unicode"), 0)
      If strTmp = "" Then strTmp = Join(varData, de_Field) Else strTmp = strTmp & de_Record & Join(varData, de_Field)
      sqlLabel.MoveNext
   Wend
   
   GetLabel = strTmp
   Exit Function
Error:
   Log.Writelog "GetLabel", "ERROR", Err.Description, "NULL"
End Function


Public Function GetTVSeriesName(ByVal ID As Long) As String

   On Error GoTo Error
   
   Dim sqlQuery As cRecordset
   Set sqlQuery = mConnection.OpenRecordset("SELECT Name, Unicode FROM TVSeries WHERE ID=" & ID)
   If Not sqlQuery.EOF Then GetTVSeriesName = sqlQuery.Fields("Name")
   
   Set sqlQuery = Nothing
   Exit Function
Error:
   Log.Writelog "GetTVSeriesName", "ERROR", Err.Description, "ID = " & ID
End Function


Public Function CompareStartLetter(ByVal Source As String, ByVal Find As String, Langue As String, bUnicode As Boolean) As Boolean
   Dim i As Integer
   Dim sChar As String
   Dim loc_LanArray() As String
   Dim loc_Language As String
   
   
   Source = Trim(Source)
   If Langue <> "" Then
      loc_LanArray = Split(Langue, " ")
      loc_Language = StrConv(loc_LanArray(0), vbProperCase)
   Else
      loc_Language = "English"
   End If
   If bUnicode Then
      If Find = "#" Then
         CompareStartLetter = InStr(ChrW(&H17E0) & ChrW(&H17E1) & ChrW(&H17E2) & ChrW(&H17E3) & ChrW(&H17E4) & ChrW(&H17E5) & ChrW(&H17E6) & ChrW(&H17E7) & ChrW(&H17E8) & ChrW(&H17E9), Left(Source, Len(Find))) > 0 And Source <> ""
      Else
         Select Case loc_Language
            Case "Khmer"
               CompareStartLetter = Left(Source, Len(Find)) = Find
            Case "English", "Vietnamese"
               CompareStartLetter = StrComp(Left(Source, Len(Find)), Find, vbTextCompare) = 0
            Case Else
               CompareStartLetter = StrComp(Left(Source, Len(Find)), Find, vbBinaryCompare) = 0
         End Select
      End If
   Else
      If Find = "#" Then
         CompareStartLetter = InStr("0123456789", Left(Source, Len(Find))) > 0 And Source <> ""
      Else
         CompareStartLetter = StrComp(Left(Source, Len(Find)), Find, vbTextCompare) = 0
      End If
   End If
End Function

Public Function GetActorName(IDs As String, boolUnicode As Boolean) As String

   On Error GoTo Error
   Dim sQueryID As String
   Dim sqlQuery As cRecordset
   Dim sQueryString As String
   
   sQueryID = Trim(Replace(IDs, ";", " "))
   sQueryID = Replace(sQueryID, " ", ",")
   
   
   sQueryString = "SELECT group_concat(CASE WHEN (Name ='' OR Name IS NULL) THEN Unicode ELSE Name END,' + ') as Name, " & _
            "       group_concat(CASE WHEN (Unicode='' OR Unicode IS NULL) THEN Name ELSE Unicode END,' + ') as Unicode  FROM  Actor WHERE ID IN(" & sQueryID & ") " & _
            "ORDER BY Name"
   Set sqlQuery = mConnection.OpenRecordset(sQueryString)
   
   If Not sqlQuery.EOF Then
      If boolUnicode Then GetActorName = sqlQuery.Fields("Name") & "%%" & sqlQuery.Fields("Unicode") Else GetActorName = sqlQuery.Fields("Name")
   End If
   
   Exit Function
Error:
   Log.Writelog "GetActorName", "ERROR", Err.Description, "IDs = " & IDs & " ,boolUnicode=" & boolUnicode

End Function
Public Function GetDirectorName(IDs As String, Optional maxName As Integer = 2) As String

   On Error GoTo Error
   Dim sQueryID As String

   'Dim sDirectorName As String
   Dim sQueryString As String
   Dim sqlQuery As cRecordset
   
   sQueryID = Trim(Replace(IDs, ";", " "))
   sQueryID = Replace(sQueryID, " ", ",")
   sQueryString = "SELECT group_concat(Name,'+') as Name, group_concat(Unicode,' + ') as Unicode FROM Director WHERE ID IN(" & sQueryID & ") ORDER BY Name LIMIT " & maxName
   Set sqlQuery = mConnection.OpenRecordset(sQueryString)
   If Not sqlQuery.EOF Then GetDirectorName = sqlQuery.Fields("Name")
   
   Exit Function
Error:
   Log.Writelog "GetDirectorName", "ERROR", Err.Description, "IDs = " & IDs & " ,maxName=" & maxName & " ,sQueryString=" & sQueryString

End Function

Public Function GetAlbumField(ByVal albumID As Long, FieldName As String) As Variant
   On Error GoTo Error
   

   Dim sQueryString As String
   Dim sqlQuery As cRecordset
   
   sQueryString = "SELECT " & FieldName & " FROM Movie WHERE ID=" & albumID
   Set sqlQuery = mConnection.OpenRecordset(sQueryString)
   If Not sqlQuery.EOF Then GetAlbumField = sqlQuery.Fields(FieldName)
   
   Exit Function
Error:
   Log.Writelog "GetAlbumField", "ERROR", Err.Description, "albumID = " & albumID & " ,FieldName=" & FieldName & " ,sQueryString=" & sQueryString
End Function


Public Function GetLabelName(ByVal ID As Long) As String

   On Error GoTo Error
   Dim sQueryString As String
   Dim sqlQuery As cRecordset
   
   sQueryString = "SELECT Name FROM Label WHERE ID=" & ID
   Set sqlQuery = mConnection.OpenRecordset(sQueryString)
   If Not sqlQuery.EOF Then GetLabelName = sqlQuery.Fields("Name")
   
   Exit Function
Error:
   Log.Writelog "GetLabelName", "ERROR", Err.Description, "ID = " & ID & " ,sQueryString=" & sQueryString

End Function

Public Function getAlbumAvailableQuery() As String
   If selectAlbumType = All Then
      getAlbumAvailableQuery = "1,0"
   ElseIf selectAlbumType = Available Then
      getAlbumAvailableQuery = "1"
   ElseIf selectAlbumType = NotAvailable Then
      getAlbumAvailableQuery = "0"
   End If
End Function



Public Function GetTVSeriesAvailable() As String
   On Error GoTo Error
   Dim varData As Variant
   Dim mReturnString As String
   Dim sQueryString As String
   Dim sqlQuery As cRecordset
   
   sQueryString = "SELECT TVSeries.ID, TVSeries.Name, TVSeries.Unicode, TVSeries.Seasoncount FROM TVSeries INNER JOIN Movie ON TVSeries.ID=Movie.TVID WHERE Movie.Available IN (" & getAlbumAvailableQuery & ") ORDER BY TVSeries.Name"
   Set sqlQuery = mConnection.OpenRecordset(sQueryString)
   While Not sqlQuery.EOF
      varData = Array(sqlQuery.Fields("ID"), sqlQuery.Fields("Name"), sqlQuery.Fields("Unicode"), sqlQuery.Fields("Seasoncount"))
      If mReturnString = "" Then mReturnString = Join(varData, de_Field) Else mReturnString = mReturnString & de_Record & Join(varData, de_Field)
      sqlQuery.MoveNext
   Wend
   GetTVSeriesAvailable = mReturnString
   Exit Function
Error:
   Log.Writelog "GetTVSeriesAvailable", "ERROR", Err.Description, "sQueryString=" & sQueryString
 
End Function

Public Function AlbumByArtist(ArtistID As Long, Optional selectType As AlbumEnum = All, Optional ArrangeBy As SortEnum = stName, Optional Favorite As Boolean, Optional CDType As CDTypeEnum = AllCD, Optional SearchPage As Boolean = False) As String
   On Error GoTo Error
   Dim i As Long
   Dim fieldAvailable As Integer
   Dim fieldFlag As Integer
   Dim fieldArtists As Integer
   Dim count As Integer
   Dim SortKey As Integer
   Dim boolSelect As Boolean
   Dim var_tmpAlbum() As Variant
   Dim loc_AlbumID As Long
   
   
   If SearchPage = False Then Erase var_Album
   
   Dim loc_Available As Integer
   Dim loc_Flag As Integer
   Dim loc_Artists As String
   
   
   With PV_TableMovie
      fieldAvailable = .FieldToIndex("Available")
      fieldFlag = .FieldToIndex("Flag")
      fieldArtists = .FieldToIndex("ActorID")
      
      
      boolSelect = True
      For i = 0 To .RecordCount - 1
         loc_Flag = FlagCDOrder(.AbsoluteField(i, fieldFlag))
         loc_Artists = .AbsoluteField(i, fieldArtists)
         If InStr(1, loc_Artists, ";" & ArtistID & ";", vbTextCompare) > 0 Then
            loc_AlbumID = .AbsoluteField(i, 0)
            loc_Available = getAlbumAvailable(prv_Country, loc_AlbumID)
            If IsCDAvailable(loc_Available, selectType) And IsValidCDType(loc_Flag, selectFlagType) Then
               ReDim Preserve var_tmpAlbum(count)
               var_tmpAlbum(count) = GetAlbumInfos(i, loc_Available, SortKey, True, loc_Flag)
               count = count + 1
            End If
         End If
         
      Next
   End With
   
   Dim sortIndices As Variant
  
   sortIndices = ReturnSortArray(3, True, 11, True, 1, ArrangeBy = stName, 4, ArrangeBy = stDate)
   If count > 0 Then QuickSort var_tmpAlbum, sortIndices, False
   If SearchPage Then var_SearchAlbum = var_tmpAlbum Else var_Album = var_tmpAlbum
   AlbumByArtist = count
   Exit Function
Error:
   Log.Writelog "AlbumByArtist", "ERROR", Err.Description, "ArtistID = " & ArtistID & " ,sQueryString=" & sQueryString & " ,selectType=" & selectType & " ,ArrangeBy=" & ArrangeBy
   'WriteEvent "AlbumByArtist", Err.Description, Err.Source
End Function

Public Function GetAlbumInfos(Index As Long, albumAvailable As Integer, SortKey As Integer, LeaveYear As Boolean, AlbumFlag As Integer) As Variant
   On Error GoTo Error
   Dim albumID As Long, albumName As String, albumUnicode As String, albumYear As String
   Dim AlbumRating As Long, AlbumFavorite As Long, AlbumRank As Long
   Dim AlbumActor As String, AlbumActorUnicode As String
   Dim AlbumDirector As String
   Dim ALbumLabel As String
   Dim AlbumArtist  As String
   
   Dim AlbumGroupCDType As Integer
   Dim AlbumDate As Date
   Dim AlbumAudioIDs As String
   Dim AlbumArtistID As Long
   Dim TVSeriesID As Long
   Dim numSeason As Long
   'modified 2005-08-23/seng
   
   
   
   Dim sqlQuery As cRecordset
   'Set sqlQuery = mConnection.OpenRecordset("SELECT ID, Name, Unice, Year, Rating, Favorite, Rank, Label, LabelUnicode, TypeCDID, AlbumArtistID, ActorName, TVSeriesName, numSeason, AudioIDs, DirectorID, ActorID FROM Movie WHERE ID=" & ID)
   
   
   
'''   With PV_TableMovie
'''      albumID = .AbsoluteField(Index, "ID")
'''      albumName = .AbsoluteField(Index, "Name")
'''      albumUnicode = .AbsoluteField(Index, "Unicode")
'''      albumYear = .AbsoluteField(Index, "Year")
'''      AlbumRating = .AbsoluteField(Index, "Rating")
'''      AlbumFavorite = .AbsoluteField(Index, "Favorite")
'''      AlbumRank = .AbsoluteField(Index, "Rank")
'''      ALbumLabel = .AbsoluteField(Index, "Label")                   'Use label instead of Singer
'''      AlbumActorUnicode = .AbsoluteField(Index, "LabelUnicode")           'Use label instead of Singer
'''      AlbumGroupCDType = .AbsoluteField(Index, "TypeCDID")               'Use label instead of Singer
'''
'''      AlbumArtistID = .AbsoluteField(Index, "AlbumArtistID")
'''      TVSeriesID = .AbsoluteField(Index, "TVID")
'''      If TVSeriesID > 0 Then
'''         numSeason = .AbsoluteField(Index, "numSeason")
'''         AlbumArtist = GetTVSeriesName(TVSeriesID) & "\Season # " & (numSeason)
'''      ElseIf AlbumArtistID > 0 Then
'''         AlbumArtist = GetActorName(";" & AlbumArtistID & ";", False)
'''      End If
'''
'''      If AlbumArtist <> "" Then ALbumLabel = AlbumArtist
'''      AlbumDate = ChangeDate(albumYear, LeaveYear)
'''      AlbumAudioIDs = .AbsoluteField(Index, "AudioIDs")
'''      ' AlbumFlag=
'''      AlbumDirector = GetDirectorName(.AbsoluteField(Index, "DirectorIDs"), 2)
'''      AlbumActor = GetActorName(.AbsoluteField(Index, "ActorID"), False)
'''      Call isValidGroupType(AlbumGroupCDType)
'''                              '0          1           2           3        4              5              6           7                 8       9     10           11                   12             13          14             15             16
'''      GetAlbumInfos = Array(albumID, albumName, albumUnicode, SortKey, AlbumDate, ALbumLabel, AlbumRating, albumAvailable, AlbumFavorite, 0, AlbumRank, AlbumGroupCDType, AlbumActorUnicode, AlbumActor, AlbumDirector, AlbumAudioIDs, AlbumFlag)
'''
'''   End With
   Exit Function
Error:
   Log.Writelog "GetAlbumInfos", "ERROR", Err.Description, "Index = " & Index & " ,albumAvailable=" & albumAvailable & " ,SortKey=" & SortKey & " ,LeaveYear=" & LeaveYear & " ,AlbumFlag=" & AlbumFlag
End Function



'Public Function AlbumByTVSeries(TVID As Long, SeasonNumber As Integer, Optional selectType As AlbumEnum = All, Optional CDType As CDTypeEnum = AllCD) As Long
'
'   On Error GoTo Error
'   Dim i As Long
'
'   Dim loc_fieldTVID As Integer, loc_fieldAvailable As Integer, loc_fieldFlag As Integer, loc_fieldSeasonNum As Integer
'   Dim loc_Available As Integer, loc_Flag As Integer
'
'   Dim loc_TVID As Long, loc_SeasonNumber As Long, loc_Count As Long
'
'   Erase var_Album
'
'   Dim sqlQuery As cRecordset
'
'
'   With PV_TableMovie
'      loc_fieldTVID = .FieldToIndex("TVID")
'      loc_fieldSeasonNum = .FieldToIndex("numSeason")
'      loc_fieldAvailable = .FieldToIndex("Available")
'      loc_fieldFlag = .FieldToIndex("Flag")
'      For i = 0 To .RecordCount - 1
'         loc_TVID = .AbsoluteField(i, loc_fieldTVID)
'         loc_SeasonNumber = .AbsoluteField(i, loc_fieldSeasonNum)
'         loc_Available = .AbsoluteField(i, loc_fieldAvailable)
'         loc_Flag = FlagCDOrder(.AbsoluteField(i, loc_fieldFlag))
''         If loc_TVID <> 0 Then
''            Debug.Assert False
''         End If
'         If loc_TVID = TVID And (SeasonNumber = loc_SeasonNumber Or SeasonNumber = 0) Then
'            If IsCDAvailable(loc_Available, selectType) And IsValidCDType(loc_Flag, CDType) Then
'               ReDim Preserve var_Album(loc_Count)
'               var_Album(loc_Count) = GetAlbumInfos(i, loc_Available, 0, False, loc_Flag)
'               loc_Count = loc_Count + 1
'            End If
'         End If
'      Next
'   End With
'   If loc_Count > 0 Then
'      'sort by name only
'      QuickSort var_Album, Array(1), False
'   End If
'   AlbumByTVSeries = loc_Count
'   Exit Function
'Error:
'   WriteEvent "AlbumByTVSeries", Err.Description, Err.Source
'End Function
'

Public Sub CloseAllDatabase()
   On Error GoTo Error
   'Do nothing since no close connectio
   Exit Sub
Error:
   Log.Writelog "CloseAllDatabase", "ERROR", Err.Description, ""
End Sub
