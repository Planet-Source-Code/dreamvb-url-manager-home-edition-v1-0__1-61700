Attribute VB_Name = "DBMod"
Type bkDb
    Sig As String * 3
    NoOfTables As Long
    Version As String * 1
End Type

Type STRUCT_FILE
    TableName() As String
    TableContents() As String
End Type

Type Edit_URL
    TSiteName As String     ' Name of the bookmark
    TSiteURL As String      ' Address URL for the bookmark
    TDateAdded As String    ' Bookmark added date
    TAddLastVis As String   ' Last visted date
    THitCnt As Integer      ' Bookmark hits counter
    TSiteDescription As String  ' Description of the bookmark
    TRated As Long              ' Bookmark rated status
    TIcon As Integer            ' Icon the bookmark has
    TVieded As Integer          ' Reteuns if link has been viewed 1 = yes 0 = no
    TWebCap As String           ' Screenshot of the bookmark page
End Type


Public d_base As Database
Public Recored_Set As Recordset
Public Q_Def As QueryDef
Public T_def As TableDef
Public T_Feild As Field

Public BookCount As Long

Public EdURL As Edit_URL

Public dbStruct As STRUCT_FILE
Public dbkHead As bkDb

Public Sub Initcbocat(cbocat As ComboBox)
' This sub will add all the table names to a combo box control
    cbocat.Clear
    cbocat.AddItem "All categories"
    For Each T_def In d_base.TableDefs
        If T_def.Attributes = 0 Then
            cbocat.AddItem T_def.Name
        End If
    Next
    
End Sub

Function AddTable(TableName As String) As Boolean
Dim mTable(1 To 11), mTableType(1 To 11) As DataTypeEnum, Cnt As Long, mTableSize(1 To 11) As Integer
On Error Resume Next

    ' Table feild names
    mTable(1) = "URLID"
    mTable(2) = "URLName"
    mTable(3) = "URLLink"
    mTable(4) = "DateAdd"
    mTable(5) = "LastVis"
    mTable(6) = "URLClicks"
    mTable(7) = "URLDescription"
    mTable(8) = "Rated"
    mTable(9) = "Icon"
    mTable(10) = "Viewed"
    mTable(11) = "Screenshot"
    
    ' Table Feild types
    mTableType(1) = 4
    mTableType(2) = dbText
    mTableType(3) = dbText
    mTableType(4) = dbDate
    mTableType(5) = dbDate
    mTableType(6) = dbInteger
    mTableType(7) = dbMemo
    mTableType(8) = dbInteger
    mTableType(9) = dbInteger
    mTableType(10) = dbInteger
    mTableType(11) = dbText
    
    'Table Sizes
    mTableSize(2) = 50
    mTableSize(3) = 50
    mTableSize(4) = 8
    mTableSize(5) = 8
    mTableSize(6) = 2
    mTableSize(7) = 50
    mTableSize(8) = 2
    mTableSize(9) = 2
    mTableSize(10) = 2
    mTableSize(11) = 128
    
    Set T_def = d_base.CreateTableDef(TableName)
    For Cnt = 1 To 11
        Set T_Feild = T_def.CreateField(mTable(Cnt), mTableType(Cnt), mTableSize(Cnt))
        If Cnt = 1 Then T_Feild.Attributes = 49
        T_def.Fields.Append T_Feild
    Next
    
    Cnt = 0
    
    d_base.TableDefs.Append T_def
    
    If Err Then
        AddTable = False
        Exit Function
    Else
        AddTable = True
    End If
    
End Function
Function DeleteTable(TableName As String)
On Error Resume Next
' This is used to delete a table from the database
    Set Recored_Set = Nothing
    d_base.TableDefs.Delete TableName
    If Err.Number = 3265 Then Err.Clear
    
End Function

Function RenameTable(TableName As String, NewName As String) As Long
On Error Resume Next
' This is used to delete a table from the database
    Set Recored_Set = Nothing
    d_base.TableDefs(TableName).Name = NewName
    RenameTable = 1
    If Err.Number = 3265 Then
        Err.Clear
        RenameTable = 0
    End If
End Function

Function AddNewUrl(RecoredSet As String)
' This adds a new bookmark to the database
On Error Resume Next
    Set Recored_Set = d_base.OpenRecordset(RecoredSet)

    With Recored_Set
        .AddNew ' Add a new recored
            !Urlname = EdURL.TSiteName
            !UrlLink = EdURL.TSiteURL
            !DateAdd = EdURL.TDateAdded
            !LastVis = EdURL.TAddLastVis
            !URLClicks = EdURL.THitCnt
            !URLDescription = EdURL.TSiteDescription
            !Rated = EdURL.TRated
            !Icon = EdURL.TIcon
            !Viewed = EdURL.TVieded
            !Screenshot = EdURL.TWebCap
        .Update
    End With
    
    Set Recored_Set = Nothing
    If Err Then Err.Clear
    
End Function
Function EditSite(lzUrlId As Long, RecoredSet As String)
Dim StrSql As String
On Error Resume Next

    StrSql = "SELECT URLID,URLName,URLLink,DateAdd," _
    & "LastVis,URLClicks,URLDescription,Rated,Icon,Viewed,Screenshot " & _
    "FROM " & RecoredSet & " WHERE URLID Like '*" & lzUrlId & "*'"
    
    Set Recored_Set = d_base.OpenRecordset(StrSql)
    If Recored_Set.RecordCount = 0 Then Exit Function
    
    With Recored_Set
        .Edit
            !Urlname = EdURL.TSiteName
            !UrlLink = EdURL.TSiteURL
            !DateAdd = EdURL.TDateAdded
            !LastVis = EdURL.TAddLastVis
            !URLClicks = EdURL.THitCnt
            !URLDescription = EdURL.TSiteDescription
            !Rated = EdURL.TRated
            !Icon = EdURL.TIcon
            !Viewed = EdURL.TVieded
            !Screenshot = EdURL.TWebCap
        .Update
    End With
    
    Set Recored_Set = Nothing
    StrSql = ""

End Function
Function DeleteURL(lzUrlId As Long, RecoredSet As String)
Dim StrSql As String
On Error Resume Next
    
    StrSql = "SELECT URLID,URLName,URLLink,DateAdd,LastVis," _
    & "URLClicks,URLDescription,Rated,Icon,Viewed,Screenshot " & _
    "FROM " & RecoredSet & " WHERE URLID Like '*" & lzUrlId & "*'"
    
    Set Recored_Set = d_base.OpenRecordset(StrSql)
    If Recored_Set.RecordCount = 0 Then Exit Function
    
    With Recored_Set
        .Delete
    End With
    
    Set Recored_Set = Nothing
    StrSql = ""

End Function

Function ShowInfo(RecoredSet As String, IDNum As Long) As String
Dim StrA As String
  On Error Resume Next

    StrSql = "SELECT URLDescription,URLLink,Rated,Screenshot,Icon,Rated,Viewed " & "FROM " & RecoredSet & " WHERE URLID Like '*" & IDNum & "*'"
    Set Recored_Set = d_base.OpenRecordset(StrSql)

    With Recored_Set
        StrA = StrConv(LoadResData(106, "CUSTOM"), vbUnicode)
        StrA = Replace(StrA, "$URL$", !UrlLink)
        StrA = Replace(StrA, "$URL_NAME$", !UrlLink)
        StrA = Replace(StrA, "<!--IMG-->", GetImage(Val(!Rated)))
        StrA = Replace(StrA, "$NUM$", !Rated)
        StrA = Replace(StrA, "$DES$", !URLDescription)
        
        TBookMarkDes = !URLDescription
        TBookSnapShot = !Screenshot
        TvIcon = !Icon
        TViewed = !Viewed
        TRate = !Rated
        OutputHtml StrA, "tmpxtygtp4.html"
        
    End With
    If Err Then Err.Clear
    
    Set Recored_Set = Nothing
    StrA = ""
    StrSql = ""
    
End Function

Sub LoadSites(RecoredSet As String)
Dim Icnt As Long

On Error Resume Next
    With d_base
        Set Recored_Set = .OpenRecordset(RecoredSet)
        If Recored_Set.RecordCount = 0 Then Exit Sub
        With Recored_Set
                While Not Recored_Set.EOF
                    Icnt = Icnt + 1
                    Set lstItem = frmmain.lstsites.ListItems.Add(, "a" & !urlid, !Urlname, Val(!Icon), Val(!Icon))
                    lstItem.SubItems(1) = !UrlLink
                    lstItem.SubItems(2) = Format(!DateAdd, "Medium Date")
                    lstItem.SubItems(3) = Format(!LastVis, "Medium Date")
                    lstItem.SubItems(4) = !URLClicks
                    
                    If Val(!Viewed) = 1 Then
                        frmmain.lstsites.ListItems(Icnt).ForeColor = Config.NewItems
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(1).ForeColor = Config.NewItems
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(2).ForeColor = Config.NewItems
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(3).ForeColor = Config.NewItems
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(4).ForeColor = Config.NewItems
                    Else
                        frmmain.lstsites.ListItems(Icnt).ForeColor = vbBlack
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(1).ForeColor = vbBlack
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(2).ForeColor = vbBlack
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(3).ForeColor = vbBlack
                        frmmain.lstsites.ListItems(Icnt).ListSubItems(4).ForeColor = vbBlack
                    End If
                    .MoveNext
                Wend
            End With
        End With
        
        Icnt = 0
        BookCount = Recored_Set.RecordCount
        Set Recored_Set = Nothing
    
        ResizeLstHeader frmmain.lstsites, 0
        ResizeLstHeader frmmain.lstsites, 1
 
        
End Sub
Function TableCount() As Long
On Error Resume Next
    ' This will return the number of tables foudn in the database
    Dim tblCnt As Long, TotalRecored As Long
   BookCount = 0
    For Each T_def In d_base.TableDefs
        If T_def.Attributes = 0 Then
            tblCnt = tblCnt + 1
            Set Recored_Set = d_base.OpenRecordset(T_def.Name)
            BookCount = BookCount + Recored_Set.RecordCount
            Set Recored_Set = Nothing
        End If
    Next
    TableCount = tblCnt
    tblCnt = 0
    
End Function

Function GenHtmlPage(RecoredSet As String) As String
Dim StrA As String, StrB As String, StrC As String

    Set Recored_Set = d_base.OpenRecordset(RecoredSet)
    StrA = StrConv(LoadResData(109, "CUSTOM"), vbUnicode)

    While Not Recored_Set.EOF
        With Recored_Set
            StrB = StrA
            StrB = Replace(StrB, "$URL_NAME$", !Urlname)
            StrB = Replace(StrB, "$URL$", !UrlLink)
            StrC = StrC & StrB
            StrB = ""
            .MoveNext
        End With
    Wend
    StrA = StrConv(LoadResData(108, "CUSTOM"), vbUnicode)
    StrA = Replace(StrA, "$CAT$", RecoredSet)
    StrA = StrA & StrC & "<hr size=""1"">"
    StrB = ""
    StrC = ""
    GenHtmlPage = StrA
    StrA = ""
    
End Function
Function FindSite(SiteName As String, RecoredSet As String)
Dim StrSql As String
On Error Resume Next

    StrSql = "SELECT URLID,UrlLink,URLName,DateAdd,LastVis,URLClicks " & _
    "FROM " & RecoredSet & " WHERE URLName Like '*" & SiteName & "*'"
    
    Set Recored_Set = d_base.OpenRecordset(StrSql)
    If Recored_Set.RecordCount = 0 Then SiteFound = False: Exit Function
    
    frmserach.lstfind.ListItems.Clear
    
    With Recored_Set
        While Not Recored_Set.EOF
            Set lstItem = frmserach.lstfind.ListItems.Add(, "a" & !urlid, !Urlname, 1, 1)
            lstItem.SubItems(1) = !UrlLink
            lstItem.SubItems(2) = Format(!DateAdd, "Medium Date")
            lstItem.SubItems(3) = Format(!LastVis, "Medium Date")
            lstItem.SubItems(4) = !URLClicks
            .MoveNext
        Wend
    End With
    
    Set Recored_Set = Nothing
    StrSql = ""
    SiteFound = True
    
End Function

Function UpdateDB(lzUrlId As Long, RecoredSet As String)
Dim StrSql As String
On Error Resume Next

    StrSql = "SELECT LastVis,URLClicks,Viewed " & _
    "FROM " & RecoredSet & " WHERE URLID Like '*" & lzUrlId & "*'"
    
    Set Recored_Set = d_base.OpenRecordset(StrSql)
    If Recored_Set.RecordCount = 0 Then Exit Function
    
    With Recored_Set
        .Edit
            !LastVis = Format$(Date, "Medium Date")
            !URLClicks = !URLClicks + 1
            !Viewed = 0
        .Update
    End With
    
    Set Recored_Set = Nothing
    StrSql = ""

End Function

Function ExportToIE(RecoredName As String) As String

    Set Recored_Set = d_base.OpenRecordset(RecoredName)
   
    While Not Recored_Set.EOF
        With Recored_Set
            StrA = StrA & !Urlname & Chr(128) & !UrlLink & vbCrLf
            .MoveNext
        End With
    Wend
    
    Set Recored_Set = Nothing
    ExportToIE = StrA
    StrA = ""
End Function

Public Function TMoveToUrl(OldRecored As String, lzUrlId As Long, NewRecored As String) As Long


Dim StrSql As String

    StrSql = "SELECT URLID,URLName,URLLink,DateAdd," _
    & "LastVis,URLClicks,URLDescription,Rated,Icon,Viewed,Screenshot " & _
    "FROM " & OldRecored & " WHERE URLID Like '*" & lzUrlId & "*'"

    Set Recored_Set = d_base.OpenRecordset(StrSql)
    If Recored_Set.RecordCount = 0 Then TMoveToUrl = -1: Exit Function
    
    With Recored_Set
        EdURL.TSiteName = !Urlname
        EdURL.TSiteURL = !UrlLink
        EdURL.TDateAdded = !DateAdd
        EdURL.TAddLastVis = !LastVis
        EdURL.THitCnt = !URLClicks
        EdURL.TSiteDescription = !URLDescription
        EdURL.TRated = !Rated
        EdURL.TIcon = !Icon
        EdURL.TVieded = !Viewed
        EdURL.TWebCap = !Screenshot
    End With
    
    AddNewUrl NewRecored
    Set Recored_Set = Nothing
    StrSql = ""
    TMoveToUrl = 1
    
End Function
' Back Up of database code
Public Function GetRecoredInfo(RecoredName As String) As String
Dim sBuff As String

    Set Recored_Set = d_base.OpenRecordset(RecoredName)
    
    Do While Not Recored_Set.EOF
        With Recored_Set
           sBuff = sBuff & En(!Urlname & Chr(128) & !UrlLink & Chr(128) & !DateAdd & Chr(128) & !LastVis & Chr(128) & !URLClicks & Chr(128) & !URLDescription & Chr(128) & !Rated & Chr(128) & !Icon & Chr(128) & !Viewed & Chr(128) & !Screenshot & Chr(128) & "[END]")
            .MoveNext
        End With
    Loop
    
    Set Recored_Set = Nothing
    GetRecoredInfo = sBuff
    sBuff = ""
    
End Function

Public Function BackupDB(dbBackFile As String)
On Error Resume Next
Dim StrB As String, I As Long, nFile As Long
    
    frmmain.MousePointer = vbHourglass ' set the mouse pointer to busy
    
    ReDim dbStruct.TableName(0) ' Resize array
    ReDim dbStruct.TableContents(0) ' Resize array
    
    For Each T_def In d_base.TableDefs
        If T_def.Attributes = 0 Then
            I = I + 1 ' Add one to our counter
            dbkHead.NoOfTables = I ' set the number of tables in db
            StrB = GetRecoredInfo(T_def.Name) ' get the current recored info
            ReDim Preserve dbStruct.TableName(dbkHead.NoOfTables)
            ReDim Preserve dbStruct.TableContents(dbkHead.NoOfTables)
            dbStruct.TableName(dbkHead.NoOfTables) = T_def.Name
            dbStruct.TableContents(dbkHead.NoOfTables) = StrB
            StrB = ""
        End If
        DoEvents
    Next
    
    I = 0 ' Reset counter
    dbkHead.Sig = "DBK" ' Sig name for the file
    dbkHead.Version = Chr$(1)
    nFile = FreeFile ' Pointer to free file
    
    Open dbBackFile For Binary As #nFile ' Open the file in binary mode
        Put #nFile, , dbkHead   ' File Head info
        Put #nFile, , dbStruct  ' File data info table
    Close #1
    
    I = 0
    Erase dbStruct.TableName()
    Erase dbStruct.TableContents()
    frmmain.MousePointer = vbNormal ' reset the mouse cursor state
    
End Function
' End Back Up of database code

Public Function RestoreDb(lzBackFile As String, mDataBase As String)

On Error Resume Next
Dim iFile As Long, iCounter As Long, jCounter As Long
Dim StrB As String, lzStr As String, vStr As Variant, iVal As Long, bStr As Variant

    
    iFile = FreeFile ' Pointer to free file
    frmrestore.MousePointer = vbHourglass ' set mouse pointer to busy
    
    Open lzBackFile For Binary As #iFile
        Get #iFile, , dbkHead ' File header info
        Get #iFile, , dbStruct ' File info table
    Close #iFile
    
    d_base.Close ' Close the main programs database
    DBEngine.Workspaces(0).CreateDatabase mDataBase, dbLangGeneral ' Create a new database
    Set d_base = OpenDatabase(mDataBase) ' Open the new created database
    
    For iCounter = 1 To dbkHead.NoOfTables
        lzStr = dbStruct.TableName(iCounter) ' Get the table names
        StrB = En(dbStruct.TableContents(iCounter)) ' Get recored info
        AddTable lzStr
        vStr = Split(StrB, "[END]") ' Find the end
        iVal = UBound(vStr) - 1 ' Get total number of bookmarks
        
        For jCounter = 0 To iVal
            bStr = Split(vStr(jCounter), Chr$(128))
            If Len(bStr(jCounter)) > 0 Then
                EdURL.TSiteName = bStr(0)
                EdURL.TSiteURL = bStr(1)
                EdURL.TDateAdded = bStr(2)
                EdURL.TAddLastVis = bStr(3)
                EdURL.THitCnt = Val(bStr(4))
                EdURL.TSiteDescription = bStr(5)
                EdURL.TRated = Val(bStr(6))
                EdURL.TIcon = Val(bStr(7))
                EdURL.TVieded = Val(bStr(8))
                EdURL.TWebCap = bStr(9)
                AddNewUrl lzStr
            End If
        Next
        DoEvents
    Next
    
    ' Reset vars
    iCounter = 0
    jCounter = 0
    lzStr = ""
    StrB = ""
    Erase vStr
    Erase bStr
    
    d_base.Close ' Close the database
    frmrestore.MousePointer = vbNormal ' Reset mouse pointer
    RestoreDb = 1
End Function

Public Function UpdateViewStat(lzUrlId As Long, RecoredSet As String, mViewedStat As Long)
Dim StrSql As String
On Error Resume Next

    StrSql = "SELECT Viewed FROM " & RecoredSet & " WHERE URLID Like '*" & lzUrlId & "*'"

    Set Recored_Set = d_base.OpenRecordset(StrSql)
    If Recored_Set.RecordCount = 0 Then Exit Function
    
    With Recored_Set
        .Edit
            !Viewed = mViewedStat
        .Update
    End With
    
    Set Recored_Set = Nothing
    StrSql = ""
    If Err Then MsgBox Err.Description
    
End Function
