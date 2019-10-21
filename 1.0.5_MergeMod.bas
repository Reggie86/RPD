Attribute VB_Name = "MergeMod"
Option Compare Database
Option Explicit
Public g_Runkey As String
Public MergeModVer As String
Public ModuleDCVer As String
'=================================================================================================================================================================
' Name        : MergeMod
' Author      : David Campbell
' Date        : 17 OCT 2019
' Description : Version 1.0.5 - Added X_Merge_Run_Log table and sub to write here with Merge Success/Failure  message
' Description : Version 1.0.4 - Added code to get, compare table counts of EXT_SITE db and internal db and ignore tables with no corresponding table in the other environment.
' Desc cont.  : This is aware of extra tables in the EXT_SITE tables and tables in the CurrentDB not matched in the EXT_SITE tables
' Description : Version 1.0.3 - Added Get_Merge_Counts_Validation
' Description : Version 1.0.2 - cleans up start and stop messages for subs, and calls the the NonMV_Table_Inserts sub for a second time after the MVInserts and Appends
' to work around a dependency problem. Adds and Optional parameter to NonMV_Table_Inserts so the sub is aware its being called for the second time. Also cleaned up the runid to use the global
' var as I intended and have one guid in the XMergeactivity for 1 run. Add new public var MergeMod so Shaggy can add that to a table.
' Description : Version 1.0.1 Baseline
'=================================================================================================================================================================
Sub Initialize_Vars()
    MergeModVer = "1.0.5"
    ModuleDCVer = "1.0.2"
End Sub

Sub Load_BE_Tables()
    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim sExtDbPath As String

    sExtDbPath = ModuleDC.selectFile
    Set db = OpenDatabase(sExtDbPath)
    
    Call DC_Print("Loading tables ...standyby")
    
    For Each tdf In db.TableDefs                 'Loop through all the table in the external database
        If (Left(tdf.Name, 4) <> "MSys") And (tdf.Name Like "tbl*") Then 'Exclude System Tables
            On Error Resume Next
            Access.DoCmd.TransferDatabase acImport, "Microsoft Access", sExtDbPath, _
                                          acTable, tdf.Name, "EXT_SITE_" & tdf.Name, False
            Debug.Print tdf.Name
        End If
    Next tdf
    db.Close
    Set db = Nothing                             'Cleanup after ourselves
    Application.RefreshDatabaseWindow            ' Excellent!
    Debug.Print "External Tables Loaded"
    Call DC_Print("External Tables Loaded")

End Sub

Sub Delete_BE_Tables()
    Dim t As TableDef
    Dim mycount As Integer
    mycount = 0
    For Each t In CurrentDb.TableDefs
        If t.Name Like "EXT_SITE_tbl*" Then
            DoCmd.RunSQL ("DROP TABLE " & t.Name)
            mycount = mycount + 1
        End If
    Next
    Debug.Print "Delete Merge Tables: " & mycount & " Tables deleted"
    Application.RefreshDatabaseWindow ' Excellent!
    Debug.Print "External Tables Deleted"
    Call DC_Print("External Tables Deleted")
End Sub
Sub Build_Table_List()
'Builds a list of all tables in the CurrentDB that start with "tbl"
'9/6/2019

Dim dbs As Database
Dim localtblcount As Integer
Dim matchingEXTcount As Integer

Dim Exttblcount As Integer
Dim matchingLOCALcount As Integer

Dim LocalDelta As Integer
Dim EXTDelta As Integer
   
    
    
    Call DC_Print("Building List Of tables...standby")
    Set dbs = CurrentDb
    CurrentDb.Execute " delete * from X_mergetables ;"
    'New
    CurrentDb.Execute " delete * from X_EXT_mergetables "
    
    'NEW Get a list of the EXT_SITE_Tables
    CurrentDb.Execute " INSERT into X_EXT_mergetables (MergeTable) SELECT Name FROM [MSysObjects] WHERE Type IN (1,4,6) AND (Name Like 'EXT_SITE_tbl*') ;"
    CurrentDb.Execute " INSERT into X_mergetables (MergeTable, Flags) SELECT Name, Flags FROM [MSysObjects] WHERE Type IN (1,4,6) AND (Name Like 'tbl*') ;"
    
    CurrentDb.Execute " UPDATE X_mergetables SET KeyField = PrimKey(MergeTable)"
    CurrentDb.Execute " UPDATE X_mergetables SET AutoMerge = NO where flags <> 0"
    CurrentDb.Execute " UPDATE X_MergeMap INNER JOIN X_mergetables ON X_MergeMap.szReferencedObject = X_mergetables.MergeTable SET X_mergetables.Distance = X_MergeMap.Distance where X_MergeMap.Distance > 0"
    
    'NEW - this matches the tbl tables to the EXT_SITE_tables to find any tables in our DB missing from ext Site collection
    CurrentDb.Execute " UPDATE X_MergeTables t1 LEFT JOIN X_EXT_mergetables t2 ON Mid(t2.Mergetable, 10, Len(t2.Mergetable)+1) = t1.Mergetable SET t1.Matching_EXT_Table = NO  WHERE t2.Mergetable IS Null "
    CurrentDb.Execute " UPDATE X_EXT_MergeTables t1 LEFT JOIN X_MergeTables t2 ON  t2.Mergetable = Mid(t1.Mergetable, 10, Len(t1.Mergetable)+1) SET t1.Matching_Local_Table = NO  WHERE t2.Mergetable IS Null "
           
    localtblcount = FindRecordCount("SELECT * FROM X_mergetables")
    matchingEXTcount = FindRecordCount("SELECT * FROM X_mergetables where Matching_EXT_Table = YES ")
    LocalDelta = (localtblcount - matchingEXTcount)
  
    Exttblcount = FindRecordCount("SELECT * FROM X_EXT_mergetables")
    matchingLOCALcount = FindRecordCount("SELECT * FROM X_EXT_mergetables where Matching_Local_Table = YES ")
    EXTDelta = (Exttblcount - matchingLOCALcount)
  
  
    If LocalDelta > 0 Then
        Debug.Print "PROBLEM: Local database has " & LocalDelta & " tables Not in EXT_SITE_ Tables"
        
        If MsgBox("PROBLEM: Local database has " & LocalDelta & " tables Not in EXT_SITE_ Tables. Unmatched tables will be safely ignored. Continue?", vbYesNo + vbQuestion) = vbYes Then
            'Continue processing
        Else
            Debug.Print "Stopping Build_Table_List() Please Look at X_mergetables.Matching_EXT_Table column for more information."
            Exit Sub
        End If
     End If
    
    If EXTDelta > 0 Then
        Debug.Print "PROBLEM: EXT_SITE_ database has " & EXTDelta & " tables NOT in Local Tables"
        
        If MsgBox("PROBLEM: EXT_SITE_ database has " & EXTDelta & " tables NOT in Local Tables. Unmatched tables will be safely ignored. Continue?", vbYesNo + vbQuestion) = vbYes Then
            'Continue processing
        Else
            Debug.Print "Stopping Build_Table_List() Please Look at X_EXT_mergetables.Matching_Local_Table column for more information."
            Exit Sub
        End If
     End If
    
     Debug.Print "Loaded Merge Tables Names: " & Exttblcount & " Merge Tables Added to List"
     Call DC_Print("Loaded Merge Tables Names: " & Exttblcount & " Merge Tables Added to List")
    
    dbs.Close
    Application.RefreshDatabaseWindow 'Excellent
   
End Sub
Sub Load_Column_Names()
' For all tables

Dim dbs As Database
Dim rs As Recordset
Dim crs As Recordset
Dim myTable As String
Dim myColumn As String
Dim tblCount As Integer
Dim colcount As Integer
Dim ii As Integer

    Call DC_Print("Building List Of Column names...standby")
    Set dbs = CurrentDb
    CurrentDb.Execute " delete * from X_mergetablecolumns "
   
    Set rs = dbs.OpenRecordset("SELECT MergeTable FROM X_mergetables WHERE (Matching_EXT_Table = YES) ORDER BY MergeTable")
    
    tblCount = 0
    While Not rs.EOF
        tblCount = tblCount + 1
        myTable = rs.Fields("MergeTable")
        'Set rsEXT = dbs.OpenRecordset("SELECT Name FROM [MSysObjects] WHERE Type IN (1,4,6) AND (Name = 'EXT_SITE_'" & "'" & myTable & "')")
        
       ' FindRecordCount(strSQL As String)
        
        Set crs = dbs.OpenRecordset("SELECT * FROM " & myTable & "")
        
        For ii = 0 To crs.Fields.count - 1
            colcount = colcount + 1
            myColumn = crs.Fields(ii).Name
            CurrentDb.Execute " INSERT into X_mergetablecolumns ([TableName],[ColumnName]) Values('" & myTable & "','" & myColumn & "' );"
            Debug.Print myColumn
        Next ii
        rs.MoveNext
        CurrentDb.Execute " UPDATE X_MV_Mappings INNER JOIN X_mergetablecolumns ON (X_MV_Mappings.ColumnName = X_mergetablecolumns.ColumnName) AND (X_MV_Mappings.SourceTable = X_mergetablecolumns.TableName) SET MVColumn = YES;"
        
    Wend
    CurrentDb.Execute " UPDATE X_mergetablecolumns INNER JOIN (MSysComplexColumns INNER JOIN MSysObjects ON MSysComplexColumns.ConceptualTableID = MSysObjects.Id) ON " & _
              " (X_mergetablecolumns.ColumnName = MSysComplexColumns.ColumnName) AND (X_mergetablecolumns.TableName = MSysObjects.Name) SET X_mergetablecolumns.AttachmentCol = Yes" & _
              "  WHERE (((MSysObjects.Name) Like 'tbl*') AND ((MSysComplexColumns.ComplexTypeObjectID)=39));"
              
    CurrentDb.Execute "UPDATE X_mergetables set HasAttachments = YES where ID in (SELECT DISTINCT X_mergetables.ID " & _
                      "FROM X_mergetablecolumns INNER JOIN X_mergetables ON X_mergetablecolumns.TableName = X_mergetables.MergeTable " & _
                      "WHERE (((X_mergetablecolumns.AttachmentCol)=Yes)));"
              
              
    rs.Close
    crs.Close
    Set rs = Nothing
    Set crs = Nothing

    Debug.Print "Loaded Merge Column Names: " & colcount & " Tables Columns Added to List from " & tblCount & " Tables"
    Call DC_Print("Loaded Merge Column Names: " & colcount & " Tables Columns Added to List from " & tblCount & " Tables")
    Call DC_Print("Completed Load Of Column names.")
    dbs.Close
    Application.RefreshDatabaseWindow 'Excellent
    
End Sub
Sub Load_MV_Mappings()

Dim dbs As Database
    
    Set dbs = CurrentDb
    dbs.Execute "delete * from X_MV_Mappings", Options:=dbFailOnError + dbSeeChanges
    dbs.Execute "INSERT into X_MV_Mappings (SourceTable, ColumnName) SELECT MSysObjects.Name AS SourceTable, MSysComplexColumns.ColumnName " & _
                 "FROM MSysComplexColumns INNER JOIN MSysObjects ON MSysComplexColumns.ConceptualTableID = MSysObjects.Id " & _
                 "WHERE (((MSysObjects.Name) Like 'tbl*') AND ((MSysComplexColumns.ComplexTypeObjectID) In (33,37))) ORDER BY MSysObjects.Name;"
        
    Debug.Print "MV Mapping data captured"
    Call DC_Print("MV Mapping data captured")
    dbs.Close
    Application.RefreshDatabaseWindow 'Excellent
End Sub
Sub NonMV_Table_Inserts(Optional ByVal Passed_Runid As String)
Dim dbCurrent As DAO.Database
Dim rs As DAO.Recordset
Dim rsNonDups As DAO.Recordset
Dim rsc As DAO.Recordset

Dim arrResults(1000) As Variant
Dim mySrcTable As String
Dim myTgtTable As String
Dim myColumn As String
Dim myPkey As String
Dim strSQL As String
Dim strSQL2 As String
Dim strDelete As String
Dim strSQLCounts As String
Dim RecCount As Integer
Dim PreRecCount As Integer
Dim PostRecCount As Integer
Dim TableCount As Integer
Dim NewRectoDB As Integer
Dim SecondPass As Boolean
Dim RunLabel As String
Dim MergeFlag As Boolean
Dim MatchExtFlag As Boolean
Dim j As Integer

    If Passed_Runid = "" Then
        g_Runkey = ModuleDC.GetGUID  ' this is where the Global variable g_Runkey is initially set - all other modules can see this
        Call DC_Print("Beginning NonMV Table Inserts...standby")
        RunLabel = "NonMV Table Inserts"
    Else
        g_Runkey = Passed_Runid
        Call DC_Print("Beginning Remaining NonMV Table Inserts...standby")
        Debug.Print "Beginning Remaining NonMV Table Inserts...standby"
        RunLabel = "Remaining NonMV Table Inserts"
        SecondPass = True
    End If
    
    Set dbCurrent = CurrentDb
    Set rs = dbCurrent.OpenRecordset("SELECT MergeTable, AutoMerge FROM X_mergetables WHERE (Matching_EXT_Table = YES) ORDER BY Distance DESC")
    CurrentDb.Execute "UPDATE X_Mergetables SET PreRecCount = 0, PostRecCount = 0, NewRecords = 0, NewDataFound = NO "
    
    TableCount = 0
    NewRectoDB = 0
    
    rs.MoveFirst
    While Not rs.EOF
        TableCount = TableCount + 1
        
        RecCount = 0
        PreRecCount = 0
        PostRecCount = 0

         mySrcTable = "EXT_SITE_" & rs.Fields("MergeTable")
         myTgtTable = rs.Fields("MergeTable")
         MergeFlag = rs.Fields("AutoMerge")
         
         myPkey = PrimKey(myTgtTable)

         'Debug.Print mySrcTable
         'Debug.Print myPkey, "++++++>", myTgtTable
         
            strSQL = "SELECT t1." & myPkey & " FROM " & mySrcTable & " t1 " & _
                     "LEFT JOIN " & myTgtTable & " t2 ON t2." & myPkey & " = t1." & myPkey & " " & _
                     "WHERE t2." & myPkey & " IS Null"
            Debug.Print strSQL
            'Call DC_Print(strsql)
               
            ' This gets a count of the records found in the query above
            Set rsNonDups = dbCurrent.OpenRecordset(strSQL)  ' count records found here
            RecCount = FindRecordCount(strSQL)
            
            'This gets a count of the records in the table before the update happens as a Pre Count
            strSQLCounts = " SELECT Count(*) AS Total FROM " & myTgtTable & ""
            Set rsc = dbCurrent.OpenRecordset(strSQLCounts)
            PreRecCount = rsc!Total
            CurrentDb.Execute "UPDATE X_Mergetables SET PreRecCount = " & PreRecCount & " WHERE MergeTable = '" & myTgtTable & "'"
                              
            If (RecCount > 0) Then
               CurrentDb.Execute "UPDATE X_Mergetables SET NewDataFound = YES, NewRecords = " & RecCount & " WHERE MergeTable = '" & myTgtTable & "'"
               Debug.Print RecCount & " new records found in " & mySrcTable & ""
               Call DC_Print(RecCount & " new records found in " & mySrcTable & "")
            End If
            
            strSQL2 = "INSERT INTO " & myTgtTable & " SELECT * FROM " & mySrcTable & " WHERE " & myPkey & " IN (" & strSQL & ")"
         
         'On Error GoTo ErrorHandler
    
         If (RecCount > 0) And (MergeFlag = True) Then    ' Do not try to update tables with MV columns here
         
            Debug.Print strSQL2
            'Begin Transaction - any problems and all changes made are rolled back
            
            NewRectoDB = NewRectoDB + 1
            Call DC_Print("Begin Transaction")
            Debug.Print "Begin Transaction"
            DAO.DBEngine.BeginTrans
            On Error GoTo tran_Err
                CurrentDb.Execute strSQL2
            DAO.DBEngine.CommitTrans
            
            Call DC_Print("Transaction Committed")
            Debug.Print "Transaction Committed"
            
            strSQLCounts = " SELECT Count(*) AS Total FROM " & myTgtTable & ""
            Set rsc = dbCurrent.OpenRecordset(strSQLCounts)
            PostRecCount = rsc!Total
            Call AtivityLog(g_Runkey, RunLabel, mySrcTable, myTgtTable, strSQL2, PreRecCount, PostRecCount)
           
            j = 0
            Do Until rsNonDups.EOF
              j = j + 1
               arrResults(j) = rsNonDups.Fields(0)
              'Debug.Print arrResults(j), strSQL2
              'Call DC_Print("Delete Statement Written to Table")
              strDelete = "DELETE * FROM " & myTgtTable & " WHERE " & myPkey & " = " & rsNonDups.Fields(0) & ""
              Call Save_Delete_Statements(g_Runkey, myTgtTable, strDelete)
              rsNonDups.MoveNext
            Loop
        End If
       ' Get the Post Record Counts updated or not
       
        strSQLCounts = " SELECT Count(*) AS Total FROM " & myTgtTable & ""
        Set rsc = dbCurrent.OpenRecordset(strSQLCounts)
        PostRecCount = rsc!Total
        CurrentDb.Execute "UPDATE X_Mergetables SET PostRecCount = " & PostRecCount & " WHERE MergeTable = '" & myTgtTable & "'"
        
        Erase arrResults
        strSQL = ""
        strSQL2 = ""
        strSQLCounts = ""
        strDelete = ""
        rs.MoveNext
    Wend
        
    dbCurrent.Close
    Set rs = Nothing
    Set rsNonDups = Nothing
    Set rsc = Nothing

If SecondPass = True Then
    Debug.Print "Remaining NonMV Table Processing Completed"
    Call DC_Print("Remaining NonMV Table Processing Completed")
Else
    Debug.Print "NonMV Table Processing Completed"
    Call DC_Print("NonMV Table Processing Completed")
End If

Call DC_Print(TableCount & " NonMV Tables examined")
Call DC_Print(NewRectoDB & " Tables contained new records")


Exit Sub

tran_Err:
    
    Call DC_Print("Rolling Back Transaction " & Err.Description)
    Call DC_Print(strSQL2)
    Call MergeLogError(g_Runkey, Err.Description, strSQL2, mySrcTable, myPkey, myTgtTable)

    DAO.DBEngine.Rollback

    'Debug.Print strSQL2
    
    dbCurrent.Close
    Set rs = Nothing
    Set rsNonDups = Nothing
    Set rsc = Nothing

End Sub

Sub MV_Table_Inserts()
'=====================================
' Name        : MV_Table_Inserts
' Author      : David Campbell
' Copyright   : 2019
' Call command: Call MV_Table_Inserts()
' Description :
'===================================
Dim dbCurrent As DAO.Database
Dim rs As DAO.Recordset
Dim rsNonDups As DAO.Recordset
Dim rsc As DAO.Recordset
Dim rsMVTableCols As DAO.Recordset

Dim arrResults(2500) As Variant
Dim arrNonMVFields(2500) As Variant

Dim mySrcTable As String
Dim myTgtTable As String
Dim myColumn As String
Dim myPkey As String
Dim strSQL As String
Dim strSQL2 As String
Dim strSQLCounts As String
Dim strInsert As String

Dim strDelete As String

Dim strFieldName As String
Dim strValues As String
Dim strComma As String

Dim RecCount As Integer
Dim PreRecCount As Integer
Dim PostRecCount As Integer
Dim TableCount As Integer

Dim MergeFlag As Boolean
Dim AttachmentFlag As Boolean

Dim j, k As Integer
Dim intNewRecs As Integer
Dim stopflag1 As Integer
Dim stopflag2 As Integer
Dim Lenstr1 As Integer
Dim Lenstr2 As Integer
Dim myRecordkey As String

    'runkey = ModuleDC.GetGUID
    Set dbCurrent = CurrentDb
    Set rs = dbCurrent.OpenRecordset("SELECT MergeTable, AutoMerge, HasAttachments FROM X_mergetables WHERE (AutoMerge = NO) AND (Matching_EXT_Table = YES) ORDER by Distance DESC")
    CurrentDb.Execute "UPDATE X_Mergetables SET PreRecCount = 0, PostRecCount = 0, NewRecords = 0, NewDataFound = NO "
    
    Call DC_Print("Beginning MVTable Inserts...standby")
    Call ClearWorkList  'prepares X_MVWorklist for next run
    
    'delete * from X_MVChanges  - to get ready for run and clean out old recs - should old go to archive?
    
    rs.MoveFirst
    While Not rs.EOF
        RecCount = 0
        PreRecCount = 0
        PostRecCount = 0

         mySrcTable = "EXT_SITE_" & rs.Fields("MergeTable")
         myTgtTable = rs.Fields("MergeTable")
         MergeFlag = rs.Fields("AutoMerge")
         AttachmentFlag = rs.Fields("HasAttachments")
         myPkey = PrimKey(myTgtTable)

         

         ' this statement identifies the changes found in external tables and will be used in an INSERT statement where its used in a subquery as the criteria
         strSQL = "SELECT t1." & myPkey & " FROM " & mySrcTable & " t1 " & _
                  "LEFT JOIN " & myTgtTable & " t2 ON t2." & myPkey & " = t1." & myPkey & " " & _
                  "WHERE t2." & myPkey & " IS Null"
          'Debug.Print strSQL ' this debug is a little confusing when used
          'Call DC_Print(strSQL)
            
         ' This gets a count of the records found in the query above
         Set rsNonDups = dbCurrent.OpenRecordset(strSQL)  ' count records found here
         'RecCount = rsNonDups.RecordCount
         RecCount = FindRecordCount(strSQL)
         'Debug.Print "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$Records found: " & RecCount
         'Call DC_Print("Records found: " & RecCount)
         If (RecCount > 0) Then
            rsNonDups.MoveFirst
            
            intNewRecs = 0
            Do While Not rsNonDups.EOF
                intNewRecs = intNewRecs + 1
                arrResults(intNewRecs) = rsNonDups.Fields(0)
                'Debug.Print intNewRecs & " " & rsNonDups.Fields(0)
                rsNonDups.MoveNext
             Loop
            rsNonDups.MoveFirst
            
            'This gets a count of the records in the table before the update happens as a Pre Count
            strSQLCounts = " SELECT Count(*) AS Total FROM " & myTgtTable & ""
            Set rsc = dbCurrent.OpenRecordset(strSQLCounts)
            PreRecCount = rsc!Total
            CurrentDb.Execute "UPDATE X_Mergetables SET PreRecCount = " & PreRecCount & " WHERE MergeTable = '" & myTgtTable & "'"
            
            CurrentDb.Execute "UPDATE X_Mergetables SET NewDataFound = YES, NewRecords = " & RecCount & " WHERE MergeTable = '" & myTgtTable & "'"
            Debug.Print RecCount & " new records found in " & mySrcTable & ""
            Call DC_Print(RecCount & " new records found in " & mySrcTable & "")
         
            'this query brings back the NON complex columns for the insert statement
            Set rsMVTableCols = dbCurrent.OpenRecordset("SELECT TableName, ColumnName from X_mergetablecolumns where (mvcolumn = NO) AND (attachmentcol = NO) AND tablename = '" & myTgtTable & "'")
                                                                                                                              
            strInsert = "INSERT into  " & myTgtTable & "("
            strDelete = ""
            strComma = ", "
            strValues = " Select "
            rsMVTableCols.MoveFirst
            
            k = 0
            stopflag1 = 0
            stopflag2 = 0
            
            'Inserts a record with all Non MV and Non Attachment fields into target table
            While Not rsMVTableCols.EOF    ' begin processing one of the MV tables
               k = k + 1
               ' build insert with non mv fields first
               'Debug.Print "--------------->>>>>>>>>>>>>>>" & myTgtTable, rsMVTableCols.Fields("ColumnName")
               strFieldName = rsMVTableCols.Fields("ColumnName")
               strInsert = strInsert & strFieldName & ""
               strValues = strValues & strFieldName & ""
               
               rsMVTableCols.MoveNext
               If Not rsMVTableCols.EOF Then
                  strInsert = strInsert & strComma
                  strValues = strValues & strComma
                  If k = 10 Then
                    strInsert = strInsert & vbCrLf
                    strValues = strValues & vbCrLf
                    k = 0
                  End If
               Else
                  strInsert = strInsert & ") "
                  strValues = strValues & " FROM " & mySrcTable & " WHERE " & myPkey & " IN (" & strSQL & ")"
               End If
             Wend
            'Debug.Print strInsert
            'Debug.Print strValues
        End If
         
        strSQL2 = strInsert & " " & strValues
        'Debug.Print strSQL2
        'Call DC_Print(strSQL2)
'
'         'On Error GoTo ErrorHandler
         If (RecCount > 0) And (MergeFlag = False) Then  ' Only do the partial insert into target from mv tables
            
            'Begin Transaction - any problems and all changes made are rolled back
            'Call DC_Print("Begin Transaction")
            DAO.DBEngine.BeginTrans
            On Error GoTo tran_Err
                CurrentDb.Execute strSQL2, dbFailOnError ' all records not in base tables inserted in this statement - does not include MV fields - done in another step.
            'Debug.Print strSQL2 & "<<<<<<<<<<<<<-----------------------------------------------------------------"
            DAO.DBEngine.CommitTrans
            Debug.Print "Committed Transaction"
            'Call DC_Print("Transaction Committed")
            
            CurrentDb.Execute "UPDATE X_Mergetables SET PostRecCount = " & PostRecCount & " WHERE MergeTable = '" & myTgtTable & "'"
            
            strSQLCounts = " SELECT Count(*) AS Total FROM " & myTgtTable & ""
            Set rsc = dbCurrent.OpenRecordset(strSQLCounts)
            PostRecCount = rsc!Total
            Call AtivityLog(g_Runkey, "MV Table Inserts", mySrcTable, myTgtTable, strSQL2, PreRecCount, PostRecCount)
            
            Dim MyGUID As String
           '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            For j = 1 To intNewRecs
              'arrResults(j) = rsNonDups.Fields(0) 'get the guid from new Recordset
              'Debug.Print arrResults(j), j
              MyGUID = arrResults(j)
              Call MVUPDATEWorklist(g_Runkey, mySrcTable, MyGUID, myPkey, myTgtTable)
              strDelete = "DELETE * FROM " & myTgtTable & " WHERE " & myPkey & " = " & arrResults(j) & ""
              Call Save_Delete_Statements(g_Runkey, myTgtTable, strDelete)
            Next
        End If
       ' Get the Post Record Counts updated or not

       strSQLCounts = " SELECT Count(*) AS Total FROM " & myTgtTable & ""
       Set rsc = dbCurrent.OpenRecordset(strSQLCounts)
       PostRecCount = rsc!Total
       CurrentDb.Execute "UPDATE X_Mergetables SET PostRecCount = " & PostRecCount & " WHERE MergeTable = '" & myTgtTable & "'"
       
       'Debug.Print "Call MV_Field_Appends(" & runkey, mySrcTable, myTgtTable & ")"
       If RecCount > 0 And (AttachmentFlag = False) Then
          'Call DC_Print("Call MV_Field_Appends(" & runkey, mySrcTable, myTgtTable & ")")
          Call MV_Field_Appends(g_Runkey, mySrcTable, myTgtTable)
       End If
               
        Erase arrResults
        strSQL = ""
        strSQL2 = ""
        strInsert = ""
        strValues = ""
        strDelete = ""
        strSQLCounts = ""
        rs.MoveNext
    Wend
   
        
    dbCurrent.Close
    Set rs = Nothing
    Set rsNonDups = Nothing
    Set rsc = Nothing
    Set rsMVTableCols = Nothing
    
    
Debug.Print "MV Table Processing Completed"
Call DC_Print("MV Table Inserts and Appends Completed")
Call NonMV_Table_Inserts(g_Runkey)
Call Merge_Run_Log(g_Runkey, "Merge Successful")
Call Get_Merge_Counts_Validation(g_Runkey)
Call DC_Print("Merge Successful")
Debug.Print "Merge Successful"

g_Runkey = "" ' reset this global to ""

Exit Sub

tran_Err:
    
    Call DC_Print("Rolling Back Transaction " & Err.Description)
    Debug.Print strSQL2
    Call MergeLogError(g_Runkey, Err.Description, strSQL2, mySrcTable, myPkey, myTgtTable)
    Call Merge_Run_Log(g_Runkey, "Merge Failed")
    Call DC_Print("Merge Failed")
    Debug.Print "Merge Failed"
    
    DAO.DBEngine.Rollback

    Debug.Print "Transaction failed. Error: " & Err.Description
    
    dbCurrent.Close
    Set rs = Nothing
    Set rsNonDups = Nothing
    Set rsc = Nothing
    Set rsMVTableCols = Nothing
    
End Sub

Public Sub MV_Field_Appends(runkey As String, src As String, tgt As String)
'=====================================
' Name        : MV_Field_Appends
' Author      : David Campbell
' Copyright   : 2019
' Call command: Call MV_Field_Appends(runkey, mySrcTable, myTgtTable)
' Description :
'===================================
' Insert rows into T1 from T2 and then insert the Multiple entries into the
' MV field from T2 into T1 by ID field Using a SQL statement

Dim db As DAO.Database
Dim RecCount As Long
Dim IDGuid As Variant
'Dim Workstr As String
Dim arrInsertStrings(2500) As String
Dim MVGuids(2500) As Variant  ' Array

' Main Recordset Contains a Multi-Value Field
Dim rsWorklist As DAO.Recordset
Dim rsMVFieldList As DAO.Recordset

Dim rsSource As DAO.Recordset
Dim rsTarget As DAO.Recordset

' Now Define the Multi-Value Fields as a RecordSet
Dim rsProgramMultiValue1 As DAO.Recordset2


' The Values of the Field Are Contained in a Field Object
Dim fldProgramMultiValue1 As DAO.Field2

Dim i As Long
Dim j As Long
Dim myRSstr As String

Set db = CurrentDb()

myRSstr = "Select sourcekey, mykeyfield from X_MVWorklist where (RunID = " & runkey & ") AND (tgtTable = '" & tgt & "')"
Debug.Print myRSstr

Set rsWorklist = db.OpenRecordset(myRSstr) ' this could be multiple records
rsWorklist.MoveFirst


Set rsMVFieldList = db.OpenRecordset("Select ColumnName from X_mergetablecolumns where (MVColumn = YES) AND (TableName ='" & tgt & "')")
rsMVFieldList.MoveFirst
     
     
Dim ii As Integer
Dim myMVColumn As String
Dim myColcount As Integer
Dim myWorkGuid As String
Dim myWorkKeyFieldName As String

   myColcount = 0
   
Do Until rsWorklist.EOF
       myWorkGuid = rsWorklist!sourcekey.Value
       myWorkKeyFieldName = rsWorklist!mykeyfield.Value
       
       'Call DC_Print("Active Work Key: " & myWorkGuid)
       'Debug.Print "Active Work Key: " & myWorkGuid
       'Call DC_Print("Active Key Field: " & myWorkKeyFieldName)
       'Debug.Print "Active Key Field: " & myWorkKeyFieldName
       
       Do Until rsMVFieldList.EOF
          ' Print the name of the task to the Immediate window.
          myMVColumn = rsMVFieldList!ColumnName.Value
          
          'Debug.Print "Active MV Field: " & myMVColumn
          'Call DC_Print("Active MV Field: " & myMVColumn)
          
          myColcount = myColcount + 1
    
          Dim strtemp1 As String
          strtemp1 = "Select * from " & src & " WHERE " & myWorkKeyFieldName & " = " & myWorkGuid & ""
          'Debug.Print strtemp1
          Set rsSource = db.OpenRecordset(strtemp1)  ' An EXT_SITE Table
          
          Dim strtemp2 As String
          strtemp2 = "Select * from " & tgt & " WHERE " & myWorkKeyFieldName & " = " & myWorkGuid & ""
          'Debug.Print strtemp2
          Set rsTarget = db.OpenRecordset(strtemp2)  'A currentDB table
        
    
          'RecCount = DCount("*", src)
    
          'Debug.Print "Imported " & db.RecordsAffected & " of " & RecCount & " Records."
        
        ' Set The Multi-Value Field
        
         Set fldProgramMultiValue1 = rsSource(myMVColumn)
        ' Set fldProgramMultiValue2 = rsTarget(myMVColumn)
        
    
        rsSource.MoveFirst
        
        Dim rsString1 As String
        Dim rsString2 As String
        
        Do Until rsSource.EOF
           ' Print the name of the task to the Immediate window.
           'Debug.Print "ID: " & rsSource!ID.Value, "Site: " & rsSource!Site.Value
           
           'rstCrtDts.Fields("ShipDate").Value
           
           'Debug.Print "KEYFIELD: " & rsSource.Fields(myWorkKeyFieldName).Value, "KeyGuid: " & rsSource.Fields(myWorkGuid).Value
           
           ' This
           Set rsProgramMultiValue1 = rsSource.Fields(myMVColumn).Value
                Do Until rsProgramMultiValue1.EOF
                    rsProgramMultiValue1.MoveFirst
                    i = 0
                    j = 0
                          ' Loop through the records in the child recordset.
                          Do Until rsProgramMultiValue1.EOF
                              i = i + 1
                              'Debug.Print i & " MV:     ", rsProgramMultiValue1!Value.Value
                              MVGuids(i) = rsProgramMultiValue1!Value.Value ' iterate through and add guids to an array
                              rsProgramMultiValue1.MoveNext
                          Loop
                 Loop
             'Debug.Print "i value: " & i
             If Len(MVGuids(i)) > 0 Then
                For j = 1 To i
                   'Workstr = "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsTarget!ID.Value & ""
                   'Debug.Print Workstr
                   
                   Dim strTemp3 As String
                   'Call DC_Print(tgt)
                   'Call DC_Print(myMVColumn)
                   'Call DC_Print(myWorkKeyFieldName)
                   'Call DC_Print(myWorkGuid)
                   
                   'Debug.Print tgt
                   'Debug.Print myMVColumn
                   'Debug.Print myWorkKeyFieldName
                   'Debug.Print myWorkGuid
                   
                   strTemp3 = "INSERT INTO [" & tgt & "] ( [" & myMVColumn & "].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [" & myWorkKeyFieldName & "] = " & myWorkGuid & ""
                   
                   Debug.Print strTemp3
                   'Call DC_Print(strTemp3)
                   
                   db.Execute strTemp3
                   
                Next j
            End If
        rsSource.MoveNext
        Loop
Erase MVGuids
rsMVFieldList.MoveNext
Loop

rsMVFieldList.Requery
    rsMVFieldList.MoveFirst
rsWorklist.MoveNext
Loop

rsWorklist.Close
rsMVFieldList.Close

'rsProgramMultiValue1.Close

rsSource.Close
rsTarget.Close

Set rsWorklist = Nothing
Set rsMVFieldList = Nothing

Set rsSource = Nothing
Set rsTarget = Nothing

'Set rsProgramMultiValue1 = Nothing

End Sub

Public Sub AtivityLog(myrunid As String, RunLabel As String, srctable As String, tgttable As String, strSQL As String, precount As Integer, postcount As Integer)
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
   
   Set dbs = CurrentDb
  
   CurrentDb.Execute "INSERT INTO X_mergeactivity (RunID, runlabel, srcTable, tgttable, SQLStatement, precount, postcount) VALUES( """ & myrunid & """,""" & RunLabel & """,""" & srctable & """, """ & tgttable & """, """ & strSQL & """, " & precount & ", " & postcount & ")"
   Debug.Print "Activity Logged in X_mergeactivity"
   dbs.Close
 
   Set dbs = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

Public Sub MVUPDATEWorklist(myrunid As String, srctable As String, srckey As String, mykeyfield As String, tgttable As String)
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String
    
    ' NOTE: Sourcekey is also the tgtkey because we want to import the3 records as is including guids
    strSQL = "INSERT INTO X_MVWorklist(runid, srcTable, sourcekey, mykeyfield, tgtTable, targetkey) VALUES( """ & myrunid & """,""" & srctable & """, """ & srckey & """,""" & mykeyfield & """,""" & tgttable & """, """ & srckey & """)"
    'Debug.Print strSQL
    
   'Debug.Print "Worklist Item Created in X_MVWorklist"
   
   Set dbs = CurrentDb

   CurrentDb.Execute strSQL
      
   dbs.Close
    
   Set dbs = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

Public Sub Save_Delete_Statements(myrunid As String, tgttable As String, sqlCode As String)
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
   
   Set dbs = CurrentDb
  
   CurrentDb.Execute "INSERT INTO X_UnMerge_Code (RunID, TargetTable, sqlstatement) VALUES( """ & myrunid & """,""" & tgttable & """,""" & sqlCode & """)"
   'Debug.Print "Rollback SQL Statement Logged in X_UnMerge_Code"
   dbs.Close
 
   Set dbs = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

Public Sub MergeLogError(myrunid As String, Errormsg As String, sqlStatement As String, srctable As String, srckey As String, tgttable As String)
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
   
   Set dbs = CurrentDb

   CurrentDb.Execute "INSERT INTO X_MergeErrorLog (RunID, ErrorMsg, sqlStatement, srcTable, srckey, tgttable) VALUES( """ & myrunid & """,""" & Errormsg & """,""" & sqlStatement & """,""" & srctable & """,""" & srckey & """,""" & tgttable & """)"
   Debug.Print "Error Logged in X_MergeErrorLog!"
   
   dbs.Close
 
   Set dbs = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub

Sub ClearWorkList()

Dim dbs As Database
Dim rs As Recordset
Dim myTable As String
Dim ctlListbox As control

    Set dbs = CurrentDb
    
    'dbs.Execute "insert into X_MVWorklistArchive (runid, srctable, sourcekey, tgttable, createdon) Select RunID, srctable, sourcekey, tgttable, createdon from X_MVWorklist", Options:=dbFailOnError + dbSeeChanges
    
    dbs.Execute "insert into X_MVWorklistArchive Select * from X_MVWorklist", Options:=dbFailOnError + dbSeeChanges
    
    Debug.Print "Copied Records to X_MVWorklistArchive"

    dbs.Execute "delete * from X_MVWorklist", Options:=dbFailOnError + dbSeeChanges
    Debug.Print "X_MVWorklist prepared for next run"
    
    Set dbs = CurrentDb
       
    dbs.Close
    Set rs = Nothing
End Sub

Public Sub Subtester()

 'Call Get_Merge_Counts_Validation("{00253F02-A5DC-48D6-AAF2-E61B8216051F}")

End Sub

Public Sub DC_Print(myText As String)
Dim strMsg As String

strMsg = myText
Forms!frmDbAdministrationDash!sbfrmDbMerge.Form.UpdateProgress strMsg
Forms!frmDbAdministrationDash!sbfrmDbMerge.Form.Refresh

End Sub

Sub Get_Merge_Counts_Validation(Optional ByVal Passed_Runid As String)
' For all tables

Dim dbs As Database
Dim rs As Recordset
Dim myTable As String
Dim tblCount As Integer
Dim ii As Integer
Dim ActivityPostCount As Integer
Dim ActualTableCount As Integer
Dim myrunlabel As String
Dim CountMessage As String
Dim tableLen As Integer
Dim Labellen As Integer

Dim NameSpacelen As Integer
Dim LabelSpaceLen As Integer

Dim j, k As Integer
Dim NameSpacer As String
Dim LabelSpacer As String

    Call DC_Print("Validating Merge Work RowCounts...standby")
    
    Set dbs = CurrentDb
    
    If Passed_Runid = "" Then
        Set rs = dbs.OpenRecordset("SELECT runlabel,tgtTable,postcount FROM X_mergeactivity WHERE RunID = " & g_Runkey & " ORDER BY tgttable")
    Else
        Set rs = dbs.OpenRecordset("SELECT runlabel,tgtTable,postcount FROM X_mergeactivity WHERE RunID = " & Passed_Runid & " ORDER BY tgttable")
    End If
    
    'RecCount = FindRecordCount(strSQL)
    tblCount = 0
    tableLen = 0
    Labellen = 0
    j = 0
    k = 0
    
    While Not rs.EOF
        tblCount = tblCount + 1
        myrunlabel = rs.Fields("runlabel")
        myTable = rs.Fields("tgtTable")
        ActivityPostCount = rs.Fields("postcount")
        ActualTableCount = FindRecordCount("SELECT * FROM " & myTable & "")
        tableLen = Len(myTable)
        Labellen = Len(myrunlabel)
        NameSpacelen = 30 - tableLen
        LabelSpaceLen = 35 - Labellen
        
        NameSpacer = ""
        LabelSpacer = ""
        
        For j = 1 To NameSpacelen
           NameSpacer = NameSpacer + " "
        Next j
        
        For k = 1 To LabelSpaceLen
           LabelSpacer = LabelSpacer + " "
        Next k
        
        If ActivityPostCount = ActualTableCount Then
           CountMessage = "Counts Match!"
           Debug.Print myrunlabel & LabelSpacer & " : " & myTable & NameSpacer & " : Activity Count Says " & ActivityPostCount & ", Table Count Says : " & ActualTableCount & " " & CountMessage & ""
        Else
           CountMessage = "Counts Do Not Match!"
           If myrunlabel <> "NonMV Table Inserts" And CountMessage = "Counts Do Not Match!" Then
              Debug.Print myrunlabel & LabelSpacer & " : " & myTable & NameSpacer & " : Activity Count Says " & ActivityPostCount & ", Table Count Says : " & ActualTableCount & " " & CountMessage & ""
           End If
        End If

        rs.MoveNext
     Wend
        
    rs.Close
    Set rs = Nothing

    'Debug.Print "Loaded Merge Column Names: " & colcount & " Tables Columns Added to List from " & tblCount & " Tables"
    'Call DC_Print("Loaded Merge Column Names: " & colcount & " Tables Columns Added to List from " & tblCount & " Tables")
    'Call DC_Print("Completed Load Of Column names.")
    dbs.Close
    Call DC_Print("Validating Merge Work RowCounts Completed")
    
End Sub

Public Function PrimKey(tblName As String)
'*******************************************
'Purpose: Programatically determine a
' table's primary key
'Coded by: raskew
'Inputs: from Northwind's debug window:
' Call PrimKey("Products")
'Output: "ProductID"
'*******************************************

Dim db As Database
Dim td As TableDef
Dim idxLoop As index

Set db = CurrentDb
Set td = db.TableDefs(tblName)
For Each idxLoop In td.Indexes
If idxLoop.Primary = True Then
'Debug.Print Mid(idxLoop.Fields, 2)
Exit For
End If
Next idxLoop


PrimKey = MID(idxLoop.Fields, 2)
db.Close
Set db = Nothing
End Function

Function FindRecordCount(strSQL As String) As Long
 
Dim dbs As DAO.Database
Dim rstRecords As DAO.Recordset
 
On Error GoTo ErrorHandler
 
   Set dbs = CurrentDb
 
   Set rstRecords = dbs.OpenRecordset(strSQL)
 
   If rstRecords.EOF Then
      FindRecordCount = 0
   Else
      rstRecords.MoveLast
      FindRecordCount = rstRecords.RecordCount
   End If
 
   rstRecords.Close
   dbs.Close
 
   Set rstRecords = Nothing
   Set dbs = Nothing
 
Exit Function
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Sub Merge_Run_Log(myrunid As String, MyDesc As String)
Dim dbs As DAO.Database
Dim rs As DAO.Recordset
   
   Set dbs = CurrentDb
  
   CurrentDb.Execute "INSERT INTO X_Merge_Run_Log (RunID, description) VALUES( """ & myrunid & """,""" & MyDesc & """)"
   Debug.Print "Activity Logged in X_Merge_Run_Log"
   dbs.Close
 
   Set dbs = Nothing
 
Exit Sub
 
ErrorHandler:
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
End Sub
