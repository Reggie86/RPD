Attribute VB_Name = "ModuleDC"
Option Compare Database
Option Explicit

'==========================================================================================================================
'ModuleDC
'08/09/2019
'Version 1.000
'==========================================================================================================================
Sub NewGuidBankEntries()
    Dim j As Integer
    Dim db  As DAO.Database
   
   Set db = CurrentDb()
   
For j = 1 To 20000

    db.Execute "INSERT INTO X_Guidbank(used) VALUES (NO)"
Next j

End Sub

Public Function GetGUID() As String
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim MyGUID As String
  Dim mygkey As String
  Dim strSQL As String
  
  Set db = CurrentDb
  Set rs = CurrentDb.OpenRecordset(" Select  TOP 1 ID, gkey from X_GuidBank")
  MyGUID = rs.Fields(0)
  mygkey = rs.Fields(1)
  
 
  strSQL = "Delete * FROM X_GuidBank WHERE gkey = " & mygkey & ""
  CurrentDb.Execute strSQL
  'Debug.Print strSQL
  'Debug.Print "This is a guid: """ & Mid$(myGuid, 7, 38) & """"
  'Debug.Print mygkey
  
  rs.Close
  Set rs = Nothing
  
  
  
  GetGUID = MID$(MyGUID, 7, 38)
  
End Function

Function selectFile()
Dim fd As FileDialog, fileName As String

On Error GoTo ErrorHandler

Set fd = Application.FileDialog(msoFileDialogFilePicker)

fd.AllowMultiSelect = False

If fd.Show = True Then
    If fd.SelectedItems(1) <> vbNullString Then
        fileName = fd.SelectedItems(1)
    End If
Else
    'Exit code if no file is selected
    End
End If

'Return Selected FileName
selectFile = fileName

Set fd = Nothing

Exit Function

ErrorHandler:
Set fd = Nothing
MsgBox "Error " & Err & ": " & Error(Err)

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
'
'Function FindRowCount(strSQL As String) As Long
'
'Dim dbs As DAO.Database
'Dim rstRecords As DAO.Recordset
'
'On Error GoTo ErrorHandler
'
'   Set dbs = CurrentDb
'
'   Set rstRecords = dbs.OpenRecordset(strSQL)
'
'   If rstRecords.EOF Then
'      FindRowCount = 0
'   Else
'      rstRecords.MoveLast
'      FindRowCount = rstRecords.RecordCount
'   End If
'
'   rstRecords.Close
'   dbs.Close
'
'   Set rstRecords = Nothing
'   Set dbs = Nothing
'
'Exit Function
'
'ErrorHandler:
'   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description
'End Function


'Sub Get_Table_Relationships()
'
'Dim dbs As Database
'Dim rs As Recordset
'
'    Set dbs = CurrentDb
'
'    'Clear Table First
'    dbs.Execute "delete * from X_MergeMap", Options:=dbFailOnError + dbSeeChanges
'
'    'Create List
'    CurrentDb.Execute " INSERT into X_MergeMap (szColumn, szObject, szReferencedColumn, szReferencedObject) SELECT szColumn, szObject, szReferencedColumn, szReferencedObject FROM [MSysRelationships] WHERE (szObject NOT Like 'MSysNavPaneGroup*') AND (szObject NOT LIKE 'tbluser') ORDER BY szObject, szReferencedObject;"
'
'    Debug.Print "Table Relationships Collected and Stored"
'
'    dbs.Close
'    Set rs = Nothing
'End Sub

'
'Sub ClearTestTable()
'
'Dim dbs As Database
'Dim rs As Recordset
'Dim myTable As String
'Dim ctlListbox As control
'
'    Set dbs = CurrentDb
'
'    dbs.Execute "delete * from TEST_Table", Options:=dbFailOnError + dbSeeChanges
'
'    Set dbs = CurrentDb
'
'    Call ModuleDC.RequeryList
'
'    dbs.Close
'    Set rs = Nothing
'End Sub


'Sub BrowseMultiValueField()
'
'    Dim db  As DAO.Database
'    Dim tdf As DAO.TableDef
'
'    Dim rs As Recordset
'    Dim childRS As Recordset
'
'   'Manipulate multivalued fields with DAO
'
'   Set db = CurrentDb()
'
'   ' Open a Recordset for the Tasks table.
'   Set rs = db.OpenRecordset("Select * from AAA_MV_Test_2 where Site = 'CSTARS Baltimore'")
'   rs.MoveFirst
'
'   Do Until rs.EOF
'      ' Print the name of the task to the Immediate window.
'      Debug.Print rs!ID.Value, rs!Site.Value
'
'      ' Open a Recordset for the multivalued field.
'
'      Set childRS = rs!MV.Value
'
'         ' Exit the loop if the multivalued field contains no records.
'         Do Until childRS.EOF
'             childRS.MoveFirst
'
'             ' Loop through the records in the child recordset.
'             Do Until childRS.EOF
'                 ' Print the owner(s) of the task to the Immediate
'                 ' window.
'                 Debug.Print "      ", childRS!Value.Value
'                 childRS.MoveNext
'             Loop
'         Loop
'      rs.MoveNext
'   Loop
'End Sub

'Sub PrintsMultiValueField()
'
'    Dim db  As DAO.Database
'    Dim rs As Recordset
'    Dim childRS As DAO.Recordset2
'
'   'Manipulate multivalued fields with DAO
'
'   Set db = CurrentDb()
'
'   ' Open a Recordset for the Tasks table.
'   Set rs = db.OpenRecordset("Select * from AAA_MV_Test_2 where Site = 'CSTARS Baltimore'")
'   rs.MoveFirst
'
'   Do Until rs.EOF
'      ' Print the name of the task to the Immediate window.
'      Debug.Print rs!ID.Value, rs!Site.Value
'
'      ' Open a Recordset for the multivalued field.
'
'      Set childRS = rs!MV.Value
'
'         ' Exit the loop if the multivalued field contains no records.
'         Do Until childRS.EOF
'             childRS.MoveFirst
'
'             ' Loop through the records in the child recordset.
'             Do Until childRS.EOF
'                 ' Print the owner(s) of the task to the Immediate
'                 ' window.
'                 Debug.Print "      ", childRS!Value.Value
'                 childRS.MoveNext
'             Loop
'         Loop
'      rs.MoveNext
'   Loop
'End Sub

'Sub Test_DAO_Insert()
'Dim db As DAO.Database, RecCount As Long
'Dim rsComplex As DAO.Recordset2
'
''StandardProperties ("AAA_MV_Test_2")
'
'
''Get the total number of records in your import table to compare later
'RecCount = DCount("*", "AAA_MV_Test_2")
'
''This line is IMPORTANT! each time you call CurrentDb a new db object is returned
''  that would cause problems for us later
'Set db = CurrentDb
'
'
''Add the records, being sure to use our db object, not CurrentDb
'db.Execute "INSERT INTO AAA_MV_Test_1 (ID, Site) " & _
'           "SELECT ID, Site " & _
'           "FROM AAA_MV_Test_2 " & _
'           "WHERE Site = 'CSTARS Baltimore'", dbFailOnError
'
''db.RecordsAffected now contains the number of records that were inserted above
''  since CurrentDb returns a new db object, CurrentDb.RecordsAffected always = 0
''If RecCount = db.RecordsAffected Then
'    Debug.Print "Imported " & db.RecordsAffected & " of " & RecCount & " Records."
'    'db.Execute "DELETE * FROM TBL_ImportTable", dbFailOnError
''End If
'End Sub
'
'Public Sub InsertIntoMultiValueField()
'' Insert rows into T1 from T2 and then insert the Multiple entries into the
'' MV field from T2 into T1 by ID field Using a SQL statement
'
'Dim db As DAO.Database, RecCount As Long, IDGuid As Variant
'Dim Workstr As String
'Dim MVGuids(30) As Variant  ' Array
'
'
'' Main Recordset Contains a Multi-Value Field
'Dim rsMVT1 As DAO.Recordset
'Dim rsMVT2 As DAO.Recordset
'
'' Now Define the Multi-Value Fields as a RecordSet
'Dim rsProgramMultiValue1 As DAO.Recordset2
'Dim rsProgramMultiValue2 As DAO.Recordset2
'
'' The Values of the Field Are Contained in a Field Object
'Dim fldProgramMultiValue1 As DAO.Field2
'Dim fldProgramMultiValue2 As DAO.Field2
'Dim i As Long
'Dim j As Long
'
'
'' Open the Parent File
'
'Set db = CurrentDb()
'Set rsMVT1 = db.OpenRecordset("Select * from AAA_MV_Test_1")
'Set rsMVT2 = db.OpenRecordset("Select * from AAA_MV_Test_2 where Site = 'CSTARS Baltimore'")
'
'RecCount = DCount("*", "AAA_MV_Test_2")
'
''Add the records, being sure to use our db object, not CurrentDb
'db.Execute "INSERT INTO AAA_MV_Test_1 (ID, Site) " & _
'           "SELECT ID, Site " & _
'           "FROM AAA_MV_Test_2 " & _
'           "WHERE Site = 'CSTARS Baltimore'", dbFailOnError
'
'Debug.Print "Imported " & db.RecordsAffected & " of " & RecCount & " Records."
'
'' Set The Multi-Value Field
'
'Set fldProgramMultiValue1 = rsMVT1("MV")
'Set fldProgramMultiValue2 = rsMVT2("MV")
'
'
'   rsMVT2.MoveFirst
'
'   Do Until rsMVT2.EOF
'      ' Print the name of the task to the Immediate window.
'      Debug.Print "ID: " & rsMVT2!ID.Value, "Site: " & rsMVT2!Site.Value
'
'      Set rsProgramMultiValue2 = rsMVT2!MV.Value
'
'         Do Until rsProgramMultiValue2.EOF
'             rsProgramMultiValue2.MoveFirst
'             i = 0
'             j = 0
'             ' Loop through the records in the child recordset.
'             Do Until rsProgramMultiValue2.EOF
'                 i = i + 1
'                 Debug.Print i & " MV:     ", rsProgramMultiValue2!Value.Value
'                 MVGuids(i) = rsProgramMultiValue2!Value.Value ' iterate through and add guids to an array
'                 rsProgramMultiValue2.MoveNext
'             Loop
'         Loop
'         Debug.Print i & "i value"
'         For j = 1 To i
'            'Workstr = "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsMVT2!ID.Value & ""
'            'Debug.Print Workstr
'
'            db.Execute "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsMVT2!ID.Value & ""
'
'         Next j
'
'   rsMVT2.MoveNext
'   Loop
'
'rsMVT1.Close
'rsMVT2.Close
'Set rsMVT1 = Nothing
'Set rsMVT2 = Nothing
'
'' Now Define the Multi-Value Fields as a RecordSet
'
'Set rsProgramMultiValue1 = Nothing
'Set rsProgramMultiValue2 = Nothing
'
'End Sub

'Public Sub MergeMaster()
'' Insert rows into T1 from T2 and then insert the Multiple entries into the
'' MV field from T2 into T1 by ID field Using a SQL statement
'
'Dim db As DAO.Database, RecCount As Long, IDGuid As Variant
''Dim Workstr As String
'Dim arrInsertStrings() As String
'
'' Main Recordset Contains a Multi-Value Field
'Dim rsTablelist As DAO.Recordset
'Dim rsFieldlist As DAO.Recordset
'Dim rsSource As DAO.Recordset
'Dim rsTarget As DAO.Recordset
'
'' Now Define the Multi-Value Fields as a RecordSet
'Dim rsProgramMultiValue1 As DAO.Recordset2
'Dim rsProgramMultiValue2 As DAO.Recordset2
'
'' The Values of the Field Are Contained in a Field Object
'Dim fldProgramMultiValue1 As DAO.Field2
'Dim fldProgramMultiValue2 As DAO.Field2
'Dim i As Long
'Dim j As Long
'
'Set db = CurrentDb()
'
''Freshen the table list
''Call ModuleDC.Load_Merge_Tables
''Call ModuleDC.Prep_Merge_Tables
''Call ModuleDC.Load_Z_Tables
'
'' Open the Parent File
'
'
'Set rsTablelist = db.OpenRecordset("Select MergeTable from X_mergetable where flags=0")
'rsTablelist.MoveFirst
'
'Dim ii As Integer
'Dim myColumn As String
'Dim mytblcount As Integer
'Dim myInsert As String
'myInsert = "INSERT into " & rsTablelist!ID.MergeTable & "( ["
'Dim mycolumnlist As String
'mycolumnlist = ""
'
'   mytblcount = 0
'   Do Until rsTablelist.EOF
'      ' Print the name of the task to the Immediate window.
'      Debug.Print "ID: " & rsTablelist!ID.MergeTable
'      mytblcount = mytblcount + 1
'      For ii = 0 To rsTablelist.Fields.count - 1
'          myColumn = rsTablelist.Fields(ii).Name
'          mycolumnlist = mycolumnlist & "[" & myColumn & "],"
'          Debug.Print myColumn
'      Next ii
'      myInsert = myInsert & mycolumnlist & " )"
'
'
''Workstr = "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsMVT2!ID.Value & ""
'
'
''Set rsFieldlist = db.OpenRecordset("Select * from AAA_MV_Test_1")
'Set rsSource = db.OpenRecordset("Select * from AAA_MV_Test_1")
'Set rsTarget = db.OpenRecordset("Select * from AAA_MV_Test_2 where Site = 'CSTARS Baltimore'")
'
'RecCount = DCount("*", "AAA_MV_Test_2")
'
''Add the records, being sure to use our db object, not CurrentDb
'
'
'
'
''db.Execute "INSERT INTO AAA_MV_Test_1 (ID, Site) " & _
''           "SELECT ID, Site " & _
''           "FROM AAA_MV_Test_2 " & _
''           "WHERE Site = 'CSTARS Baltimore'", dbFailOnError
'
'
'
'
'
'Debug.Print "Imported " & db.RecordsAffected & " of " & RecCount & " Records."
'
'' Set The Multi-Value Field
'
'Set fldProgramMultiValue1 = rsSource("MV")
'Set fldProgramMultiValue2 = rsTarget("MV")
'
'' Check to Make Sure it is Multi-Value
'
''If Not (fldDBAInStatesMultiValue.IsComplex) Then
'    'MsgBox ("Not A Multi-Value Field")
'    'rsBusiness.Close
'    'Set rsBusiness = Nothing
'    'Set fldDBAInStatesMultiValue = Nothing
'    'Exit Function
''End If
''On Error Resume Next
'
'' Loop Through
'
'   rsTarget.MoveFirst
'
'   Do Until rsTarget.EOF
'      ' Print the name of the task to the Immediate window.
'      Debug.Print "ID: " & rsTarget!ID.Value, "Site: " & rsTarget!Site.Value
'
'      Set rsProgramMultiValue2 = rsTarget!MV.Value
'
'         Do Until rsProgramMultiValue2.EOF
'             rsProgramMultiValue2.MoveFirst
'             i = 0
'             j = 0
'             ' Loop through the records in the child recordset.
'             Do Until rsProgramMultiValue2.EOF
'                 i = i + 1
'                 Debug.Print i & " MV:     ", rsProgramMultiValue2!Value.Value
'                 MVGuids(i) = rsProgramMultiValue2!Value.Value ' iterate through and add guids to an array
'                 rsProgramMultiValue2.MoveNext
'             Loop
'         Loop
'         Debug.Print i & "i value"
'         For j = 1 To i
'            'Workstr = "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsTarget!ID.Value & ""
'            'Debug.Print Workstr
'
'            db.Execute "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsTarget!ID.Value & ""
'
'         Next j
'
'   rsMVT2.MoveNext
'   Loop
'
'rsSource.Close
'rsTarget.Close
'Set rsSource = Nothing
'Set rsTarget = Nothing
'
'' Now Define the Multi-Value Fields as a RecordSet
'
'Set rsProgramMultiValue1 = Nothing
'Set rsProgramMultiValue2 = Nothing
'
'End Sub

'Public Sub Parse_Column_Names()
'Dim db As DAO.Database
'Dim arrInsertStrings() As String
'
'Dim rsTablelist As DAO.Recordset
'Dim rsFieldlist As DAO.Recordset
'
'Set db = CurrentDb()
'
''Freshen the table list
'Call ModuleDC.Load_Merge_Tables
'
'Set rsTablelist = dbs.OpenRecordset("Select MergeTable from X_mergetable where flags=0")
'rsTablelist.MoveFirst
'
'Dim ii As Integer
'Dim myColumn As String
'Dim mytblcount As Integer
'Dim myInsert As String
'Set myInsert = "INSERT into " & rsTablelist!ID.MergeTable & "( "
'Dim mycolumnlist As String
'Set mycolumnlist = Nothing
'Dim myValues As String
'Set myValues = "VALUES ("
'
'   mytblcount = 0
'   Do Until rsTablelist.EOF
'      ' Print the name of the task to the Immediate window.
'      Debug.Print "ID: " & rsTablelist!ID.MergeTable
'      mytblcount = mytblcount + 1
'      For ii = 0 To rsTablelist.Fields.count - 1
'          myColumn = rs.Fields(ii).Name
'          mycolumnlist = mycolumnlist & myColumn & ", "
'          myValues
'          Debug.Print myColumn
'      Next ii
'      myInsert = myInsert & mycolumnlist & " )"
'
''Workstr = "INSERT INTO [AAA_MV_Test_1] ( [MV].[Value] ) VALUES ( " & MVGuids(j) & ") WHERE [ID] = " & rsMVT2!ID.Value & ""
''    CurrentDb.Execute " INSERT into X_MergeMap (szColumn, szObject, szReferencedColumn, szReferencedObject) SELECT szColumn, szObject, szReferencedColumn, szReferencedObject FROM [MSysRelationships] WHERE (szObject NOT Like 'MSysNavPaneGroup*') AND (szObject NOT LIKE 'tbluser') ORDER BY szObject, szReferencedObject;"
'
'rsTablelist.Close
'rsFieldlist.Close
'Set rsTablelist = Nothing
'Set rsFieldlist = Nothing
'
'
'
'End Sub

