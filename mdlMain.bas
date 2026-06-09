Attribute VB_Name = "mdlMain"
Option Explicit


'-----------------------------------------------------------
' iLink
'
' Purpose: The purpose of iLink is to collect pen level failure information from
'          external sources and to calculate an overall failcode for a pen.
'          It then updates Nextcap with an overall failcode for the pen.  for further details consult the iLink documntation.
'
' Author:  Mark Davis, DIMO
'
' Revision History:
'          10 May  02 - TESTED - Initial Version for use with inspector
'          15 June 02 - TESTED - Improved logging, various bug fixes
'          20 June 02 - TESTED - Added new business rules for classifying failures
'          22 June 02 - TESTED - Made into a service so that it was less obvious to the user
'          24 June 02 - TESTED - Added some bug fixes
'          26 July 02 - TESTED - Added support for Wildflowers
'          26 Nove 03 - TESTED - Added support for multiple test types (major upgrade for magma)
'          14 Apri 04 - TESTED - Added performance improvements
'          10 Sept 04 - TESTED - Added ability to split ini file into 2 - global/local
'          01 Octo 04 -  - Added option to recalculate the failcodes for pens in open lots
'-----------------------------------------------------------
'Global constants
Public Const CN_ILINK_COMMENT = "Inspector Automated Operator"

' Global Types
Public Type ScannerBank
    Path As String
    Alias As String
    Active As Boolean
End Type

Public Type CapDecisionRule
    FailCodes As String
    SetFailCode As String
End Type

Public Type TestResult
    TestType As String          'The type of the test - e.g. "Inspector", "E-Test"
    PrintSampleID As String     'The ID of the print sample in the case of Inspector
    PenID As String             'Pen ID
    Comment As String           'The Lot ID or comment
    Failcode As String          'The failcode returned by this result
    ScanDatetime As Date        'The date according to the test tool
    PageNumber As Integer       'The page number in the case of INSPECTOR
    DataFile As String          'The name of the File in the pase of Inspector
End Type

Public Type TestLevel
    TestType As String
    FailcodePriorities() As String
End Type

Public Type PenResult
    PenID As String
    OverallFailcode As String
End Type

'Global Settings Variables - populated from ini file
Private m_arrCapDecisionRules() As CapDecisionRule
Private m_arrTestLevels() As TestLevel
Private m_arrPriorityFailures() As String 'obsoleting this
Private m_fileExtension As String
Private m_codesToIgnore As String
Private m_numIPASColumns As Integer
Private m_numDaysData As Integer
Private m_numCapRetryAttempts As Integer
Private m_doubleCheckNextcap As Boolean
Private m_numdaysLogFile As Integer
Private m_fileStyle As String
Private m_MfgDBType As String
Private m_Debug As Boolean

'Connection Strings
Private m_ILinkConnectionString As String
Public g_NextcapConnectionString As String
Private m_GradebookConnectionString As String

'Etester Settings - populated from ini file
Private m_EtesterSQL As String
Private m_CheckEtestGen1 As Boolean
Private m_CheckEtestGen2 As Boolean

'Kahuna Settings - populated from ini file
Private m_KahunaSQL As String

'Test_types to check for
Private m_CheckForInspectorResults As Boolean
Private m_CheckForEtesterResults As Boolean
Private m_CheckForEtesterRegionFails As Boolean
Private m_CheckForKahunaRegionFails As Boolean
Private m_RegionCheckEngineeringLots As Boolean

Public g_arrScanners() As ScannerBank
Public g_loadToNextcap As Boolean
Public g_refreshInterval As Integer
Public g_logEventsToFile As Boolean
Public g_capDefectLevel As Integer
Public g_goodPenCode As String

'Other local variables
Private m_sOpenNextcapPens As String
Private m_sExistingEtestStamps As String
Private m_dctProdRefIDFetMktg As New Dictionary
Private m_dctProdRefIDFetProd As New Dictionary
Private m_dctProdRefVentLab As New Dictionary

Private m_dctPenIDRuntype As Dictionary
Private m_dctPenIDProductNumber As Dictionary

Private m_bShareBanks As Boolean


'Global variables - db connections, etc.
Public g_cnxDb As ADODB.Connection
Public g_lastDay As Integer

'-----------------------------------------------------------
' Main Subroutine
'
' Various setup events when application is started..
'-----------------------------------------------------------
Sub Main()
On Error GoTo EH
   'mark todays date
    g_lastDay = Day(Now())
    
    'Mark the logfile to say we've started up.
    LogEvent "---------    STARTING iLINK   ---------------"
    
   'freeze the user controls
   ToggleControls (False)

   'Load the form
   frmMain.Show
   
   'read settings from ini file
   Call LoadSettings
         
   'Initialise the scanner shares
   Call InitialiseScanners
   
   'initialise CAP server if necessary/possible
    If g_loadToNextcap Then
        LogEvent "Connecting to Nextcap Server..."
        If SetupNextCap() = False Then
            'Don't bother trying nextcap loads if I can't connect
            g_loadToNextcap = False
        Else
            LogEvent "Connected to Nextcap Successfully!"
        End If
    End If
      
   'miscellaneous setups
   If g_logEventsToFile = True Then
        frmMain.mnuLogResults.Checked = True
        frmMain.chkLogEvents.Value = vbChecked
    Else
        frmMain.mnuLogResults.Checked = False
        frmMain.chkLogEvents.Value = vbUnchecked
   End If
   
   frmMain.tmrMain.Interval = g_refreshInterval
   
   'unfreeze the user controls
   ToggleControls (True)
   LogEvent "Ready!"

   Exit Sub
EH:
    LogEvent "** ERROR : main - " & Err.Number & " - " & Err.Description, True
End Sub


Public Sub CheckForNewData()
    '-----------------------------------------------------------
    ' CheckForNewData Subroutine
    '
    ' This will run when an interval triggers and do the following
    ' 1. Check for any previously failed uploads & attempt to upload again
    ' 2. Check for new scanner data
    ' 3. Insert it into the iLink database
    ' 4. Make any necessary updates to Nextcap
    '-----------------------------------------------------------
    On Error GoTo EH
    
    frmMain.tmrMain.Enabled = False
    ToggleControls (False)
       
    'backup & clear the logfiles if necessary
    DailyMaintenance
   
    LogEvent ("Checking for new Data...")
    
    'connect to db for next steps
    If dbConnect() Then
        '1. Try for previously failed updates
        checkForCapLoadFailures
        
        '2. If necessary compile a list of pens to check for
        If m_bShareBanks Or m_CheckForEtesterResults Or m_CheckForEtesterRegionFails Or m_CheckForKahunaRegionFails Then
            m_sOpenNextcapPens = getOpenNextcapPens()
        End If
    
        If m_CheckForInspectorResults Then
            'Read results from Inspector if required
            Call checkForInspectorResults
        End If
        
        If m_CheckForEtesterResults Then
            'Read results from ETester if required
            If m_CheckEtestGen1 Then Call checkForEtestResults("etestgen1_d")
            If m_CheckEtestGen2 Then Call checkForEtestResults("etestgen2_d")
        End If
        
        If m_CheckForKahunaRegionFails Then
            'Read results from Kahuna if required
            Call checkForKahunaResults
        End If
                            
        ' don't need db connection any more
        Call dbClose
    Else
        LogEvent "**ERROR! Can't Connect to Database", True
    End If
    
    frmMain.tmrMain.Enabled = True
    ToggleControls (True)
    LogEvent ("Ready!")
    Exit Sub
EH:
    LogEvent "** ERROR : CheckForNewData - " & Err.Number & " - " & Err.Description, True
End Sub


Private Sub checkForInspectorResults()
'-------------------------------------------------------------------
' Load results from Inspector
'-------------------------------------------------------------------
On Error GoTo EH
    Dim TestResults() As TestResult
    Dim NewPens() As String
    Dim PenResults() As PenResult
    Dim iNumResults As Integer
    LogEvent ("Checking for INSPECTOR Results... ")

    iNumResults = ReadDataFromBanks(TestResults)
        
    If iNumResults > 0 Then 'can quit now if there are no new results
        NewPens = getUniquePens(TestResults)
       
        '1. Insert the new data into the finishedData table
        If addTestResults(TestResults) Then
        Call DeleteDataFiles(TestResults)
        
        '2. Determine and set the overall failmode for the pens
        PenResults = setOverallFailmode(NewPens)
                   
        '3. Load results to Nextcap if necessary
        Call LoadResultsToNextCAP(PenResults)
        End If
    End If
    
    Exit Sub
EH:
    LogEvent ("** ERROR : checkForInspectorResults - " & Err.Number & " - " & Err.Description)
End Sub

Private Sub checkForEtestResults(sEtestTable As String)
'-------------------------------------------------------------------
' Load results from the Etester
'
' Note: In order for this to work the view vw_etest_genx needs to be created
'       at the target DB.
'
' Operation:
'       1. Connect to the gradebook database
'       2. Query all new results lots since the last update
'       3. Apply changes to pens in open lots
'-------------------------------------------------------------------
On Error GoTo EH
    LogEvent ("Checking for " & sEtestTable & " ETESTER Results... ")
    Dim TestResults() As TestResult
    Dim NewPens() As String
    Dim PenResults() As PenResult
    Dim iNumRelevantResults As Long
    Dim sFailcode As String
    Dim sPreviousDateTime As String
    Dim sDateRegistryKey As String
    Dim sLatestDateTime As String
    Dim sSQL As String
    Dim sExistingStamps As String
    Dim iNumRecords As Long
    
    'If we don't have any mids open exit the function without checking
    If Len(m_sOpenNextcapPens) < 16 Then
        LogEvent "No Pens in Open Lots.  Etest Search Skipped."
        Exit Sub
    End If
    
    'Look for last date stamp
    sDateRegistryKey = "Etest" & sEtestTable
    sPreviousDateTime = GetSetting("iLink", "QueryDates", sDateRegistryKey)
    
    ' The SQL depends on the m_MfgDBType
    
    If sPreviousDateTime = "" Then
        sExistingStamps = getExistingStamps("ETESTER", m_sOpenNextcapPens)
        If m_MfgDBType = "SQL-SERVER" Then
            sPreviousDateTime = "dateadd(day, -20, getdate()) AND process_id in (" & m_sOpenNextcapPens & ") AND d.stamp_link not in (" & sExistingStamps & ")"
        Else
            sPreviousDateTime = "TODAY - 20 AND process_id in (" & m_sOpenNextcapPens & ") AND d.stamp_link not in (" & sExistingStamps & ")"
        End If
    Else
        sPreviousDateTime = "'" & sPreviousDateTime & "'"
    End If
 
    'build the SQL String
    If m_MfgDBType = "SQL-SERVER" Then
        sSQL = "select d.stamp_link, l.insert_dttm, l.insert_dttm + '' txt_dttm, d.process_id, d.pass_fail_cd, d.id_fet, cast(year(date_test) as varchar) + '-' + cast(month(date_test)as varchar)  + '-' + cast(day(date_test)as varchar) + ' ' + d.time_test etest_dm etest_dm " & _
            " from insert_time_log l, @TABLE@ d where @DATEFILTER@ AND @INIFILTER@ and l.unique_id=d.stamp_link order by insert_dttm"
    Else
          sSQL = "select d.stamp_link, l.insert_dttm, l.insert_dttm || '' txt_dttm, d.process_id, d.pass_fail_cd, d.id_fet, extend (d.date_test, year to day) || ' ' || d.time_test etest_dm " & _
            " from insert_time_log l, @TABLE@ d where @DATEFILTER@ AND @INIFILTER@ and l.unique_id=d.stamp_link order by insert_dttm"
    End If
      
     
    sSQL = Replace(sSQL, "@DATEFILTER@", "l.insert_dttm > " & sPreviousDateTime)
     
    If Trim(m_EtesterSQL) = "" Then
        sSQL = Replace(sSQL, "@INIFILTER@", "1=1")
    Else
        sSQL = Replace(sSQL, "@INIFILTER@", m_EtesterSQL)
    End If
    
    sSQL = Replace(sSQL, "@TABLE@", sEtestTable)
     
    Dim cnxDb As ADODB.Connection
    Dim rs As ADODB.Recordset
     
    'Log the details if in debug mode
    If m_Debug = True Then
        Call LogEvent(sSQL)
    End If
     
    '1. Connect to the database
    Set cnxDb = New ADODB.Connection
    cnxDb.Mode = adModeRead
    cnxDb.CursorLocation = adUseClient
    cnxDb.Open m_GradebookConnectionString
   
    Set rs = cnxDb.Execute(sSQL)
            
    iNumRelevantResults = 0
    iNumRecords = 0
    While Not rs.EOF
        iNumRecords = iNumRecords + 1
        If InStr(m_sOpenNextcapPens, "'" & rs.Fields("process_id") & "'") Then
            iNumRelevantResults = iNumRelevantResults + 1
            ReDim Preserve TestResults(iNumRelevantResults)
            'Determine the failcode
            sFailcode = ""
            
            'Check for Etest failcodes
            sFailcode = IIf(Trim(LCase(rs.Fields("pass_fail_cd"))) = "pass", "NTF", "FET")
            
            'If we have no etest failcode then check for region mismatches
            If m_CheckForEtesterRegionFails And sFailcode = "NTF" Then
                sFailcode = getRegionFailcode(rs.Fields("process_id"), ReverseText("" & rs.Fields("id_fet")))
            End If
            
            sFailcode = IIf(UCase(sFailcode) = "NTF", "NTF_E", UCase(sFailcode))
            
            TestResults(iNumRelevantResults - 1).TestType = "ETESTER"
            TestResults(iNumRelevantResults - 1).PenID = rs.Fields("process_id")
            TestResults(iNumRelevantResults - 1).ScanDatetime = rs.Fields("etest_dm")
            TestResults(iNumRelevantResults - 1).Failcode = sFailcode
            TestResults(iNumRelevantResults - 1).Comment = rs.Fields("pass_fail_cd")
            TestResults(iNumRelevantResults - 1).PrintSampleID = rs.Fields("stamp_link")
            sLatestDateTime = CStr(rs.Fields("txt_dttm"))
        End If
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    cnxDb.Close
    Set cnxDb = Nothing
    
    If iNumRecords > 0 Then
        'record the last data time if we got a result
        Call SaveSetting("iLink", "QueryDates", sDateRegistryKey, sLatestDateTime)
    End If
   
    LogEvent ("Found " & iNumRelevantResults & " results... ")
    If iNumRelevantResults > 0 Then 'can quit now if there are no new results
        NewPens = getUniquePens(TestResults)
       
        '1. Insert the new data into the finishedData table
        Call addTestResults(TestResults)
        
        '2. Determine and set the overall failmode for the pens
        PenResults = setOverallFailmode(NewPens)
                   
        '3. Load results to Nextcap if necessary
        Call LoadResultsToNextCAP(PenResults)
    End If
    
    Exit Sub
EH:
    LogEvent ("** ERROR : checkForEtestResults - " & Err.Number & " - " & Err.Description & " - sSQL=" & sSQL)
End Sub

Private Sub checkForKahunaResults()
'-------------------------------------------------------------------
' Load results from Kahuna
'
' Note: In order for this to work the view vw_kahuna_genx needs to be created
'       at the target DB.
'
' Operation:
'       1. Connect to the gradebook database
'       2. Query all new results for pens in open lots since the last update
'-------------------------------------------------------------------
On Error GoTo EH
    LogEvent ("Checking for KAHUNA Results... ")
    Dim TestResults() As TestResult
    Dim NewPens() As String
    Dim PenResults() As PenResult
    Dim iNumRelevantResults As Long
    Dim sFailcode As String
    Dim sPreviousDateTime As String
    Dim sDateRegistryKey As String
    Dim sLatestDateTime As String
    Dim sExistingKahunaStamps As String
    Dim iNumRecords As Long

    
    Dim cnxDb As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    'If we don't have any mids open exit the function without checking
    If Len(m_sOpenNextcapPens) < 16 Then
        LogEvent "No Pens in Open Lots.  Kahuna Search Skipped."
        Exit Sub
    End If
    
    'Look for last date stamp
    sDateRegistryKey = "Kahuna"
    sPreviousDateTime = GetSetting("iLink", "QueryDates", sDateRegistryKey)
    
    If sPreviousDateTime = "" Then
        sExistingKahunaStamps = getExistingStamps("KAHUNA", m_sOpenNextcapPens)
        
        If m_MfgDBType = "SQL-SERVER" Then
            sPreviousDateTime = "dateadd(day, -20, getdate()) AND test_pen_id in (" & m_sOpenNextcapPens & ") AND d.stamp not in (" & sExistingKahunaStamps & ")"
        Else
            sPreviousDateTime = "TODAY - 20 AND test_pen_id in (" & m_sOpenNextcapPens & ") AND d.stamp not in (" & sExistingKahunaStamps & ")"
        End If
    Else
        sPreviousDateTime = "'" & sPreviousDateTime & "'"
    End If
      
    'build the SQL String
    If m_MfgDBType = "SQL-SERVER" Then
        sSQL = "select distinct d.stamp, d.print_sample_id, d.test_pen_id, d.measure_tx, l.insert_dttm kahuna_dm, l.insert_dttm, l.insert_dttm + '' txt_dttm " & _
            " from insert_time_log l, wst_sample_value d where d.measure_nm = 'pen_id_bits' AND @DATEFILTER@ AND @INIFILTER@ and l.unique_id=d.stamp order by insert_dttm"
    Else
        sSQL = "select distinct d.stamp, d.print_sample_id, d.test_pen_id, d.measure_tx, extend (l.insert_dttm, year to second) kahuna_dm, l.insert_dttm, l.insert_dttm || '' txt_dttm " & _
            " from insert_time_log l, wst_sample_value d where d.measure_nm = 'pen_id_bits' AND @DATEFILTER@ AND @INIFILTER@ and l.unique_id=d.stamp order by insert_dttm"
    End If
   
    
    sSQL = Replace(sSQL, "@DATEFILTER@", "l.insert_dttm > " & sPreviousDateTime)
     
    If Trim(m_KahunaSQL) = "" Then
        sSQL = Replace(sSQL, "@INIFILTER@", "1=1")
    Else
        sSQL = Replace(sSQL, "@INIFILTER@", m_KahunaSQL)
    End If

    'Log the details if in debug mode
    If m_Debug = True Then
        Call LogEvent(sSQL)
    End If
         
    '1. Connect to the database
    Set cnxDb = New ADODB.Connection
    cnxDb.Mode = adModeRead
    cnxDb.CursorLocation = adUseClient
    cnxDb.Open m_GradebookConnectionString
           
    Set rs = cnxDb.Execute(sSQL)
    
    iNumRelevantResults = 0
    While Not rs.EOF
        iNumRecords = iNumRecords + 1
        If InStr(m_sOpenNextcapPens, "'" & rs.Fields("test_pen_id") & "'") Then
            iNumRelevantResults = iNumRelevantResults + 1
            ReDim Preserve TestResults(iNumRelevantResults)
            'Determine the failcode
            sFailcode = ""
            
            'If necessary look for region failcodes
            If m_CheckForKahunaRegionFails Then
                sFailcode = getRegionFailcode(rs.Fields("test_pen_id"), "" & rs.Fields("measure_tx"))
            End If
            
            'If we have no region code then assign an etester failcode
            If sFailcode = "" Or sFailcode = "NTF" Then
                sFailcode = "NTF_K"
            End If
            
            TestResults(iNumRelevantResults - 1).TestType = "KAHUNA"
            TestResults(iNumRelevantResults - 1).PenID = rs.Fields("test_pen_id")
            TestResults(iNumRelevantResults - 1).ScanDatetime = rs.Fields("kahuna_dm")
            TestResults(iNumRelevantResults - 1).Failcode = sFailcode
            TestResults(iNumRelevantResults - 1).Comment = "NA"
            TestResults(iNumRelevantResults - 1).PrintSampleID = rs.Fields("stamp")
            sLatestDateTime = CStr(rs.Fields("txt_dttm"))
        End If
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    cnxDb.Close
    Set cnxDb = Nothing
    
    If iNumRecords > 0 Then
        'record the last data time if we got a result
        Call SaveSetting("iLink", "QueryDates", sDateRegistryKey, sLatestDateTime)
    End If
    
    LogEvent ("Found " & iNumRelevantResults & " results... ")
    If iNumRelevantResults > 0 Then 'can quit now if there are no new results
        NewPens = getUniquePens(TestResults)
       
        '1. Insert the new data into the finishedData table
        Call addTestResults(TestResults)
        
        '2. Determine and set the overall failmode for the pens
        PenResults = setOverallFailmode(NewPens)
                   
        '3. Load results to Nextcap if necessary
        Call LoadResultsToNextCAP(PenResults)
    End If
    
    Exit Sub
EH:
    LogEvent ("** ERROR : checkForKahunaResults - " & Err.Number & " - " & Err.Description)
End Sub

Private Function getRegionFailcode(PenID As String, IdFet As String) As String
'-------------------------------------------------------------------
' Returns the region failcode for this pen or NTF if it checks okay
' Note VB is not a very efficient way of doing this - if performance is an issue
' then this function is one that might yield benefits by moving it to stored proc
'-------------------------------------------------------------------
    Dim sSQL As String
   
    On Error GoTo EH
    Dim sProduct As String
    Dim sMarketing As String
    
    'Extract the product/marketing elements from the fet these will
    sProduct = Mid(IdFet, 51, 5)
    sMarketing = Mid(IdFet, 25, 5)
       
    Dim sProductNumber As String
    Dim sRunType As String
            
    sProductNumber = m_dctPenIDProductNumber(PenID)
    sRunType = m_dctPenIDRuntype(PenID)
    
    'Exit point 1: Product number not available => error
    If sProductNumber = "" Or sRunType = "" Then
        getRegionFailcode = "FRG_NA"
        GoTo EndOfFunction
    End If
    
    'Exit point 2: No bits => region test not possible => good pen
    If (UCase(m_RegionCheckEngineeringLots) = "FALSE" And UCase(sRunType) = "ENGINEERING") Or Trim(m_dctProdRefIDFetProd(sProductNumber)) = "NA" Or Trim(m_dctProdRefIDFetProd(sProductNumber)) = "" Then
        getRegionFailcode = "NTF"
        GoTo EndOfFunction
    End If
    
    'Exit point 3: Region Mismatch (product bits)
    If m_dctProdRefIDFetProd(sProductNumber) <> sProduct Then
        getRegionFailcode = "FRG_PR"
        GoTo EndOfFunction
    End If
    
    'Exit point 4: Region Mismatch (marketing bits)
    If m_dctProdRefIDFetMktg(sProductNumber) <> sMarketing Then
        getRegionFailcode = "FRG_RG"
        GoTo EndOfFunction
    End If
    
    'Else all is good
    getRegionFailcode = "NTF"
    
EndOfFunction:
    Exit Function
EH:
    getRegionFailcode = "NTF"
    LogEvent ("** ERROR : getRegionFailcode - " & Err.Number & " - " & Err.Description & " - " & sSQL)
End Function


Private Function getTimeofLastResult(sTestType As String) As String
'-------------------------------------------------------------------
' Returns a database formatted string stating the most recent result received for this test
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim sSQL As String
    Dim rs As ADODB.Recordset
        
    sSQL = "SELECT max(scan_birthdate) from finisheddata where test_type = '" & sTestType & "'"
    
    Set rs = g_cnxDb.Execute(sSQL)
    
    If Not rs.EOF Then
        getTimeofLastResult = "'" & Format(rs.Fields(0), "YYYY-MM-DD hh:mm:ss") & "'"
    Else
        getTimeofLastResult = "'" & Format(DateAdd("d", -30, Now()), "YYYY-MM-DD hh:mm:ss") & "'"
    End If
       
    Exit Function
EH:
    LogEvent ("** ERROR : getTimeofLastResult - " & Err.Number & " - " & Err.Description)
End Function



Private Sub checkForCapLoadFailures()
'-------------------------------------------------------------------
' check for results which failed to load
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim sSQL As String
    Dim rsLoadFailures As ADODB.Recordset
    Dim oPenResults() As PenResult
    Dim i As Integer
        
    sSQL = "SELECT pen_id, failcode from NextCapFailures where page<" & m_numCapRetryAttempts & ";"
    
    
    Set rsLoadFailures = g_cnxDb.Execute(sSQL)
    
    i = 0
    While Not rsLoadFailures.EOF
        ReDim Preserve oPenResults(i)
        oPenResults(i).PenID = rsLoadFailures.Fields("pen_id")
        oPenResults(i).OverallFailcode = rsLoadFailures.Fields("failcode")
        i = i + 1
        rsLoadFailures.MoveNext
    Wend
    
    Call LoadResultsToNextCAP(oPenResults)
        
    Exit Sub
EH:
    LogEvent ("** ERROR : checkForCapLoadFailures - " & Err.Number & " - " & Err.Description)
End Sub


Private Function getExistingStamps(sTestType, sPidList) As String
'-------------------------------------------------------------------
' Compiles a list of the stamps for the currently open pens
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim sSQL As String
    Dim sReturn As String
    Dim rs As ADODB.Recordset
    
    sReturn = ""
       
    sSQL = "select distinct print_sample_id from finishedData " & _
            " Where test_type='" & sTestType & "' and pen_id in (" & sPidList & ") "
            
    Set rs = g_cnxDb.Execute(sSQL)
    While Not rs.EOF
        sReturn = sReturn & "'" & Trim(rs.Fields("print_sample_id")) & "'"
        rs.MoveNext
        If Not rs.EOF Then
            sReturn = sReturn & ","
        End If
    Wend
    Set rs = Nothing
    
    If Len(sReturn) < 2 Then
        sReturn = "'NA'"
    End If
    
    getExistingStamps = sReturn
    Exit Function
EH:
    LogEvent "** ERROR : getExistingStamps - " & sTestType & " - " & Err.Number & " - " & Err.Description, True
End Function

Public Sub reCalculateOpenPenFailures()
'-------------------------------------------------------------------
' Recalculates the iLink failcodes for the pens in open lots
'-------------------------------------------------------------------
    
    LogEvent ("Recalculating failcodes in OPEN lots... ")
    On Error GoTo EH
    
    Dim AllOpenPens() As String
    Dim PenResults() As PenResult

    'connect to db for next 2 steps
    If dbConnect() Then
        '1. Get list of pens in open lots into an array
        AllOpenPens = Split(getOpenNextcapPens("", True), ",")
            
        '2. Determine and set the overall failmode for the pens
        PenResults = setOverallFailmode(AllOpenPens)
                       
        '3. Load results to Nextcap if necessary
        Call LoadResultsToNextCAP(PenResults)
                            
        ' don't need db connection any more
        Call dbClose
    Else
        LogEvent "**ERROR! Can't Connect to Database", True
    End If

    Exit Sub
EH:
    LogEvent "** ERROR : reCalculateOpenPenFailures - " & Err.Number & " - " & Err.Description, True
End Sub


Private Function getOpenNextcapPens(Optional sRunType As String = "", Optional bReturnCSV As Boolean = False) As String
'-------------------------------------------------------------------
' Compiles a list of the pens in the currently open nextcap Lots
'-------------------------------------------------------------------
    LogEvent ("Compiling list of nextcap pens... " & sRunType)
    On Error GoTo EH
    Dim cnxDb As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim sReturn As String
    
    sReturn = ""
    
    Set m_dctPenIDRuntype = New Dictionary
    Set m_dctPenIDProductNumber = New Dictionary
        
    'Set up the Database Connection
    mdlNextcap.checkForNextcapLock 'check that we're not doing an upload
    Set cnxDb = New ADODB.Connection
    cnxDb.Mode = adModeRead
    cnxDb.CursorLocation = adUseClient
    cnxDb.Open g_NextcapConnectionString
    
    sSQL = "select pens.penid, pens.runtype, pens.productnumber from pens, lots " & _
            " Where pens.LotId = lots.LotId " & _
            "and pens.linetype = lots.linetype " & _
            "and pens.linenumber = lots.linenumber " & _
            "and pens.source = lots.source " & _
            "and pens.birthday = lots.birthday " & _
            "and lots.materialstatus <> 'CLOSED' "
            
    If sRunType <> "" Then
        sSQL = sSQL & "and pens.RunType = '" & sRunType & "'"
    End If
    Set rs = cnxDb.Execute(sSQL)
    While Not rs.EOF
        If Not bReturnCSV Then
            sReturn = sReturn & "'" & rs.Fields("PenId") & "'"
        Else
            sReturn = sReturn & rs.Fields("PenId")
        End If
        
        m_dctPenIDRuntype(Trim(CStr(rs.Fields("PenId")))) = rs.Fields("RunType")
        m_dctPenIDProductNumber(Trim(CStr(rs.Fields("PenId")))) = rs.Fields("ProductNumber")
        rs.MoveNext
        If Not rs.EOF Then
            sReturn = sReturn & ","
        End If
    Wend
    Set rs = Nothing
    cnxDb.Close
    Set cnxDb = Nothing
    
    getOpenNextcapPens = sReturn
    Exit Function
EH:
    LogEvent "** ERROR : getOpenNextcapPens - " & sRunType & " - " & Err.Number & " - " & Err.Description, True
End Function




Private Function setOverallFailmode(PenIDs() As String) As PenResult()
'-------------------------------------------------------------------
' Determine the overall failmode for a given pen id
'-------------------------------------------------------------------
    Dim iPen As Integer
    Dim sFailures() As String
    Dim iRule As Integer
    Dim iPriority As Integer
    Dim sSetFailcode As String
    Dim oPenResults() As PenResult
    Dim iFailure As Integer
    Dim iArraysize As Integer
    
    LogEvent ("Calculating Failmodes... ")

    For iPen = 0 To UBound(PenIDs)
        '1. Get the failures for each pen into an array (quicker than recordsets)
        sFailures = getPenFailures(PenIDs(iPen))
            
        On Error Resume Next 'an undefined array produces an error
        iArraysize = UBound(sFailures)
        If Err = 9 Then
            'If there are no results for the pen then it must be good from an iLink perspective
            sSetFailcode = g_goodPenCode
            GoTo Found
        End If
        On Error GoTo EH
    
        '2. See if we can match a rule for the pen
        On Error Resume Next 'an undefined array produces an error
        iArraysize = UBound(m_arrCapDecisionRules)
        If Err = 9 Then
            iArraysize = -1
        End If
        On Error GoTo EH
        
        For iRule = 0 To iArraysize
            If MatchRule(sFailures, m_arrCapDecisionRules(iRule).FailCodes) Then
                sSetFailcode = m_arrCapDecisionRules(iRule).SetFailCode
                LogEvent ("Matched on Rule number: " & iRule + 1)
                GoTo Found
            End If
        Next iRule
        
        '3. Find the most recent failcode of highest priority
        sSetFailcode = getHigestPriorityFailcode(sFailures, m_arrPriorityFailures)
        If sSetFailcode <> "" Then
            LogEvent ("Matched priority failcode..")
            GoTo Found
        End If
        
        '4. Try to assign the highest priority failcode of the highest level test
        '   provided that test did not return a "good" value
        On Error Resume Next 'an undefined array produces an error
        iArraysize = UBound(m_arrTestLevels)
        If Err = 9 Then
            iArraysize = -1
        End If
        On Error GoTo EH
        Dim iTestLevel As Integer
        Dim sTestFailures() As String
        
        For iTestLevel = 0 To iArraysize
            sTestFailures = getPenFailures(PenIDs(iPen), m_arrTestLevels(iTestLevel).TestType)
            sSetFailcode = getHigestPriorityFailcode(sTestFailures, m_arrTestLevels(iTestLevel).FailcodePriorities)
            If sSetFailcode <> "" And InStr(sSetFailcode, g_goodPenCode) <= 0 Then
                'If we get a result that's not an NTF then this is the result we want to use.
                GoTo Found
            End If
        Next iTestLevel
        
        '5. If all else fails, just assign a good failcode
        sSetFailcode = g_goodPenCode
        GoTo Found
                    
Found:
        ReDim Preserve oPenResults(iPen)
        oPenResults(iPen).PenID = PenIDs(iPen)
        oPenResults(iPen).OverallFailcode = sSetFailcode
        If updOverallFailcode(oPenResults(iPen).PenID, oPenResults(iPen).OverallFailcode) Then
            LogEvent ("Pen ID:" & oPenResults(iPen).PenID & " has overall failmode of: " & oPenResults(iPen).OverallFailcode)
        Else
            LogEvent ("** Error updating Result for Pen ID:" & oPenResults(iPen).PenID & " of overall failmode of: " & oPenResults(iPen).OverallFailcode)
        End If
    Next iPen
    
    setOverallFailmode = oPenResults
Exit Function
EH:
    LogEvent "** ERROR : setOverallFailmode - " & Err.Number & " - " & Err.Description, True
End Function

Private Function getHigestPriorityFailcode(sFailcodeArray() As String, sPriorityArray() As String) As String
'-------------------------------------------------------------------
' Given the ordered list of failcodes and priorities it will return
' the most appropriate failcode
'-------------------------------------------------------------------
    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    Dim iPriority As Integer
    Dim iFailure As Integer
    Dim sSetFailcode As String
    
    sSetFailcode = ""
    
    iArraysize = UBound(sPriorityArray)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
    
    For iPriority = 0 To iArraysize
        For iFailure = 0 To UUBound(sFailcodeArray)
            If InStr(sPriorityArray(iPriority), "'" & sFailcodeArray(iFailure) & "'") > 0 Then
                sSetFailcode = sFailcodeArray(iFailure)
                    GoTo Found
                End If
            Next iFailure
        Next iPriority
Found:
    getHigestPriorityFailcode = sSetFailcode
Exit Function
EH:
    LogEvent "** ERROR : getHigestPriorityFailcode - " & Err.Number & " - " & Err.Description, True
End Function

Private Function updOverallFailcode(PenID As String, Failcode As String) As Boolean
'-------------------------------------------------------------------
' Set the overall failcode in the iLink database to the new setfailcode
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim i As Integer
    Dim sSQL As String

    
    sSQL = "UPDATE finishedData set set_failcode='" & Failcode & "' " _
           & " WHERE pen_id='" & PenID & "';"

    g_cnxDb.Execute (sSQL)

    
    updOverallFailcode = True
    Exit Function
EH:
    updOverallFailcode = False
    LogEvent "** ERROR : updOverallFailcode - " & Err.Number & " - " & Err.Description & " - " & sSQL, True
End Function


Private Function LoadResultsToNextCAP(PenResults() As PenResult)
'-------------------------------------------------------------------
' Load any pen results to nextcap
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim i As Integer
    Dim sPenID As String
    Dim sCapFailcode As String
    Dim sSuccess As String
   
    
    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    iArraysize = UBound(PenResults)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
    
    For i = 0 To iArraysize
        sCapFailcode = getCapMatch(PenResults(i).OverallFailcode)
        sPenID = PenResults(i).PenID
        
        If (InStr(m_codesToIgnore, "|" & sCapFailcode & "|") > 0) Then
            'ignore these failcodes
            LogEvent ("Ignoring failcode :" & sCapFailcode)
        Else
            'Update NextCAP with the results
            If Not doubleCheckNextcap(sPenID, sCapFailcode) Then
                'Only update if the failcode is not already the intended one
                sSuccess = UpdatePen(sPenID, sCapFailcode)
                
                If sSuccess = "OK" Then
                    If m_doubleCheckNextcap Then
                        If Not doubleCheckNextcap(sPenID, sCapFailcode) Then
                            sSuccess = "LOAD ERROR - Contact Support"
                            LogEvent "**WARNING - " & sPenID & " (" & sCapFailcode & ") FAILED Double Check.", True
                        Else
                            LogEvent sPenID & " (" & sCapFailcode & ") Double Checked - OK. "
                        End If
                    End If
                    'this might be a retry attempt so clear this pen from the nextcap failures table if it's up okay
                    g_cnxDb.Execute ("DELETE FROM NextCapFailures where pen_id='" & PenResults(i).PenID & "';")
                Else
                    'Update of nextcap is unsuccessful - so increment the issue table
                    Call updNextcapIssueTable(sPenID, PenResults(i).OverallFailcode, _
                                          sSuccess)
                End If
                
            Else
                LogEvent ("Pen status not changed - no need to update NextCAP: " & sPenID)
                'this might be a retry attempt so clear this pen from the nextcap failures table if it's up okay
                g_cnxDb.Execute ("DELETE FROM NextCapFailures where pen_id='" & PenResults(i).PenID & "';")
            End If
            
        End If
    Next i
    
    Exit Function
EH:
    LogEvent "** ERROR : LoadResultsToNextCAP - " & Err.Number & " - " & Err.Description, True
End Function


Private Function doubleCheckNextcap(sPenID, sFailcode) As Boolean
'-------------------------------------------------------------------
' This might seem like overkill, but I've been asked to double check
' all nextcap updates to make sure that the nextcap result is okay.
' we've seen cases where nextcap says the pen is x but iLink says it's
' y so we have to make sure this doesn't happen
'-------------------------------------------------------------------
On Error GoTo EH
    Dim cnxDb As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    'Set up the Database Connection
    mdlNextcap.checkForNextcapLock 'check that we're not doing an upload
    Set cnxDb = New ADODB.Connection
    cnxDb.Mode = adModeRead
    cnxDb.CursorLocation = adUseClient
    cnxDb.Open g_NextcapConnectionString
    
    'assume the worst
    doubleCheckNextcap = False
    'If it's an NTF then make sure there are no iLink defects against that pen, otherwise
    'check that the specified failcode is against the pen
    If sFailcode = g_goodPenCode Then
        sSQL = "SELECT * from PenDefects where PenId='" & sPenID & "' AND DefectComment='" & _
        CN_ILINK_COMMENT & "' AND SyncState NOT IN ('REMOVE', 'DELETE')"
        Set rs = cnxDb.Execute(sSQL)
        If rs.EOF Then doubleCheckNextcap = True
    Else
        sSQL = "SELECT * from PenDefects where PenId='" & sPenID & "' AND code" & g_capDefectLevel & "='" & _
            sFailcode & "' AND SyncState NOT IN ('REMOVE', 'DELETE')"
        Set rs = cnxDb.Execute(sSQL)
        If Not rs.EOF Then doubleCheckNextcap = True
    End If
    
    Set rs = Nothing
    cnxDb.Close
    Set cnxDb = Nothing
    
    Exit Function
EH:
    doubleCheckNextcap = False
    LogEvent "** ERROR : doubleCheckNextcap - " & Err.Number & " - " & Err.Description, True
End Function

Private Function updNextcapIssueTable(PenID As String, Failcode As String, LoadError As String)
'-------------------------------------------------------------------
' Update or insert the number of nextcap failures for this pen in the load errors table
'-------------------------------------------------------------------
On Error GoTo EH
    Dim sSQL As String
    Dim iNumRetries As Integer
    Dim rs As ADODB.Recordset
    
    sSQL = "select page from NextCapFailures where pen_id='" & PenID & "'"
    
    Set rs = g_cnxDb.Execute(sSQL)
   
    If rs.EOF Then
        sSQL = "insert into NextCapFailures(pen_id, failcode, page, serverText, scan_birthdate) " _
            & "VALUES ('" & PenID & "', '" & Failcode & "', 1, '" & LoadError & "', '" & Now() & "');"
        LogEvent "Adding " & PenID & " (" & Failcode & ") To error Table."
    Else
        iNumRetries = CInt(rs.Fields(0).Value & "")
        sSQL = "UPDATE NextCapFailures SET " _
            & " failcode = '" & Failcode & "', " _
            & " serverText = '" & LoadError & "', " _
            & " page = '" & iNumRetries + 1 & "' " _
             & " WHERE pen_id='" & PenID & "';"
        LogEvent "Updating retries on " & PenID & " (" & Failcode & ") To " & iNumRetries & " in error Table."
    End If
    
    g_cnxDb.Execute (sSQL)
    
Exit Function
EH:
    LogEvent "** ERROR : updNextcapIssueTable - " & Err.Number & " - " & Err.Description, True

End Function


Private Function MatchRule(arrFailures() As String, sRule As String) As Boolean
'-------------------------------------------------------------------
' See if a list of failures matches a given rule
'-------------------------------------------------------------------
 Dim bReturn As Boolean
 
  
 '1. Prepare the rule expression
 'If there are no quotes in the rule then it is of the form used in earlier versions of iLink
 If Not InStr(sRule, "'") > 0 Then
    'Old-style rule - convert to the newer style to maintain backward compatibility
    sRule = "'" & Trim(Replace(sRule, "+", "'&'")) & "'"
 End If
 
 Dim i As Integer
 Dim iNumQuotes As Double
 Dim arrFailsInRule() As String
 Dim sCurrentChar As String
 Dim sTempFail
 Dim iNumFails As Integer
 
 '1. Compile an array of failcodes or warn if the quotes do not match up
 iNumQuotes = 0
 For i = 1 To Len(sRule)
    sCurrentChar = Mid(sRule, i, 1)
    If sCurrentChar = "'" Then
        'If we find a quote then take action
        If sTempFail <> "" Then
            iNumFails = iNumFails + 1
            ReDim Preserve arrFailsInRule(iNumFails)
            arrFailsInRule(iNumFails - 1) = sTempFail
            sTempFail = ""
        End If
        iNumQuotes = iNumQuotes + 1
        
    ElseIf iNumQuotes > 0 Then
        If iNumQuotes Mod 2 > 0 Then
            'If there is an uneven number of quotes then we are totting up a failcode
            sTempFail = sTempFail & sCurrentChar
        End If
    End If

 Next i
 
 'Check that our parsing was valid
 If (iNumQuotes Mod 2 > 0) Then
    LogEvent "** ERROR : MatchRule - Rule has uneven number of quotes! " & sRule, True
    Exit Function
 End If
 
  
 '2. Take the failures array and place it into a string for easier checking via instr
 Dim sFailsOccuring As String
 For Each sTempFail In arrFailures
    sFailsOccuring = sFailsOccuring & "|" & sTempFail & "|"
 Next
 
 '3. Format the rule into a valid boolean algebra statement
 For Each sTempFail In arrFailsInRule
    If sTempFail <> "" Then
        If InStr(sFailsOccuring, "|" & sTempFail & "|") Then
            sRule = Replace(sRule, "'" & sTempFail & "'", "1")
        Else
            sRule = Replace(sRule, "'" & sTempFail & "'", "0")
        End If
    End If
 Next
 
 'At this point sRule should be a valid boolean statement so we just have to parse it
 Dim oEval As New clsEval
 Dim iResult As Integer
 
 iResult = oEval.Evaluate(sRule)
 
 If iResult = 1 Then
    MatchRule = True
Else
    MatchRule = False
End If

Exit Function
EH:
    LogEvent "** ERROR : MatchRule - " & Err.Number & " - " & Err.Description, True
End Function


Private Function getPenFailures(PenID As String, Optional TestType As String = "") As String()
'-------------------------------------------------------------------
' Return a list of failures which have been assigned to a pen
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim iNumFailures As Integer
    Dim sSQL As String
    Dim rsFailures As ADODB.Recordset
    Dim sFailures() As String
    iNumFailures = 0

    sSQL = "select pen_failcode, page_count, scan_birthdate from finishedData where pen_id='" & PenID & "'"
    If TestType <> "" Then
        sSQL = sSQL & " AND test_type='" & TestType & "'"
    End If
    sSQL = sSQL & " order by page_count desc, scan_birthdate desc;"
    Set rsFailures = g_cnxDb.Execute(sSQL)
    
    While Not rsFailures.EOF
        ReDim Preserve sFailures(iNumFailures)
        sFailures(iNumFailures) = rsFailures.Fields("pen_failcode")
        iNumFailures = iNumFailures + 1
        rsFailures.MoveNext
    Wend
    rsFailures.Close
    Set rsFailures = Nothing
    
    getPenFailures = sFailures
    Exit Function
EH:

    LogEvent "** ERROR : getPenFailures - " & Err.Number & " - " & Err.Description & " - " & sSQL, True
End Function

Private Function addTestResults(TestResults() As TestResult) As Boolean
'-------------------------------------------------------------------
' Add new results to iLink database
' Am using Stuffer database schema for backward compatibility
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim i As Integer
    Dim sSQL As String

    sSQL = ""
    
    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    iArraysize = UBound(TestResults)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
        
    For i = 0 To iArraysize
        If TestResults(i).Failcode <> "" Then   'Results without a failcode are no good
            LogEvent ("Inserting ps: " & TestResults(i).PrintSampleID & " pen_id: " & TestResults(i).PenID _
                    & " Failcode: " & TestResults(i).Failcode & " to finishedData Table.")
            sSQL = "INSERT INTO finishedData(test_type, print_sample_id, pen_id, lot_id, " _
                 & "pen_birthdate, scan_birthdate, pen_failcode, page_count) VALUES ('" & TestResults(i).TestType & "', '" _
                & TestResults(i).PrintSampleID & "', '" & TestResults(i).PenID & "', '" _
                & TestResults(i).Comment & "', '" & Now() & "', '" & TestResults(i).ScanDatetime & "', '" _
                & TestResults(i).Failcode & "', " & TestResults(i).PageNumber & ");"
            g_cnxDb.Execute (sSQL)
        End If
     Next i

    
    addTestResults = True
    Exit Function
EH:
    addTestResults = False
    LogEvent "** ERROR : addTestResults - " & Err.Number & " - " & Err.Description & " - " & sSQL, True
End Function

Private Function getUniquePens(TestResults() As TestResult) As String()
'-------------------------------------------------------------------
' Returns an array of unique pen_id's from a set of results
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim i As Integer
    Dim sFoundPens As String
    Dim sPenID As String
    Dim iUniquePens As Integer
    
    iUniquePens = 0
    sFoundPens = ""
    
    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    iArraysize = UBound(TestResults)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
        
    For i = 0 To iArraysize
        sPenID = TestResults(i).PenID
         If Not InStr(sFoundPens, sPenID & "|") > 0 Then
            sFoundPens = sFoundPens & sPenID & "|"
         End If
    Next i
    
    If Len(sFoundPens) > 0 Then
        sFoundPens = Left(sFoundPens, Len(sFoundPens) - 1) 'remove the last pipe
    End If
    
    getUniquePens = Split(sFoundPens, "|")
    Exit Function
EH:
    LogEvent "** ERROR : getUniquePens - " & Err.Number & " - " & Err.Description, True
End Function


Private Function ReadDataFromBanks(ByRef TestResults() As TestResult) As Long
'-------------------------------------------------------------------
' Checks all of the active banks for results
'-------------------------------------------------------------------
   ' 7 fields
   ' 0 = psID, 1 = psLot, 2 = psFailCode, 3 = psPenId
   ' 4 = psPage, 5 & 6 = psDate
        
    On Error GoTo EH
    Dim i As Integer
    Dim sPath As String
    Dim sFullFilePath As String
    Dim arrTestResult() As TestResult
    Dim sFiles() As String
    Dim iFileNumber As Long
    Dim iFreefile As Integer
    Dim sFileData As String
    Dim iNumFieldsInFile As Integer
    Dim sFieldData() As String
    Dim iNumResultsInFile As Integer
    Dim iCtr As Integer
    Dim iNumResultsOverall As Integer
    Dim bIgnoreResult As Boolean
    
    Dim sTestType As String
    Dim sFailcode As String
    
    Dim CryptRef As New Crypto
    
    iNumResultsOverall = 0
    
    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    iArraysize = UBound(g_arrScanners)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
    
    For i = 0 To iArraysize
        'loop through the banks
       If g_arrScanners(i).Active Then
            If Not TestShare(g_arrScanners(i).Path) Then
                MsgBox App.Title & " cannot contact scanner: " & g_arrScanners(i).Alias & " and am disabling it.  " & vbCrLf _
                    & vbCrLf & "Re-enable it manually in iLink when it is working.", vbCritical
                g_arrScanners(i).Active = False
                Dim lstItem As ListItem
                For Each lstItem In frmMain.lstScanners.ListItems
                    If lstItem.text = g_arrScanners(i).Alias Then
                        lstItem.Checked = False
                    End If
                Next
                
                
            Else
                sPath = g_arrScanners(i).Path & "*." & m_fileExtension
                sFiles = AllFiles(sPath)
                
                If sFiles(0) <> "" Then
                For iFileNumber = 0 To UBound(sFiles)
                    sFullFilePath = g_arrScanners(i).Path & sFiles(iFileNumber)
                    'loop through all the files you can find at this bank
                    iFreefile = FreeFile()
                    LogEvent ("Opening - " & sFullFilePath)
                    Open sFullFilePath For Input As iFreefile
                            sFileData = Input(LOF(iFreefile), iFreefile)
                            'Panter4.6 release - Decrypt the file
                            sFileData = CryptRef.DecryptString(sFileData)
                    Close iFreefile
                
                    sFileData = Replace(AllTrim(sFileData), Chr(10), ",")
                    sFileData = Replace(AllTrim(sFileData), Chr(13), "")
                    'Trim out the trailing comma
                    If Right(sFileData, 1) = "," Then
                        sFileData = Left(sFileData, Len(sFileData) - 1)
                    End If
                    
                    sFieldData = Split(sFileData, ",", -1, 1)
                
                    iNumFieldsInFile = UBound(sFieldData) + 1
                    If (iNumFieldsInFile Mod (m_numIPASColumns) = 0) Then
                        'we have some results
                        iNumResultsInFile = iNumFieldsInFile / (m_numIPASColumns)
                        For iCtr = 0 To iNumResultsInFile - 1
                                    ' If the banks are shared with other nextcap PC's then disregard this result if it is
                                    ' not from a pen that has been entered to this nextcap station
                                    bIgnoreResult = False
                                    If m_bShareBanks Then
                                        Select Case m_fileStyle
                                            Case "Wildflowers"
                                                If InStr(m_sOpenNextcapPens, sFieldData(0 + (m_numIPASColumns * iCtr))) < 0 Then
                                                    bIgnoreResult = True
                                                End If
                                            
                                            Case Else
                                                If InStr(m_sOpenNextcapPens, sFieldData(3 + (m_numIPASColumns * iCtr))) < 0 Then
                                                    bIgnoreResult = True
                                                End If
                                        End Select
                                    End If
                                    
                                    If bIgnoreResult = False Then
                                        iNumResultsOverall = iNumResultsOverall + 1
                                        ReDim Preserve arrTestResult(iNumResultsOverall - 1)
                                        Select Case m_fileStyle
                                            Case "Wildflowers"
                                                '2 Fields
                                                ' 0 = PenID, 1= Failcode
                                                arrTestResult(iNumResultsOverall - 1).TestType = "WILDFLOWERS"
                                                arrTestResult(iNumResultsOverall - 1).PrintSampleID = "NA"
                                                arrTestResult(iNumResultsOverall - 1).Comment = "NA"
                                                arrTestResult(iNumResultsOverall - 1).Failcode = sFieldData(1 + (m_numIPASColumns * iCtr))
                                                arrTestResult(iNumResultsOverall - 1).PenID = sFieldData(0 + (m_numIPASColumns * iCtr))
                                                arrTestResult(iNumResultsOverall - 1).PageNumber = 1
                                                arrTestResult(iNumResultsOverall - 1).ScanDatetime = Now()
                                            Case Else 'Defaults to Inspector format
                                                ' 7 fields (Inspector)
                                                ' 0 = psID, 1 = psLot, 2 = psFailCode, 3 = psPenId
                                                ' 4 = psPage, 5 & 6 = psDate
                                                
                                                ' NOTE: Need to find more generic way of identifyiny glossy paper source
                                                '       perhaps Steve Gaddi can return a flag via IPass
                                                '
                                                sFailcode = sFieldData(2 + (m_numIPASColumns * iCtr))
                                                If Len(sFailcode) > 3 And Right$(sFailcode, 1) = "G" Then
                                                    sTestType = "GLOSSY"
                                                Else
                                                    sTestType = "PLAIN"
                                                End If
                                                
                                                arrTestResult(iNumResultsOverall - 1).TestType = sTestType
                                                arrTestResult(iNumResultsOverall - 1).PrintSampleID = sFieldData(0 + (m_numIPASColumns * iCtr))
                                                arrTestResult(iNumResultsOverall - 1).Comment = sFieldData(1 + (m_numIPASColumns * iCtr))
                                                arrTestResult(iNumResultsOverall - 1).Failcode = sFailcode
                                                arrTestResult(iNumResultsOverall - 1).PenID = sFieldData(3 + (m_numIPASColumns * iCtr))
                                                arrTestResult(iNumResultsOverall - 1).PageNumber = sFieldData(4 + (m_numIPASColumns * iCtr))
                                                arrTestResult(iNumResultsOverall - 1).ScanDatetime = CDate(sFieldData(5 + (m_numIPASColumns * iCtr)) _
                                                                            & " " & sFieldData(6 + (m_numIPASColumns * iCtr)))
                                        End Select
                                        arrTestResult(iNumResultsOverall - 1).DataFile = sFullFilePath
                                End If
                        Next iCtr
                    
                    Else
                        LogEvent ("** Warning! ** - File " & sFullFilePath & " has incorrect number of columns")
                    End If

                Next iFileNumber
                End If 'has files

            End If 'share is working
        End If 'share is supposed to be active
    Next i
    LogEvent ("Found " & iNumResultsOverall & " results.")
    If iNumResultsOverall > 0 Then
        TestResults = arrTestResult
        ReadDataFromBanks = iNumResultsOverall
    End If
    Exit Function
EH:
    LogEvent "** ERROR : ReadDataFromBanks - " & Err.Number & " - " & Err.Description, True
End Function

Private Function DeleteDataFiles(TestResults() As TestResult) As Boolean
'-------------------------------------------------------------------
' remove the datafiles that results were extracted from
' should only be called after the results have been put somewhere - i.e. iLink database
'-------------------------------------------------------------------
    On Error GoTo EH
    
    Dim sCurrentFile As String
    Dim sAlreadyRemoved As String
    Dim i As Integer


    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    iArraysize = UBound(TestResults)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
    
    sAlreadyRemoved = ""
    For i = 0 To iArraysize
        sCurrentFile = TestResults(i).DataFile
        If Not InStr(sAlreadyRemoved, sCurrentFile) > 0 Then
            LogEvent ("Deleting Datafile:" & sCurrentFile)
            Kill (sCurrentFile)
            sAlreadyRemoved = sAlreadyRemoved & "|" & sCurrentFile
        End If
    Next i
        
    DeleteDataFiles = True
    Exit Function
EH:
    LogEvent "** ERROR : DeleteDataFiles - " & Err.Number & " - " & Err.Description, True
    DeleteDataFiles = False
End Function


Public Function LoadSettings(Optional sIniFile As String = "") As Boolean
'-------------------------------------------------------------------
' Loads settings from application ini file
'-------------------------------------------------------------------
    Dim sRuleIniFile As String
    On Error GoTo EH
    Dim i As Integer
    Dim j As Integer
    
    If sIniFile = "" Then
        sIniFile = App.Path & "\" & App.Title & ".ini"
    End If
    
    LogEvent ("Reading general Settings from ini file - " & sIniFile)
    '1. Load the General Settings
    g_logEventsToFile = CBool(ReadIniString("LogEventsToFile", "False", "GENERAL", sIniFile))
    g_loadToNextcap = CBool(ReadIniString("LoadToNextcap", "False", "GENERAL", sIniFile))
    g_refreshInterval = CInt(ReadIniString("RefreshInterval", "10", "GENERAL", sIniFile)) * 1000
    g_goodPenCode = ReadIniString("GoodPenCode", "NTF", "GENERAL", sIniFile)
    m_numDaysData = CInt(ReadIniString("NumDaysDataToStore", "12", "GENERAL", sIniFile))
    m_numCapRetryAttempts = CInt(ReadIniString("NumCapRetryAttempts", "5", "GENERAL", sIniFile))
    m_doubleCheckNextcap = CBool(ReadIniString("DoubleCheckNextcap", "False", "GENERAL", sIniFile))
    m_numdaysLogFile = CInt(ReadIniString("NumDaysLogFile", "3", "GENERAL", sIniFile))
    g_capDefectLevel = CInt(ReadIniString("CapDefectLevel", "2", "GENERAL", sIniFile))
    sRuleIniFile = App.Path & "\" & ReadIniString("RuleIniFile", App.Title & ".ini", "GENERAL", sIniFile)
    m_MfgDBType = ReadIniString("MfgDBType", "INFORMIX", "GENERAL", sIniFile)
    m_Debug = CBool(ReadIniString("Debug", "False", "GENERAL", sIniFile))
    
    'Connection Strings are generally optional
    g_NextcapConnectionString = ReadIniString("NextcapConnectionString", "File Name=c:\program files\nextcap\server\udls\local.udl", "DATABASE", sIniFile)
    m_ILinkConnectionString = ReadIniString("ILinkConnectionString", "Driver={Microsoft Access Driver (*.mdb)};Dbq=c:\program files\iTools\iLink.mdb;Uid=admin;Pwd=", "DATABASE", sIniFile)
    m_GradebookConnectionString = ReadIniString("GradebookConnectionString", "File Name=c:\program files\nextcap\server\udls\mfg.udl", "DATABASE", sIniFile)
        
    'Test Types to check
    m_CheckForInspectorResults = CBool(ReadIniString("CheckForInspectorResults", "False", "TestTypes", sIniFile))
    m_CheckForEtesterResults = CBool(ReadIniString("CheckForEtesterResults", "False", "TestTypes", sIniFile))
    m_CheckForEtesterRegionFails = CBool(ReadIniString("CheckForEtesterRegionFails", "False", "TestTypes", sIniFile))
    m_CheckForKahunaRegionFails = CBool(ReadIniString("CheckForKahunaRegionFails", "False", "TestTypes", sIniFile))
    m_RegionCheckEngineeringLots = CBool(ReadIniString("RegionCheckEngineeringLots", "True", "TestTypes", sIniFile))
    
    'Inspector Specific Settings
    m_numIPASColumns = CInt(ReadIniString("numOfColumns", "7", "INSPECTOR", sIniFile))
    m_fileExtension = ReadIniString("FileExtension", "txt", "INSPECTOR", sIniFile)
    m_fileStyle = ReadIniString("FileStyle", "Inspector", "INSPECTOR", sIniFile)
    m_bShareBanks = ReadIniString("ShareBanks", "False", "INSPECTOR", sIniFile)
   
    '2. Load the scanner bank information
    LogEvent ("Reading Bank information from ini file - " & sIniFile)
    Dim numScanners As Integer
    numScanners = CInt(ReadIniString("numOfBanks", "0", "INSPECTOR", sIniFile))
    
    If numScanners > 0 Then
        ReDim g_arrScanners(numScanners - 1)
            For i = 0 To numScanners - 1
                g_arrScanners(i).Path = ReadIniString("Bank" & i + 1, "Missing", "INSPECTOR", sIniFile)
                If Right(g_arrScanners(i).Path, 1) <> "\" Then
                    g_arrScanners(i).Path = g_arrScanners(i).Path & "\"
                End If
                g_arrScanners(i).Alias = ReadIniString("Bank" & i + 1 & "Alias", "Missing", "INSPECTOR", sIniFile)
                If g_arrScanners(i).Alias = "Missing" Then g_arrScanners(i).Alias = g_arrScanners(i).Path
                g_arrScanners(i).Active = False
        Next i
    End If
    
    'Etester Specific Settings
    m_EtesterSQL = ReadIniString("SQL", "", "ETESTER", sIniFile)
    m_CheckEtestGen1 = CBool(ReadIniString("CheckEtestGen1", "True", "ETESTER", sIniFile))
    m_CheckEtestGen2 = CBool(ReadIniString("CheckEtestGen2", "True", "ETESTER", sIniFile))
    
    'Kahuna Specific Settings
    m_KahunaSQL = ReadIniString("SQL", "", "KAHUNA", sIniFile)
    
    '3. Load the CAP Decision Rules
    LogEvent ("Reading CAP Decision rules from ini file - " & sRuleIniFile)
    Dim numRules As Integer
    numRules = CInt(ReadIniString("numOfRules", "0", "RULES", sRuleIniFile))
    
    If numRules > 0 Then
        ReDim m_arrCapDecisionRules(numRules - 1)
        For i = 0 To numRules - 1
            m_arrCapDecisionRules(i).FailCodes = ReadIniString("Failcodes" & i + 1, "Missing", "RULES", sRuleIniFile)
            m_arrCapDecisionRules(i).SetFailCode = ReadIniString("SetFailcode" & i + 1, "Missing", "RULES", sRuleIniFile)
        Next i
    End If
        
    '4. load the priority failmodes
    LogEvent ("Reading Priority Failures from ini file - " & sRuleIniFile)
    Dim numPriorities As Integer
    numPriorities = CInt(ReadIniString("numOfPriorities", "0", "SpecialFailcodes", sRuleIniFile))
    
    If numPriorities > 0 Then
        ReDim m_arrPriorityFailures(numPriorities - 1)
        For i = 0 To numPriorities - 1
            m_arrPriorityFailures(i) = "'" & ReadIniString("Priority" & i + 1, "Missing", "SpecialFailcodes", sRuleIniFile) & "'"
        Next i
    End If
    m_codesToIgnore = CStr(ReadIniString("CodesToIgnore", "", "SpecialFailcodes", sRuleIniFile))
       
    '5. load the "test" level priorities
    LogEvent ("Reading Test Levels from ini file - " & sRuleIniFile)
    Dim numLevels As Integer
    numLevels = CInt(ReadIniString("numTestLevels", "0", "RULES", sRuleIniFile))
    
    If numLevels > 0 Then
        ReDim m_arrTestLevels(numLevels - 1)
        For i = 0 To numLevels - 1
            m_arrTestLevels(i).TestType = ReadIniString("TestLevel" & i + 1, "Missing", "RULES", sRuleIniFile)
            numPriorities = CInt(ReadIniString("numOfPriorities", "0", m_arrTestLevels(i).TestType, sRuleIniFile))
            If numPriorities > 0 Then
                ReDim m_arrTestLevels(i).FailcodePriorities(numPriorities - 1)
                For j = 0 To numPriorities - 1
                    m_arrTestLevels(i).FailcodePriorities(j) = "'" & ReadIniString("Priority" & j + 1, "Missing", m_arrTestLevels(i).TestType, sRuleIniFile) & "'"
                Next j
            End If
        Next i
    End If
    
    'If we are to do region checking, then load the product reference information into local variables for quick access
    If m_CheckForEtesterRegionFails Or m_CheckForKahunaRegionFails Then
        Dim cnxDb As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim sProductNumber As String
      
        '1. Connect to the database
        Set cnxDb = New ADODB.Connection
        cnxDb.Mode = adModeRead
        cnxDb.CursorLocation = adUseClient
        cnxDb.Open m_GradebookConnectionString
        
        Set rs = cnxDb.Execute("select * from product_ref_llk")
    
        While Not rs.EOF
            sProductNumber = rs("inv_item_lk_nr")
            m_dctProdRefIDFetMktg(sProductNumber) = rs("id_fet_marketing")
            m_dctProdRefIDFetProd(sProductNumber) = rs("id_fet_product")
            m_dctProdRefVentLab(sProductNumber) = rs("pica_cd")
            rs.MoveNext
        Wend
        Set rs = Nothing
        cnxDb.Close
        Set cnxDb = Nothing
        
    End If
    
    LoadSettings = True
    Exit Function
EH:
    LoadSettings = False
    LogEvent "** ERROR : LoadSettings - " & Err.Number & " - " & Err.Description, True
End Function

Private Function InitialiseScanners() As Boolean
'-------------------------------------------------------------------
' Tests and sets up the scanners
'-------------------------------------------------------------------
On Error GoTo EH
    Dim lstElement As ListItem
    Dim sAlias As String
    Dim sPath As String
    Dim sLabel As String
    Dim i As Integer
        
    LogEvent ("Initialising Scanners..")
    
    
    On Error Resume Next 'an undefined array produces an error
    Dim iArraysize As Integer
    iArraysize = UBound(g_arrScanners)
    If Err = 9 Then
        iArraysize = -1
    End If
    On Error GoTo EH
        
    
    For i = 0 To UBound(g_arrScanners)
        sAlias = g_arrScanners(i).Alias
        sPath = g_arrScanners(i).Path
        
        Set lstElement = frmMain.lstScanners.ListItems.Add()
        lstElement.text = sAlias
        If TestShare(sPath) Then
            lstElement.Checked = True
            g_arrScanners(i).Active = True
        Else
            lstElement.Checked = False
            g_arrScanners(i).Active = False
        End If
    Next i
    
    InitialiseScanners = True
    Exit Function
EH:
    InitialiseScanners = False
    LogEvent "** ERROR : InitialiseScanners - " & Err.Number & " - " & Err.Description, True
End Function

Public Sub DailyMaintenance()
'-------------------------------------------------------------------
' check if it's necessary to archive the logfile and if so, archive it
'-------------------------------------------------------------------
Dim i As Integer
Dim sFilename As String
Dim iFileNum As Integer

On Error GoTo EH:
    If Day(Now()) <> g_lastDay Then
        'the day has changed - archive the logfiles
        LogEvent "Archiving Logfiles... "
        sFilename = App.Path & "\" & App.Title
        On Error Resume Next
        For i = m_numdaysLogFile To 1 Step -1
            If i > 1 Then
                FileCopy sFilename & "-" & i - 1 & ".log", sFilename & "-" & i & ".log"
            Else
                FileCopy sFilename & ".log", sFilename & "-" & i & ".log"
            End If
        Next i
        
        'clear the default log file
        iFileNum = FreeFile()
        Open App.Path & "\" & App.Title & ".log" For Output As iFileNum
            Print #iFileNum, Now() & " : Initialising New Logfile .. "
        Close #iFileNum
        
        'Do some database maintenance
        Call CleanUpDB(m_numDaysData)
        
    End If
        g_lastDay = Day(Now())
Exit Sub
EH:
    LogEvent "** ERROR : DailyMaintenance - " & Err.Number & " - " & Err.Description, True
End Sub


Public Sub LogEvent(Message As String, Optional LogError As Boolean = False)
'-------------------------------------------------------------------
' Logs an event to the status bar or logfile depending on the settings.
'-------------------------------------------------------------------
    Dim dummy As Integer
    Dim iFileNum As Integer
    
    frmMain.StatusBar1.Panels(1).text = Now()
    frmMain.StatusBar1.Panels(2).text = Message
    dummy = DoEvents()
    
    If g_logEventsToFile = True Then
        iFileNum = FreeFile()
        Open App.Path & "\" & App.Title & ".log" For Append As iFileNum
            Print #iFileNum, Now() & " : " & Message
        Close #iFileNum
        
    End If
    
    If LogError Then
        iFileNum = FreeFile()
        Open App.Path & "\" & App.Title & "Err.log" For Append As iFileNum
            Print #iFileNum, Now() & " : " & Message
        Close #iFileNum
    End If
End Sub


Private Function getCapMatch(sFailcode As String) As String
'-------------------------------------------------------------------
' Find corresponding cap failcode for a pen
'-------------------------------------------------------------------
On Error GoTo EH
    Dim sResult As String
    
    sResult = ReadIniString(sFailcode, "Missing", "MatchList")
    
    If sResult = "Missing" Then
        sResult = sFailcode
    End If
    
    getCapMatch = sResult
    
    Exit Function
EH:
    LogEvent "** ERROR : getCapMatch - " & Err.Number & " - " & Err.Description, True
End Function

Private Sub ToggleControls(bEnabled As Boolean)
'-------------------------------------------------------------------
' enables or disables controls on the user interface.
'-------------------------------------------------------------------
On Error GoTo EH
    With frmMain
        .mnuTools.Enabled = bEnabled
        .cmdViewLog.Enabled = bEnabled
        .btnHide.Enabled = bEnabled
        .lstScanners.Enabled = bEnabled
        .CmdExit.Enabled = bEnabled
        .mnuAbout.Enabled = bEnabled
        .chkLogEvents.Enabled = bEnabled
        .mnuFile.Enabled = bEnabled
    End With
    Exit Sub
EH:
    LogEvent "** ERROR : ToggleControls - " & Err.Number & " - " & Err.Description, True
End Sub


Private Function CleanUpDB(iNumDays As Integer)
'-------------------------------------------------------------------
' Trim out any old data in the database to stop it growing indefinitely
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim i As Integer
    Dim sSQL As String
   
    If dbConnect() Then
    
        LogEvent ("Deleting old data...")
        sSQL = "DELETE from finishedData where pen_birthdate < #" & Now - iNumDays & "#;"
        g_cnxDb.Execute (sSQL)

        sSQL = "DELETE from NextCapFailures where scan_birthdate < #" & Now - iNumDays & "#;"
        g_cnxDb.Execute (sSQL)
    
        LogEvent ("Compacting and repairing database...")
    
        '1. find the path to the access db
        Dim sConnectionString As String
        Dim sDBPath As String
        Dim sDBQString As String
    
        sDBQString = "DBQ="
    
        sConnectionString = g_cnxDb.ConnectionString
        Call dbClose
    
        sDBPath = Right(sConnectionString, Len(sConnectionString) - InStr(sConnectionString, sDBQString) - Len(sDBQString) + 1)
        sDBPath = Left(sDBPath, InStr(sDBPath, ";") - 1)

        mdlUtil.CompactRepair (sDBPath)
        CleanUpDB = True
    End If
    
    Exit Function
EH:
    CleanUpDB = False
    LogEvent "** ERROR : CleanUpDB - " & Err.Number & " - " & Err.Description & " - " & sSQL, True
End Function


Private Function dbConnect() As Boolean
'-------------------------------------------------------------------
' Connects to the database
'-------------------------------------------------------------------

On Error GoTo EH
    Set g_cnxDb = Nothing
    Set g_cnxDb = New ADODB.Connection
    'Set up the Database Connection
    g_cnxDb.Mode = adModeReadWrite
    g_cnxDb.CursorLocation = adUseClient
    g_cnxDb.Open m_ILinkConnectionString
   
    dbConnect = True
    
    Exit Function

EH:
    dbConnect = False
    LogEvent "** ERROR: dbConnect: " & Err.Description, True
End Function


Private Function dbClose() As Boolean
'-------------------------------------------------------------------
' DisConnects from the Database
'-------------------------------------------------------------------
On Error GoTo EH
    g_cnxDb.Close
    Set g_cnxDb = Nothing
    
    dbClose = True
    Exit Function
EH:
    dbClose = False
    LogEvent "** ERROR: dbClose: " & Err.Description, True
End Function


Private Function AllTrim(ByVal s As String) As String
'-------------------------------------------------------------------
' trims all leading and trailing non-alphanumeric characters from a string
'-------------------------------------------------------------------
    Dim i As Long
    Dim iStrLen As Double
    Dim cOneCharacter As String
    
    While Not Left(s, 1) Like "[0-9A-Za-z]"
       s = Right(s, Len(s) - 1)
    Wend
    
    While Not Right(s, 1) Like "[0-9A-Za-z]"
       s = Left(s, Len(s) - 1)
    Wend
        
    AllTrim = s
End Function
