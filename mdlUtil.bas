Attribute VB_Name = "mdlUtil"
Option Explicit

'/*Global error log
Public Const g_sErrorLog As String = "Error.txt"

'/*Collection of LotManagers
Public gcol_LotManagers As Collection

'/*Global Supervisor object
Public go_Supervisor As NextCapServer.clsSupervisor
'/*Boot object required to get a connection to the supervisor
Public go_BusinessServer As NextCapServer.clsBoot
'/*LotManager, there is one of these for each valid
'/*station configuration in the NextCap client
Public go_ActiveLotManager As NextCapServer.clsLotManager
Public gn_LotManagerIndex As Integer

'/*Constant to define server error command
Public Const cServerError As String = "ERROR"

'/*Minimum interval for WFlower form timer
Public Const c_lngTimerMin As Long = 60000

Public bVisible As Boolean

'/*Public concurrency counter for shutdown procedure
Public gn_ShutdownCount As Integer

'/*-------------------------------------------------------------------
'/*Win32 API to get computer UNC ref. Microsoft KB article Q148835
'/*-------------------------------------------------------------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH As Long = 15&

'/*-------------------------------------------------------------
'/API for thread sleep function
'/*-------------------------------------------------------------
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'/*--------------------------------------------------------------------
'/*API declarations for Readprofile and Writeprofile functions
'/*--------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal appName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$) As Long
  
  
Public Function SetupNextCap() As Boolean
On Error GoTo EH
    Set gcol_LotManagers = New Collection
    EstablishConnection
    RecvStations
    ShowFeedback "Status: System Ready"
    SetupNextCap = True
    Exit Function
EH:
    MsgBox "Error connecting to nextcap application - check that you are using the correct software versions."
    LogEvent "** ERROR:  SetupNextCap: " & Err.Description, True
    SetupNextCap = False
End Function

'
'============================================================
'Routine: mdlTools.CurrentMachineName()
'Purpose: This queries the machine name using Win32 API.
'         More references to similar calls can be found
'         at the Microsoft Knowledge Base in article
'         Q148835 which lists API for WorkGroup, Domain etc.
'
'Globals:None
'
'Input:None
'
'Return: String - This returns the UNC for a PC.
'
'Tested:
'   11-12-1998 Tested by hand. Chris Barker
'
'Modifications:
'   11-12-1998 As written for Pass1.3
'
'
'============================================================
Public Function CurrentMachineName() As String
Dim lSize As Long
Dim sBuffer As String
    sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
    lSize = Len(sBuffer)
 
    If GetComputerName(sBuffer, lSize) Then
        CurrentMachineName = Left$(sBuffer, lSize)
    End If
End Function


'
'===============================================================
'Routine: mdlMain.EstablishConnection()
'Purpose: Tis creates an instance of the primary
'exposed COM object that will pass back a Supervisor
'object to us. The supervisor can be queried for LotManagers
'
'
'===============================================================
Public Function EstablishConnection() As Boolean
    'ShowFeedback "Status: Connecting to server..."
    Set go_BusinessServer = CreateObject("NextCapServer.clsBoot")
    If Not (go_BusinessServer Is Nothing) Then
        '/*Set our Sink object in the form with the supervisor events
        Set frmMain.m_BusinessServer = go_BusinessServer
        '/*Tell the server to initiate its processes to come on-line
        Call go_BusinessServer.WakeUp(mdlUtil.CurrentMachineName)
        EstablishConnection = True
    Else
        'ShowFeedback "Status: Connection failed, make sure station configuration is valid."
    End If
Exit Function
'..............................................
    logToFile g_sErrorLog, "mdlUtil.EstablishConnection:" & Err.Description & "-" & CStr(Err.Number)
    Err.Clear
    Exit Function
End Function

'
'=======================================================
'Routine: mdlMain.RecvStations(o)
'Purpose: This requests the stations assinged to this
'         NextCap PC.
'
'Globals:None
'
'Input: oSupervisor - A reference to the Supevisor object.
'
'Return: Boolean - true if the retrieval resulted in no
'        errors.
'
'Modifications:
'   11-18-1998 As written for Pass1.4
'
'
'=======================================================
Public Function RecvStations() As Boolean
Dim vrtArray As Variant
Dim lngItem As Long

On Error GoTo RecvStations_Err
    '/*ensure this is set
    If Not (go_Supervisor Is Nothing) Then
        ShowFeedback "Status: Retrieving LotManagers"
        
        vrtArray = go_Supervisor.GetStations(mdlUtil.CurrentMachineName)
        '/*An array of dimension (0,0) indicates an error occured.
        If vrtArray(0, 0) = cServerError Then
            '/*Route the error through the centralized handler
            'ReportServerError "mdlMain.RecvStations", vrtArray
        Else
            '/*Loop through each item
            For lngItem = 0 To UBound(vrtArray, 2)
                '/*add the station key
                '/*and create a new instance of a frmLotManagerSink
                '/*These allows for a by pointer collection
                '/*of LotManagers
                GenerateLotManager CStr(vrtArray(0, lngItem)), _
                                     CInt(vrtArray(1, lngItem)), _
                                     CStr(vrtArray(2, lngItem))
            Next lngItem
            '/*Set the result value
            If lngItem > 0 Then
                Set go_ActiveLotManager = gcol_LotManagers(1).GetLotManager()
                gn_LotManagerIndex = 1
                RecvStations = True
            End If
        End If
    End If
    ShowFeedback "Status: LotManagers Fetch Completed"
Exit Function
'........................................................
RecvStations_Err:
    logToFile g_sErrorLog, "mdlUtil.RecvStations:" & Err.Description & "-" & CStr(Err.Number)
    On Error GoTo 0
    Exit Function
End Function


'
'=======================================================
'Routine: mdlMain.GenerateLotManager()
'Purpose: This attaches to the LotManager and sets the
'current global LotManager object. The Context for the
'Open Lot is also set at this time.
'
'Globals: go_ActiveLotManager - The global LotManager object.
'
'Input:None
'
'Return: LotManager - A reference to a LotManager
'        connection on the business server.
'
'Tested: hand tested 1-6-1999. Always returned the same
'        lot manager regardless of the parameters passed.
'
'Modifications:
'   01-06-1999 As written for Pass1.6
'
'
'=======================================================
Public Function GenerateLotManager(ByVal strLine As String, ByVal nLineNumber As Integer, ByVal strSource As String) As NextCapServer.clsLotManager
Dim oTemporary As NextCapServer.clsLotManager
Dim frmTemporary As frmLotManagerSink

On Error GoTo GenerateLotManager_Err
    '/*Insure that the Supervisor object is initilaized
    If Not (go_Supervisor Is Nothing) Then
        '/*Attach the Lot Manager handle
        Set oTemporary = go_Supervisor.GetLotManager(strLine, nLineNumber, strSource)
        '/*Make sure we recieved a valid reference
        If Not (oTemporary Is Nothing) Then
            '/*Generate a new form
            Set frmTemporary = New frmLotManagerSink
            '/*Make sure that the form is functional
            '/*This will trigger the open lot and transfer to the Context object
            Load frmTemporary
            '/*Now set the instance of the LotManager
            frmTemporary.SetLotManager oTemporary
            '/*Set the id of the manager
            frmTemporary.Source = strSource
            frmTemporary.LineType = strLine
            frmTemporary.LineNumber = nLineNumber
            frmTemporary.LotManagerName = "INST" & vtoa(gcol_LotManagers.Count + 1)
            '/*Add the Form to the global collection
            '/*of business server sinks so that we
            '/*have an Index formed for searching
            gcol_LotManagers.Add frmTemporary
        End If
    End If

    '/*Destroy the instance
    Set oTemporary = Nothing
Exit Function
'.....................................................................
GenerateLotManager_Err:
    logToFile g_sErrorLog, "mdlUtil.GenerateLotManager:" & Err.Description & "-" & CStr(Err.Number)
    Err.Clear
    Exit Function
End Function


'
'========================================================
'Routine: LogToFile(filename,msg)
'Purpose: This logs a string to a file and insures that
'the file size does not exceed 20k.
'
'Globals:None
'
'Input: sFile - the file name to log to
'       sMsg  - the string to write to the file
'
'Return:None
'
'Tested:
'
'Modifications:
'   01-24-2000 As written
'
'
'=======================================================
Public Sub logToFile(sFile As String, sMsg As String)
Dim nFp As Integer
Const nFileMax As Integer = 20000
Dim sDir As String

On Error GoTo LogToFile_Err

    '/*Location where file will exist
    sDir = App.Path & "\" & sFile
    If FileExist(sDir) Then
        If FileLen(sDir) > nFileMax Then
            Kill sDir
        End If
    End If

    '/*Output to file
    nFp = FreeFile()
    Open sDir For Append As #nFp
        Print #nFp, Format(Now(), "mm-dd hh:mm ") & sMsg
    Close #nFp
LogToFile_Err:
    Err.Clear
    Exit Sub
End Sub
'
'=======================================================
'Routine: vtoa(vrt)
'Purpose: Converts a varinat to string in a safe manner.
'
'Globals:None
'
'Input: vrtIn - the varinat to convert.
'
'Return: String - the resulting string.
'
'Tested:
'   01-05-1999 via mdlTestDebug script; OK -C Barker
'
'Modifications:
'   11-25-1998 As written for Pass1.5
'
'
'=======================================================
Public Function vtoa(ByRef vrtIn As Variant) As String
On Error GoTo vtoa_Err
    If IsNull(vrtIn) = False Then vtoa = CStr(vrtIn)
vtoa_Err:
    Err.Clear
    Exit Function
End Function
'
'=======================================================
'Routine: vtoi(vrt)
'Purpose: Converts a varinat to integer in a safe manner.
'
'Globals:None
'
'Input: vrtIn - the varinat to convert.
'
'Return: String - the resulting string.
'
'Tested:
'   01-05-1999 via mdlTestDebug script; OK -C Barker
'
'Modifications:
'   11-25-1998 As written for Pass1.5
'
'
'=======================================================
Public Function vtoi(ByRef vrtIn As Variant) As Integer
On Error GoTo vtoi_Err
    If IsNull(vrtIn) = False Then
        If IsNumeric(vrtIn) Then vtoi = CInt(vrtIn)
    End If
vtoi_Err:
    Err.Clear
    Exit Function
End Function
'
'=======================================================
'Routine: vtol(vrt)
'Purpose: Converts a varinat to Long in a safe manner.
'
'Globals:None
'
'Input: vrtIn - the varinat to convert.
'
'Return: Long - the resulting Long.
'
'Tested:
'   01-05-1999 via mdlTestDebug script; OK -C Barker
'
'Modifications:
'   11-25-1998 As written for Pass1.5
'
'
'=======================================================
Public Function vtol(ByRef vrtIn As Variant) As Long
On Error GoTo vtol_Err
    If IsNull(vrtIn) = False Then
        If IsNumeric(vrtIn) Then vtol = CLng(vrtIn)
    End If
vtol_Err:
    Err.Clear
    Exit Function
End Function

'
'=======================================================
'Routine: vtob(vrt)
'Purpose: Converts a varinat to boolean in a safe manner.
'
'Globals:None
'
'Input: vrtIn - the varinat to convert.
'
'Return: String - the resulting string.
'
'Tested:
'   01-05-1999 via mdlTestDebug script; OK -C Barker
'
'Modifications:
'   11-25-1998 As written for Pass1.5
'
'
'=======================================================
Public Function vtob(ByRef vrtIn As Variant) As Boolean
On Error GoTo vtob_Err
    If IsNull(vrtIn) = False Then vtob = CBool(vrtIn)
vtob_Err:
    Err.Clear
    Exit Function
End Function
'
'=======================================================
'Routine: showFeedback(s)
'Purpose: Sends message to Monitor form for
'display.
'
'=======================================================
Public Sub ShowFeedback(ByRef sMsg As String)
    If bVisible Then
        MsgBox sMsg
        'frmMonitor.lblMain.Caption = sMsg
        DoEvents
    End If
End Sub
'
'=======================================================
'Routine: mdlUtil.Shutdown()
'Purpose: execute the required shutdown procedure.
'
'
'=======================================================
Public Sub Shutdown(Optional ByRef sForm As Form)
Dim oForm As Form

    If gn_ShutdownCount > 1 Then Exit Sub
    
    For Each oForm In Forms
        If oForm.Name <> sForm.Name Then
            Unload oForm
        End If
    Next oForm
End Sub
'
'=======================================================
'Routine: FileExist()
'Purpose: Test to see if a file exists at a given
'         pathway.
'
'Globals:None
'
'Input:None
'
'Return: Boolean - True = File Exists
'                  False = No such file
'
'Tested:
'   11-24-1998 mdlTools
'
'Modifications:
'   9-24-1998 As written for Pass1.0
'
'   06-23-1999 [Bug] - This was returning true when
'   a string such as c:\temp\ was passed (because it
'   is a valid directory) so we need to make sure
'   that the incoming string is trimed and that it
'   does not end in "\".
'=======================================================
Public Function FileExist(ByVal strFile As String) As Boolean
On Error GoTo FileExist_Err
    '/*Make sure the string is for a file and not
    '/*a directory
    strFile = Trim$(strFile)
    If Len(strFile) > 0 Then
        If Right$(strFile, 1) = "\" Then
            '/*This will error if the file len=1, but
            '/*that is fine since it would not be a
            '/*valid file in this case anyway.
            strFile = Left$(strFile, Len(strFile) - 1)
        End If
    End If
    '/*Test the actual file
    If Len(strFile) > 0 And Len(Dir(strFile, vbNormal)) Then FileExist = True
FileExist_Err:
    Err.Clear
    Exit Function
End Function


Public Function UUBound(a As Variant) As Long
' Unfortunately this doesn't work with UDT's
' since they can't be passed as variant parameters outside of a class
'
    On Error Resume Next
    Dim size As Long
    size = UBound(a)
    If Err = 9 Then
        UUBound = -1
    Else
        UUBound = size
    End If
End Function

Public Function ReverseText(text As String) As String
' reverses a string
    Dim TextLen As Integer
    Dim i As Integer
    TextLen = Len(text)
    For i = TextLen To 1 Step -1
        ReverseText = ReverseText & Mid(text, i, 1)
    Next i

End Function


Public Function CompactRepair(sDestination As String)
  ' Compacts and Repairs a Microsoft Access Database
  ' Function receives a string equal to the full path and filename of the DB
  Dim JRO       As New JRO.JetEngine    ' jet engine object
  Dim sTempFile As String               ' temp file
  Dim sExt      As String               ' extension of the file

  ' get the file extension
  sExt = Right$(sDestination, Len(sDestination) - InStrRev(sDestination, ".") - 1)
  ' set the temp file name
  sTempFile = Left$(sDestination, Len(sDestination) - Len(sExt)) & "-Compacting" & sExt
  ' rename the destination to the temp file name
  Name sDestination As sTempFile

  ' compact the old database into the destination database file name
  JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sTempFile, _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDestination & ";Jet OLEDB:Engine Type=5"
  ' remove the temp file
  Kill sTempFile
End Function




