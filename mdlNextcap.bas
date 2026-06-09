Attribute VB_Name = "mdlNextcap"
Option Explicit

'/* Speed up the use of this string by making it constant
Private Const c_strGroupId = "UNKNOWN"
'/*Same for this date used in retrieving pens
Private Const c_dtBirthDay = #1/1/1980#

Public mutexCounter As Integer

'/*Constant to tune milliseconds of delay for
'/*failed calls to add pen.
'/* 01-18-2000 - Set at 2 seconds
Private Const m_lSecondsWait As Long = 2000

'/*This flags when the Database is offline for maintanance
Public g_dbLocked As Boolean
'

Public Sub checkForNextcapLock(Optional PollIntervalMS As Long = 250)
    'if the db is locked wait until it is unlocked.
    Dim dummy As Integer
    Dim i As Integer
    
    If g_dbLocked Then
        LogEvent ("NEXTCAP:checkForNextcapLock: Nextcap is Busy - iLink waiting for it to finish. ")
        
        While g_dbLocked
            Sleep (PollIntervalMS)
            For i = 1 To 5
                dummy = DoEvents
            Next i
        Wend
        LogEvent ("NEXTCAP:checkForNextcapLock: Nextcap ready - resuming iLink. ")
    End If
End Sub


Public Function UpdatePen(sPenID As String, sFailcode As String) As String
On Error GoTo EH
    Dim bFound As Boolean
    Dim nIndex As Integer
    Dim vrtPen As Variant
    Dim sError As String
    Dim sCurrentFailcode As String
    Dim iOldMutex
    
    checkForNextcapLock
   
    'STEP 1 - Connect to a Lot Manager and Check the Pen out.
    iOldMutex = mutexCounter
    mutexCounter = mutexCounter + 1
   
    'assume all is good
    sError = "OK"

    'CHECK 1 - Make sure that we're not already trying to do this
    If mutexCounter > 1 Then
        LogEvent ("NEXTCAP:UpdatePen: routine was entered before last run finished, or an error kicked us out.")
        sError = "Function overlap"
        GoTo FINISHED
    End If
    
    bFound = RetrievePen(sPenID, vrtPen)
    
    checkForNextcapLock
    
    If Not bFound Then
        For nIndex = 1 To gcol_LotManagers.Count
        '/*Make sure we are not repeating the same LotManager
            If nIndex <> gn_LotManagerIndex Then
                Set go_ActiveLotManager = gcol_LotManagers(nIndex).GetLotManager
                gn_LotManagerIndex = nIndex
                '/*If found bail on the loop
                checkForNextcapLock
                bFound = RetrievePen(sPenID, vrtPen)
                If bFound Then Exit For
            End If
        Next nIndex
    End If
        
    'CHECK 2 - Make sure that we have a pen
    If Not bFound Then
        LogEvent ("NEXTCAP:UpdatePen ERROR " & sPenID & ". Pen Not in Nextcap.")
        sError = "NotInCap or Duplicate"
        GoTo FINISHED
    End If
    
    
    'CHECK 3 - Make sure it's a valid pen object
    If Not UBound(vrtPen) > 18 Then
        LogEvent ("Incorrect Pen Array Size Retrieved from CAP for Pen: " & sPenID)
        sError = "ArraySize"
        GoTo FINISHED
    End If
    
    'Check for any previous inspector failcodes
    sCurrentFailcode = getCurrentInspectorFailcode(vrtPen)

    If sCurrentFailcode = sFailcode Then
        'no change
        LogEvent ("Pen status not changed - no need to update NextCAP: " & sPenID)
        If Not ReleaseUnit(vrtPen) Then
            sError = "ReleaseUnit"
        End If
        GoTo FINISHED
    ElseIf sFailcode = g_goodPenCode Then
        'make pen good
        LogEvent ("Removing Inspector Defects From: " & sPenID)
        If Not removeAllInspectorDefects(vrtPen) Then
            sError = "removeAllInspectorDefects"
        End If
    ElseIf sCurrentFailcode = g_goodPenCode Then
            'add new defect
            LogEvent ("Adding New Defect(" & sFailcode & ") To: " & sPenID)
            If Not addNewInspectorDefect(vrtPen, sFailcode) Then
                sError = "addNewInspectorDefect"
            End If
    Else
            LogEvent ("Updating existing defect on " & sPenID & " to:" & sFailcode)
            'update existing defect
            If Not updExistingInspectorDefect(vrtPen, sFailcode) Then
                sError = "updExistingInspectorDefect"
            End If
    End If
            
    checkForNextcapLock
    If (SendPenBack(vrtPen)) Then
        LogEvent ("Update OK.")
    Else
        LogEvent ("Update Error.")
        sError = "SendPenBack"
    End If


FINISHED:
    UpdatePen = sError
    mutexCounter = mutexCounter - 1
    
    Exit Function

EH:
    mutexCounter = iOldMutex
    If sError = "SendPenBack" Or sError = "ReleaseUnit" Or sError = "ArraySize" Then
        ReleaseUnit vrtPen
    End If
    UpdatePen = " UnhandledException: " & sError & Err.Number & " - " & Err.Description
    LogEvent "** ERROR : UpdatePen - " & Err.Number & " - " & Err.Description, True
End Function


Private Function getCurrentInspectorFailcode(vrtPen As Variant) As String
    On Error GoTo EH
    Dim iNumDefects As Integer
    Dim sReturn As String
    Dim i As Integer
        
    sReturn = ""
    iNumDefects = (UBound(vrtPen) - 20) / 9

    If iNumDefects <> 0 Then
        For i = 0 To iNumDefects
            If vrtPen(24 + (i - 1) * 9, 0) = CN_ILINK_COMMENT Then
                sReturn = vrtPen(25 + g_capDefectLevel + (i - 1) * 9, 0)
            End If
        Next i
    End If
        
    If sReturn = "" Then 'no inspector defects found
        sReturn = g_goodPenCode
    End If
    
    getCurrentInspectorFailcode = sReturn
    
    Exit Function
EH:
    LogEvent "** ERROR : getCurrentInspectorFailcode - " & Err.Number & " - " & Err.Description, True
End Function


Private Function getNextcapClassParentCode(ByVal sFailcode As String, iDefectLevel As Integer, ByRef sParentCode As String) As String
'-------------------------------------------------------------------
' Compiles a list of the pens in the currently open nextcap Lots
'-------------------------------------------------------------------
    On Error GoTo EH
    Dim cnxDb As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim sClass As String

    
    'Set up the Database Connection
    mdlNextcap.checkForNextcapLock 'check that we're not doing an upload
    Set cnxDb = New ADODB.Connection
    cnxDb.Mode = adModeRead
    cnxDb.CursorLocation = adUseClient
    cnxDb.Open g_NextcapConnectionString
    
        
    If iDefectLevel = 1 Then
        'Level one failcode
        sSQL = "SELECT top 1 ClassName FROM levelonedescriptions WHERE code1='" & sFailcode & "'"
    Else
        'Level two failcode
        sSQL = "SELECT top 1 l1.ClassName, l2.Code1 FROM leveltwodescriptions l2, levelonedescriptions l1 WHERE l2.code2='" & sFailcode & "' and l1.Code1 = l2.Code1 and l1.LineType=l2.LineType and l1.linenumber=l2.linenumber and l1.Source=l2.Source"
    End If
    

    Set rs = cnxDb.Execute(sSQL)
    If Not rs.EOF Then
        If iDefectLevel = 1 Then
            'Level one failcode
            sClass = rs.Fields("ClassName")
        Else
            'Level two failcode
            sClass = rs.Fields("ClassName")
            sParentCode = rs.Fields("Code1")
        End If
    End If
    Set rs = Nothing
    cnxDb.Close
    Set cnxDb = Nothing
    
    getNextcapClassParentCode = sClass
    Exit Function
EH:
    LogEvent "** ERROR : getNextcapClassParentCode - " & Err.Number & " - " & Err.Description, True
End Function


Private Function updExistingInspectorDefect(ByRef vrtPen As Variant, sFailcode As String) As Boolean
    'Warning you MUST have verified that the pen has an inspector failcode before you call this function
    On Error GoTo EH
    Dim iNumDefects As Integer
    Dim bReturn As Boolean
    Dim i As Integer
    Dim sClass As String
    Dim sCode1 As String
    Dim sCode2 As String
        
    'queries nextcap for defect level + code1
    sClass = getNextcapClassParentCode(sFailcode, g_capDefectLevel, sCode1)
    If g_capDefectLevel = 1 Then
        sCode1 = sFailcode
    Else
        sCode2 = sFailcode
    End If
    
        
    bReturn = False
    iNumDefects = (UBound(vrtPen) - 20) / 9
      
    If iNumDefects > 0 Then
        For i = 0 To iNumDefects
            vrtPen(23 + (i - 1) * 9, 0) = False
           If vrtPen(24 + (i - 1) * 9, 0) = CN_ILINK_COMMENT Then
                 vrtPen(22 + (i - 1) * 9, 0) = sClass
                 vrtPen(26 + (i - 1) * 9, 0) = sCode1
                 vrtPen(27 + (i - 1) * 9, 0) = sCode2
                 vrtPen(23 + (i - 1) * 9, 0) = True
                bReturn = True
            End If
        Next i
    End If
    
    updExistingInspectorDefect = bReturn
    Exit Function
EH:
    updExistingInspectorDefect = False
    LogEvent "** ERROR : updExistingInspectorDefect - " & Err.Number & " - " & Err.Description, True
End Function



Private Function addNewInspectorDefect(ByRef vrtPen As Variant, sFailcode As String) As Boolean
    On Error GoTo EH
    Dim iNumDefects As Integer
    Dim bReturn As Boolean
    Dim vrtReturn() As Variant
    Dim i As Integer
    Dim sClass As String
    Dim sCode1 As String
    Dim sCode2 As String
    Dim iArraysize As Integer
    
    'queries nextcap for defect level + code1
    sClass = getNextcapClassParentCode(sFailcode, g_capDefectLevel, sCode1)
    If g_capDefectLevel = 1 Then
        sCode1 = sFailcode
    Else
        sCode2 = sFailcode
    End If
        
    bReturn = False


    iArraysize = UBound(vrtPen)
    ReDim vrtReturn(iArraysize + 9, 0)
    
    'add the existing defects to our new array
    For i = 0 To iArraysize
        vrtReturn(i, 0) = vrtPen(i, 0)
    Next i
    
    iNumDefects = (iArraysize - 20) / 9
    'first make sure there are no other primary defects
    If iNumDefects > 0 Then
        For i = 0 To iNumDefects
            vrtReturn(23 + (i) * 9, 0) = False
        Next i
    End If
    
   
    vrtReturn(22 + ((iNumDefects) * 9), 0) = sClass
    vrtReturn(23 + ((iNumDefects) * 9), 0) = True 'make the new defect the primary
    vrtReturn(24 + ((iNumDefects) * 9), 0) = CN_ILINK_COMMENT
    vrtReturn(26 + ((iNumDefects) * 9), 0) = sCode1
    vrtReturn(27 + ((iNumDefects) * 9), 0) = sCode2
    vrtReturn(0, 0) = vrtPen(0, 0) + 1 'increment the failcode count
    
    vrtPen = vrtReturn
    bReturn = True
    
    addNewInspectorDefect = bReturn
    
    Exit Function
EH:
    addNewInspectorDefect = bReturn
    LogEvent "** ERROR : addNewInspectorDefect - " & Err.Number & " - " & Err.Description, True
End Function




Private Function removeAllInspectorDefects(ByRef vrtPen As Variant) As Boolean
    On Error GoTo EH
    Dim iNumDefects As Integer
    Dim bReturn As Boolean
    Dim vrtReturn() As Variant
    Dim i As Integer
    Dim sClass As String
    Dim sCode1 As String
    Dim iArraysize As Integer
    Dim iInspectorDefects As Integer
    Dim iNewDefect As Integer
    Dim iDefectData As Integer
         
    bReturn = False

    iArraysize = UBound(vrtPen)
    
    'count the number of inspector defects on the pen
    iNumDefects = (UBound(vrtPen) - 20) / 9
    iInspectorDefects = 0

    If iNumDefects <> 0 Then
        For i = 0 To iNumDefects
            If vrtPen(24 + (i - 1) * 9, 0) = CN_ILINK_COMMENT Then
                iInspectorDefects = iInspectorDefects + 1
            End If
        Next i
    End If
        
    'create a variant array of the correct size
    ReDim vrtReturn(iArraysize - (iInspectorDefects * 9), 0)

    'add the basic pen information
    For i = 0 To 20
        vrtReturn(i, 0) = vrtPen(i, 0)
    Next i
    vrtReturn(0, 0) = vrtReturn(0, 0) - iInspectorDefects
    
    'add the non-inspector defects
    iNewDefect = 0
    If iNumDefects <> 0 Then
        For i = 1 To iNumDefects
            If vrtPen(24 + (i - 1) * 9, 0) <> CN_ILINK_COMMENT Then
                iNewDefect = iNewDefect + 1
                For iDefectData = 1 To 9
                    vrtReturn(20 + iDefectData + (iNewDefect - 1) * 9, 0) = vrtPen(20 + iDefectData + (i - 1) * 9, 0)
                Next iDefectData
            End If
        Next i
    End If
    
    If iNewDefect > 0 Then 'if we have a defect, make sure at least one defect is primary
        vrtReturn(23, 0) = True
    End If
    
    vrtPen = vrtReturn
    bReturn = True
    
    removeAllInspectorDefects = bReturn
    
    Exit Function
EH:
    removeAllInspectorDefects = bReturn
    LogEvent "** ERROR : removeAllInspectorDefects - " & Err.Number & " - " & Err.Description, True
End Function




'=======================================================
'Routine: mdlCreatePen.RetrievePen(str,dt,str)
'Purpose: This requests an already entered Unit back
'from the B:server for editing or deleteing.
'
'Globals:None
'
'Input: strId - The unique ID of the Unit.
'
'       dtBirthday - The birth date of the Units
'       group/Lot
'
'       strGroupId - The name of the group that this
'       Unit is in.
'
'Return: vrtPen - The variant array version
'        of a Pen.
'
'
'Modifications:
'   10-14-1998 As written for Pass1.1
'
'   02-09-1999 Added switching BirthDay and GroupId
'   to the Active Lot in the Context if one is
'   not provided by the calling funciton.
'=======================================================
Public Function RetrievePen(ByRef strId As String, ByRef vrtPen As Variant) As Boolean
    On Error GoTo EH

    '/*Insure that there is some where to go to
    If go_ActiveLotManager Is Nothing Then
        '/*Do nothing for now
    Else
        '/*Request the unit from the bussiness server
        vrtPen = go_ActiveLotManager.GetPen(c_strGroupId, c_dtBirthDay, strId)
        
        '/*Insure that there were no errors fetching the pen
        If vrtPen(0, 0) <> cServerError Then
            '/*Unpack the retieved pen array
            RetrievePen = True
        Else
            RetrievePen = False
        End If
    End If
    Exit Function
EH:
    RetrievePen = False
    LogEvent "** ERROR : RetrievePen - " & Err.Number & " - " & Err.Description, True
End Function




'=======================================================
'Routine: mdlCreatePen.ReleaseUnit(o)
'Purpose: This makes the call to the Business Server
'         to release a pen that we had checked out
'         for editing.
'
'Globals:None
'
'Input: oClsPen - A pen object that we were handling.
'
'Return: Boolean - feedback the succes of the attempted
'        transaction.
'
'Modifications:
'   12-03-1998 As written for Pass1.5
'
'   01-18-2000 Placed sleep() call between consecutive
'   calls to release when the first one fails.
'=======================================================
Public Function ReleaseUnit(ByRef vrtPen As Variant) As Boolean
Dim strGroupId As String
Dim dtBirth As Date
Dim strId As String
Dim timeOut As Boolean, iCount As Integer

On Error GoTo EH
    '/*Set the needed items
    strGroupId = vrtPen(4, 0)  'The batch name
    dtBirth = vrtPen(5, 0) 'The date of birth for this unit
    strId = vrtPen(6, 0)  'The test/units unique id

    '/*Make sure we have some where to go with this
    If go_ActiveLotManager Is Nothing Then
        '/*Do nothing for now
    Else
        iCount = 0
        timeOut = False
        '/*Send out the relaese request via Global LotManager object
        Do While Not (go_ActiveLotManager.ReleasePen(strGroupId, dtBirth, strId)) And Not (timeOut)
            '/*Wait before second release attempt
            Sleep m_lSecondsWait
            iCount = iCount + 1

        Loop
        
        If Not (timeOut) Then
            ReleaseUnit = True
        Else
            ReleaseUnit = False
        End If

    End If
    Exit Function
    
EH:
    ReleaseUnit = False
    LogEvent "** ERROR : ReleaseUnit - " & Err.Number & " - " & Err.Description, True
End Function





Public Function SendPenBack(ByRef vrtArray As Variant) As Boolean
    
On Error GoTo EH

    '/*Make sure we have some where to go with this
    If go_ActiveLotManager Is Nothing Then
        '/*Do nothing for now
    Else
        '/*Send out pen data via Global LotManager object
        If Not (go_ActiveLotManager.UpdatePen(vrtArray)) Then
            '/*Wait
            Sleep m_lSecondsWait
            '/*An error occured logging the pen;
            '/*attempt to deal with it
            If go_ActiveLotManager.UpdatePen(vrtArray) Then
                '/*Return succes
                SendPenBack = True
            End If
        Else
            '/*Return succes
            SendPenBack = True
        End If
    End If
    Exit Function
    
EH:
    SendPenBack = False
    LogEvent "** ERROR : SendPenBack - " & Err.Number & " - " & Err.Description, True
End Function

