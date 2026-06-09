Attribute VB_Name = "mdlIniFiles"
' Windows API for reading INI files
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function ReadIniString(keyNm As String, DefVal As String, Optional sectionNm As String = "Missing", Optional iniNm As String = "Missing") As String
    Dim retStr As String * 512
    
    If sectionNm = "Missing" Then sectionNm = App.Title
    If iniNm = "Missing" Then iniNm = App.Path + "\" + App.Title + ".ini"
    Call GetPrivateProfileString(sectionNm, keyNm, DefVal, retStr, 512, iniNm)
    If InStr(retStr, Chr$(0)) Then retStr = Left$(retStr, InStr(retStr, Chr$(0)) - 1)
    ReadIniString = Trim$(retStr)
End Function
Public Sub WriteIniString(sectionNm As String, keyNm As String, KeyVal As String)
    Call WritePrivateProfileString(sectionNm, keyNm, KeyVal, App.Path + "\temp.ini")
End Sub

Public Function TestShare(sDirectoryName As String) As Boolean
    On Error GoTo EH
    If Len(sDirectoryName) <> 0 Then
        If Right(sDirectoryName, 1) <> "\" Then
            sDirectoryName = sDirectoryName & "\"
        End If
        Dir (sDirectoryName)
    TestShare = True
    Else
        GoTo EH
    End If
    Exit Function
EH:
    TestShare = False
End Function



Public Function AllFiles(ByVal DirPath As String) As String()
'***************************************************
'PURPOSE: RETURN AN ARRAY CONTAINING NAME OF ALL FILES IN
'DIRECTORY SPECIFIED BY DIR PATH

'PARAMETER:

  'DIRPATH: A VALID DRIVE OR SUBDIRECTORY ON YOUR SYSTEM,
  'ENDING WITH FORWARD SLASH (\) CHARACTER, OR
  'A DRIVE OR SUBDIRECTORY FOLLOWED BY A WILD CARD
  'STRING (e.g., C:\WINDOWS\*.txt)

'RETURNS: A STRING ARRAY WITH THE NAMES OF ALL FILENAMES
'IN THE DIRECTORY, INCLUDING HIDDEN, SYSTEM, AND READ-ONLY FILES
'THE FUNCTION IS NON RECURSIVE, I.E., IT DOES NOT SEARCH
'SUBDIRECTORIES UNDERNEATH DIRPATH

'EXAMPLE
'Dim sFiles() As String
'Dim lCtr As Long

'sFiles = AllFiles("C:\windows\")
'For lCtr = 0 To UBound(sFiles)
'    Debug.Print sFiles(lCtr)
'Next
'********************************************************

Dim sFile As String
Dim lElement As Long
Dim sAns() As String
ReDim sAns(0) As String

sFile = Dir(DirPath, vbNormal + vbHidden + vbReadOnly + _
   vbSystem + vbArchive)
If sFile <> "" Then
sAns(0) = sFile
    Do
        sFile = Dir
        If sFile = "" Then Exit Do
        lElement = IIf(sAns(0) = "", 0, UBound(sAns) + 1)
        ReDim Preserve sAns(lElement) As String
        sAns(lElement) = sFile
    Loop
End If
AllFiles = sAns
End Function
