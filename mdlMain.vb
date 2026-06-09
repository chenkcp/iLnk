Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports encryptxt
Imports IniParser
Imports IniParser.Model

Module mdlMain

    Public Const CN_ILINK_COMMENT = "Inspector Automated Operator"
    Public g_arrScanners() As ScannerBank
    Public m_fileExtension As String
    Public m_numIPASColumns As Integer
    Public m_bShareBanks As Boolean
    Public m_fileStyle As String
    Public m_sOpenNextcapPens As String
    Public g_logEventsToFile As Boolean
    Public g_loadToNextcap As Boolean
    Public g_refreshInterval As Integer
    Public g_goodPenCode As String
    Public m_numDaysData As Integer
    Public m_numCapRetryAttempts As Integer
    Public m_doubleCheckNextcap As Boolean
    Public m_numdaysLogFile As Integer
    Public g_capDefectLevel As Integer
    Public m_MfgDBType As String
    Public m_Debug As Boolean
    Public g_cnxDb As OdbcConnection
    Public mutexCounter As Integer

    Public g_NextcapConnectionString As String
    Public m_ILinkConnectionString As String
    Public m_GradebookConnectionString As String

    Public m_CheckForInspectorResults As Boolean
    Public m_CheckForEtesterResults As Boolean
    Public m_CheckForEtesterRegionFails As Boolean
    Public m_CheckForKahunaRegionFails As Boolean
    Public m_RegionCheckEngineeringLots As Boolean

    Public m_EtesterSQL As String
    Public m_CheckEtestGen1 As Boolean
    Public m_CheckEtestGen2 As Boolean
    Public m_KahunaSQL As String

    Public Class CapDecisionRule
        Public FailCodes As String
        Public SetFailCode As String
    End Class
    Public m_arrCapDecisionRules() As CapDecisionRule

    Public m_arrPriorityFailures() As String
    Public m_codesToIgnore As String

    Public Class TestLevel
        Public Property TestType As String
        Public Property FailcodePriorities() As String()
    End Class
    Public m_arrTestLevels() As TestLevel

    Public m_dctProdRefIDFetMktg As New Dictionary(Of String, String)
    Public m_dctProdRefIDFetProd As New Dictionary(Of String, String)
    Public m_dctProdRefVentLab As New Dictionary(Of String, String)

    Public Class ScannerBank
        Public Property Active As Boolean
        Public Property Path As String
        Public Property ScannerAlias As String
    End Class

    Public Structure TestResult
        Public Property TestType As String
        Public Property PrintSampleID As String
        Public Property PenID As String
        Public Property Comment As String
        Public Property Failcode As String
        Public Property ScanDatetime As Date?
        Public Property PageNumber As Integer
        Public Property DataFile As String
    End Structure

    Public Class PenResult
        Public Property PenID As String
        Public Property OverallFailcode As String
    End Class

    Public Sub Main()
        LoadSettings()
        CheckForNewData()
    End Sub

    Public Sub CheckForNewData()
        If dbConnect() Then
            'm_sOpenNextcapPens = getOpenNextcapPens()
            m_sOpenNextcapPens = "2757509304798700,2757509304798715,2757509304798799"

            ' read file, calculate fail code, load to nextcap
            checkForInspectorResults()
        End If

    End Sub

    Public Sub checkForInspectorResults()

        Dim TestResults() As TestResult
        Dim NewPens() As String
        Dim PenResults() As PenResult
        Dim iNumResults As Integer

        Try
            LogEvent("Checking for INSPECTOR Results... ")

            ' read and parse inspector data
            iNumResults = ReadDataFromBanks(TestResults)

            If iNumResults > 0 Then
                NewPens = getUniquePens(TestResults)
                'PenResults = setOverallFailmode(NewPens)  ' --- move from below block because addTestResults is false
                ' insert ReadDataFromBanks TestResults into table finishedData
                If addTestResults(TestResults) Then

                    ' delete source file from ReadDataFromBanks, avoid duplicate
                    DeleteDataFiles(TestResults)

                    ' calculate pen fail code
                    PenResults = setOverallFailmode(NewPens)

                    ' load to nextcap
                    'LoadResultsToNextCAP(PenResults)
                End If
            End If

        Catch ex As Exception
            LogEvent($"** ERROR : checkForInspectorResults - {ex.HResult} - {ex.Message}")
        End Try
    End Sub

    ' VB.NET 适配版（兼容 .NET Framework/.NET Core/.NET 5+）
    Private Function ReadDataFromBanks(ByRef TestResults() As TestResult) As Long
        '-------------------------------------------------------------------
        ' Checks all of the active banks for results
        '-------------------------------------------------------------------
        ' 7 fields
        ' 0 = psID, 1 = psLot, 2 = psFailCode, 3 = psPenId
        ' 4 = psPage, 5 & 6 = psDate

        ' 声明变量（VB6 类型映射 + .NET 类型优化）
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

        ' VB6 New Crypto → VB.NET 需确保 Crypto 类已迁移（后续补充实现）
        Dim CryptRef As New Crypto

        iNumResultsOverall = 0

        Try
            ' VB6 On Error Resume Next + UBound 检查 → VB.NET 安全数组检查
            Dim iArraysize As Integer = -1
            If g_arrScanners IsNot Nothing AndAlso g_arrScanners.Length > 0 Then
                iArraysize = g_arrScanners.GetUpperBound(0) ' 替代 UBound
            End If

            For i = 0 To iArraysize
                ' 遍历所有激活的扫描器
                If g_arrScanners(i).Active Then
                    If Not TestShare(g_arrScanners(i).Path) Then
                        ' VB6 MsgBox → VB.NET MessageBox（需引用 System.Windows.Forms）
                        MessageBox.Show($"{Application.ProductName} cannot contact scanner: {g_arrScanners(i).ScannerAlias} and am disabling it.{vbCrLf}{vbCrLf}Re-enable it manually in iLink when it is working.",
                                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        g_arrScanners(i).Active = False

                        ' VB6 ListItem → VB.NET ListViewItem（需适配 frmMain.lstScanners 控件）
                        'For Each lstItem As ListViewItem In frmMain.lstScanners.Items
                        '    If lstItem.Text = g_arrScanners(i).ScannerAlias Then
                        '        lstItem.Checked = False
                        '    End If
                        'Next
                    Else
                        ' 拼接文件路径（兼容 VB6 通配符）
                        sPath = Path.Combine(g_arrScanners(i).Path, $"*.{m_fileExtension}")
                        sFiles = AllFiles(sPath) ' 后续补充 AllFiles 方法实现

                        If sFiles IsNot Nothing AndAlso sFiles.Length > 0 AndAlso sFiles(0) <> "" Then
                            For iFileNumber = 0 To sFiles.GetUpperBound(0) ' 替代 UBound
                                sFullFilePath = Path.Combine(g_arrScanners(i).Path, sFiles(iFileNumber))
                                LogEvent($"Opening - {sFullFilePath}")

                                ' ========== VB6 文件操作 → .NET 安全文件读取 ==========
                                ' 方式1：兼容 FreeFile（不推荐，仅过渡用）
                                'iFreefile = FreeFile()
                                'FileOpen(iFreefile, sFullFilePath, OpenMode.Input)
                                'sFileData = InputString(iFreefile, LOF(iFreefile))
                                'FileClose(iFreefile)

                                ' 方式2：.NET 标准写法（推荐，更安全）
                                If File.Exists(sFullFilePath) Then
                                    sFileData = File.ReadAllText(sFullFilePath, Encoding.Default)
                                    ' 解密文件数据（Crypto 类需迁移）
                                    sFileData = CryptRef.DecryptString(sFileData)
                                End If

                                ' 清理文件数据格式
                                sFileData = Replace(AllTrim(sFileData), Chr(10), ",")
                                sFileData = Replace(AllTrim(sFileData), Chr(13), "")
                                If sFileData.EndsWith(",") Then
                                    sFileData = sFileData.Substring(0, sFileData.Length - 1)
                                End If

                                If String.IsNullOrEmpty(sFileData) Then
                                    sFieldData = New String() {}
                                Else
                                    sFieldData = sFileData.Split(New Char() {","c}, StringSplitOptions.None)
                                End If
                                iNumFieldsInFile = sFieldData.Length

                                ' 检查字段数量是否符合要求
                                If (iNumFieldsInFile Mod m_numIPASColumns) = 0 Then
                                    iNumResultsInFile = iNumFieldsInFile \ m_numIPASColumns ' 整数除法

                                    For iCtr = 0 To iNumResultsInFile - 1
                                        bIgnoreResult = False

                                        ' 检查是否需要忽略重复采集的结果
                                        If m_bShareBanks Then
                                            Select Case m_fileStyle
                                                Case "Wildflowers"
                                                    If InStr(m_sOpenNextcapPens, sFieldData(0 + (m_numIPASColumns * iCtr))) = 0 Then
                                                        bIgnoreResult = True
                                                    End If
                                                Case Else
                                                    If InStr(m_sOpenNextcapPens, sFieldData(3 + (m_numIPASColumns * iCtr))) = 0 Then
                                                        bIgnoreResult = True
                                                    End If
                                            End Select
                                        End If

                                        If Not bIgnoreResult Then
                                            iNumResultsOverall += 1
                                            ' 动态扩容数组（替代 ReDim Preserve）
                                            ReDim Preserve arrTestResult(iNumResultsOverall - 1)

                                            Select Case m_fileStyle
                                                Case "Wildflowers"
                                                    ' Wildflowers 格式数据解析
                                                    arrTestResult(iNumResultsOverall - 1).TestType = "WILDFLOWERS"
                                                    arrTestResult(iNumResultsOverall - 1).PrintSampleID = "NA"
                                                    arrTestResult(iNumResultsOverall - 1).Comment = "NA"
                                                    arrTestResult(iNumResultsOverall - 1).Failcode = sFieldData(1 + (m_numIPASColumns * iCtr))
                                                    arrTestResult(iNumResultsOverall - 1).PenID = sFieldData(0 + (m_numIPASColumns * iCtr))
                                                    arrTestResult(iNumResultsOverall - 1).PageNumber = 1
                                                    arrTestResult(iNumResultsOverall - 1).ScanDatetime = DateTime.Now
                                                Case Else
                                                    ' Inspector 格式数据解析
                                                    sFailcode = sFieldData(2 + (m_numIPASColumns * iCtr))
                                                    ' 判断测试类型（GLOSSY/PLAIN）
                                                    If sFailcode.Length > 3 AndAlso sFailcode.EndsWith("G") Then
                                                        sTestType = "GLOSSY"
                                                    Else
                                                        sTestType = "PLAIN"
                                                    End If

                                                    arrTestResult(iNumResultsOverall - 1).TestType = sTestType
                                                    arrTestResult(iNumResultsOverall - 1).PrintSampleID = sFieldData(0 + (m_numIPASColumns * iCtr))
                                                    arrTestResult(iNumResultsOverall - 1).Comment = sFieldData(1 + (m_numIPASColumns * iCtr))
                                                    arrTestResult(iNumResultsOverall - 1).Failcode = sFailcode
                                                    arrTestResult(iNumResultsOverall - 1).PenID = sFieldData(3 + (m_numIPASColumns * iCtr))
                                                    ' 类型转换：String → Integer
                                                    arrTestResult(iNumResultsOverall - 1).PageNumber = CInt(sFieldData(4 + (m_numIPASColumns * iCtr)))
                                                    ' 拼接日期时间并转换
                                                    Dim dateStr As String = $"{sFieldData(5 + (m_numIPASColumns * iCtr))} {sFieldData(6 + (m_numIPASColumns * iCtr))}"
                                                    arrTestResult(iNumResultsOverall - 1).ScanDatetime = CDate(dateStr)
                                            End Select
                                            arrTestResult(iNumResultsOverall - 1).DataFile = sFullFilePath
                                        End If
                                    Next iCtr
                                Else
                                    LogEvent($"** Warning! ** - File {sFullFilePath} has incorrect number of columns")
                                End If
                            Next iFileNumber
                        End If ' 有文件
                    End If ' 共享正常
                End If ' 扫描器激活
            Next i

            LogEvent($"Found {iNumResultsOverall} results.")
            If iNumResultsOverall > 0 Then
                TestResults = arrTestResult
                ReadDataFromBanks = iNumResultsOverall
            End If

        Catch ex As Exception
            ' VB6 Err.Number/Description → .NET Exception 属性
            LogEvent($"** ERROR : ReadDataFromBanks - {ex.HResult} - {ex.Message}", True)
            ReadDataFromBanks = 0
        End Try
    End Function

    Private Function AllTrim(ByVal s As String) As String
        '-------------------------------------------------------------------
        ' trims all leading and trailing non-alphanumeric characters from a string
        '-------------------------------------------------------------------
        If String.IsNullOrEmpty(s) Then
            Return ""
        End If

        ' check first character
        Dim startIndex As Integer = 0
        While startIndex < s.Length AndAlso Not Char.IsLetterOrDigit(s(startIndex))
            startIndex += 1
        End While

        ' check last character
        Dim endIndex As Integer = s.Length - 1
        While endIndex >= startIndex AndAlso Not Char.IsLetterOrDigit(s(endIndex))
            endIndex -= 1
        End While

        If startIndex > endIndex Then
            Return ""
        End If

        Return s.Substring(startIndex, endIndex - startIndex + 1)
    End Function


    Private Sub LogEvent(ByVal message As String, Optional ByVal isError As Boolean = False)
        Dim logLine As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {(If(isError, "ERROR: ", ""))}{message}"
        Console.WriteLine(logLine)
        Dim logPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "iLink.log")
        File.AppendAllText(logPath, logLine & Environment.NewLine, Encoding.Default)
    End Sub

    Public Function LoadSettings(Optional sIniFile As String = "") As Boolean
        '-------------------------------------------------------------------
        ' Loads settings from application ini file
        '-------------------------------------------------------------------
        Dim sRuleIniFile As String
        Dim i As Integer
        Dim j As Integer

        Try
            ' 1. ini path
            If String.IsNullOrEmpty(sIniFile) Then
                sIniFile = Path.Combine(Application.StartupPath, $"{Application.ProductName}.ini")
            End If

            LogEvent($"Reading general Settings from ini file - {sIniFile}")

            ' 2. get common config
            g_logEventsToFile = Convert.ToBoolean(mdlIniFiles.ReadIniString("LogEventsToFile", "False", "GENERAL", sIniFile))
            g_loadToNextcap = Convert.ToBoolean(mdlIniFiles.ReadIniString("LoadToNextcap", "False", "GENERAL", sIniFile))
            g_refreshInterval = Convert.ToInt32(mdlIniFiles.ReadIniString("RefreshInterval", "10", "GENERAL", sIniFile)) * 1000
            g_goodPenCode = mdlIniFiles.ReadIniString("GoodPenCode", "NTF", "GENERAL", sIniFile)
            m_numDaysData = Convert.ToInt32(mdlIniFiles.ReadIniString("NumDaysDataToStore", "12", "GENERAL", sIniFile))
            m_numCapRetryAttempts = Convert.ToInt32(mdlIniFiles.ReadIniString("NumCapRetryAttempts", "5", "GENERAL", sIniFile))
            m_doubleCheckNextcap = Convert.ToBoolean(mdlIniFiles.ReadIniString("DoubleCheckNextcap", "False", "GENERAL", sIniFile))
            m_numdaysLogFile = Convert.ToInt32(mdlIniFiles.ReadIniString("NumDaysLogFile", "3", "GENERAL", sIniFile))
            g_capDefectLevel = Convert.ToInt32(mdlIniFiles.ReadIniString("CapDefectLevel", "2", "GENERAL", sIniFile))
            sRuleIniFile = Path.Combine(Application.StartupPath, mdlIniFiles.ReadIniString("RuleIniFile", Application.ProductName & ".ini", "GENERAL", sIniFile))
            m_MfgDBType = mdlIniFiles.ReadIniString("MfgDBType", "INFORMIX", "GENERAL", sIniFile)
            m_Debug = Convert.ToBoolean(mdlIniFiles.ReadIniString("Debug", "False", "GENERAL", sIniFile))

            ' 3. get db info
            g_NextcapConnectionString = mdlIniFiles.ReadIniString("NextcapConnectionString", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\work\project\nextcap\NextCap2025\capui\bin\x64\Debug\nextcap.mdb;Uid=Admin;Pwd=;", "DATABASE", sIniFile)
            m_ILinkConnectionString = mdlIniFiles.ReadIniString("ILinkConnectionString", "Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\work\project\nextcap\iLink\iLink2026\iLink2026\bin\Debug\iLink.mdb;Uid=admin;Pwd=", "DATABASE", sIniFile)
            m_GradebookConnectionString = mdlIniFiles.ReadIniString("GradebookConnectionString", "File Name=c:\program files\nextcap\server\udls\mfg.udl", "DATABASE", sIniFile)

            ' 4. get test type
            m_CheckForInspectorResults = Convert.ToBoolean(mdlIniFiles.ReadIniString("CheckForInspectorResults", "False", "TestTypes", sIniFile))
            m_CheckForEtesterResults = Convert.ToBoolean(mdlIniFiles.ReadIniString("CheckForEtesterResults", "False", "TestTypes", sIniFile))
            m_CheckForEtesterRegionFails = Convert.ToBoolean(mdlIniFiles.ReadIniString("CheckForEtesterRegionFails", "False", "TestTypes", sIniFile))
            m_CheckForKahunaRegionFails = Convert.ToBoolean(mdlIniFiles.ReadIniString("CheckForKahunaRegionFails", "False", "TestTypes", sIniFile))
            m_RegionCheckEngineeringLots = Convert.ToBoolean(mdlIniFiles.ReadIniString("RegionCheckEngineeringLots", "True", "TestTypes", sIniFile))

            ' 5. get Inspector settings
            m_numIPASColumns = Convert.ToInt32(mdlIniFiles.ReadIniString("numOfColumns", "7", "INSPECTOR", sIniFile))
            m_fileExtension = mdlIniFiles.ReadIniString("FileExtension", "txt", "INSPECTOR", sIniFile)
            m_fileStyle = mdlIniFiles.ReadIniString("FileStyle", "Inspector", "INSPECTOR", sIniFile)
            m_bShareBanks = Convert.ToBoolean(mdlIniFiles.ReadIniString("ShareBanks", "False", "INSPECTOR", sIniFile))

            ' 6. 读取扫描器银行信息
            LogEvent($"Reading Bank information from ini file - {sIniFile}")
            Dim numScanners As Integer = Convert.ToInt32(mdlIniFiles.ReadIniString("numOfBanks", "0", "INSPECTOR", sIniFile))
            If numScanners > 0 Then
                ReDim g_arrScanners(numScanners - 1)
                For i = 0 To numScanners - 1
                    g_arrScanners(i) = New ScannerBank()
                    ' 读取银行路径并补全末尾 \
                    g_arrScanners(i).Path = mdlIniFiles.ReadIniString("Bank" & (i + 1), "Missing", "INSPECTOR", sIniFile)
                    If Not g_arrScanners(i).Path.EndsWith("\") Then
                        g_arrScanners(i).Path &= "\"
                    End If
                    ' 读取银行别名（处理 Missing 兜底）
                    g_arrScanners(i).ScannerAlias = mdlIniFiles.ReadIniString("Bank" & (i + 1) & "Alias", "Missing", "INSPECTOR", sIniFile)
                    If g_arrScanners(i).ScannerAlias = "Missing" Then
                        g_arrScanners(i).ScannerAlias = g_arrScanners(i).Path
                    End If
                    g_arrScanners(i).Active = True
                Next i
            End If

            ' 7. 读取 Etester 专属设置
            m_EtesterSQL = mdlIniFiles.ReadIniString("SQL", "", "ETESTER", sIniFile)
            m_CheckEtestGen1 = Convert.ToBoolean(mdlIniFiles.ReadIniString("CheckEtestGen1", "True", "ETESTER", sIniFile))
            m_CheckEtestGen2 = Convert.ToBoolean(mdlIniFiles.ReadIniString("CheckEtestGen2", "True", "ETESTER", sIniFile))

            ' 8. 读取 Kahuna 专属设置
            m_KahunaSQL = mdlIniFiles.ReadIniString("SQL", "", "KAHUNA", sIniFile)

            ' 9. 读取 CAP 决策规则
            LogEvent($"Reading CAP Decision rules from ini file - {sRuleIniFile}")
            Dim numRules As Integer = Convert.ToInt32(mdlIniFiles.ReadIniString("numOfRules", "0", "RULES", sRuleIniFile))
            If numRules > 0 Then
                ReDim m_arrCapDecisionRules(numRules - 1)
                For i = 0 To numRules - 1
                    m_arrCapDecisionRules(i) = New CapDecisionRule()
                    m_arrCapDecisionRules(i).FailCodes = mdlIniFiles.ReadIniString("Failcodes" & (i + 1), "Missing", "RULES", sRuleIniFile)
                    m_arrCapDecisionRules(i).SetFailCode = mdlIniFiles.ReadIniString("SetFailcode" & (i + 1), "Missing", "RULES", sRuleIniFile)
                Next i
            End If

            ' 10. 读取优先级失败模式
            LogEvent($"Reading Priority Failures from ini file - {sRuleIniFile}")
            Dim numPriorities As Integer = Convert.ToInt32(mdlIniFiles.ReadIniString("numOfPriorities", "0", "SpecialFailcodes", sRuleIniFile))
            If numPriorities > 0 Then
                ReDim m_arrPriorityFailures(numPriorities - 1)
                For i = 0 To numPriorities - 1
                    m_arrPriorityFailures(i) = $"'{mdlIniFiles.ReadIniString("Priority" & (i + 1), "Missing", "SpecialFailcodes", sRuleIniFile)}'"
                Next i
            End If
            m_codesToIgnore = Convert.ToString(mdlIniFiles.ReadIniString("CodesToIgnore", "", "SpecialFailcodes", sRuleIniFile))

            ' 11. 读取测试级别优先级
            LogEvent($"Reading Test Levels from ini file - {sRuleIniFile}")
            Dim numLevels As Integer = Convert.ToInt32(mdlIniFiles.ReadIniString("numTestLevels", "0", "RULES", sRuleIniFile))
            If numLevels > 0 Then
                ReDim m_arrTestLevels(numLevels - 1)
                For i = 0 To numLevels - 1
                    m_arrTestLevels(i) = New TestLevel()
                    m_arrTestLevels(i).TestType = mdlIniFiles.ReadIniString("TestLevel" & (i + 1), "Missing", "RULES", sRuleIniFile)
                    numPriorities = Convert.ToInt32(mdlIniFiles.ReadIniString("numOfPriorities", "0", m_arrTestLevels(i).TestType, sRuleIniFile))
                    If numPriorities > 0 Then
                        ReDim m_arrTestLevels(i).FailcodePriorities(numPriorities - 1)
                        For j = 0 To numPriorities - 1
                            m_arrTestLevels(i).FailcodePriorities(j) = $"'{mdlIniFiles.ReadIniString("Priority" & (j + 1), "Missing", m_arrTestLevels(i).TestType, sRuleIniFile)}'"
                        Next j
                    End If
                Next i
            End If

            ' 12. 读取产品参考信息（替换 VB6 ADODB → .NET OleDb）
            If m_CheckForEtesterRegionFails Or m_CheckForKahunaRegionFails Then
                ' 1. 替换 OleDbConnection 为 OdbcConnection
                Using cnxDb As New OdbcConnection(m_GradebookConnectionString)
                    Try
                        cnxDb.Open()
                        Dim sql As String = "select * from product_ref_llk"

                        ' 2. 修复语法错误：OleDbConnection → OdbcCommand（原代码笔误）
                        Using cmd As New OdbcCommand(sql, cnxDb)
                            ' 3. 替换 OleDbDataReader 为 OdbcDataReader
                            Using rs As OdbcDataReader = cmd.ExecuteReader()
                                ' 遍历结果集（保持原有逻辑）
                                While rs.Read()
                                    ' 读取字段值（兼容 DBNull 处理）
                                    Dim sProductNumber As String = If(rs.IsDBNull(rs.GetOrdinal("inv_item_lk_nr")), "", Convert.ToString(rs("inv_item_lk_nr")))

                                    ' 写入字典（避免空键报错）
                                    If Not String.IsNullOrEmpty(sProductNumber) Then
                                        m_dctProdRefIDFetMktg(sProductNumber) = If(rs.IsDBNull(rs.GetOrdinal("id_fet_marketing")), "", Convert.ToString(rs("id_fet_marketing")))
                                        m_dctProdRefIDFetProd(sProductNumber) = If(rs.IsDBNull(rs.GetOrdinal("id_fet_product")), "", Convert.ToString(rs("id_fet_product")))
                                        m_dctProdRefVentLab(sProductNumber) = If(rs.IsDBNull(rs.GetOrdinal("pica_cd")), "", Convert.ToString(rs("pica_cd")))
                                    End If
                                End While
                            End Using
                        End Using
                    Catch ex As Exception
                        ' 新增异常处理（可选）
                        Debug.WriteLine($"查询 product_ref_llk 失败：{ex.Message}")
                    End Try
                End Using ' Using 自动释放连接，无需手动 Close/Dispose
            End If

            LoadSettings = True
        Catch ex As Exception
            ' 替换 VB6 Err.Number/Description → .NET Exception 属性
            LoadSettings = False
            LogEvent($"** ERROR : LoadSettings - {ex.HResult} - {ex.Message}", True)
        End Try
    End Function

    Private Function getUniquePens(TestResults() As TestResult) As String()
        '-------------------------------------------------------------------
        ' Returns an array of unique pen_id's from a set of results
        '-------------------------------------------------------------------
        Try
            ' 1. 空数组/空引用容错（替代 VB6 On Error Resume Next + UBound 检查）
            If TestResults Is Nothing OrElse TestResults.Length = 0 Then
                Return New String() {} ' 返回空数组，避免后续报错
            End If

            ' 方式1：高效版本（推荐，使用 HashSet 自动去重，性能远优于字符串拼接）
            Dim uniquePens As New HashSet(Of String)()
            For Each result As TestResult In TestResults
                ' 过滤空的 PenID，避免无效数据
                If Not String.IsNullOrEmpty(result.PenID) Then
                    uniquePens.Add(result.PenID)
                End If
            Next

            ' 将 HashSet 转换为字符串数组（返回结果）
            Return uniquePens.ToArray()

            ' 方式2：完全匹配原 VB6 逻辑（字符串拼接+Split，仅兼容用）
            ' Dim sFoundPens As String = ""
            ' For i As Integer = 0 To TestResults.GetUpperBound(0)
            '     Dim sPenID As String = TestResults(i).PenID
            '     If Not String.IsNullOrEmpty(sPenID) AndAlso Not sFoundPens.Contains(sPenID & "|") Then
            '         sFoundPens &= sPenID & "|"
            '     End If
            ' Next

            ' If sFoundPens.Length > 0 Then
            '     ' 移除最后一个 |
            '     sFoundPens = sFoundPens.Substring(0, sFoundPens.Length - 1)
            ' End If

            ' ' 拆分字符串为数组（匹配原 VB6 Split 逻辑）
            ' If String.IsNullOrEmpty(sFoundPens) Then
            '     Return New String() {}
            ' Else
            '     Return sFoundPens.Split("|"c, StringSplitOptions.RemoveEmptyEntries)
            ' End If

        Catch ex As Exception
            ' 替换 VB6 Err.Number/Description → .NET Exception 属性
            LogEvent($"** ERROR : getUniquePens - {ex.HResult} - {ex.Message}", True)
            ' 异常时返回空数组，避免调用方崩溃
            Return New String() {}
        End Try
    End Function

    Private Function addTestResults(TestResults() As TestResult) As Boolean
        '-------------------------------------------------------------------
        ' Add new results to iLink database
        ' Am using Stuffer database schema for backward compatibility
        '-------------------------------------------------------------------
        Dim i As Integer
        Dim sSQL As String = ""

        Try
            ' 1. 空数组/空引用容错（替代 VB6 On Error Resume Next + UBound 检查）
            If TestResults Is Nothing OrElse TestResults.Length = 0 Then
                LogEvent("No test results to insert - array is empty or null.")
                addTestResults = True ' 无数据视为执行成功
                Return addTestResults
            End If

            ' 2. 验证数据库连接是否有效
            If g_cnxDb Is Nothing OrElse g_cnxDb.State <> ConnectionState.Open Then
                Throw New Exception("Database connection is not open")
            End If

            ' 3. 遍历测试结果，批量插入（使用参数化查询防 SQL 注入）
            For i = 0 To TestResults.GetUpperBound(0)
                Dim result As TestResult = TestResults(i)

                ' 过滤无失败码的无效结果（保留原逻辑）
                If Not String.IsNullOrEmpty(result.Failcode) Then
                    LogEvent($"Inserting ps: {result.PrintSampleID} pen_id: {result.PenID} " &
                             $"Failcode: {result.Failcode} to finishedData Table.")

                    ' 【关键修复】参数化 SQL（替代字符串拼接，防注入+避免特殊字符报错）
                    sSQL = "INSERT INTO finishedData(
                            test_type, print_sample_id, pen_id, lot_id, 
                            pen_birthdate, scan_birthdate, pen_failcode, page_count
                        ) VALUES (
                            ?, ?, ?, ?,
                            ?, ?, ?, ?
                        );"

                    ' 创建命令对象（Using 自动释放资源）
                    Using cmd As New OdbcCommand(sSQL, g_cnxDb)
                        ' ========== 关键适配：ODBC 不支持 @ 命名参数，需先修改 SQL 语句中的参数占位符为 ? ==========
                        ' 假设你的原始 sSQL 是：INSERT INTO table (TestType, PrintSampleID, ...) VALUES (@TestType, @PrintSampleID, ...)
                        ' 需先替换为：INSERT INTO table (TestType, PrintSampleID, ...) VALUES (?, ?, ...)
                        ' （参数顺序必须和 Add 顺序完全一致！）

                        ' 1. 添加参数（ODBC 按顺序匹配，无需参数名，仅需类型和值）
                        ' 注意：Add 顺序必须和 SQL 中 ? 的顺序完全一致！
                        cmd.Parameters.Add(CreateOdbcParameter(If(String.IsNullOrEmpty(result.TestType), DBNull.Value, result.TestType), OdbcType.VarChar))
                        cmd.Parameters.Add(CreateOdbcParameter(If(String.IsNullOrEmpty(result.PrintSampleID), DBNull.Value, result.PrintSampleID), OdbcType.VarChar))
                        cmd.Parameters.Add(CreateOdbcParameter(If(String.IsNullOrEmpty(result.PenID), DBNull.Value, result.PenID), OdbcType.VarChar))
                        cmd.Parameters.Add(CreateOdbcParameter(If(String.IsNullOrEmpty(result.Comment), DBNull.Value, result.Comment), OdbcType.VarChar)) ' lot_id 映射到 Comment
                        cmd.Parameters.Add(CreateOdbcParameter(DateTime.Now, OdbcType.DateTime)) ' pen_birthdate = Now()
                        cmd.Parameters.Add(CreateOdbcParameter(If(result.ScanDatetime = DateTime.MinValue, DBNull.Value, result.ScanDatetime), OdbcType.DateTime))
                        cmd.Parameters.Add(CreateOdbcParameter(result.Failcode, OdbcType.VarChar))
                        cmd.Parameters.Add(CreateOdbcParameter(If(result.PageNumber = 0, DBNull.Value, result.PageNumber), OdbcType.Int))

                        ' 2. 执行插入（ExecuteNonQuery 用法不变）
                        cmd.ExecuteNonQuery()
                    End Using
                End If
            Next i

            addTestResults = True
        Catch ex As Exception
            ' 替换 VB6 Err 信息，记录错误 SQL（若有）
            LogEvent($"** ERROR : addTestResults - {ex.HResult} - {ex.Message} - SQL: {sSQL}", True)
            addTestResults = False
        End Try

        Return addTestResults
    End Function

    Private Function CreateOdbcParameter(value As Object, odbcType As OdbcType) As OdbcParameter
        Dim param As New OdbcParameter() With {
        .OdbcType = odbcType,
        .Value = If(value Is Nothing, DBNull.Value, value)
    }
        Return param
    End Function

    Private Function DeleteDataFiles(TestResults() As TestResult) As Boolean
        '-------------------------------------------------------------------
        ' remove the datafiles that results were extracted from
        ' should only be called after the results have been put somewhere - i.e. iLink database
        '-------------------------------------------------------------------
        Try
            ' 1. 空数组/空引用容错（替代 VB6 On Error Resume Next + UBound 检查）
            If TestResults Is Nothing OrElse TestResults.Length = 0 Then
                LogEvent("No test results to process - no files to delete.")
                DeleteDataFiles = True ' 无数据视为执行成功
                Return DeleteDataFiles
            End If

            ' 2. 用 HashSet 存储已删除的文件（替代字符串拼接，避免重复删除+性能优化）
            Dim deletedFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) ' 忽略文件名大小写

            ' 3. 遍历测试结果，删除文件
            For i As Integer = 0 To TestResults.GetUpperBound(0)
                Dim sCurrentFile As String = TestResults(i).DataFile

                ' 过滤空文件名 + 已删除的文件
                If Not String.IsNullOrEmpty(sCurrentFile) AndAlso Not deletedFiles.Contains(sCurrentFile) Then
                    LogEvent($"Deleting Datafile: {sCurrentFile}")

                    ' 【关键】替换 VB6 Kill 语句 → .NET File.Delete，增加文件存在性检查
                    If File.Exists(sCurrentFile) Then
                        Try
                            ' 设置文件属性为正常（避免只读文件删除失败）
                            File.SetAttributes(sCurrentFile, FileAttributes.Normal)
                            File.Delete(sCurrentFile) ' 删除文件
                            deletedFiles.Add(sCurrentFile) ' 标记为已删除
                        Catch ex As IOException
                            ' 捕获文件占用/权限不足等异常（不中断循环，继续删除其他文件）
                            LogEvent($"** WARNING : Failed to delete {sCurrentFile} - {ex.Message}", True)
                        End Try
                    Else
                        LogEvent($"** WARNING : File {sCurrentFile} does not exist - skipped", True)
                        deletedFiles.Add(sCurrentFile) ' 标记为已处理，避免重复提示
                    End If
                End If
            Next i

            DeleteDataFiles = True
        Catch ex As Exception
            ' 替换 VB6 Err 信息，记录全局异常
            LogEvent($"** ERROR : DeleteDataFiles - {ex.HResult} - {ex.Message}", True)
            DeleteDataFiles = False
        End Try

        Return DeleteDataFiles
    End Function

    Private Function getPenFailures(PenID As String, Optional TestType As String = "") As String()
        '-------------------------------------------------------------------
        ' Return a list of failures which have been assigned to a pen
        '-------------------------------------------------------------------
        Dim sFailures As New List(Of String)() ' 替代动态数组，避免 ReDim Preserve 性能损耗
        Dim sSQL As String = ""

        Try
            ' 1. 空值容错：PenID 为空直接返回空数组
            If String.IsNullOrEmpty(PenID) Then
                LogEvent("** WARNING : PenID is empty - no failures to query")
                Return sFailures.ToArray()
            End If

            ' 2. 验证数据库连接状态
            If g_cnxDb Is Nothing OrElse g_cnxDb.State <> ConnectionState.Open Then
                Throw New Exception("Database connection is not open")
            End If

            ' 3. 构建参数化 SQL（防注入，替代字符串拼接）
            sSQL = "SELECT pen_failcode, page_count, scan_birthdate 
                FROM finishedData 
                WHERE pen_id = ?"

            ' 可选：添加 TestType 过滤条件
            If Not String.IsNullOrEmpty(TestType) Then
                sSQL &= " AND test_type = ?"
            End If

            sSQL &= " ORDER BY page_count DESC, scan_birthdate DESC;"

            ' 4. 创建命令对象（Using 自动释放资源）
            Using cmd As New OdbcCommand(sSQL, g_cnxDb)
                ' ========== 关键适配：ODBC 不支持 @参数名，需将 SQL 中的 @XXX 替换为 ?，且按顺序添加参数 ==========
                ' 假设原始 sSQL 示例：SELECT pen_failcode FROM table WHERE PenID = @PenID AND (TestType = @TestType OR @TestType IS NULL)
                ' 需先修改为：SELECT pen_failcode FROM table WHERE PenID = ? AND (TestType = ? OR ? IS NULL)
                ' （参数添加顺序必须和 SQL 中 ? 的顺序严格一致！）

                ' 1. 添加参数（按 SQL 中 ? 的顺序，ODBC 仅认顺序不认名称）
                ' 第一个 ? 对应原 @PenID
                cmd.Parameters.AddWithValue("?", PenID)

                ' 第二个/第三个 ? 对应原 @TestType（根据 SQL 写法调整，以下是通用适配）
                If Not String.IsNullOrEmpty(TestType) Then
                    cmd.Parameters.AddWithValue("?", TestType)
                Else
                    ' 空值时传 DBNull，避免参数数量不匹配
                    cmd.Parameters.AddWithValue("?", DBNull.Value)
                End If

                ' 2. 执行查询（替换 OleDbDataReader 为 OdbcDataReader）
                Using reader As OdbcDataReader = cmd.ExecuteReader()
                    ' 遍历结果集（逻辑完全不变，仅对象类型替换）
                    While reader.Read()
                        ' 读取失败码，空值则跳过（兼容 DBNull 处理）
                        Dim failCode As String = ""
                        ' 先检查字段是否为 DBNull，避免转换报错
                        If Not reader.IsDBNull(reader.GetOrdinal("pen_failcode")) Then
                            failCode = reader("pen_failcode").ToString().Trim()
                        End If

                        If Not String.IsNullOrEmpty(failCode) Then
                            sFailures.Add(failCode)
                        End If
                    End While
                End Using
            End Using

            ' 6. 转换为字符串数组返回（匹配原函数返回值类型）
            Return sFailures.ToArray()

        Catch ex As Exception
            ' 记录错误日志（包含 SQL 和异常信息）
            LogEvent($"** ERROR : getPenFailures - {ex.HResult} - {ex.Message} - SQL: {sSQL}", True)
            ' 异常时返回空数组，避免调用方崩溃
            Return New String() {}
        End Try
    End Function

    ''' <summary>
    ''' See if a list of failures matches a given rule
    ''' </summary>
    ''' <param name="arrFailures">失败码数组</param>
    ''' <param name="sRule">规则表达式（支持旧格式+新格式，如 "F001+F002" 或 "'F001'&'F002'"）</param>
    ''' <returns>是否匹配规则（True=匹配，False=不匹配）</returns>
    Private Function MatchRule(arrFailures() As String, sRule As String) As Boolean
        Dim bReturn As Boolean = False

        Try
            ' 空值容错
            If String.IsNullOrWhiteSpace(sRule) Then
                LogEvent("** WARNING : MatchRule - Empty rule string", True)
                Return False
            End If

            ' 1. Prepare the rule expression
            ' 旧格式转换（无单引号 → 转换为新格式，如 "F001+F002" → "'F001'&'F002'"）
            If Not sRule.Contains("'") Then
                sRule = $"'{sRule.Trim().Replace("+", "'&'")}'"
            End If

            Dim iNumQuotes As Integer = 0 ' 替换 Double → Integer（计数用整型更合理）
            Dim arrFailsInRule As New List(Of String)() ' 替换动态数组 → List，避免 ReDim Preserve
            Dim sCurrentChar As String = String.Empty
            Dim sTempFail As String = String.Empty

            ' 2. 解析规则中的失败码（提取单引号内的失败码）
            For i As Integer = 0 To sRule.Length - 1 ' VB6 Mid(i,1) → .NET Substring(i,1)，索引从0开始
                sCurrentChar = sRule.Substring(i, 1)

                If sCurrentChar = "'" Then
                    ' 遇到单引号：若缓存有失败码则加入数组
                    If Not String.IsNullOrEmpty(sTempFail) Then
                        arrFailsInRule.Add(sTempFail)
                        sTempFail = String.Empty
                    End If
                    iNumQuotes += 1
                ElseIf iNumQuotes > 0 Then
                    ' 单引号数量为奇数时，拼接失败码
                    If iNumQuotes Mod 2 = 1 Then
                        sTempFail &= sCurrentChar
                    End If
                End If
            Next

            ' 检查单引号是否成对
            If iNumQuotes Mod 2 <> 0 Then
                LogEvent($"** ERROR : MatchRule - Rule has uneven number of quotes! {sRule}", True)
                Return False
            End If

            ' 3. 将失败码数组转换为分隔符包裹的字符串（便于 InStr 检查）
            Dim sFailsOccuring As String = "|" ' 首尾加|，避免部分匹配（如 F001 匹配 F0011）
            If arrFailures IsNot Nothing Then
                For Each sTempFail In arrFailures
                    If Not String.IsNullOrEmpty(sTempFail) Then
                        sFailsOccuring &= $"{sTempFail}|"
                    End If
                Next
            End If

            ' 4. 将规则中的失败码替换为 1（存在）/0（不存在），生成布尔表达式
            For Each sTempFail In arrFailsInRule
                If Not String.IsNullOrEmpty(sTempFail) Then
                    Dim searchStr As String = $"|{sTempFail}|"
                    If sFailsOccuring.Contains(searchStr) Then
                        sRule = sRule.Replace($"'{sTempFail}'", "1")
                    Else
                        sRule = sRule.Replace($"'{sTempFail}'", "0")
                    End If
                End If
            Next

            ' 5. 解析布尔表达式（复用 clsEval 类）
            Dim oEval As New clsEval()
            Dim dResult As Double = oEval.Evaluate(sRule) ' Evaluate 返回 Double

            ' 结果判断：1=匹配，0=不匹配
            bReturn = (dResult = 1.0)

        Catch ex As Exception
            ' 替换 VB6 Err 信息，记录异常
            LogEvent($"** ERROR : MatchRule - {ex.HResult} - {ex.Message} - Rule: {sRule}", True)
            bReturn = False
        End Try

        Return bReturn
    End Function

    ''' <summary>
    ''' Given the ordered list of failcodes and priorities it will return
    ''' the most appropriate failcode (highest priority match)
    ''' </summary>
    ''' <param name="sFailcodeArray">待匹配的失败码数组</param>
    ''' <param name="sPriorityArray">按优先级排序的规则数组（高优先级在前）</param>
    ''' <returns>匹配到的最高优先级失败码（无匹配返回空字符串）</returns>
    Private Function getHigestPriorityFailcode(sFailcodeArray() As String, sPriorityArray() As String) As String
        Dim sSetFailcode As String = String.Empty

        Try
            ' 1. 空数组容错（替代 VB6 On Error Resume Next + UBound 检查）
            ' 优先级数组为空 → 直接返回空
            If sPriorityArray Is Nothing OrElse sPriorityArray.Length = 0 Then
                LogEvent("** WARNING : getHigestPriorityFailcode - Priority array is empty", True)
                Return sSetFailcode
            End If

            ' 失败码数组为空 → 直接返回空
            If sFailcodeArray Is Nothing OrElse sFailcodeArray.Length = 0 Then
                LogEvent("** WARNING : getHigestPriorityFailcode - Failcode array is empty", True)
                Return sSetFailcode
            End If

            ' 2. 遍历优先级数组（高优先级在前，匹配到立即返回）
            For Each priorityRule As String In sPriorityArray
                ' 跳过空的优先级规则
                If String.IsNullOrWhiteSpace(priorityRule) Then Continue For

                ' 遍历失败码数组，匹配当前优先级规则
                For Each failcode As String In sFailcodeArray
                    ' 跳过空失败码
                    If String.IsNullOrWhiteSpace(failcode) Then Continue For

                    ' 匹配规则：检查优先级规则中是否包含 '失败码'（原 VB6 InStr 逻辑）
                    Dim searchStr As String = $"'{failcode}'"
                    If priorityRule.Contains(searchStr) Then
                        sSetFailcode = failcode
                        ' 匹配到最高优先级，立即退出所有循环
                        Return sSetFailcode
                    End If
                Next
            Next

        Catch ex As Exception
            ' 替换 VB6 Err 信息，记录异常
            LogEvent($"** ERROR : getHigestPriorityFailcode - {ex.HResult} - {ex.Message}", True)
            sSetFailcode = String.Empty
        End Try

        ' 返回匹配结果（无匹配则为空）
        Return sSetFailcode
    End Function

    ''' <summary>
    ''' Set the overall failcode in the iLink database to the new setfailcode
    ''' </summary>
    ''' <param name="PenID">笔ID</param>
    ''' <param name="Failcode">要更新的整体失败码</param>
    ''' <returns>更新是否成功（True=成功，False=失败）</returns>
    Private Function updOverallFailcode(PenID As String, Failcode As String) As Boolean
        Dim isSuccess As Boolean = False
        Dim sSQL As String = String.Empty

        Try
            ' 1. 空值容错：PenID/Failcode 为空直接返回失败
            If String.IsNullOrWhiteSpace(PenID) Then
                LogEvent("** WARNING : updOverallFailcode - PenID is empty", True)
                Return False
            End If
            If String.IsNullOrWhiteSpace(Failcode) Then
                LogEvent("** WARNING : updOverallFailcode - Failcode is empty", True)
                Return False
            End If

            ' 2. 验证数据库连接状态
            If g_cnxDb Is Nothing OrElse g_cnxDb.State <> ConnectionState.Open Then
                Throw New Exception("Database connection is not open")
            End If

            ' 3. 构建参数化 SQL（防注入，替代字符串拼接）
            sSQL = "UPDATE finishedData 
                SET set_failcode = ? 
                WHERE pen_id = ?;"

            ' 4. 创建命令对象（Using 自动释放资源，避免内存泄漏）
            Using cmd As New OdbcCommand(sSQL, g_cnxDb)
                ' ========== 关键适配：ODBC 不支持 @参数名，需先修改 SQL 中的 @XXX 为 ? ==========
                ' 假设原始 sSQL 示例：UPDATE table SET failcode = @Failcode WHERE PenID = @PenID
                ' 需同步修改为：UPDATE table SET failcode = ? WHERE PenID = ?
                ' （参数添加顺序必须和 SQL 中 ? 的顺序严格一致！）

                ' 1. 添加参数（ODBC 按顺序匹配 ?，名称无意义，仅需保证顺序）
                ' 第一个 ? 对应原 @Failcode
                cmd.Parameters.AddWithValue("?", Failcode)
                ' 第二个 ? 对应原 @PenID
                cmd.Parameters.AddWithValue("?", PenID)

                ' 2. 执行更新（ExecuteNonQuery 用法与 OleDb 完全一致）
                Dim affectedRows As Integer = cmd.ExecuteNonQuery()

                ' 3. 验证更新结果（保留原有逻辑）
                isSuccess = True ' 执行无异常即视为成功（兼容原 VB6 逻辑）
                ' 若需严格验证是否有记录被更新，可改为：
                ' isSuccess = (affectedRows > 0)
            End Using

        Catch ex As Exception
            ' 记录错误日志（包含异常信息+SQL语句）
            LogEvent($"** ERROR : updOverallFailcode - {ex.HResult} - {ex.Message} - SQL: {sSQL}", True)
            isSuccess = False
        End Try

        Return isSuccess
    End Function

    ''' <summary>
    ''' Determine the overall failmode for a given pen id
    ''' </summary>
    ''' <param name="PenIDs">笔ID数组</param>
    ''' <returns>包含每支笔整体失败码的 PenResult 数组</returns>
    Private Function setOverallFailmode(PenIDs() As String) As PenResult()
        ' 替换动态数组为 List，避免 ReDim Preserve 性能损耗
        Dim penResults As New List(Of PenResult)()

        Try
            LogEvent("Calculating Failmodes... ")

            ' 空数组容错：PenIDs 为空直接返回空数组
            If PenIDs Is Nothing OrElse PenIDs.Length = 0 Then
                LogEvent("** WARNING : setOverallFailmode - PenIDs array is empty", True)
                Return penResults.ToArray()
            End If

            ' 遍历每支笔，计算整体失败码
            For iPen As Integer = 0 To PenIDs.GetUpperBound(0)
                Dim penID As String = PenIDs(iPen)
                Dim sSetFailcode As String = String.Empty
                Dim sFailures As String() = Nothing

                ' 跳过空的 PenID
                If String.IsNullOrWhiteSpace(penID) Then
                    LogEvent($"** WARNING : setOverallFailmode - Empty PenID at index {iPen}", True)
                    Continue For
                End If

                ' 1. 获取该笔的所有失败码（替代 VB6 On Error Resume Next 空数组检查）
                sFailures = getPenFailures(penID)
                If sFailures Is Nothing OrElse sFailures.Length = 0 Then
                    ' 无失败码 → 标记为 "好笔"
                    sSetFailcode = g_goodPenCode
                    GoTo Found ' 保留原 GoTo 逻辑，确保流程一致
                End If

                ' 2. 匹配判定规则（CapDecisionRules）
                If m_arrCapDecisionRules IsNot Nothing AndAlso m_arrCapDecisionRules.Length > 0 Then
                    For iRule As Integer = 0 To m_arrCapDecisionRules.GetUpperBound(0)
                        Dim rule As CapDecisionRule = m_arrCapDecisionRules(iRule)
                        If rule IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(rule.FailCodes) Then
                            If MatchRule(sFailures, rule.FailCodes) Then
                                sSetFailcode = rule.SetFailCode
                                LogEvent($"Matched on Rule number: {iRule + 1}")
                                GoTo Found
                            End If
                        End If
                    Next
                End If

                ' 3. 匹配最高优先级失败码
                sSetFailcode = getHigestPriorityFailcode(sFailures, m_arrPriorityFailures)
                If Not String.IsNullOrWhiteSpace(sSetFailcode) Then
                    LogEvent("Matched priority failcode..")
                    GoTo Found
                End If

                ' 4. 按测试级别匹配最高优先级失败码（排除 "好笔" 码）
                If m_arrTestLevels IsNot Nothing AndAlso m_arrTestLevels.Length > 0 Then
                    For iTestLevel As Integer = 0 To m_arrTestLevels.GetUpperBound(0)
                        Dim testLevel As TestLevel = m_arrTestLevels(iTestLevel)
                        If testLevel IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(testLevel.TestType) Then
                            ' 获取该测试类型的失败码
                            Dim sTestFailures As String() = getPenFailures(penID, testLevel.TestType)
                            ' 匹配该级别优先级失败码
                            sSetFailcode = getHigestPriorityFailcode(sTestFailures, testLevel.FailcodePriorities)
                            ' 排除 "好笔" 码（g_goodPenCode）
                            If Not String.IsNullOrWhiteSpace(sSetFailcode) AndAlso Not sSetFailcode.Contains(g_goodPenCode) Then
                                GoTo Found
                            End If
                        End If
                    Next
                End If

                ' 5. 所有匹配失败 → 标记为 "好笔"
                sSetFailcode = g_goodPenCode

Found: ' 结果汇总（保留原标签，确保流程一致）
                ' 创建 PenResult 对象并添加到列表
                Dim penResult As New PenResult() With {
                    .PenID = penID,
                    .OverallFailcode = sSetFailcode
                }
                penResults.Add(penResult)

                ' 更新数据库中的整体失败码
                If updOverallFailcode(penResult.PenID, penResult.OverallFailcode) Then
                    LogEvent($"Pen ID:{penResult.PenID} has overall failmode of: {penResult.OverallFailcode}")
                Else
                    LogEvent($"** Error updating Result for Pen ID:{penResult.PenID} of overall failmode of: {penResult.OverallFailcode}", True)
                End If
            Next

        Catch ex As Exception
            ' 全局异常捕获，记录详细日志
            LogEvent($"** ERROR : setOverallFailmode - {ex.HResult} - {ex.Message}", True)
        End Try

        ' 转换为数组返回（匹配原函数返回值类型）
        Return penResults.ToArray()
    End Function

    Private Function getCapMatch(sFailcode As String) As String
        Dim sResult As String = String.Empty

        Try
            ' 1. 空值容错：传入的失败码为空时直接返回空
            If String.IsNullOrWhiteSpace(sFailcode) Then
                LogEvent("** WARNING : getCapMatch - Input failcode is empty", True)
                Return String.Empty
            End If

            ' 2. 调用INI读取函数（需实现VB6 ReadIniString的等效逻辑）
            sResult = ReadIniString(sFailcode, "Missing", "MatchList")

            ' 3. 匹配不到时返回原失败码（兼容原逻辑）
            If sResult = "Missing" Or String.IsNullOrWhiteSpace(sResult) Then
                sResult = sFailcode
            End If

        Catch ex As Exception
            ' 替换VB6 On Error GoTo EH，记录详细异常
            LogEvent($"** ERROR : getCapMatch - {ex.HResult} - {ex.Message} (Failcode: {sFailcode})", True)
            ' 异常时返回原失败码，保证流程不中断
            sResult = sFailcode
        End Try

        Return sResult
    End Function

    Private Function doubleCheckNextcap(sPenID As String, sFailcode As String) As Boolean
        Dim isConsistent As Boolean = False ' 默认校验失败（最坏情况）

        Try
            ' 1. 空值容错
            If String.IsNullOrWhiteSpace(sPenID) Then
                LogEvent("** WARNING : doubleCheckNextcap - PenID is empty", True)
                Return False
            End If
            If String.IsNullOrWhiteSpace(sFailcode) Then
                LogEvent("** WARNING : doubleCheckNextcap - Failcode is empty", True)
                Return False
            End If

            ' 2. 检查Nextcap上传锁（复用原逻辑）
            'mdlNextcap.checkForNextcapLock()

            ' 3. 数据库操作（替换ADODB为OleDb，使用Using自动释放资源）
            Using cnxDb As New OdbcConnection(g_NextcapConnectionString)
                Try
                    ' 打开连接（ODBC 无需额外配置 Mode，兼容原 VB6 adModeRead）
                    cnxDb.Open()

                    Dim sSQL As String = String.Empty
                    Using cmd As New OdbcCommand()
                        cmd.Connection = cnxDb
                        cmd.CommandType = CommandType.Text

                        ' ========== 关键适配：ODBC 不支持 @参数名，替换为 ? 占位符 ==========
                        If sFailcode = g_goodPenCode Then
                            ' 场景1：无故障码 → 检查无iLink缺陷记录
                            ' SQL 中 @PenID → ? 、@ILinkComment → ?
                            sSQL = "SELECT * FROM PenDefects 
                    WHERE PenId = ? 
                    AND DefectComment = ? 
                    AND SyncState NOT IN ('REMOVE', 'DELETE')"
                            ' 按 SQL 中 ? 顺序添加参数（PenID → ILinkComment）
                            cmd.Parameters.AddWithValue("?", sPenID)
                            cmd.Parameters.AddWithValue("?", CN_ILINK_COMMENT)
                        Else
                            ' 场景2：有故障码 → 检查匹配的缺陷记录
                            Dim codeField As String = $"code{g_capDefectLevel}" ' 拼接缺陷级别字段（如code1）
                            ' SQL 中 @PenID → ? 、@Failcode → ?
                            sSQL = $"SELECT * FROM PenDefects 
                    WHERE PenId = ? 
                    AND [{codeField}] = ? 
                    AND SyncState NOT IN ('REMOVE', 'DELETE')"
                            ' 按 SQL 中 ? 顺序添加参数（PenID → Failcode）
                            cmd.Parameters.AddWithValue("?", sPenID)
                            cmd.Parameters.AddWithValue("?", sFailcode)
                        End If

                        ' 赋值最终SQL到Command（易漏点！）
                        cmd.CommandText = sSQL

                        ' 执行查询并校验结果（替换 OleDbDataReader 为 OdbcDataReader）
                        Using reader As OdbcDataReader = cmd.ExecuteReader()
                            If sFailcode = g_goodPenCode Then
                                ' 无故障码：无记录 → 校验通过
                                isConsistent = reader.IsClosed OrElse Not reader.HasRows
                            Else
                                ' 有故障码：有记录 → 校验通过
                                isConsistent = reader.HasRows
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    ' 异常处理：校验失败 + 记录错误
                    isConsistent = False
                    Debug.WriteLine($"数据库校验失败：{ex.Message}")
                    ' 可选：日志记录
                    ' LogEvent($"PenDefects校验异常：{ex.Message}", True)
                End Try
                ' Using 自动关闭/释放连接，无需手动 Close
            End Using

        Catch ex As Exception
            ' 替换VB6 On Error GoTo EH，记录详细异常
            LogEvent($"** ERROR : doubleCheckNextcap - {ex.HResult} - {ex.Message} (PenID: {sPenID}, Failcode: {sFailcode})", True)
            isConsistent = False ' 异常时默认校验失败
        End Try

        Return isConsistent
    End Function

    Private Function updNextcapIssueTable(PenID As String, Failcode As String, LoadError As String) As Object
        Try
            ' 1. 空值容错：核心参数为空直接记录警告并退出
            If String.IsNullOrWhiteSpace(PenID) Then
                LogEvent("** WARNING : updNextcapIssueTable - PenID is empty", True)
                Return Nothing
            End If
            If String.IsNullOrWhiteSpace(Failcode) Then
                LogEvent("** WARNING : updNextcapIssueTable - Failcode is empty", True)
                Return Nothing
            End If

            ' 2. 检查数据库连接状态
            If g_cnxDb Is Nothing OrElse g_cnxDb.State <> ConnectionState.Open Then
                Throw New Exception("Global database connection (g_cnxDb) is not open")
            End If

            Dim iNumRetries As Integer = 0
            Dim isNewRecord As Boolean = True ' 是否为新增记录

            ' 3. 查询是否已存在该笔的失败记录（参数化查询防注入）
            Using cmdCheck As New OdbcCommand()
                cmdCheck.Connection = g_cnxDb ' g_cnxDb 需是 OdbcConnection 类型
                ' ========== 关键适配：ODBC 不支持 @参数名，替换为 ? 占位符 ==========
                cmdCheck.CommandText = "SELECT page FROM NextCapFailures WHERE pen_id = ?"

                ' 添加参数（ODBC 按顺序匹配 ?，名称无意义，仅需保证顺序）
                ' 唯一 ? 对应原 @PenID，参数值空值处理（避免 DBNull 报错）
                Dim penIDValue As Object = If(String.IsNullOrEmpty(PenID), DBNull.Value, PenID)
                cmdCheck.Parameters.AddWithValue("?", penIDValue)

                ' 执行查询并获取重试次数（ExecuteScalar 用法与 OleDb 完全一致）
                Dim pageObj As Object = cmdCheck.ExecuteScalar()

                ' 空值判断逻辑（完全保留，兼容 DBNull/Null）
                If pageObj IsNot DBNull.Value AndAlso pageObj IsNot Nothing Then
                    iNumRetries = CInt(pageObj)
                    isNewRecord = False ' 存在记录，标记为更新
                End If
            End Using

            ' 4. 构建新增/更新SQL（参数化，避免注入）
            Using cmdUpdate As New OdbcCommand()
                cmdUpdate.Connection = g_cnxDb ' g_cnxDb 需是 OdbcConnection 类型

                If isNewRecord Then
                    ' ========== 场景1：新增记录（适配 ODBC 占位符 ?）==========
                    ' SQL 中 @PenID/@Failcode/@LoadError/@ScanBirthdate → 替换为 ?，按顺序匹配
                    cmdUpdate.CommandText = "INSERT INTO NextCapFailures(pen_id, failcode, page, serverText, scan_birthdate) 
                             VALUES (?, ?, 1, ?, ?);"

                    ' 添加参数（顺序必须和 SQL 中 ? 完全一致：PenID → Failcode → LoadError → ScanBirthdate）
                    cmdUpdate.Parameters.AddWithValue("?", PenID)
                    cmdUpdate.Parameters.AddWithValue("?", Failcode)
                    cmdUpdate.Parameters.AddWithValue("?", If(String.IsNullOrWhiteSpace(LoadError), "", LoadError))
                    cmdUpdate.Parameters.AddWithValue("?", DateTime.Now) ' 替换 VB6 Now()

                    LogEvent($"Adding {PenID} ({Failcode}) To error Table.")
                Else
                    ' ========== 场景2：更新记录（适配 ODBC 占位符 ?）==========
                    ' SQL 中 @Failcode/@LoadError/@Page/@PenID → 替换为 ?，按顺序匹配
                    cmdUpdate.CommandText = "UPDATE NextCapFailures SET 
                             failcode = ?, 
                             serverText = ?, 
                             page = ? 
                             WHERE pen_id = ?;"

                    ' 添加参数（顺序必须和 SQL 中 ? 完全一致：Failcode → LoadError → Page → PenID）
                    cmdUpdate.Parameters.AddWithValue("?", Failcode)
                    cmdUpdate.Parameters.AddWithValue("?", If(String.IsNullOrWhiteSpace(LoadError), "", LoadError))
                    cmdUpdate.Parameters.AddWithValue("?", iNumRetries + 1)
                    cmdUpdate.Parameters.AddWithValue("?", PenID)

                    LogEvent($"Updating retries on {PenID} ({Failcode}) To {iNumRetries + 1} in error Table.")
                End If

                ' 执行 SQL（ExecuteNonQuery 用法与 OleDb 完全一致）
                cmdUpdate.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            ' 替换VB6 On Error GoTo EH，记录详细异常
            LogEvent($"** ERROR : updNextcapIssueTable - {ex.HResult} - {ex.Message} (PenID: {PenID}, Failcode: {Failcode})", True)
        End Try

        Return Nothing
    End Function

    'Private Function LoadResultsToNextCAP(PenResults() As PenResult) As Object
    '    Try
    '        ' 1. 空数组容错（替代VB6 On Error Resume Next + UBound检查）
    '        Dim iArraysize As Integer = -1
    '        If PenResults IsNot Nothing AndAlso PenResults.Length > 0 Then
    '            iArraysize = PenResults.GetUpperBound(0)
    '        Else
    '            LogEvent("** WARNING : LoadResultsToNextCAP - PenResults array is empty", True)
    '            Return Nothing
    '        End If

    '        ' 2. 遍历每支笔的结果
    '        For i As Integer = 0 To iArraysize
    '            ' 空元素容错：跳过数组中为空的PenResult对象
    '            If PenResults(i) Is Nothing Then
    '                LogEvent($"** WARNING : LoadResultsToNextCAP - Empty PenResult at index {i}", True)
    '                Continue For
    '            End If

    '            Dim sPenID As String = PenResults(i).PenID
    '            Dim sOverallFailcode As String = PenResults(i).OverallFailcode

    '            ' 空值容错：跳过PenID/失败码为空的记录
    '            If String.IsNullOrWhiteSpace(sPenID) Then
    '                LogEvent($"** WARNING : LoadResultsToNextCAP - Empty PenID at index {i}", True)
    '                Continue For
    '            End If
    '            If String.IsNullOrWhiteSpace(sOverallFailcode) Then
    '                LogEvent($"** WARNING : LoadResultsToNextCAP - Empty OverallFailcode for PenID {sPenID}", True)
    '                Continue For
    '            End If

    '            ' 3. 获取Cap失败码（复用已迁移的getCapMatch）
    '            Dim sCapFailcode As String = getCapMatch(sOverallFailcode)

    '            ' 4. 检查是否为需忽略的失败码
    '            If Not String.IsNullOrWhiteSpace(m_codesToIgnore) AndAlso m_codesToIgnore.Contains($"|{sCapFailcode}|") Then
    '                LogEvent($"Ignoring failcode :{sCapFailcode}")
    '                Continue For ' 跳过当前笔，处理下一个
    '            End If

    '            ' 5. 校验Nextcap与iLink数据一致性（复用已迁移的doubleCheckNextcap）
    '            If Not doubleCheckNextcap(sPenID, sCapFailcode) Then
    '                ' 5.1 不一致则更新Nextcap
    '                Dim sSuccess As String = UpdatePen(sPenID, sCapFailcode)

    '                If sSuccess = "OK" Then
    '                    ' 5.2 启用双重校验时再次验证
    '                    If m_doubleCheckNextcap Then
    '                        If Not doubleCheckNextcap(sPenID, sCapFailcode) Then
    '                            sSuccess = "LOAD ERROR - Contact Support"
    '                            LogEvent($"**WARNING - {sPenID} ({sCapFailcode}) FAILED Double Check.", True)
    '                        Else
    '                            LogEvent($"{sPenID} ({sCapFailcode}) Double Checked - OK. ")
    '                        End If
    '                    End If

    '                    ' 5.3 更新成功/校验通过：删除失败记录表中的记录（参数化防注入）
    '                    Using cmdDelete As New OleDbCommand()
    '                        cmdDelete.Connection = g_cnxDb
    '                        cmdDelete.CommandText = "DELETE FROM NextCapFailures WHERE pen_id = @PenID;"
    '                        cmdDelete.Parameters.AddWithValue("@PenID", sPenID)
    '                        cmdDelete.ExecuteNonQuery()
    '                    End Using
    '                Else
    '                    ' 5.4 更新失败：更新失败记录表（复用已迁移的updNextcapIssueTable）
    '                    Call updNextcapIssueTable(sPenID, sOverallFailcode, sSuccess)
    '                End If
    '            Else
    '                ' 6. 数据一致：无需更新Nextcap，清理失败记录
    '                LogEvent($"Pen status not changed - no need to update NextCAP: {sPenID}")
    '                Using cmdDelete As New OleDbCommand()
    '                    cmdDelete.Connection = g_cnxDb
    '                    cmdDelete.CommandText = "DELETE FROM NextCapFailures WHERE pen_id = @PenID;"
    '                    cmdDelete.Parameters.AddWithValue("@PenID", sPenID)
    '                    cmdDelete.ExecuteNonQuery()
    '                End Using
    '            End If
    '        Next

    '    Catch ex As Exception
    '        ' 替换VB6 On Error GoTo EH，记录详细异常
    '        LogEvent($"** ERROR : LoadResultsToNextCAP - {ex.HResult} - {ex.Message}", True)
    '    End Try

    '    Return Nothing
    'End Function

    '    Public Function UpdatePen(sPenID As String, sFailcode As String) As String
    '        Dim bFound As Boolean = False
    '        Dim nIndex As Integer
    '        Dim vrtPen As Object ' 替代VB6的Variant，存储笔对象数组
    '        Dim sError As String = "OK" ' 默认执行成功
    '        Dim sCurrentFailcode As String = String.Empty
    '        Dim iOldMutex As Integer = mutexCounter ' 保存原始互斥计数器

    '        Try
    '            ' 1. 空值容错
    '            If String.IsNullOrWhiteSpace(sPenID) Then
    '                LogEvent("** WARNING : UpdatePen - PenID is empty")
    '                Return "EmptyPenID"
    '            End If
    '            If String.IsNullOrWhiteSpace(sFailcode) Then
    '                LogEvent("** WARNING : UpdatePen - Failcode is empty for PenID: " & sPenID)
    '                Return "EmptyFailcode"
    '            End If

    '            ' 2. 检查Nextcap上传锁
    '            checkForNextcapLock()

    '            ' 3. 互斥计数器自增（防止函数重叠执行）
    '            mutexCounter += 1

    '            ' CHECK 1 - 防止函数重叠执行
    '            If mutexCounter > 1 Then
    '                LogEvent("NEXTCAP:UpdatePen: routine was entered before last run finished, or an error kicked us out.")
    '                sError = "Function overlap"
    '                GoTo FINISHED ' 保留原标签，保证流程一致
    '            End If

    '            ' 4. 尝试获取笔对象
    '            bFound = RetrievePen(sPenID, vrtPen)
    '            checkForNextcapLock()

    '            ' 5. 遍历LotManager集合，重新获取笔对象（未找到时）
    '            If Not bFound AndAlso gcol_LotManagers IsNot Nothing AndAlso gcol_LotManagers.Count > 0 Then
    '                For nIndex = 0 To gcol_LotManagers.Count - 1 ' VB.NET集合索引从0开始
    '                    ' 跳过当前激活的LotManager
    '                    If nIndex <> gn_LotManagerIndex Then
    '                        go_ActiveLotManager = DirectCast(gcol_LotManagers(nIndex), ILotManager).GetLotManager()
    '                        gn_LotManagerIndex = nIndex
    '                        checkForNextcapLock()
    '                        bFound = RetrievePen(sPenID, vrtPen)
    '                        If bFound Then Exit For
    '                    End If
    '                Next
    '            End If

    '            ' CHECK 2 - 验证笔对象是否存在
    '            If Not bFound Then
    '                LogEvent("NEXTCAP:UpdatePen ERROR " & sPenID & ". Pen Not in Nextcap.")
    '                sError = "NotInCap or Duplicate"
    '                GoTo FINISHED
    '            End If

    '            ' CHECK 3 - 验证笔对象数组大小（替代VB6 UBound）
    '            Dim penArray As Array = TryCast(vrtPen, Array)
    '            If penArray Is Nothing OrElse penArray.GetUpperBound(0) <= 18 Then
    '                LogEvent("Incorrect Pen Array Size Retrieved from CAP for Pen: " & sPenID)
    '                sError = "ArraySize"
    '                GoTo FINISHED
    '            End If

    '            ' 6. 获取当前Inspector失败码
    '            sCurrentFailcode = getCurrentInspectorFailcode(vrtPen)

    '            ' 7. 分场景更新失败码
    '            If sCurrentFailcode = sFailcode Then
    '                ' 场景1：失败码无变化，释放笔对象
    '                LogEvent("Pen status not changed - no need to update NextCAP: " & sPenID)
    '                If Not ReleaseUnit(vrtPen) Then
    '                    sError = "ReleaseUnit"
    '                End If
    '                GoTo FINISHED
    '            ElseIf sFailcode = g_goodPenCode Then
    '                ' 场景2：标记为好笔，移除所有Inspector缺陷
    '                LogEvent("Removing Inspector Defects From: " & sPenID)
    '                If Not removeAllInspectorDefects(vrtPen) Then
    '                    sError = "removeAllInspectorDefects"
    '                End If
    '            ElseIf sCurrentFailcode = g_goodPenCode Then
    '                ' 场景3：新增缺陷（原无缺陷）
    '                LogEvent("Adding New Defect(" & sFailcode & ") To: " & sPenID)
    '                If Not addNewInspectorDefect(vrtPen, sFailcode) Then
    '                    sError = "addNewInspectorDefect"
    '                End If
    '            Else
    '                ' 场景4：更新现有缺陷
    '                LogEvent("Updating existing defect on " & sPenID & " to:" & sFailcode)
    '                If Not updExistingInspectorDefect(vrtPen, sFailcode) Then
    '                    sError = "updExistingInspectorDefect"
    '                End If
    '            End If

    '            ' 8. 检查锁并发送笔对象回Nextcap
    '            checkForNextcapLock()
    '            If SendPenBack(vrtPen) Then
    '                LogEvent("Update OK.")
    '            Else
    '                LogEvent("Update Error.")
    '                sError = "SendPenBack"
    '            End If

    'FINISHED: ' 结果汇总
    '            UpdatePen = sError
    '            mutexCounter -= 1 ' 恢复互斥计数器

    '        Catch ex As Exception
    '            ' 异常处理：恢复原始互斥计数器，释放笔对象
    '            mutexCounter = iOldMutex
    '            If {"SendPenBack", "ReleaseUnit", "ArraySize"}.Contains(sError) Then
    '                ReleaseUnit(vrtPen)
    '            End If

    '            ' 构造异常返回值
    '            sError = $"UnhandledException: {sError} {ex.HResult} - {ex.Message}"
    '            UpdatePen = sError
    '            LogEvent($"** ERROR : UpdatePen - {ex.HResult} - {ex.Message} (PenID: {sPenID}, Failcode: {sFailcode})", True)
    '        End Try

    '        Return sError
    '    End Function

    Private Function dbConnect() As Boolean
        ' 1. 释放原有连接（避免资源泄漏）
        If g_cnxDb IsNot Nothing Then
            If g_cnxDb.State = ConnectionState.Open Then
                g_cnxDb.Close()
            End If
            g_cnxDb.Dispose()
            g_cnxDb = Nothing
        End If

        Try
            ' 2. 核心修复：改用 OdbcConnection（适配 ODBC 格式连接字符串）
            g_cnxDb = New OdbcConnection(m_ILinkConnectionString)

            ' 3. 打开连接（无需额外配置 Mode/CursorLocation，ODBC 自动适配）
            g_cnxDb.Open()

            ' 4. 验证连接状态
            If g_cnxDb.State = ConnectionState.Open Then
                dbConnect = True
            Else
                dbConnect = False
                LogEvent("** WARNING: dbConnect: 连接对象已创建但未打开", True)
            End If

        Catch ex As Exception
            ' 5. 异常处理
            dbConnect = False
            LogEvent($"** ERROR: dbConnect: {ex.Message}", True)
        End Try

        Return dbConnect
    End Function

End Module
