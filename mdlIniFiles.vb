Imports System.IO
Imports System.Windows.Forms
Imports IniParser
Imports IniParser.Model
Imports IniParser.Model.Configuration

Module mdlIniFiles
    <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Function GetPrivateProfileString(
        ByVal lpSectionName As String,
        ByVal lpKeyName As String,
        ByVal lpDefault As String,
        ByVal lpReturnedString As System.Text.StringBuilder,
        ByVal nSize As Integer,
        ByVal lpFileName As String
    ) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Function WritePrivateProfileString(
        ByVal lpSectionName As String,
        ByVal lpKeyName As String,
        ByVal lpString As String,
        ByVal lpFileName As String
    ) As Boolean
    End Function

    Public Function ReadIniString(keyNm As String, DefVal As String, Optional sectionNm As String = "Missing", Optional iniNm As String = "Missing") As String
        ' 1. 处理默认节名/文件路径（兼容原有逻辑）
        If sectionNm = "Missing" Then
            sectionNm = Application.ProductName
        End If
        If iniNm = "Missing" Then
            iniNm = Path.Combine(Application.StartupPath, $"{Application.ProductName}.ini")
        End If

        Try
            ' 2. 初始化基础解析器（低版本兼容）
            Dim parser As New FileIniDataParser()

            ' 3. 预处理 INI 文件（核心：解决###注释/Tab分隔问题）
            ' 读取文件并清理非标准格式，生成临时文件供解析
            Dim tempIniPath As String = Path.GetTempFileName()
            Using sw As New StreamWriter(tempIniPath, False, System.Text.Encoding.Default)
                For Each line As String In File.ReadAllLines(iniNm, System.Text.Encoding.Default)
                    Dim cleanLine As String = line.Trim()

                    ' 处理###注释：替换为标准;注释（解析器可识别）
                    If cleanLine.StartsWith("###") OrElse cleanLine.StartsWith("#") Then
                        sw.WriteLine($";{cleanLine.Substring(1)}") ' # → ;
                    Else
                        ' 处理Tab分隔符：替换Tab为等号前的空格
                        cleanLine = cleanLine.Replace(vbTab, " ")
                        ' 移除行尾注释（避免解析干扰）
                        If cleanLine.Contains("#") Then
                            cleanLine = cleanLine.Substring(0, cleanLine.IndexOf("#")).Trim()
                        End If
                        sw.WriteLine(cleanLine)
                    End If
                Next
            End Using

            ' 4. 读取预处理后的临时INI文件（低版本解析器可正常解析）
            Dim iniData As IniData = parser.ReadFile(tempIniPath)

            ' 5. 删除临时文件
            File.Delete(tempIniPath)

            ' 6. 严谨读取键值（兼容节/键不存在）
            Dim value As String = DefVal.Trim()
            If iniData.Sections.ContainsSection(sectionNm) Then
                If iniData(sectionNm).ContainsKey(keyNm) Then
                    value = iniData(sectionNm)(keyNm).Trim()
                End If
            End If

            Return value

        Catch ex As Exception
            ' 输出调试信息（可选）
            Debug.WriteLine($"INI 解析异常：{ex.Message}")
            Return DefVal.Trim()
        End Try
    End Function

    Public Sub WriteIniString(sectionNm As String, keyNm As String, KeyVal As String)
        Dim iniPath As String = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "temp.ini")
        WritePrivateProfileString(sectionNm, keyNm, KeyVal, iniPath)
    End Sub

    Public Function TestShare(sDirectoryName As String) As Boolean
        Try
            If Not String.IsNullOrEmpty(sDirectoryName) Then
                If Not sDirectoryName.EndsWith("\") Then
                    sDirectoryName = sDirectoryName & "\"
                End If
                If System.IO.Directory.Exists(sDirectoryName) Then
                    TestShare = True
                Else
                    TestShare = False
                End If
            Else
                TestShare = False
            End If
        Catch ex As Exception
            TestShare = False
        End Try
    End Function

    Public Function AllFiles(ByVal DirPath As String) As String()
        Dim fileList As New List(Of String)()

        Try
            ' ========== 核心修复：拆分目录路径和搜索模式 ==========
            ' 1. 提取纯目录路径（去掉末尾的 *.txt 等通配符）
            Dim pureDirPath As String = System.IO.Path.GetDirectoryName(DirPath)
            ' 2. 提取搜索模式（如 *.txt，默认 *.*）
            Dim searchPattern As String = System.IO.Path.GetFileName(DirPath)
            If String.IsNullOrEmpty(searchPattern) OrElse searchPattern = "." Then
                searchPattern = "*.*"
            End If

            Dim searchOption As System.IO.SearchOption = System.IO.SearchOption.TopDirectoryOnly

            ' 3. 正确调用 EnumerateFiles（目录路径 + 搜索模式 分离）
            For Each filePath As String In System.IO.Directory.EnumerateFiles(pureDirPath, searchPattern, searchOption)
                ' 过滤文件属性（可选：仅保留指定属性的文件）
                Dim attr As System.IO.FileAttributes = System.IO.File.GetAttributes(filePath)
                If (attr And (System.IO.FileAttributes.Normal Or
                          System.IO.FileAttributes.Hidden Or
                          System.IO.FileAttributes.ReadOnly Or
                          System.IO.FileAttributes.System Or
                          System.IO.FileAttributes.Archive)) <> 0 Then
                    fileList.Add(System.IO.Path.GetFileName(filePath))
                End If
            Next
        Catch ex As Exception
            ' 可选：添加异常日志，方便排查问题
            System.Diagnostics.Debug.WriteLine($"读取文件失败：{ex.Message}")
        End Try

        Return fileList.ToArray()
    End Function
End Module