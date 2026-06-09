' 需确保已导入核心命名空间（文件顶部）
Imports System
Imports System.Collections.Generic

''' <summary>
''' Class to evaluate generic arithmetic / boolean expressions
''' Written to parse iLink rules
''' </summary>
Public Class clsEval
    ' State constants (替换 VB6 Const 为 .NET Const，类型显式声明)
    Private Const STATE_NONE As Integer = 0
    Private Const STATE_OPERAND As Integer = 1
    Private Const STATE_OPERATOR As Integer = 2
    Private Const STATE_UNARYOP As Integer = 3
    Private Const UNARY_NEG As String = "(-)"

    ' 私有字段（保持原命名，增加 _ 前缀符合 .NET 命名规范，可选）
    Private m_sErrMsg As String = String.Empty

    ''' <summary>
    ''' Evaluates the expression and returns the result.
    ''' </summary>
    ''' <param name="sExpression">算术/布尔表达式（支持 AND/OR/NOT、四则运算、括号）</param>
    ''' <returns>计算结果（布尔结果：1=真，0=假；算术结果=具体数值）</returns>
    ''' <exception cref="Exception">表达式解析错误时抛出</exception>
    Public Function Evaluate(ByVal sExpression As String) As Double
        Dim sBuffer As String = String.Empty
        Dim nErrPosition As Integer

        ' 空表达式容错
        If String.IsNullOrWhiteSpace(sExpression) Then
            Throw New Exception("Empty expression is not valid")
        End If

        ' Replace English readable booleans with one-symbol equivalents
        sExpression = sExpression.ToUpperInvariant() ' 替换 ToUpper → ToUpperInvariant（文化无关，更稳定）
        sExpression = sExpression.Replace("AND", "&")
        sExpression = sExpression.Replace("OR", "|")
        sExpression = sExpression.Replace("NOT", "!")

        ' Convert to postfix expression (中缀转后缀/逆波兰表示)
        nErrPosition = InfixToPostfix(sExpression, sBuffer)

        ' Raise exception if error in expression
        If nErrPosition <> 0 Then
            Throw New Exception($"{m_sErrMsg} : Column {nErrPosition}")
        End If

        ' Evaluate postfix expression
        Return DoEvaluate(sBuffer)
    End Function

    ''' <summary>
    ''' Converts an infix expression to a postfix expression
    ''' that contains exactly one space following each token.
    ''' </summary>
    ''' <param name="sExpression">中缀表达式</param>
    ''' <param name="sBuffer">输出：后缀表达式（空格分隔）</param>
    ''' <returns>错误位置（0=无错误）</returns>
    Private Function InfixToPostfix(ByVal sExpression As String, ByRef sBuffer As String) As Integer
        Dim i As Integer = 1
        Dim ch As String = String.Empty
        Dim sTemp As String = String.Empty
        Dim nCurrState As Integer = STATE_NONE
        Dim nParenCount As Integer = 0
        Dim bDecPoint As Boolean = False
        Dim stkTokens As New Stack(Of String)() ' 替换 VB6 Stack → .NET Generic Stack(Of String)

        Do Until i > sExpression.Length
            ' Get next character in expression (替换 VB6 Mid → Substring)
            ch = sExpression.Substring(i - 1, 1)

            ' Respond to character type
            Select Case ch
                Case "("
                    ' Cannot follow operand
                    If nCurrState = STATE_OPERAND Then
                        m_sErrMsg = "Operator expected"
                        Return i
                    End If

                    ' Allow additional unary operators after "("
                    If nCurrState = STATE_UNARYOP Then
                        nCurrState = STATE_OPERATOR
                    End If

                    ' Push opening parenthesis onto stack
                    stkTokens.Push(ch)
                    nParenCount += 1

                Case ")"
                    ' Must follow operand
                    If nCurrState <> STATE_OPERAND Then
                        m_sErrMsg = "Operand expected"
                        Return i
                    End If

                    ' Must have matching open parenthesis
                    If nParenCount = 0 Then
                        m_sErrMsg = "Closing parenthesis without matching open parenthesis"
                        Return i
                    End If

                    ' Pop all operators until matching "(" found
                    Do
                        sTemp = stkTokens.Pop()
                        If sTemp = "(" Then Exit Do
                        sBuffer &= $"{sTemp} "
                    Loop
                    nParenCount -= 1

                Case "+", "-", "*", "/", "^", "&", "|", "!"
                    ' Handle unary operators
                    If nCurrState = STATE_OPERAND Then
                        ' Pop operators with precedence >= operator in ch
                        While stkTokens.Count > 0 AndAlso GetPrecedence(stkTokens.Peek()) >= GetPrecedence(ch)
                            sBuffer &= $"{stkTokens.Pop()} "
                        End While

                        ' Push new operand
                        stkTokens.Push(ch)
                        nCurrState = STATE_OPERATOR
                    ElseIf nCurrState = STATE_UNARYOP Then
                        ' Don't allow two unary operators in a row
                        m_sErrMsg = "Operand expected"
                        Return i
                    Else
                        ' Test for unary operator
                        Select Case ch
                            Case "-"
                                stkTokens.Push(UNARY_NEG)
                                nCurrState = STATE_UNARYOP
                            Case "+"
                                nCurrState = STATE_UNARYOP
                            Case "!"
                                stkTokens.Push("!")
                                nCurrState = STATE_UNARYOP
                            Case Else
                                m_sErrMsg = "Operand expected"
                                Return i
                        End Select
                    End If

                Case "0" To "9", "."
                    ' Cannot follow other operand
                    If nCurrState = STATE_OPERAND Then
                        m_sErrMsg = "Operator expected"
                        Return i
                    End If

                    sTemp = String.Empty
                    bDecPoint = False

                    ' 提取完整数字（包括小数点）
                    Do While "0123456789.".Contains(ch)
                        If ch = "." Then
                            If bDecPoint Then
                                m_sErrMsg = "Operand contains multiple decimal points"
                                Return i
                            Else
                                bDecPoint = True
                            End If
                        End If

                        sTemp &= ch
                        i += 1
                        If i > sExpression.Length Then Exit Do
                        ch = sExpression.Substring(i - 1, 1)
                    Loop

                    i -= 1

                    ' Error if number contains decimal point only
                    If sTemp = "." Then
                        m_sErrMsg = "Invalid operand"
                        Return i
                    End If

                    sBuffer &= $"{sTemp} "
                    nCurrState = STATE_OPERAND

                Case Else
                    ' Unexpected character (过滤空白字符，原 VB6 未处理)
                    If Char.IsWhiteSpace(ch(0)) Then
                        i += 1
                        Continue Do
                    End If
                    m_sErrMsg = $"Unexpected character encountered: '{ch}'"
                    Return i
            End Select

            i += 1
        Loop

        ' Expression cannot end with operator
        If nCurrState = STATE_OPERATOR OrElse nCurrState = STATE_UNARYOP Then
            m_sErrMsg = "Operand expected"
            Return i
        End If

        ' Check for balanced parentheses
        If nParenCount > 0 Then
            m_sErrMsg = "Closing parenthesis expected"
            Return i
        End If

        ' Retrieve remaining operators from stack
        While stkTokens.Count > 0
            sBuffer &= $"{stkTokens.Pop()} "
        End While

        ' Indicate no error
        Return 0
    End Function

    ''' <summary>
    ''' Returns a number that indicates the relative precedence of an operator.
    ''' </summary>
    ''' <param name="ch">运算符</param>
    ''' <returns>优先级数值（越大优先级越高）</returns>
    Private Function GetPrecedence(ByVal ch As String) As Integer
        Select Case ch
            Case "+", "-"
                Return 3
            Case "*", "/"
                Return 4
            Case "^"
                Return 5
            Case UNARY_NEG
                Return 10
            Case "&", "|"
                Return 1
            Case "!"
                Return 9
            Case Else
                Return 0
        End Select
    End Function

    ''' <summary>
    ''' Evaluates the given postfix expression and returns the result.
    ''' </summary>
    ''' <param name="sExpression">后缀表达式（空格分隔）</param>
    ''' <returns>计算结果</returns>
    ''' <exception cref="Exception">无效运算符/栈异常时抛出</exception>
    Private Function DoEvaluate(ByVal sExpression As String) As Double
        Dim stkTokens As New Stack(Of Double)()
        ' 修复 Split 报错的核心代码
        Dim rawTokens As String() = sExpression.Split(New Char() {" "c}, StringSplitOptions.None)
        Dim tokens As New List(Of String)()
        For Each token As String In rawTokens
            If Not String.IsNullOrWhiteSpace(token) Then
                tokens.Add(token)
            End If
        Next
        Dim tokenArray As String() = tokens.ToArray()

        For Each token In tokenArray ' 遍历修复后的 token 数组
            If Double.TryParse(token, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, Nothing) Then
                stkTokens.Push(Double.Parse(token, System.Globalization.CultureInfo.InvariantCulture))
            Else
                Dim op1, op2 As Double

                Select Case token
                    Case "+"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '+' operator")
                        stkTokens.Push(stkTokens.Pop() + stkTokens.Pop())
                    Case "-"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '-' operator")
                        op1 = stkTokens.Pop()
                        op2 = stkTokens.Pop()
                        stkTokens.Push(op2 - op1)
                    Case "*"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '*' operator")
                        stkTokens.Push(stkTokens.Pop() * stkTokens.Pop())
                    Case "/"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '/' operator")
                        op1 = stkTokens.Pop()
                        If op1 = 0 Then Throw New DivideByZeroException("Division by zero is not allowed")
                        op2 = stkTokens.Pop()
                        stkTokens.Push(op2 / op1)
                    Case "^"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '^' operator")
                        op1 = stkTokens.Pop()
                        op2 = stkTokens.Pop()
                        stkTokens.Push(Math.Pow(op2, op1))
                    Case "&"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '&' (AND) operator")
                        op1 = stkTokens.Pop()
                        op2 = stkTokens.Pop()
                        stkTokens.Push(If(op1 <> 0 AndAlso op2 <> 0, 1.0, 0.0))
                    Case "|"
                        If stkTokens.Count < 2 Then Throw New Exception("Insufficient operands for '|' (OR) operator")
                        op1 = stkTokens.Pop()
                        op2 = stkTokens.Pop()
                        stkTokens.Push(If(op1 <> 0 OrElse op2 <> 0, 1.0, 0.0))
                    Case "!"
                        If stkTokens.Count < 1 Then Throw New Exception("Insufficient operands for '!' (NOT) operator")
                        op1 = stkTokens.Pop()
                        stkTokens.Push(If(op1 <> 0, 0.0, 1.0))
                    Case UNARY_NEG
                        If stkTokens.Count < 1 Then Throw New Exception("Insufficient operands for unary '-' operator")
                        stkTokens.Push(-stkTokens.Pop())
                    Case Else
                        Throw New Exception($"Bad token in Evaluate: '{token}'")
                End Select
            End If
        Next

        If stkTokens.Count = 0 Then
            Return 0.0
        ElseIf stkTokens.Count > 1 Then
            Throw New Exception("Invalid expression: too many operands")
        Else
            Return stkTokens.Pop()
        End If
    End Function
End Class