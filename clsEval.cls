VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------
' Class to evaluate generic arithmetic / boolean expressions
'
' Written to parse iLink rules
'-------------------------------------------------------------------
'State constants
Private Const STATE_NONE = 0
Private Const STATE_OPERAND = 1
Private Const STATE_OPERATOR = 2
Private Const STATE_UNARYOP = 3
Private Const UNARY_NEG = "(-)"

Private m_sErrMsg As String

'Evaluates the expression and returns the result.
Public Function Evaluate(sExpression As String) As Double
    Dim sBuffer As String
    Dim nErrPosition As Integer
    
    'First replace the english readable booleans with one-symbol equivalents
    sExpression = UCase(sExpression)
    sExpression = Replace(sExpression, "AND", "&")
    sExpression = Replace(sExpression, "OR", "|")
    sExpression = Replace(sExpression, "NOT", "!")
    
    'Convert to postfix expression
    nErrPosition = InfixToPostfix(sExpression, sBuffer)
    'Raise trappable error if error in expression
    If nErrPosition Then
        Err.Raise vbObjectError + 1001, , m_sErrMsg & " : Column " & CStr(nErrPosition)
    End If
    'Evaluate postfix expression
    Evaluate = DoEvaluate(sBuffer)
End Function

'Converts an infix expression to a postfix expression
'that contains exactly one space following each token.
Private Function InfixToPostfix(sExpression As String, sBuffer As String) As Integer
    Dim i As Integer, ch As String, sTemp As String
    Dim nCurrState As Integer, nParenCount As Integer
    Dim bDecPoint As Boolean
    Dim stkTokens As New clsStack



    nCurrState = STATE_NONE
    nParenCount = 0
    i = 1
    Do Until i > Len(sExpression)
        'Get next character in expression
        ch = Mid$(sExpression, i, 1)
        'Respond to character type
        Select Case ch
            Case "("
                'Cannot follow operand
                If nCurrState = STATE_OPERAND Then
                    m_sErrMsg = "Operator expected"
                    GoTo EvalError
                End If
                'Allow additional unary operators after "("
                If nCurrState = STATE_UNARYOP Then
                    nCurrState = STATE_OPERATOR
                End If
                'Push opening parenthesis onto stack
                stkTokens.Push ch
                'Keep count of parentheses on stack
                nParenCount = nParenCount + 1
            Case ")"
                'Must follow operand
                If nCurrState <> STATE_OPERAND Then
                    m_sErrMsg = "Operand expected"
                    GoTo EvalError
                End If
                'Must have matching open parenthesis
                If nParenCount = 0 Then
                    m_sErrMsg = "Closing parenthesis without matching open parenthesis"
                    GoTo EvalError
                End If
                'Pop all operators until matching "(" found
                sTemp = stkTokens.Pop
                Do Until sTemp = "("
                    sBuffer = sBuffer & sTemp & " "
                    sTemp = stkTokens.Pop
                Loop
                'Keep count of parentheses on stack
                nParenCount = nParenCount - 1
            Case "+", "-", "*", "/", "^", "&", "|", "!"
                'Need a bit of extra code to handle unary operators
                If nCurrState = STATE_OPERAND Then
                    'Pop operators with precedence >= operator in ch
                    Do While stkTokens.StackSize > 0
                        If GetPrecedence(stkTokens.GetPopValue) < GetPrecedence(ch) Then
                            Exit Do
                        End If
                        sBuffer = sBuffer & stkTokens.Pop & " "
                    Loop
                    'Push new operand
                    stkTokens.Push ch
                    nCurrState = STATE_OPERATOR
                ElseIf nCurrState = STATE_UNARYOP Then
                    'Don't allow two unary operators in a row
                    m_sErrMsg = "Operand expected"
                    GoTo EvalError
                Else
                    'Test for unary operator
                    If ch = "-" Then
                        'Push unary minus
                        stkTokens.Push UNARY_NEG
                        nCurrState = STATE_UNARYOP
                    ElseIf ch = "+" Then
                        'Simply ignore positive unary operator
                        nCurrState = STATE_UNARYOP
                    ElseIf ch = "!" Then
                        'Simply ignore positive unary operator
                        stkTokens.Push "!"
                        nCurrState = STATE_UNARYOP
                    Else
                        m_sErrMsg = "Operand expected"
                        GoTo EvalError
                    End If
                End If
            Case "0" To "9", "."
                'Cannot follow other operand
                If nCurrState = STATE_OPERAND Then
                    m_sErrMsg = "Operator expected"
                    GoTo EvalError
                End If
                sTemp = ""
                bDecPoint = False
                Do While InStr("0123456789.", ch)
                    If ch = "." Then
                        If bDecPoint Then
                            m_sErrMsg = "Operand contains multiple decimal points"
                            GoTo EvalError
                        Else
                            bDecPoint = True
                        End If
                    End If
                    sTemp = sTemp & ch
                    i = i + 1
                    If i > Len(sExpression) Then Exit Do
                    ch = Mid$(sExpression, i, 1)
                Loop
                'i will be incremented at end of loop
                i = i - 1
                'Error if number contains decimal point only
                If sTemp = "." Then
                    m_sErrMsg = "Invalid operand"
                    GoTo EvalError
                End If
                sBuffer = sBuffer & sTemp & " "
                nCurrState = STATE_OPERAND
            Case Is <= " "  'Ignore spaces, tabs, etc.
            Case Else
                'Unexpected character
                m_sErrMsg = "Unexpected character encountered"
                GoTo EvalError
        End Select
        i = i + 1
    Loop
    'Expression cannot end with operator
    If nCurrState = STATE_OPERATOR Or nCurrState = STATE_UNARYOP Then
        m_sErrMsg = "Operand expected"
        GoTo EvalError
    End If
    'Check for balanced parentheses
    If nParenCount > 0 Then
        m_sErrMsg = "Closing parenthesis expected"
        GoTo EvalError
    End If
    'Retrieve remaining operators from stack
    Do Until stkTokens.StackSize = 0
        sBuffer = sBuffer & stkTokens.Pop & " "
    Loop
    'Indicate no error
    InfixToPostfix = 0
    Exit Function
EvalError:
    'Report error postion
    InfixToPostfix = i
    Exit Function
End Function

'Returns a number that indicates the relative precedence of an operator.
Private Function GetPrecedence(ch As String) As Integer
    Select Case ch
        Case "+", "-"
            GetPrecedence = 3
        Case "*", "/"
            GetPrecedence = 4
        Case "^"
            GetPrecedence = 5
        Case UNARY_NEG
            GetPrecedence = 10
        Case "&", "|"
            GetPrecedence = 1
        Case "!"
            GetPrecedence = 9
        Case Else
            GetPrecedence = 0
    End Select
End Function

'Evaluates the given expression and returns the result.
'It is assumed that the expression has been converted to
'a postix expression and that a space follows each token.
Private Function DoEvaluate(sExpression As String) As Double
    Dim i As Integer, j As Integer, stkTokens As New clsStack
    Dim sTemp As String, Op1 As Variant, Op2 As Variant

    'Locate first token
    i = 1
    j = InStr(sExpression, " ")
    Do Until j = 0
        'Extract token from expression
        sTemp = Mid$(sExpression, i, j - i)
        If IsNumeric(sTemp) Then
            'If operand, push onto stack
            stkTokens.Push CDbl(sTemp)
        Else
            'If operator, perform calculations
            Select Case sTemp
                Case "+"
                    stkTokens.Push stkTokens.Pop + stkTokens.Pop
                Case "-"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push Op2 - Op1
                Case "*"
                    stkTokens.Push stkTokens.Pop * stkTokens.Pop
                Case "/"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push Op2 / Op1
                Case "^"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push Op2 ^ Op1
                Case "&"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push IIf(Op1 > 0 And Op2 > 0, 1, 0)
                Case "|"
                    Op1 = stkTokens.Pop
                    Op2 = stkTokens.Pop
                    stkTokens.Push IIf(Op1 > 0 Or Op2 > 0, 1, 0)
                Case "!"
                    Op1 = stkTokens.Pop
                    stkTokens.Push IIf(Op1 > 0, 0, 1)
                Case UNARY_NEG
                    stkTokens.Push -stkTokens.Pop
                Case Else
                    'This should never happen (bad tokens caught in InfixToPostfix)
                    Err.Raise vbObjectError + 1002, , "Bad token in Evaluate: " & sTemp
            End Select
        End If
        'Find next token
        i = j + 1
        j = InStr(i, sExpression, " ")
    Loop
    'Remaining item on stack contains result
    If stkTokens.StackSize > 0 Then
        DoEvaluate = stkTokens.Pop
    Else
        'Null expression; return 0
        DoEvaluate = 0
    End If
End Function

'Returns a boolean value that indicates if sChar is a valid
'character to be used as the first character in symbols names
Private Function IsSymbolCharFirst(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If (c >= "A" And c <= "Z") Or InStr("_", c) Then
        IsSymbolCharFirst = True
    Else
        IsSymbolCharFirst = False
    End If
End Function

'Returns a boolean value that indicates if sChar is a valid
'character to be used in symbols names
Private Function IsSymbolChar(sChar As String) As Boolean
    Dim c As String

    c = UCase(Left(sChar, 1))
    If (c >= "A" And c <= "Z") Or InStr("0123456789_", c) Then
        IsSymbolChar = True
    Else
        IsSymbolChar = False
    End If
End Function


