VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private m_StackArray() As Variant
Private m_nStackSize As Integer
Private m_nCurrPos As Integer

'Granularity for growing stack
Private Const INC_SIZE = 10

'Read-only property returns the number of items on the stack
Public Property Get StackSize() As Integer
    StackSize = m_nCurrPos
End Property

'Pushes a value onto the stack.
'The stack is made bigger if required.
Public Sub Push(Value As Variant)
    m_nCurrPos = m_nCurrPos + 1
    If m_nCurrPos > m_nStackSize Then
        m_nStackSize = m_nStackSize + INC_SIZE
        ReDim Preserve m_StackArray(1 To m_nStackSize)
    End If
    m_StackArray(m_nCurrPos) = Value
End Sub

'Returns the value on the top of the stack and removes it
'from the stack. A run-time error is raised if the stack is empty.
Public Function Pop() As Variant
    If m_nCurrPos > 0 Then
        Pop = m_StackArray(m_nCurrPos)
        m_nCurrPos = m_nCurrPos - 1
    Else
        Err.Raise vbObjectError + 1011, , "Pop method invoked on empty stack"
    End If
End Function

'Returns the value on top of the stack without removing it
'from the stack. A run-time error is raised if the stack is empty.
Public Function GetPopValue() As Variant
    If m_nCurrPos > 0 Then
        GetPopValue = m_StackArray(m_nCurrPos)
    Else
        Err.Raise vbObjectError + 1012, , "GetPopValue method invoked on empty stack"
    End If
End Function

'Initialization routine
Private Sub Class_Initialize()
    m_nStackSize = 0
    m_nCurrPos = 0
End Sub
