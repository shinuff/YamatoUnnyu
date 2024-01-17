VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'v0.04
'@Folder("Library")
'SheetCondition-(Collection)-FilterClass
Private m_Enable As Boolean
Private m_Order As Long
Private m_Criteria1 As Variant
Private m_Criteria2 As Variant
Private m_Operator As Variant
Private filterArray()

Public Sub Create(NewFilter As Filter, NewOrder As Long)
    Enable = Not (NewFilter Is Nothing)
    Order = NewOrder
    If Enable Then
        If NewFilter.Count > 0 Then
            Criteria1 = NewFilter.Criteria1
        End If
        Operator = NewFilter.Operator
        If Operator <> 0 And (NewFilter.Operator And xlFilterValues) <> xlFilterValues Then
            Criteria2 = NewFilter.Criteria2
        End If
    End If
End Sub

Public Property Get Enable() As Boolean
    Enable = m_Enable
End Property
Private Property Let Enable(ByVal vNewValue As Boolean)
    m_Enable = vNewValue
End Property

Public Property Get Order() As Long
    Order = m_Order
End Property
Public Property Let Order(ByVal vNewValue As Long)
    m_Order = vNewValue
End Property

Public Property Get Criteria2() As Variant
    Criteria2 = m_Criteria2
End Property
Private Property Let Criteria2(ByVal vNewValue As Variant)
    m_Criteria2 = vNewValue
End Property

Public Property Get Criteria1() As Variant
    If TypeName(m_Criteria1) = "Interior" Then
        Set Criteria1 = m_Criteria1
    Else
        Criteria1 = m_Criteria1
    End If
End Property
Private Property Let Criteria1(ByVal vNewValue As Variant)
    If TypeName(vNewValue) = "Interior" Then
        Set m_Criteria1 = vNewValue
    Else
        m_Criteria1 = vNewValue
    End If
End Property

Public Property Get Operator() As Variant
    Operator = m_Operator
End Property
Private Property Let Operator(ByVal vNewValue As Variant)
    m_Operator = vNewValue
End Property

Private Sub Class_Initialize()
    Operator = Empty
    Criteria1 = Empty
    Criteria2 = Empty
End Sub
