VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeliveryForYamatoUnnyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DeliverySpan As DateSpan
Private m_Modify As Date
Private m_Path As String


Public Property Get DeliverySpan() As DateSpan
    DeliverySpan = m_DeliverySpan
End Property

Public Property Let DeliverySpan(ByVal vNewValue As DateSpan)
    m_DeliverySpan = vNewValue
End Property

Public Property Get Modify() As Date
    Modify = m_Modify
End Property

Public Property Let Modify(ByVal vNewValue As Date)
    m_Modify = vNewValue
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal vNewValue As String)
    m_Path = vNewValue
End Property

Public Function ConcreteFromCSV(Path As String, Optional GenerateProxy As Boolean = False) As DeliveryForYamatoUnnyu
    
    If GenerateProxy Then Set ConcreteFromCSV = Me: Exit Function
    Dim ass()  As String, i As Long
    ass = StringSet.GetArrayFromCSV(Path, True, "Žó•t“ú")
    Dim ads() As Date: ReDim ads(UBound(ass))
    For i = 0 To uboud(ass)
        ads(i) = CFormatDateWithoutYear(ass(i), "mm/dd", FixDateEnum.NearByBase, Now())
    Next
    Set ConcreteFromCSV = Me
End Function
