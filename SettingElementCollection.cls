VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingElementCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IClassRule

Private m_IsDisposed As Boolean
Private m_Elements As Scripting.Dictionary
Private m_IsLoadCompleted As Boolean
Private m_Saved As Boolean
Public Sub Clear()
    Dim vItem As Variant
    Dim eItem As SettingElement
    For Each vItem In m_Elements.Items
        Set eItem = vItem
        eItem.Dispose True
    Next
    m_Elements.RemoveAll
End Sub
Private Property Get IClassRule_IsDisposed() As Boolean
    IClassRule_IsDisposed = m_IsDisposed
End Property
Private Property Let IClassRule_IsDisposed(ByVal IsDisposing As Boolean)
    m_IsDisposed = IsDisposing
End Property

Private Property Get IClassRule_IsLoadCompleted() As Boolean
    IClassRule_IsLoadCompleted = m_IsLoadCompleted
End Property
Private Property Let IClassRule_IsLoadCompleted(ByVal IsLoadCompleting As Boolean)
    m_IsLoadCompleted = IsLoadCompleting
End Property

Public Function Exists(Key As String) As Boolean
    If m_Elements.Exists(Key) Then
        Exists = m_Elements(Key).Description <> "Proxy"
    Else
        Exists = False
    End If
End Function
Public Function Item(Index As Variant) As SettingElement
Attribute Item.VB_UserMemId = 0
    If (VarType(Index) = vbLong) Or (VarType(Index) = vbInteger) Then
        Set Item = ItemOnNumbering(CLng(Index))
    ElseIf VarType(Index) = vbString Then
        Set Item = ItemOnKeyName(CStr(Index))
    Else
        Err.Raise vbObjectError + 204, TypeName(Me) & ".Item", "TypeName=""" & TypeName(Index) & """はインデックスの範囲外です。"
    End If
End Function
Private Function ItemOnNumbering(Index As Long) As SettingElement
    Dim lElementCount As Long
    Dim vLine As Variant
    Dim elmCast As SettingElement
    For Each vLine In m_Elements.Items
        If TypeName(vLine) = "SettingElement" Then
            If Index > lElementCount Then
                lElementCount = lElementCount + 1
            Else
                Set elmCast = vLine
                Set ItemOnNumbering = elmCast
                Exit Function
            End If
        Else
            Err.Raise vbObjectError + 104, TypeName(Me) & ".Item", "VarType=""" & TypeName(vLine) & """は許されない保存アイテムです"
        End If
    Next
    Err.Raise vbObjectError + 122, TypeName(Me) & ".Item", "指定されたIndex=" & Index & " は 0 <= Index < " & Count & "の範囲にないといけません"
End Function

Private Function ItemOnKeyName(RelationalKey As String) As SettingElement
    Dim eol As SettingElement
    If m_Elements.Exists(RelationalKey) Then
        
        If TypeName(m_Elements(RelationalKey)) = "SettingElement" Then
            Set ItemOnKeyName = m_Elements(RelationalKey)
        ElseIf VarType(m_Elements(RelationalKey)) = vbEmpty Then
            Set eol = New SettingElement
            eol.Concrete RelationalKey, Default, SettingEnum.ForString, True, "Proxy"
            Set ItemOnKeyName = eol
            Set m_Elements(RelationalKey) = eol
        ElseIf m_Elements(RelationalKey).Description = "Proxy" Then
            Set ItemOnKeyName = m_Elements(RelationalKey)
        Else
            Err.Raise 23, TypeName(Me) & ".Key()", "SettingItemに型違いが入りました"
        End If
    Else
        'ダミーproxyを立てる
        Set eol = New SettingElement
        eol.Concrete RelationalKey, Default, SettingEnum.ForString, True, "Proxy"
        Set ItemOnKeyName = eol
        m_Elements.Add RelationalKey, eol
    End If
End Function
Public Property Get Count() As Long
    Count = m_Elements.Count
End Property
Public Function Add(ByRef OneElement As SettingElement) As SettingElement

    If Not m_Elements.Exists(OneElement.Name) Then
        m_Elements.Add OneElement.Name, OneElement
        Set Add = OneElement
    Else
        Set m_Elements.Item(OneElement.Name) = OneElement
        Set Add = m_Elements(OneElement.Name)
    End If
End Function

Private Sub Class_Initialize()
    Set m_Elements = New Scripting.Dictionary
    Saved = True
End Sub

Public Sub Remove(Key As String)
    m_Elements(Key).Dispose True
    m_Elements.Remove Key
End Sub

Public Property Get Saved() As Boolean
    Dim vLine As Variant
    Dim elLine  As SettingElement
    Dim i As Long
    'xxx- m_ElementsとMe.Item(x)が連動していない
    For i = 0 To Me.Count - 1
        Set elLine = Me.Item(i)
        Saved = m_Saved And elLine.Saved 'ダウントリガ
    Next
End Property
Private Property Let Saved(ByVal vNewSavedValue As Boolean)
    m_Saved = vNewSavedValue
End Property
