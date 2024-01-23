VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IClassRule

Private m_Value As Variant
Private m_Name As String
Private m_Required As Boolean
Private m_Purpose As SettingEnum
Private m_IsLoadCompleted As Boolean
Private m_IsDisposed As Boolean
Private m_Description As String
Private m_Saved As Boolean

Private Sub Class_Terminate()
    Dispose False
End Sub
Public Sub Dispose(ByVal IsDisposing As Boolean)
    Dim CastMe As IClassRule
    Set CastMe = Me
    CastMe.IsDisposed = True
    If IsDisposing Then
        Name = ""
        Value = ""
        Purpose = SettingEnum.None
    Else
        m_Name = ""
        m_Value = Null
        m_Purpose = SettingEnum.None
    End If
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

Public Sub Concrete(ElementName As String, ElementValue As Variant, ElementPurpose As SettingEnum, IsSettingRequired As Boolean, Optional ElementDescription As String = "")
    Purpose = ElementPurpose
    Value = ElementValue
    Name = ElementName
    Required = IsSettingRequired
    Description = ElementDescription
    Saved = ElementDescription <> "Proxy"
    m_IsLoadCompleted = True
End Sub
Private Sub RegulizeValue()
    'm_Valueは初期化されていない可能性がある。この時点ではまだ読取プロパティをアクセサには使えない。
    If (Purpose = SettingEnum.None) Or (Description = "Proxy") Or ((VarType(m_Value) = vbString) And (m_Value = "")) Then m_Value = vbEmpty: Exit Sub
    Select Case Purpose
    Case SettingEnum.ForString
        'Stringの場合
        m_Value = CStr(m_Value)
    Case SettingEnum.ForPath
        Dim oFS As Scripting.FileSystemObject
        Set oFS = New Scripting.FileSystemObject
        m_Value = oFS.GetAbsolutePathName(m_Value)
        Set oFS = Nothing
    End Select
End Sub
Public Property Get Value() As Variant
    
    If Not m_IsLoadCompleted Then Err.Raise vbObjectError + 24, TypeName(Me) & ".Value", "値が入っていません"
    RegulizeValue
    Value = m_Value
End Property
Public Property Let Value(ByVal ElementValue As Variant)
    Saved = Saved And (m_Value = ElementValue)
    m_Value = ElementValue
    If Description = "Proxy" And ElementValue <> CStr(Default) Then Description = ""  ' 暫定キーから値を入れた時にProxy状態から通常に遷移
    RegulizeValue
End Property

Public Property Get Name() As String
    Name = m_Name
End Property
Private Property Let Name(ByVal ElementName As String)
    m_Name = ElementName
End Property

Public Property Get Required() As Boolean
    Required = m_Required
End Property
Public Property Let Required(ByVal IsSettingRequired As Boolean)
    m_Required = IsSettingRequired
End Property

Public Property Get Purpose() As SettingEnum
    Purpose = m_Purpose
End Property
Public Property Let Purpose(ByVal PurposeValue As SettingEnum)
    m_Purpose = PurposeValue
    RegulizeValue
End Property

Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Let Description(ByVal vNewValue As String)
    m_Description = vNewValue
End Property

Private Sub Class_Initialize()
    m_Value = ""
    Saved = True
End Sub

Public Property Get Saved() As Boolean
    Saved = m_Saved
End Property

Private Property Let Saved(ByVal NewSavedValue As Boolean)
    m_Saved = NewSavedValue
End Property
