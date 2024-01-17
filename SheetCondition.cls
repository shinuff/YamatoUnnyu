VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Library")
'CheckOut 売上実績表.xlsm(update 2019/3/20)
'CheckIn 一般貯蔵品調査表.xls(update 2010/4/18)
'Version 0.04.0
'Required>Import FilterClass.cls(v0.04upper)

Option Explicit

Private m_SCUP As Boolean
Private m_ProtectContents As Boolean
Private m_CodeName As String
Private m_ActiveSheet As Worksheet
Private m_CurrentRange As Range
Private m_SheetVisibility As XlSheetVisibility
Private m_TargetSheet As Worksheet
Private m_IsAutoFilter As Boolean
Private m_IsAdvancedFilter As Boolean
Private m_AutoFiltered  As AutoFilter
Private m_AutoFilterCollection As New Collection
Private m_FilterRange As Range
Private currentFiltRange As String

Public Sub Stock(Target As Worksheet)
    Dim shtCurrent As Worksheet
    Set shtCurrent = ActiveSheet
    m_SheetVisibility = Target.Visible
    Target.Visible = xlSheetVisible
    Target.Activate
    
    If Target.AutoFilterMode Then
        currentFiltRange = Target.AutoFilter.Range.Address
        Set m_AutoFiltered = Target.AutoFilter
        m_IsAutoFilter = True
        Dim i As Long, fil As Filter, lCounter As Long
        Dim oFil As FilterClass
        Set m_FilterRange = Target.AutoFilter.Range
        For Each fil In Target.AutoFilter.Filters
            i = i + 1
            If fil.On Then
                lCounter = lCounter + 1
                Set oFil = New FilterClass
                oFil.Create fil, lCounter
                m_AutoFilterCollection.Add oFil, "Key" & lCounter  'xxx-- ActiveSheetを使って初期化する→バグ
                m_AutoFilterCollection(lCounter).Create fil, i
            End If
        Next
        m_IsAutoFilter = True
        If Target.FilterMode Then
            Target.ShowAllData
        End If
    Else
        If Target.FilterMode Then
            'AdvancedFilterの場合
            Set m_FilterRange = Target.UsedRange.Cells(1, 1).CurrentRegion
            m_IsAdvancedFilter = True
            Target.ShowAllData
        End If
    End If
    m_CodeName = Target.CodeName
    Set m_TargetSheet = Target
    m_SCUP = Target.Application.ScreenUpdating
    Target.Application.ScreenUpdating = False
    Set m_ActiveSheet = ActiveSheet
    Set m_CurrentRange = Selection
    m_ProtectContents = Target.ProtectContents
    On Error GoTo ReEnterUnprotectKey
    If m_ProtectContents Then Target.Unprotect
    On Error GoTo 0
    shtCurrent.Activate
Exit Sub
ReEnterUnprotectKey:
    If Err.Number = 1004 Then Resume
    Err.Raise Err.Number, , Err.Description
End Sub

Public Sub Restore()
    If m_SheetVisibility = xlSheetVisible Then
        m_TargetSheet.Visible = m_SheetVisibility
        m_TargetSheet.Activate
        If m_ProtectContents Then m_TargetSheet.Protect
        On Error Resume Next
        m_CurrentRange.Select
        If Err.Number = 424 Then
            Cells(1, 1).Select
        End If
        On Error GoTo 0
    Else
        m_TargetSheet.Visible = xlSheetVisible
        m_TargetSheet.Activate
        If m_ProtectContents Then m_TargetSheet.Protect
        On Error Resume Next
        m_CurrentRange.Select
        If Err.Number = 424 Then
            Cells(1, 1).Select
        End If
        On Error GoTo 0
        m_TargetSheet.Visible = m_SheetVisibility
    End If
    If m_IsAutoFilter Then
        'AutoFilterの復帰
        Dim fil As FilterClass, vFil As Variant
        For Each vFil In m_AutoFilterCollection
            Set fil = vFil
            If fil.Operator Then
                If fil.Operator = xlFilterValues Then
                     m_FilterRange.AutoFilter Field:=fil.Order, Criteria1:=fil.Criteria1, Operator:=fil.Operator
                Else
                     m_FilterRange.AutoFilter Field:=fil.Order, Criteria1:=fil.Criteria1, Operator:=fil.Operator, Criteria2:=fil.Criteria2
                End If
            Else
                m_FilterRange.AutoFilter Field:=fil.Order, Criteria1:=fil.Criteria1
            End If
        Next
    ElseIf m_IsAdvancedFilter Then
        If Misc.NameExists(m_TargetSheet.Name & "!Criteria") Then
            'AdvancedFilterの復帰
            m_FilterRange.Cells(1, 1).CurrentRegion.AdvancedFilter xlFilterInPlace, m_FilterRange.Worksheet.Range("Criteria")
        End If
    End If
    If m_SCUP Then Application.ScreenUpdating = m_SCUP
    m_ActiveSheet.Activate
End Sub

Public Property Get CodeName() As String
    CodeName = m_CodeName
End Property

