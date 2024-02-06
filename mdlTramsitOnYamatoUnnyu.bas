Attribute VB_Name = "mdlTramsitOnYamatoUnnyu"
Option Explicit
Public Const c_DownloadedFile As String = "Downloaded_File"
Public Enum FixDateEnum
    NearByBase = 0
    BeforeAYearByBase = 1
End Enum
Sub OpenFacade()
    
    TookOneLatestTarget ThisWorkbook.Names("DownloadedFolder").RefersToRange.Value, ThisWorkbook.Names("").RefersToRange.Value, ThisWorkbook.Names("TargetRegs").RefersToRange.Value
    Dim sm As New SettingMemoryObject
    sm.Remember SettingFile
    '�Ō�ɓ��������t�@�C��(����̃V�[�N�G���X�ł͂Ȃ��\��������)�̃^�C���X�^���v�ƃX�g�b�J�t�H���_�̍ŐV�ƃ^�C���X�^���v����v�Ȃ玟�̕]����
    Dim oFS As New Scripting.FileSystemObject
    Dim fs As Scripting.File
    If oFS.FileExists(sm.Items("Downloaded_File").Value) Then
        Set fs = oFS.GetFile(sm.Items("Downloaded_File").Value)
        If sm.Items("Downloaded_DateTime").Value <= fs.DateLastModified Then
            
        End If
    End If
    '�[�i���̓��e��
End Sub
Private Function ValidateYamatounnyuCSV(FileName As String) As Boolean
    ValidateYamatounnyuCSV = False
    Dim oFS As New FileSystemObject
    If Not oFS.FileExists(FileName) Then Exit Function
    
    ValidateYamatounnyuCSV = True
End Function
Public Function GetSpanFromYamatounnyuCSV(FileName As String) As DateSpan
    Dim va() As String
    va = StringSet.GetArrayFromCSV(FileName, True, "��t��")
    Debug.Print va(1, 0)
    Dim da() As Date
    ReDim da(UBound(va))
    Dim sp As DateSpan
    Dim i As Long
    For i = 0 To UBound(va)
        da(i) = CFormatDateWithoutYear(va(i, 0), "mm/dd", FixDateEnum.NearByBase, Now())
    Next
    Set sp = New DateSpan
    sp.ConcreteFromArray da
    Set GetSpanFromYamatounnyuCSV = sp
End Function

Public Sub TookOneLatestTarget(SrcFolder As String, DestFolder As String, FilterReg As String)
    'DownloadedFolder�ɂ���(�ꂩ���ȓ��̍X�V�����܂����t�H�[�}�b�g�̃t�@�C�������͈͓��̔N����)�Ώۂ�����Έ����̃t�H���_��(�ړ��ł����)�ړ�����
    '����������ŐV�X�V�t�@�C���̃^�C���X�^���v�X�V���A������Â���Ζ���
    Dim sMoveSrc As String, sMoveDest As String
    Dim vLine As Variant
    Dim sm As New SettingMemoryObject
    sm mdlTramsitOnYamatoUnnyu.SettingFile
    For Each vLine In Misc.FilesFinderWithRegExp(SrcFolder, FilterReg)
        
    Next
End Sub
Private Function IsNumericB(Expression As String) As Boolean
    '�O-�X �� 0-9�݂̂𐔎��Ƃ���
    Dim i As Long
    For i = 1 To Len(Expression)
        Select Case Mid(Expression, i, 1)
        Case "�O" To "�X", "0" To "9"
        Case Else
            IsNumericB = False: Exit Function
        End Select
    Next
    IsNumericB = True
End Function
'�Ή��t�H�[�}�b�g
'mmdd,mm/dd,m/d,m/dd,mm/d
'/�͔C�ӕ�����ŉ�����(�޲�)�ł���
Public Function CFormatDateWithoutYear(MonthDateString As String, Format As String, NearType As FixDateEnum, Optional BaseDate As Date = #1/1/1900#) As Date
    Const DefaultDate As Date = #1/1/1900#
    'Format�̃o���f�[�g�B�G���[�������Ȃ�
    'dm�͖�����
    If InStr(1, Format, "d", vbTextCompare) < InStr(1, Format, "m", vbTextCompare) Then Err.Raise vbObjectError + 3033, "CFormatDateWithoutYear()", "yet implements."
    'm�Ȃ��Ad�Ȃ��͒�^�O
    If InStr(1, Format, "d", vbTextCompare) < 0 Or InStr(1, Format, "m", vbTextCompare) < 0 Then Err.Raise vbObjectError + 3034, "CFormatDateWithoutYear()", "out of range." & vbNewLine & Format
    'm,d���R�����ȏ�͖�����
    If InStr(1, Format, "ddd", vbTextCompare) > 0 Or InStr(1, Format, "mmm", vbTextCompare) > 0 Then Err.Raise vbObjectError + 3035, "CFormatDateWithoutYear()", "yet implements." & vbNewLine & Format
    
    'm�̕��������擾
    Dim lCount_m As Long, lCount_d As Long
    If InStr(1, Format, "mm", vbTextCompare) > 0 Then
        lCount_m = 2
    ElseIf InStr(1, Format, "m", vbTextCompare) > 0 Then
        lCount_m = 1
    End If
    'd�̕��������擾
    If InStr(1, Format, "dd", vbTextCompare) > 0 Then
        lCount_d = 2
    ElseIf InStr(1, Format, "d", vbTextCompare) > 0 Then
        lCount_d = 1
    End If
    Dim bHasSeparator As Boolean, sSeparator As String
    Dim sWork As String
    Dim sTargetNumStr As String
    Dim lPointIndex As Long
    lPointIndex = 1
    bHasSeparator = InStr(1, Format, "md", vbTextCompare) <= 0
    bHasSeparator = bHasSeparator And InStr(1, Format, "d", vbTextCompare) - InStr(1, Format, "m", vbTextCompare) - lCount_m >= 1
    If bHasSeparator Then
        sSeparator = Mid(Format, 1 + lCount_m, InStr(1, Format, "d", vbTextCompare) - lCount_m - 1)
        If InStr(1, MonthDateString, sSeparator, vbTextCompare) < 0 Then Err.Raise vbObjectError + 3036, "CFormatDateWithoutYear()", "not have separator."
        '�Z�p���[�^�L��Ȃ猩�������ŏ��̃Z�p���[�^����lCount_m�����O�����̕������擾
        sWork = Mid(MonthDateString, InStr(1, MonthDateString, sSeparator, vbTextCompare) - lCount_m, Len(MonthDateString))
        sMonth = Split(sWork, sSeparator)(0)
        sDay = Split(sWork, sSeparator)(1)
    Else
        '�Z�p���[�^�����Ȃ�4�����ȏ�̐����̐擪���擾
        Dim i As Long
        For i = 1 To Len(MonthDateString) - lCount_m - lCount_d + 1
            If IsNumericB(Mid(MonthDateString, i, lCount_m + lCount_d)) Then lPointIndex = i: Exit For
        Next
        If lPointIndex = 0 Then Err.Raise vbObjectError + 3036, "CFormatDateWithoutYear()", "dest not have four number_strings."
        sWork = Mid(MonthDateString, lPointIndex, lCount_m + lCount_d)
    End If
    '��^1�����Ȃ�Ώۂ�1,2����(2�����̏ꍇ1�����ڂ�1)�A��^2����(�ȏ�)�Ȃ�Ώۂ�2����(�ȏ�)
    Dim sDay As String
    Dim sMonth As String
    Dim lMonth As Long, lDay As Long
    If lCount_m = 1 Then
        If IsNumericB(Mid(sWork, 1, 2)) Then  '2�����ڂ̂ݐ����̕]���B1�����ڋ󔒂ł�����
            '2����
            sMonth = Mid(sWork, 1, 2)
        ElseIf IsNumericB(Mid(sWork, 1, 1)) Then
            '1����
            sMonth = Mid(sWork, 1, 1)
        End If
        lMonth = Val(sMonth)
    ElseIf lCount_m > 1 Then
        If True Then
            
        Else
        
        End If
    Else
        Err.Raise vbObjectError + 3036, "CFormatDateWithoutYear()", "out of range." & vbNewLine & Format '�璷�ȃo���f�[�g
    End If
    If NearType = FixDateEnum.NearByBase Then
        'mm,dd���琔�����擾
        If BaseDate = DefaultDate Then BaseDate = Now()
        
    Else
        Err.Raise vbObjectError + 3044, "CDateWithoutYear()", "yet implements."
    End If
End Function

Public Property Get SettingFile() As String
    '�ݒ�t�@�C�����Ȃ��ꍇ�E�K�{����(�ݒ�t�@�C����)�����݂����Ȃ��l�̎��A�ݒ���������B
    '�ݒ�t�@�C���ɐݒ�l�����������݂��Ă��Ă������ڂ������E�K�{���ڃZ���l���ݒ�t�@�C���̂���ƒl���Ⴄ���ɂ͐ݒ�t�@�C������ݒ�Z���l�Ɏʂ�
    Dim oFS As New Scripting.FileSystemObject, bExistsSetting As Boolean
    sSettingFile = oFS.BuildPath(ThisWorkbook.Path, csSettingPath) '���}�V���ɓ��l�̔h��������ꍇ�͔z���̃t�H���_�ɐݒ�t�@�C��������
    'sSettingFile = oFS.BuildPath(CreateObject("WshShell").Environments("AppData"), csSettingPath) '���}�V���͋��L�̐ݒ�l�������������ꍇ��AppData�t�H���_�ɐݒ�t�@�C��������
    bExistsSetting = oFS.FileExists(sSettingFile)
    SettingFile = sSettingFile
End Property

