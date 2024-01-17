Attribute VB_Name = "InitSettingHernes"
Option Explicit

'�ݒ�t�@�C���I�u�W�F�N�g���p�̂��߂̃n�[�l�X
Const csThisWorkbookKey As String = "ThisWorkbook"
Const csDataKey As String = "Download_Files" '�K�{�ݒ�̂����̈�B�Z���ɖ��O���`���Ă����B
Const csSettingPath As String = "Setting.config"
Private m_sm As New SettingMemoryObject

Sub TestLoadInit()
    LoadInit
End Sub

Function LoadInit() As Boolean
    '�p�^�[���P:�ݒ�t�@�C���D��A��ɐݒ�Z���ɐݒ�l��p��
    '���g�̃p�X���Z���lA�Ƃ��ĕۑ��BA=blank�Ȃ�Đݒ�B
    Dim sSettingFile As String
    '�ݒ�t�@�C�����Ȃ��ꍇ�E�K�{����(�ݒ�t�@�C����)�����݂����Ȃ��l�̎��A�ݒ���������B
    '�ݒ�t�@�C���ɐݒ�l�����������݂��Ă��Ă������ڂ������E�K�{���ڃZ���l���ݒ�t�@�C���̂���ƒl���Ⴄ���ɂ͐ݒ�t�@�C������ݒ�Z���l�Ɏʂ�
    Dim oFS As New Scripting.FileSystemObject, bExistsSetting As Boolean
    sSettingFile = oFS.BuildPath(ThisWorkbook.Path, csSettingPath) '���}�V���ɓ��l�̔h��������ꍇ�͔z���̃t�H���_�ɐݒ�t�@�C��������
    'sSettingFile = oFS.BuildPath(CreateObject("WshShell").Environments("AppData"), csSettingPath) '���}�V���͋��L�̐ݒ�l�������������ꍇ��AppData�t�H���_�ɐݒ�t�@�C��������
    bExistsSetting = oFS.FileExists(sSettingFile)
    Set m_sm = New SettingMemoryObject
    
    m_sm.Remember sSettingFile
    'Excel���ŃR���g���[������ݒ胊�Z�b�g�g���K(���OThisWorkbook�l���󔒂ɂ��ĕۑ����čăI�[�v��)
    If IsEmpty(ThisWorkbook.Names(csThisWorkbookKey).RefersToRange.Value) Then
        m_sm.Items(csThisWorkbookKey).Value = ""
    End If
    
    'C���L��
    'Stop
    If ThisWorkbook.Names(csDataKey).RefersToRange.Value <> m_sm(csDataKey).Value Then
        ThisWorkbook.Names(csDataKey).RefersToRange.Value = m_sm(csDataKey).Value
    End If
    Do While Not oFS.FileExists(m_sm.Items(csDataKey).Value)
        SetRequiredItem
        '�_�C�A���O��ESC�Ŕ�����
        If m_sm.Items(csDataKey).Value = False Then m_sm.Refresh: Exit Function
    Loop
    If ThisWorkbook.Names(csThisWorkbookKey).RefersToRange.Value <> ThisWorkbook.FullName Then
        ThisWorkbook.Names(csThisWorkbookKey).RefersToRange.Value = ThisWorkbook.FullName
        m_sm.Items(csThisWorkbookKey).Value = ThisWorkbook.FullName
    ElseIf m_sm.Items(csThisWorkbookKey).Value <> ThisWorkbook.FullName Then
        m_sm.Items(csThisWorkbookKey).Value = ThisWorkbook.FullName
    End If
    'D�L���H
    If bExistsSetting Then
        'Nop
    Else
        'B���݁H

        If Not ThisWorkbook.Names(csDataKey).RefersToRange.Value = m_sm.Items(csDataKey).Value Then ThisWorkbook.Names(csDataKey).RefersToRange.Value = m_sm.Items(csDataKey).Value
        m_sm.Memorize sSettingFile 'Non Proxy
        Exit Function
    End If
    LoadInit = True
ExitInitializing:
    m_sm.Memorize SettingFile
End Function

Sub SetRequiredItem()
    '�_�C�A���O�ɂ��ݒ肷��B�����Őݒ肷��l�ȊO�̉��̐ݒ�l���K�{����m��Ȃ��B
    '�����̐ݒ�Z���Ɣ�r���ĈقȂ�����ݒ�Z���ɒl�ݒ�
    '�ݒ�t�@�C���D��
    Dim sPaymentDataPath As String, vPath As Variant
    vPath = Misc.GetOpenFilenameOnInitialDir("Microsoft Access Database�t�@�C��(*.mdb),*.mdb,���ׂẴt�@�C�� (*.*),*.*", 1, "�z���䒠���w�肵�Ă��������B", "", False, ThisWorkbook.Path & "\dataStore", True)
    ThisWorkbook.Names(csDataKey).RefersToRange.Value = vPath
    m_sm(csDataKey).Value = vPath
End Sub
Sub ImportBranchCode()
'
    Sheets("�x�X").Select
    Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight)).ClearContents
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & ThisWorkbook.Path & "\" & ThisWorkbook.Names("BranchFileName").RefersToRange.Value, Destination:=Range("A1"))
        .Name = "Z_Branch"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2, 2)
        .TextFileFixedColumnWidths = Array(4, 3)
        .Refresh BackgroundQuery:=False
    End With
    Range("D1:F1").Cut Destination:=Range("A1:C1")
End Sub

Sub ImportBankCode()
    Sheets("��s").Select
    Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight)).ClearContents
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & ThisWorkbook.Path & "\" & ThisWorkbook.Names("BankFileName").RefersToRange.Value, Destination:=Range("A1"))
        .Name = "Z_Bank"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2)
        .TextFileFixedColumnWidths = Array(4)
        .Refresh BackgroundQuery:=False
    End With
    Range("C1:D1").Cut Destination:=Range("A1:B1")
End Sub
