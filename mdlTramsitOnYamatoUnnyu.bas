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
    '�Ō�ɓ��������t�@�C��(����̃V�[�N�G���X�ł͂Ȃ��\��������)�̃^�C���X�^���v�ƃX�g�b�J�t�H���_�̍ŐV�ƃ^�C���X�^���v����v�Ȃ�
    Dim oFS As New Scripting.FileSystemObject
    Dim fs As Scripting.File
    If oFS.FileExists(sm.Items("Downloaded_File").Value) Then
        Set fs = oFS.GetFile(sm.Items("Downloaded_File").Value)
        If sm.Items("Downloaded_DateTime").Value <= fs.DateLastModified Then
            
        End If
    End If
End Sub
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

Public Function CFormatDateWithoutYear(MonthDateString As String, Format As String, NearType As FixDateEnum, Optional BaseDate As Date = #1/1/1900#) As Date
    Const DefaultDate As Date = #1/1/1900#
    If BaseDate = DefaultDate Then
        
    Else
        
    End If
    Select Case NearType
    Case NearByBase
        
    Case Else
        Err.Raise vbObjectError + 3044, "CDateWithoutYear()", "yet implements."
    End Select
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

