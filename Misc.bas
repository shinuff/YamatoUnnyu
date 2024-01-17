Attribute VB_Name = "Misc"
'CheckIn ���[�����X�^�[����.xls(update 2023/7/14)(x86,x64���p)
'CheckIn ����\�Z�O�N���ѕ\.xlsm(update 2019/6/29)
'CheckIn �����e�����x���ʒm��.xls(update 2018/4/15)
'CheckIn �����U���x�������Z�p.xls(update 2018/3/4)
'CheckIn ���|�V�X�e���o�̓g�����X�R�[�_.xls(update 2016/1/23)
'CheckIn �����\�݌v�e���v���[�g.xls(update 2016/1/17)
'CheckIn ���l�����Ք�z�v�Z�\RT����.xls(update 2014/3/2)
'CheckIn ���|�V�X�e���o�̓g�����X�R�[�_.xls(update 2013/5/3)
'CheckIn ����X�`�[130212.xls(update 2013/4/28)
'CheckIn �����\�݌v�e���v���[�g.xls(update 2013/1/14)
'CheckIn ���[�����X�^�[.xls(update 2013/1/9)
'CheckIn �I���\�W�v.xls(New)(update 2012/10/8)
'CheckIn �����̌��Z.xls(update 2012/9/15)
'CheckIn ���[�����X�^�[.xls(update 2012/7/28)
'CheckIn ���}�g�^�A.xls(update 2012/6/27)
'CheckIn �I���\�W�v.xls(update 2012/3/3)
'CheckIn �I���\�W�v.xls(update 2010/9/3)
'CheckIn �I���\�W�v.xls(update 2010/6/27)
'CheckIn ���[�����X�^�[.xls(update 2010/6/3)
'CheckIn �I���\�W�v.xls(update 2010/5/16)
'CheckIn �I���\�e���v���[�g.xls(update 2010/4/8)
'CheckIn �m�[�}�b�h����.xls(update 2010/3/22)
'CheckIn ���[�����X�^�[.xls(update 2010/3/16)
'CheckIn ���ɊǗ��\2009.12.xls(update 2010/1/5)
'CheckIn ���i�����i�����\.xls(update 2009/5/10)
'CheckIn TypedList.xls
'CheckIn �����݌ɏ��i�񍐏��e���v���[�g.xls
'CheckIn ���i�����i�����\.xls
'CheckIn �����݌ɏ��i�񍐏��e���v���[�g.xls
'CheckIn Bool2.xls
'CheckIn �����݌ɏ��i�񍐏��e���v���[�g.xls

'�Q�Ɛݒ�   : Microsft ActiveX Data Objects 2.1 or later 2.5 or 2.6 or 2.7 or 2.8 Library
'           : Microsoft Scripting Runtime
'           : Windows Script Host Object Model
'           : Microsoft VBScript Regular Expressions 5.5
'           : Microsoft XML, v3.0
'Imports    : SheetCondition.cls

Option Explicit

Public Enum DecisionDirectionFlagEnum
    WithoutConcidering = 0
    SourceSide = 1
    DestinationSide = 2
    BothSide = 3
End Enum

Public Enum SettingValueTypeEnum
    None = 0
    First = 1
    ForPath = 1
    ForString = 2
    ForNumber = 3
    ForNumeric = 4
    Last = 3
End Enum

Private Type BROWSEINFO
    hwndOwner   As Long
    pidlRoot    As Long
    pszDisplayName As String
    lpszTitle   As String
    ulFlags     As Long
    lpfn        As Long
    lParam      As String   ' LPSTR�Ŏ󂯓n��
    iImage      As Long
End Type

Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTION = WM_USER + 102
Private Const BIF_RETURNONLYFSDIRS = &H1

#If VBA7 Then
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef ppsaOut() As Any) As Long
Private Declare PtrSafe Sub CopyMemoryFromArray Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef RetPointer As Long, SrcArray() As Any, Optional ByVal Length As Long = 4&)
Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare PtrSafe Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" _
        Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" _
        (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare PtrSafe Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal Msg As Long, _
        ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
Private Const VT_BYREF = &H4000
Private Const VARIANT_DATA_OFFSET As Long = 8

Private Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32.dll" (ByVal pSA As Long) As Long

Private Declare PtrSafe Function SafeArrayGetLBound Lib "oleaut32.dll" _
    (ByVal pSA As Long, _
     ByVal nDim As Long, _
     ByRef plLbound As Long) _
    As Long

Private Declare PtrSafe Function SafeArrayGetUBound Lib "oleaut32.dll" _
    (ByVal pSA As Long, _
     ByVal nDim As Long, _
     ByRef plUbound As Long) _
    As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef lpDest As Any, _
     ByRef lpSource As Any, _
     ByVal lByteLen As Long)
#Else
Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef ppsaOut() As Any) As Long
Private Declare Sub CopyMemoryFromArray Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef RetPointer As Long, SrcArray() As Any, Optional ByVal Length As Long = 4&)
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" _
        (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal Msg As Long, _
        ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Const VT_BYREF = &H4000
Private Const VARIANT_DATA_OFFSET As Long = 8

Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" _
    (ByVal pSA As Long) _
    As Long

Private Declare Function SafeArrayGetLBound Lib "oleaut32.dll" _
    (ByVal pSA As Long, _
     ByVal nDim As Long, _
     ByRef plLbound As Long) _
    As Long

Private Declare Function SafeArrayGetUBound Lib "oleaut32.dll" _
    (ByVal pSA As Long, _
     ByVal nDim As Long, _
     ByRef plUbound As Long) _
    As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef lpDest As Any, _
     ByRef lpSource As Any, _
     ByVal lByteLen As Long)

#End If
'
Public Sub ResetStatusMessageAsync() '�{����Private�ɂ�����
    Application.StatusBar = False
    Application.OnTime 0, "ResetStatusMessageAsync"
End Sub
Public Sub SetStatusMessageAsync(Message As String, TimeoutSec As Integer)
        Application.StatusBar = Message
        Dim sTime As String
        sTime = Right("00" & Hour(TimeSerial(0, 0, TimeoutSec)), 2) & ":" & Right("00" & Minute(TimeSerial(0, 0, TimeoutSec)), 2) & ":" & Right("00" & Second(TimeSerial(0, 0, TimeoutSec)), 2)
        Application.OnTime Now + TimeValue(sTime), "ResetStatusMessageAsync" 'Private�ł͌Ăяo���Ȃ�
End Sub

Public Function DistinctCode(CodeList As Range) As Variant()
    Const clCapacity As Long = 32
    Dim vList As Variant, bIsAcross As Boolean, bIsAcrossFirst As Boolean, rngOneArea As Range, lOffset As Long
    '�s�D��
    If CodeList.Columns.Count = 1 Then
        bIsAcrossFirst = False
        lOffset = CodeList.Column
    ElseIf CodeList.Rows.Count = 1 Then
        bIsAcrossFirst = True
        lOffset = CodeList.Row
    Else
        Err.Raise vbObjectError + 1004, "ArgumentOutOfRangeException", "�񂩍s�̂����ꂩ��1�łȂ��ƃ_��"
    End If
    For Each rngOneArea In CodeList.Areas
        If CodeList.Columns.Count = 1 Then
            bIsAcross = False
            If bIsAcross <> bIsAcrossFirst Or lOffset <> rngOneArea.Column Then
                Err.Raise vbObjectError + 1004, "ArgumentOutOfRangeException", "���������s��������łȂ��ƃ_��"
            End If
        ElseIf rngOneArea.Rows.Count = 1 Then
            bIsAcross = True
            If bIsAcross <> bIsAcrossFirst Or lOffset <> rngOneArea.Row Then
                Err.Raise vbObjectError + 1004, "ArgumentOutOfRangeException", "���������񂪓����s�łȂ��ƃ_��"
            End If
        Else
            Err.Raise vbObjectError + 1004, "ArgumentOutOfRangeException", "�񂩍s�̂����ꂩ��1�łȂ��ƃ_��"
        End If
        
    Next
    Dim vaRC() As Variant, lMax As Long, t As Long, i As Long
    ReDim vaRC(clCapacity - 1)
    lMax = -1
    
    If Not bIsAcross Then
        For Each rngOneArea In CodeList.Areas
            If rngOneArea.Count = 1 Then
                ReDim vList(1 To 1, 1 To 1)
                vList(1, 1) = rngOneArea.Value
            Else
                vList = rngOneArea
            End If
            For t = 1 To UBound(vList)
                For i = 0 To lMax
                    If vaRC(i) = vList(t, 1) Then GoTo DistinctCode_Continue1
                Next
                lMax = lMax + 1
                If UBound(vaRC) < lMax Then
                    ReDim Preserve vaRC((UBound(vaRC) + 1) * 2 - 1)
                End If
                vaRC(lMax) = vList(t, 1)
DistinctCode_Continue1:
            Next
        Next
    Else
        For Each rngOneArea In CodeList.Areas
            If rngOneArea.Count = 1 Then
                ReDim vList(1 To 1, 1 To 1)
                vList(1, 1) = rngOneArea.Value
            Else
                vList = rngOneArea
            End If
            For t = 1 To UBound(vList, 2)
                For i = 0 To lMax
                    If vaRC(i) = vList(1, t) Then GoTo DistinctCode_Continue2
                Next
                lMax = lMax + 1
                If UBound(vaRC) < lMax Then
                    ReDim Preserve vaRC((UBound(vaRC) + 1) * 2 - 1)
                End If
                vaRC(lMax) = vList(1, t)
DistinctCode_Continue2:
            Next
        Next
    End If
    If lMax > -1 Then
        ReDim Preserve vaRC(lMax)
    Else
        vaRC = Split("", "") ' NullArrayForStringType
    End If
    DistinctCode = vaRC
End Function

Public Function LBoundEx(ByRef vArray As Variant, Optional ByVal lDimension As Long = 1) As Long

    Dim iDataType As Integer
    Dim pSA As Long

    'Make sure an array was passed in:
    If IsArray(vArray) Then

        'Try to get the pointer:
        CopyMemory pSA, ByVal VarPtr(vArray) + VARIANT_DATA_OFFSET, 4

        If pSA Then

            'If byref then deref the pointer to get the actual pointer:
            CopyMemory iDataType, vArray, 2
            If iDataType And VT_BYREF Then
                CopyMemory pSA, ByVal pSA, 4
            End If

            If pSA Then
                If lDimension > 0 Then
                    'Make sure this is a valid array dimension:
                    If lDimension <= SafeArrayGetDim(pSA) Then
                        'Get the LBound:
                        SafeArrayGetLBound pSA, lDimension, LBoundEx
                    Else
                        LBoundEx = -1
                    End If
                Else
                    Err.Raise vbObjectError Or 10000, "LBoundEx", "Invalid Dimension"
                End If
            Else
                LBoundEx = -1
            End If
        Else
            LBoundEx = -1
        End If
    Else
        Err.Raise vbObjectError Or 10000, "LBoundEx", "Not an array"
    End If

End Function


Public Function UBoundEx(ByRef vArray As Variant, _
                         Optional ByVal lDimension As Long = 1) As Long

    Dim iDataType As Integer
    Dim pSA As Long

    'Make sure an array was passed in:
    If IsArray(vArray) Then

        'Try to get the pointer:
        CopyMemory pSA, ByVal VarPtr(vArray) + VARIANT_DATA_OFFSET, 4

        If pSA Then

            'If byref then deref the pointer to get the actual pointer:
            CopyMemory iDataType, vArray, 2
            If iDataType And VT_BYREF Then
                CopyMemory pSA, ByVal pSA, 4
            End If

            If pSA Then
                If lDimension > 0 Then
                    'Make sure this is a valid array dimension:
                    If lDimension <= SafeArrayGetDim(pSA) Then
                        'Get the UBound:
                        SafeArrayGetUBound pSA, lDimension, UBoundEx
                    Else
                        UBoundEx = -1
                    End If
                Else
                    Err.Raise vbObjectError Or 10000, "UBoundEx", "Invalid Dimension"
                End If
            Else
                UBoundEx = -1
            End If
        Else
            UBoundEx = -1
        End If
    Else
        Err.Raise vbObjectError Or 10000, "UBoundEx", "Not an array"
    End If

End Function
    
'�t�H���_��K�w�𐳋K�\���Ńq�b�g�����ŐV�̃t�@�C�����擾
Public Function GetRecentFileByReg(SearchFolder As String, SearchPattern As String) As Scripting.File
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    Dim fil As Scripting.File, filRecent As Scripting.File, datRecent As Date
    Dim oReg As New VBScript_RegExp_55.RegExp
    Set filRecent = Nothing
    oReg.Pattern = SearchPattern
    For Each fil In oFS.GetFolder(SearchFolder).Files
        If oReg.test(fil.Name) Then
            If fil.DateLastModified > datRecent Then
                Set filRecent = fil
                datRecent = fil.DateLastModified
            End If
        End If
    Next
    '�q�b�g�������
    Set GetRecentFileByReg = filRecent
End Function
'TagName="�f�[�^ - �O���f�[�^�̎�荞��"��CommandBarPopup�܂���CommandBarButton���擾�ł���
Public Function FindControlByTagName(TagName As String) As Object
    '���j���[�͊K�w���E�̂���R���g���[���ł��邱��
    '�^�O���̓L���v�V�����̑O����v�ł��邱��
    Const csSeparator As String = " - "
    '�g�b�v�̃��j���[�͗�O(Type="Object/CommandBarControl/CommandBarPopup")�Ȃ̂ŒP�ƂŎ擾
    Dim vaMenuTag As Variant
    vaMenuTag = Split(TagName, csSeparator)
    Dim vCon As Variant, rootCon As CommandBarControl
    For Each vCon In CommandBars("Worksheet Menu Bar").Controls
        Set rootCon = vCon
        If InStr(1, rootCon.Caption, CStr(vaMenuTag(0))) = 1 Then
            Exit For
        End If
    Next
    If UBound(vaMenuTag) = 0 Then
        '�����Ŏ擾���I������
        Set FindControlByTagName = rootCon
        Exit Function
    Else
        'Nop
    End If
    Dim saMenuTag() As String
    ReDim saMenuTag(UBound(vaMenuTag))
    Dim conTarget As CommandBarControl, i As Long
    Set conTarget = rootCon
    Dim copRC As CommandBarPopup, cobRC As CommandBarButton, cocRC As CommandBarComboBox
    Set copRC = Nothing
    Set cobRC = Nothing
    Set cocRC = Nothing
    i = 1
    Do
        Set rootCon = conTarget
        For Each vCon In rootCon.Controls
            Set conTarget = vCon
            If InStr(1, conTarget.Caption, vaMenuTag(i)) = 1 Then
                If i = UBound(vaMenuTag) Then
                    Select Case conTarget.Type
                    Case msoControlPopup
                        Set copRC = vCon
                    Case msoControlComboBox
                        Set cocRC = vCon
                    Case msoControlButton
                        Set cobRC = vCon
                    Case Else
                        Err.Raise vbObjectError + 310, "", "�ΏۊO�̃R�}���h�o�[���"
                    End Select
                End If
                Exit For
            End If
        Next
        i = i + 1
    Loop While i <= UBound(vaMenuTag)
    'xxx--������Ȃ��Ƃ���Nothing��Ԃ� �� ���]��
    Select Case conTarget.Type
    Case msoControlPopup
        Set FindControlByTagName = copRC
    Case msoControlComboBox
        Set FindControlByTagName = cocRC
    Case msoControlButton
        Set FindControlByTagName = cobRC
    Case Else
        Err.Raise vbObjectError + 310, "", "�ΏۊO�̃R�}���h�o�[���"
    End Select
End Function

'���b�N�̂������Ă��Ȃ��Z���̂ݒl���R�s�[����
Public Sub PasteValuesOnUnlockCells(SourceRange As Range, Destination As Worksheet, UnlockDecision As DecisionDirectionFlagEnum, Optional RowOffset As Long = 0, Optional ColumnOffset As Long = 0, Optional SkipBlank As Boolean = False)
    Dim rngSrc As Range
    If (UnlockDecision And DecisionDirectionFlagEnum.SourceSide) = DecisionDirectionFlagEnum.SourceSide Then
        Set rngSrc = ReduceRangeForUnlock(SourceRange)
    Else
        Set rngSrc = SourceRange
    End If
    If rngSrc Is Nothing Then Exit Sub
    Dim rngTemp As Range, rngDest As Range
    If (UnlockDecision And DecisionDirectionFlagEnum.DestinationSide) = DecisionDirectionFlagEnum.DestinationSide Then
        Set rngTemp = ReflectRangeOverWorksheet(rngSrc, Destination)
        Set rngDest = ReduceRangeForUnlock(rngTemp)
        Set rngTemp = ReflectRangeOverWorksheet(rngDest, SourceRange.Worksheet)
        Set rngSrc = Application.Intersect(rngSrc, rngTemp)
    Else
        'Nop
    End If
    Debug.Assert rngSrc.Cells.Count <> 0
    PasteValues rngSrc, Destination, RowOffset, ColumnOffset, SkipBlank
End Sub

'�l���R�s�[(�����Z���͍���1�Z���̂ݒl���R�s�[)
'Src�͌���͈͂𐄏�
Public Sub PasteValues(src As Range, Destination As Worksheet, Optional RowOffset As Long = 0, Optional ColumnOffset As Long = 0, Optional SkipBlank As Boolean = False)
    Dim rngOneArea As Range, vWorkerValue() As Variant, lCounter As Long
    If SkipBlank Then Err.Raise vbObjectError + 300, "PasteValues()", "SkipBlank=True�͖������ł��B"
    lCounter = 0
    ReDim vWorkerValue(src.Areas.Count - 1)
    For Each rngOneArea In src.Areas
        vWorkerValue(lCounter) = rngOneArea
        lCounter = lCounter + 1
    Next
    lCounter = 0
    For Each rngOneArea In src.Areas
        Destination.Cells(RowOffset + rngOneArea.Row, ColumnOffset + rngOneArea.Column).Resize(rngOneArea.Rows.Count, rngOneArea.Columns.Count) = vWorkerValue(lCounter)
        lCounter = lCounter + 1
    Next
End Sub

'���b�N�̂������Ă��Ȃ��Z�����擾����
Public Function ReduceRangeForUnlock(src As Range) As Range
    Dim rng As Range, rngRC As Range, rngSrcOneArea As Range
    Set rngRC = Nothing
    For Each rngSrcOneArea In src.Areas
        For Each rng In rngSrcOneArea.Cells
            If Not rng.Locked Then
                If rngRC Is Nothing Then
                    Set rngRC = rng
                Else
                    Set rngRC = Application.Union(rngRC, rng)
                End If
            End If
        Next
    Next
    Set ReduceRangeForUnlock = rngRC
End Function

' �t�H���_�̎Q��
Public Function ShowFolderDialog(Optional ByVal Title As String = "�t�H���_��I�����Ă�������", Optional ByVal InitDir As String) As String


    Dim bi      As BROWSEINFO
    Dim pidl    As Long
    Dim strBuf  As String
    Dim fExists As Boolean

    ' �f�B���N�g���̑��݃`�F�b�N
    On Error Resume Next
    fExists = GetAttr(InitDir) And vbDirectory
    On Error GoTo 0

    With bi
        .hwndOwner = GetActiveWindow()
        .lpszTitle = Title
        .ulFlags = BIF_RETURNONLYFSDIRS
        If fExists Then
            .lpfn = GetAddressOf(AddressOf BrowseCallback)
            .lParam = InitDir
        End If
    End With

    pidl = SHBrowseForFolder(bi)

    If pidl Then
        strBuf = String$(MAX_PATH, 0)
        If SHGetPathFromIDList(pidl, strBuf) Then
            ShowFolderDialog = Left$(strBuf, _
                                InStr(strBuf, vbNullChar) - 1)
        End If
        CoTaskMemFree pidl
    End If

End Function

' AddressOf���Z�q�̃��b�p
Private Function GetAddressOf(ByVal lngProcAddress As Long) As Long
    GetAddressOf = lngProcAddress
End Function
' SHBrowseForFolder�̃R�[���o�b�N
Private Function BrowseCallback( _
                    ByVal hWnd As Long, ByVal uMsg As Long, _
                    ByVal lParam As Long, ByVal lpData As Long) As Long

    ' �����t�H���_�̑I��
    If uMsg = BFFM_INITIALIZED Then
        SendMessage hWnd, BFFM_SETSELECTION, 1, ByVal lpData
    End If

End Function

' �w���DLL��API�������True��Ԃ��B
Public Function IsAPIExist(DLLName As String, APIName As String) As Boolean
    Dim lModuleHandle As Long, lResult As Long
    lModuleHandle = LoadLibrary(DLLName)
    If lModuleHandle <> 0 Then
        lResult = GetProcAddress(lModuleHandle, APIName)
        IsAPIExist = (lResult <> 0)
        lResult = FreeLibrary(lModuleHandle)
    End If
End Function
'�u�b�N���J���Ă��悤�����Ă��悤���Ԃ�l�𓯂��d�l�ł���u�b�N���̖��O�͈͂�l�����ĕԂ�
'�Ԃ�l��2��,0���̃o���A���g�^(�z��)
Public Function GetValueNamedArea(SrcBookPath As String, Optional SrcName As String = "", Optional SrcSheetName As String = "")
    If SrcName <> "" Eqv SrcSheetName <> "" Then Err.Raise vbObjectError + 395, "GetValueNamedArea", "SrcName��SrcSheetName���w�肵�Ă�������"
    Dim bOpeningSrc As Boolean
    If InStr(1, SrcBookPath, "\") > 0 Then
        bOpeningSrc = Misc.BookExists(Mid(SrcBookPath, InStrRev(SrcBookPath, "\") + 1))
    Else
        bOpeningSrc = Misc.BookExists(SrcBookPath)
    End If
    Dim oFS As New Scripting.FileSystemObject
    If Not oFS.FileExists(SrcBookPath) Then
        GetValueNamedArea = Empty
        Exit Function
    End If
    If bOpeningSrc Then
        Dim rng As Range
        If SrcName <> "" Then
            Set rng = Workbooks(Mid(SrcBookPath, InStrRev(SrcBookPath, "\") + 1)).Names(SrcName).RefersToRange
    '        If rng.Rows.Count = 1 Or rng.Columns.Count = 1 Then
    '            Dim vaWorker() As Variant
    '            ReDim vaWorker(1 To rng.Rows.Count, 1 To rng.Columns.Count)
    '            Dim t As Long
    '            For i = 1 To UBound(vaWorker)
    '                For t = 1 To UBound(vaWorker, 2)
    '                    vaWorker(i, t) = rng.Cells(i, t).Value
    '                Next
    '            Next
    '            GetValueNamedArea = vaWorker
    '            Exit Function
    '        End If
            GetValueNamedArea = rng
            Exit Function
        ElseIf SrcSheetName <> "" Then
            '�V�[�g�̎w��
            Set rng = Workbooks(Mid(SrcBookPath, InStrRev(SrcBookPath, "\") + 1)).Sheets(SrcSheetName).UsedRange
            Set rng = rng.Worksheet.Range(rng.Worksheet.Cells(1, 1), rng.Resize(1, 1).Offset(rng.Rows.Count - 1, rng.Columns.Count - 1))
            GetValueNamedArea = rng
            Exit Function
        End If
    Else
        'Nop(���s�ȍ~�ɋL��)
    End If
    '�\�[�X�̕̂݁B�J�ł͖��O��������Ȃ��ꍇ������(DataBaseEngine�̑���H)Excel�ォ��̎擾�ɕύX
    Dim cnList As New ADODB.Connection, rsList As New ADODB.Recordset
    cnList.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Data Source=" & SrcBookPath & ";" & _
               "Extended Properties=""Excel 8.0;HDR=No;IMEX=1;"""
    rsList.CursorLocation = adUseClient
    Dim sAddress As String, sSheetName As String, oReg As New VBScript_RegExp_55.RegExp
    If SrcSheetName = "" And SrcName <> "" Then
        '���O�͈�
        Err.Clear
        On Error Resume Next
        rsList.Open "SELECT * FROM " & SrcName, cnList, adOpenStatic, adLockOptimistic
        On Error GoTo 0
        '�t�@�C�����ɖ��O���Ȃ��ꍇ
        If rsList.State = adStateClosed Then
            GetValueNamedArea = Empty
            cnList.Close
            Exit Function
        End If
    ElseIf SrcName <> "" Then
        '�A�h���X�w�� �A�h���X��!���t���Ȃ�V�[�g���w����D��
        sAddress = LCase(Replace(SrcName, "$", ""))
        sSheetName = SrcSheetName
        If InStr(1, sAddress, "!") > 1 Then
            sSheetName = Left(sAddress, InStr(1, sAddress, "!") - 1)
            sAddress = Mid(sAddress, InStr(1, sAddress, "!") + 1)
        Else
            
        End If
        If Not oReg.test(sAddress) Then
            Err.Raise vbObjectError + 301, "GetValueNamedArea()", "SrcName�ɖ��O�͈͂��A�h���X�̎w��͕K�{�ł�"
        End If
        If Not IsValidSheetName(sSheetName) Then
            Err.Raise vbObjectError + 301, "GetValueNamedArea()", "SrcName�ɖ��O�͈͂��A�h���X�̎w��͕K�{�ł�"
        End If
    ElseIf SrcSheetName <> "" Then
        '�V�[�g���Ńf�[�^���擾
        rsList.Open "SELECT * FROM [" & SrcSheetName & "$];", cnList, adOpenStatic, adLockReadOnly
        GetValueNamedArea = GetTraverce(rsList.GetRows())
        rsList.Close
        cnList.Close
        Exit Function
    Else
        Err.Raise vbObjectError + 299, "GetValueNamedArea()", "SrcName�ɖ��O�͈͂��A�h���X�̎w��͕K�{�ł�"
    End If
    
    If Err.Number = vbObjectError + 3639 Then
        rsList.Close
        On Error GoTo 0
        cnList.Close
        GetValueNamedArea = Empty
        GoTo NameNotExists
    ElseIf Err.Number = 0 Then
        'Nop
    'ElseIf vbObjectError + 3604 Then
        '�s�� Err.Clear����Ζ��Ȃ�?
    Else
        'Resume
        Err.Raise Err.Number, Err.Source, Err.Description
        On Error GoTo 0
    End If
    On Error GoTo 0
    Dim vaRC() As Variant, i As Long, lRowCounter As Long
    If rsList.RecordCount = 1 And rsList.Fields.Count = 1 Then
        GetValueNamedArea = rsList.Fields(0).Value
    Else
        ReDim vaRC(1 To rsList.RecordCount, 1 To rsList.Fields.Count)
        Do Until rsList.EOF
            lRowCounter = lRowCounter + 1
            For i = 1 To UBound(vaRC, 2)
                vaRC(lRowCounter, i) = rsList.Fields(i - 1).Value
            Next
            rsList.MoveNext
        Loop
        GetValueNamedArea = vaRC
    End If
    rsList.Close
NameNotExists:
    cnList.Close
    If Not bOpeningSrc And BookExists(SrcBookPath) Then
        Misc.OpenBook(SrcBookPath).Close False
    End If
End Function

'�A�h���X�Ƃ��Ďg���邩������
Public Function IsValidRangeAddress(ByVal TestAddress As String) As Boolean
    Dim oReg As VBScript_RegExp_55.RegExp, sTestAddress As String, sSheetName As String
    Set oReg = New RegExp
    If InStr(1, TestAddress, "!") > 1 Then
        sSheetName = Left(TestAddress, InStrRev(TestAddress, "!") - 1)
        sTestAddress = UCase(Mid(TestAddress, InStrRev(TestAddress, "!") + 1))
    Else
        sTestAddress = UCase(TestAddress)
    End If
    oReg.Pattern = "^(\$?[A-I]?[A-Z]\$?[1-9]\d{1,4}):?(\$?[A-I]?[A-Z]\$?[1-9]\d{1,4})?$"
    If sSheetName <> "" Then
        IsValidRangeAddress = oReg.test(sTestAddress) And IsValidSheetName(sSheetName)
    Else
        IsValidRangeAddress = oReg.test(sTestAddress)
    End If
    Set oReg = Nothing
End Function

Private Sub testIsValidRangeAddress()
    Dim vLine As Variant, bRC As Boolean
    bRC = True
    Dim i As Long
    For Each vLine In Array("AD10", "AD10w", "$AD1")
        bRC = bRC And (Array(True, False, True)(i) = IsValidRangeAddress(CStr(vLine)))
        If Not bRC Then MsgBox CStr(vLine) & "(" & i + 1 & "���)"
        i = i + 1
    Next
End Sub
'�V�[�g���Ƃ��Ďg���邩��Ԃ�
Public Function IsValidSheetName(ByVal TestSheetName As String) As Boolean
    Dim oReg As VBScript_RegExp_55.RegExp
    Set oReg = New RegExp
    oReg.Pattern = "[\:\\\/\?\*\[\]]"
    IsValidSheetName = Not oReg.test(TestSheetName)
    Set oReg = Nothing
End Function

'�Ώۃt�@�C�������ŊY���p�X��S�Ď擾
Private Sub RecurseFilesFinderWithRegExp(TargetFolder As Scripting.Folder, RegPattern As String, ByRef AnswerValues() As String, ByRef Index As Long)
    Dim fFinder As Scripting.File
    Dim lIndex As Long
    Dim oReg As New RegExp
    oReg.Pattern = RegPattern
    For Each fFinder In TargetFolder.Files
        If oReg.test(fFinder.Name) Then
            AnswerValues(Index) = fFinder.Path
            Index = Index + 1
        End If
        '���~�b�g��2�{������
        If Index > UBound(AnswerValues) Then
            ReDim Preserve AnswerValues((UBound(AnswerValues) + 1) * 2 - 1)
        End If
    Next
    '�e���q�ւ̍ċA
    Dim folFinder As Scripting.Folder
    For Each folFinder In TargetFolder.SubFolders
        RecurseFilesFinderWithRegExp folFinder, RegPattern, AnswerValues, Index
    Next
End Sub
#If Win32 Then
Private Function NullArrayForStringType() As String()
    Dim saZero() As String, sZero As String '�N���C�A���g�R�[�h���ŃG���[�������o����UBound�֐����g����悤�ɖ���`�ł͂Ȃ���̔z��(String�^)��Ԃ��B
    Dim pSafeArray As Long
    SafeArrayAllocDescriptor 1, saZero
    CopyMemoryFromArray pSafeArray, saZero
    CopyMemory ByVal pSafeArray + 4&, LenB(sZero), 4&
    NullArrayForStringType = saZero
End Function
#End If
'�ߒl �Y���Ȃ�=������^��z��, �Y������=������z��(�ǂ�����R���N�V�����ƌ��􂹂�)
'v0.01
Public Function FilesFinderWithRegExp(RootFolder As String, RegPattern As String) As String()
    Const clCapacity As Long = 16
    Dim AnswerValues() As String, oFS As New Scripting.FileSystemObject, oFolder As Scripting.Folder, Index As Long
    Index = 0
    ReDim AnswerValues(clCapacity - 1)
    Set oFolder = oFS.GetFolder(RootFolder)
    RecurseFilesFinderWithRegExp oFolder, RegPattern, AnswerValues, Index
    If Index > 0 Then
        ReDim Preserve AnswerValues(Index - 1)
        FilesFinderWithRegExp = AnswerValues
    Else
        Dim saZeroMass() As String
        'saZeroMass = Split("", "") 'NullArrayForStringType
        FilesFinderWithRegExp = Split("", "")
    End If
End Function

'0.01: �X�g�A�pExcel�͈͋󔒂ōĐݒ�
Public Sub SetPath(RegKey As String, InitialFile As String, SetRange As Range, Optional DialogDescription As String = "�t�@�C���܂��̓t�H���_���w�肵�Ă��������B", Optional StoredPlace As String = "(registry)")
'1:SetRange�ɏ����ꂽ�p�X�ɊY�����Ȃ���΃p�X��V���ɐݒ肷��
'2:�X�g�A��xml�`���ǉ�
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    If IsEmpty(SetRange.Value) Then
        If ("(local)" = LCase(StoredPlace)) Or ("(registry)" = LCase(StoredPlace)) Then
            Dim WshShell As WshShell
            Set WshShell = New IWshRuntimeLibrary.WshShell
            On Error Resume Next
            WshShell.RegDelete RegKey ' ���݂��Ȃ��L�[��RegDelete����ƃG���[
            On Error GoTo 0
        ElseIf oFS.FileExists(StoredPlace) Then
            DeleteSettingValueFromConfigFile StoredPlace, RegKey
        Else
            'Nop
            'Err.Raise vbObjectError + 70, "�t�@�C�������݂��Ă��܂���"
        End If
    End If
    If oFS.FileExists(SetRange.Value) Then Exit Sub '���W�X�g���̐ݒ�͕K�v���H
    Dim sDataPath As String
    sDataPath = GetPath(RegKey, InitialFile, DialogDescription, StoredPlace)
    If SetRange.Value <> sDataPath Then
        SetRange.Value = sDataPath
    End If
End Sub

Public Function GetPath(RegKey As String, InitialFile As String, Optional DialogDescription As String = "�t�@�C���܂��̓t�H���_���w�肵�Ă��������B", Optional StoredPlace As String = "(registry)") As String
    '�܂��̓��W�X�g������ǂޓǂ߂�ΏI��
    Dim WshShell As IWshRuntimeLibrary.WshShell
    Set WshShell = New IWshRuntimeLibrary.WshShell
    Dim sFilePath As String, bNeedWriteReg As Boolean, bIsUseReg As Boolean
    bIsUseReg = ("(local)" = LCase(StoredPlace)) Or ("(registry)" = LCase(StoredPlace))
    If bIsUseReg Then
        sFilePath = GetPathWithRegistry(RegKey)
    Else
        sFilePath = GetSettingValueFromConfigFile(StoredPlace, RegKey, ForPath)
        If sFilePath = "None" Then
            sFilePath = ""
        ElseIf sFilePath = "False" Then
            sFilePath = ""
        Else
            'Nop
        End If
    End If
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    If sFilePath <> "" Then
        If oFS.FileExists(sFilePath) Then
            GetPath = sFilePath
            Exit Function
        ElseIf oFS.FolderExists(sFilePath) Then
            GetPath = sFilePath
            Exit Function
        End If
    Else
        bNeedWriteReg = True
    End If
    '�J���_�C�A���O�œǂޓǂ߂�΃X�g�A�ɏ������ݏI��
    Dim sExt As String
    sExt = oFS.GetExtensionName(InitialFile)
    Dim bNeedGavageInitialFile As Boolean
    Dim sInitialFile As String
    If oFS.FolderExists(oFS.GetAbsolutePathName(oFS.GetParentFolderName(InitialFile))) Then
        sInitialFile = oFS.GetAbsolutePathName(InitialFile)
    ElseIf oFS.FolderExists(oFS.GetParentFolderName(InitialFile)) Then '�E�C���X�o�X�^�[�̕s��΍�
        sInitialFile = InitialFile
    Else
        sInitialFile = oFS.BuildPath(WshShell.Environment("Process")("Temp"), oFS.GetFileName(InitialFile))
    End If
    bNeedGavageInitialFile = Not oFS.FileExists(sInitialFile) And (sExt <> "")
    If bNeedGavageInitialFile Then
        oFS.CreateTextFile(sInitialFile, False, False).Write ""
    End If
    Dim sFileType As String
    If sExt <> "" Then
        sFileType = oFS.GetFile(sInitialFile).Type
        If bNeedGavageInitialFile Then
            oFS.DeleteFile sInitialFile, False
        End If
        Dim vFilePath As Variant
        
        vFilePath = GetOpenFilenameOnInitialDir(sFileType & " (*." & sExt & "),*." & sExt & ",���ׂẴt�@�C��(*.*),*.*", 1, DialogDescription, "�J��", False, Left(sInitialFile, InStrRev(sInitialFile, "\")), True)
        If VarType(vFilePath) = vbBoolean Then
        ElseIf VarType(vFilePath) = vbString Then
            If bNeedWriteReg Then
                If bIsUseReg Then
                    WshShell.RegWrite RegKey, vFilePath
                Else
                    SetPathWithConfigFile StoredPlace, RegKey, CStr(vFilePath)
                End If
            End If
            GetPath = vFilePath
            Exit Function
        Else
            '��O����
            Err.Raise vbObjectError + 291
        End If
    Else
        Dim sFolder As String
        sFolder = ShowFolderDialog(DialogDescription, sInitialFile)
        If sFolder = "" Then Exit Function
        GetPath = sFolder
            If bNeedWriteReg Then
                If bIsUseReg Then
                    WshShell.RegWrite RegKey, sFolder
                Else
                    SetPathWithConfigFile StoredPlace, RegKey, CStr(sFolder)
                End If
            Else
                SetPathWithConfigFile StoredPlace, RegKey, CStr(sFolder)
                
            End If
        
        Exit Function
    End If
    '�󕶎���Ԃ�
    GetPath = ""
End Function

Public Function SetConfig(StoredPlace As String, RegKey As String, Value As Variant, Purpose As SettingValueTypeEnum) As Boolean
    Select Case Purpose
    Case SettingValueTypeEnum.ForNumber
        SetConfig = SetConfigFile(StoredPlace, RegKey, Value, "ForNumber") '����
    Case SettingValueTypeEnum.ForNumeric
        SetConfig = SetConfigFile(StoredPlace, RegKey, Value, "ForNumeric") '����
    Case SettingValueTypeEnum.ForPath
        SetConfig = SetConfigFile(StoredPlace, RegKey, Value, "ForPath") '�p�X������
    Case SettingValueTypeEnum.ForString
        SetConfig = SetConfigFile(StoredPlace, RegKey, Value, "ForString") '������
    Case Else
        Err.Raise vbObjectError + 3005, Err.Source, "Type=" & Purpose & "�͖������ł��B"
    End Select
End Function

Private Function SetConfigFile(StoredPlace As String, RegKey As String, Value As Variant, Purpose As String) As Boolean
    'See Setting\SettingMemoryObject.cls
    Dim elm As MSXML2.IXMLDOMElement
    Dim xmlStore As MSXML2.DOMDocument
    Set xmlStore = New MSXML2.DOMDocument
    xmlStore.async = False
    Dim sTargetXPath As String
    Dim bIsDirty As Boolean
    Dim sRegKey As String
    If InStr(1, RegKey, "\\") > 0 Then
        '\\������Ή������Ȃ�
        sRegKey = RegKey
    ElseIf InStr(1, RegKey, "\") = 0 Then
        '\���Ȃ���Ή������Ȃ�
        sRegKey = RegKey
    Else
        sRegKey = Replace(RegKey, "\", "\\")
    End If
    'ConfigFile���Ȃ����͍쐬
    If Not xmlStore.Load(StoredPlace) Then
        xmlStore.LoadXML "<?xml version=""1.0"" encoding=""Shift_JIS"" standalone=""yes""?>" & vbNewLine & _
                    "<root>" & vbNewLine & _
                    "   <Settings Original=""true"">" & vbNewLine & _
                    "       <Setting Name="""" Value="""" Purpose="""" Required=""false"" Original=""true""/>" & vbNewLine & _
                    "   </Settings>" & vbNewLine & _
                    "</root>" & vbNewLine
        bIsDirty = True
    End If
    'ComputerName�ɊY�����Ȃ�����original node���R�s�[���A�Ō�ɑ}��
    sTargetXPath = "/root"
    Dim sSelfComputerName As String
    sSelfComputerName = CreateObject("Wscript.Network").ComputerName
    Set elm = xmlStore.selectSingleNode(sTargetXPath & "/Settings[@ComputerName='" & sSelfComputerName & "']")
    Dim elmOrg As IXMLDOMElement
    If elm Is Nothing Then
        Set elmOrg = xmlStore.selectSingleNode(sTargetXPath & "/Settings[@Original='true']")
        xmlStore.selectSingleNode(sTargetXPath).appendChild xmlStore.createTextNode(vbCrLf)
        Set elm = xmlStore.selectSingleNode(sTargetXPath).appendChild(elmOrg.CloneNode(True))
        elm.setAttribute "ComputerName", sSelfComputerName
        elm.removeAttribute "Original"
        bIsDirty = True
    End If
    '���O�ɊY�����Ȃ���Βǉ�
    sTargetXPath = sTargetXPath & "/Settings[@ComputerName='" & sSelfComputerName & "']"
    Set elm = xmlStore.selectSingleNode(sTargetXPath & "/Setting[@Name='" & sRegKey & "']")
    If elm Is Nothing Then
        Set elmOrg = xmlStore.selectSingleNode(sTargetXPath & "/Setting[@Original='true']")
        xmlStore.selectSingleNode(sTargetXPath).appendChild xmlStore.createTextNode(vbCrLf)
        Set elm = xmlStore.selectSingleNode(sTargetXPath).appendChild(elmOrg.CloneNode(True))
        elm.setAttribute "Name", RegKey
        elm.setAttribute "Purpose", Purpose
        elm.setAttribute "Required", "true"
        elm.removeAttribute "Original"
        bIsDirty = True
    End If
    '�Y���̃m�[�h�ƒl���Ⴄ���͍X�V
    sTargetXPath = sTargetXPath & "/Setting[@Name='" & sRegKey & "']"
    Set elm = xmlStore.selectSingleNode(sTargetXPath)
    If elm.getAttribute("Value") <> Value Then
        elm.setAttribute "Value", CStr(Value)
        bIsDirty = True
    End If
    '�ύX������Ƃ��͕ۑ�
    If bIsDirty Then
        Dim oFS As Scripting.FileSystemObject
        Set oFS = New Scripting.FileSystemObject
        If Not oFS.FolderExists(oFS.GetParentFolderName(StoredPlace)) Then
            CreateDeepFolder oFS.GetParentFolderName(StoredPlace)
        End If
        xmlStore.Save StoredPlace
    End If
    SetConfigFile = bIsDirty
End Function
Private Sub SetValueWithConfigFile(StoredPlace As String, RegKey As String, Path As String)
    '�t�@�C�����V���N���i�C�Y���Ă������Ƃɒl��ݒ�ł���R���t�B�O�t�@�C��
    SetConfigFile StoredPlace, RegKey, Path, "ForString"
End Sub
Private Sub SetPathWithConfigFile(StoredPlace As String, RegKey As String, Path As String)
    '�t�@�C�����V���N���i�C�Y���Ă������ƂɃp�X��ݒ�ł���R���t�B�O�t�@�C��
    SetConfigFile StoredPlace, RegKey, Path, "ForPath"
End Sub
Private Sub SetNumberWithConfigFile(StoredPlace As String, RegKey As String, Value As Long)
    '�t�@�C�����V���N���i�C�Y���Ă������Ƃɐ�����ݒ�ł���R���t�B�O�t�@�C��
    SetConfigFile StoredPlace, RegKey, CStr(Value), "ForNumber"
End Sub
Public Sub SetSettingValue(StoredPlace As String, RegKey As String, Purpose As SettingValueTypeEnum, Value As Variant)
    Dim bIsUseReg As Boolean
    bIsUseReg = ("(local)" = LCase(StoredPlace)) Or ("(registry)" = LCase(StoredPlace))
    If bIsUseReg Then
        Dim WshShell As New WshShell
        Select Case Purpose
        Case SettingValueTypeEnum.ForPath, SettingValueTypeEnum.ForString
            WshShell.RegWrite RegKey, CStr(Value), "REG_SZ"
        Case SettingValueTypeEnum.ForNumber
            WshShell.RegWrite RegKey, CLng(Value), "REG_DWORD"
        Case SettingValueTypeEnum.ForNumeric
            WshShell.RegWrite RegKey, CStr(Value), "REG_SZ"
        Case Else
            Err.Raise vbObjectError + 333, "SetSettingValueFromRegistry", "PurposeType='" & Purpose & "'�͖������ł�"
         End Select
    Else
        SetConfig StoredPlace, RegKey, Value, Purpose
    End If
End Sub
Public Function GetSettingValue(StoredPlace As String, RegKey As String, Purpose As SettingValueTypeEnum) As Variant
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    If oFS.FileExists(StoredPlace) Then
        GetSettingValue = GetSettingValueFromConfigFile(StoredPlace, RegKey, Purpose)
    ElseIf StoredPlace = "(local)" Or StoredPlace = "(registry)" Then
        GetSettingValue = GetSettingValueFromRegistry(RegKey, Purpose)
    Else
        Err.Raise 53, "GetSettingValue()", "�t�@�C����������܂���B"
    End If
    If "False" = GetSettingValue Then
        '�ʏ�̓G���[��Ԃ�����
        Err.Raise 1004, "GetSettingValue()", "�G���g��" & RegKey & "������܂���B"
    End If
End Function

Private Function GetSettingValueFromRegistry(RegKey As String, Purpose As SettingValueTypeEnum) As Variant
    Dim WshShell As New WshShell
    Select Case Purpose
    Case SettingValueTypeEnum.ForPath, SettingValueTypeEnum.ForString
        GetSettingValueFromRegistry = CStr(WshShell.RegRead(RegKey))
    Case SettingValueTypeEnum.ForNumber
        GetSettingValueFromRegistry = CLng(WshShell.RegRead(RegKey))
    Case SettingValueTypeEnum.ForNumeric
        GetSettingValueFromRegistry = CDbl(WshShell.RegRead(RegKey))
    Case Else
        Err.Raise vbObjectError + 333, "GetSettingValueFromRegistry", "PurposeType='" & Purpose & "'�͖������ł�"
     End Select
End Function

Private Function GetSettingValueFromConfigFile(StoredPlace As String, RegKey As String, Purpose As SettingValueTypeEnum) As Variant
    Dim elm As MSXML2.IXMLDOMElement
    Dim xmlStore As MSXML2.DOMDocument
    Set xmlStore = New MSXML2.DOMDocument
    xmlStore.async = False
    Dim sRegKey As String
    sRegKey = Replace(RegKey, "\", "\\")
    Dim sTargetXPath As String
    'ConfigFile���Ȃ�����"None"��Ԃ�
    If Not xmlStore.Load(StoredPlace) Then
        GetSettingValueFromConfigFile = "None"
        Exit Function
    End If
    '�Y�����Ȃ��ꍇ��"False"��Ԃ�
    Set elm = xmlStore.selectSingleNode("/root/Settings[@ComputerName='" & CreateObject("WScript.Network").ComputerName & "']/Setting[@Name='" & sRegKey & "']")
    If elm Is Nothing Then
        GetSettingValueFromConfigFile = "False"
        Exit Function
    End If
    Select Case elm.getAttribute("Purpose")
    Case "ForPath", "ForString"
        GetSettingValueFromConfigFile = CStr(elm.getAttribute("Value"))
    Case "ForNumber"
        GetSettingValueFromConfigFile = CLng(elm.getAttribute("Value"))
    Case "ForNumeric"
        GetSettingValueFromConfigFile = CDbl(elm.getAttribute("Value"))
    Case Else
        Err.Raise vbObjectError + 333, "GetSettingValueFromConfigFile", "PurposeType='" & Purpose & "'�͖������ł�"
     End Select
End Function
Private Sub DeleteSettingValueFromConfigFile(StoredPlace As String, RegKey As String)
    Dim nod As MSXML2.IXMLDOMNodeList, elm2 As MSXML2.IXMLDOMElement
    Dim xmlStore As MSXML2.DOMDocument
    Set xmlStore = New MSXML2.DOMDocument
    xmlStore.async = False
    Dim sRegKey As String
    sRegKey = Replace(RegKey, "\", "\\")
    Dim sTargetXPath As String
    If Not xmlStore.Load(StoredPlace) Then
        Exit Sub
    End If
    Set nod = xmlStore.SelectNodes("/root/Settings[@ComputerName='" & CreateObject("WScript.Network").ComputerName & "']/Setting[@Name='" & sRegKey & "']")
    If nod.Length = 0 Then Exit Sub
    For Each elm2 In nod
        elm2.ParentNode.RemoveChild elm2
    Next
    xmlStore.Save StoredPlace
End Sub

Private Function GetPathWithRegistry(RegistryKey) As String
    Dim WshShell, sValue, lErr, sErr
    sValue = ""
    Set WshShell = CreateObject("WScript.Shell")
    On Error Resume Next
    sValue = WshShell.RegRead(RegistryKey)
    lErr = Err.Number
    sErr = Err.Description
    On Error GoTo 0
    If lErr = -2147024894 Then
        '�L�[�����݂��Ȃ�
        GetPathWithRegistry = ""
    ElseIf lErr = 0 Then
        '�L�[�����݂��� Nop
        GetPathWithRegistry = sValue
    Else
        '��O����
        MsgBox lErr & ":" & sErr, , "ChkVer()"
    End If
End Function

Public Function GetSaveAsFilenameOnInitialDir(InitialFilename, FileFilter, FilterIndex, Title, ButtonText, InitialDir As String, KeepCurrentDir As Boolean) As Variant
    Dim sCurDir As String
    sCurDir = GetCurrentDirectoryEx
    SetCurrentDirectory InitialDir
    GetSaveAsFilenameOnInitialDir = Application.GetSaveAsFilename(InitialFilename, FileFilter, FilterIndex, Title, ButtonText)
    If KeepCurrentDir Or sCurDir <> InitialDir Then
        SetCurrentDirectory sCurDir
    End If
End Function

'OneDrive��path�iex: https://d.docs.live.net/892147db5037df74/�h�L�������g/onefile.doc�j��^����ƃ��[�J���p�X��Ԃ�
Public Function GetOneDrivePath(Path As String) As String
     
    Dim Tgt, sPath As String, i, cnt As Long
     
    Tgt = Path
     
    'URL�̕������폜���ăt���p�X�����܂�
    If Left(Tgt, 5) = "https" Then
        For i = 1 To Len(Tgt)
            If Mid(Tgt, i, 1) = "/" Then cnt = cnt + 1
        Next i
        If cnt = 3 Then
            sPath = Environ("UserProfile") & "\OneDrive"
        Else
            cnt = 0
            For i = 1 To Len(Tgt)
                If Mid(Tgt, 1, 1) = "/" Then cnt = cnt + 1
                If cnt < 4 Then
                    Tgt = Right(Tgt, Len(Tgt) - 1)
                Else
                    Exit For
                End If
            Next i
            sPath = Replace(Environ("UserProfile") & "\OneDrive" & Tgt, "/", "\")
        End If
    End If
    GetOneDrivePath = sPath
End Function

Public Function GetOpenFilenameOnInitialDir(FileFilter, FilterIndex, Title, ButtonText, MultiSelect, InitialDir As String, KeepCurrentDir As Boolean) As Variant
    Dim sCurDir As String
    sCurDir = GetCurrentDirectoryEx
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    Dim sInitialDir As String
    If oFS.FolderExists(InitialDir) Then
        sInitialDir = InitialDir
    ElseIf oFS.FolderExists(oFS.GetParentFolderName(InitialDir)) Then
        sInitialDir = oFS.GetParentFolderName(InitialDir)
    Else
        sInitialDir = InitialDir
    End If
    SetCurrentDirectory InitialDir
    GetOpenFilenameOnInitialDir = Application.GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
    If KeepCurrentDir And sCurDir <> sInitialDir Then
        SetCurrentDirectory sCurDir
    End If
End Function
Private Function GetCurrentDirectoryEx() As String
    'API�֐�GetCurrentDirectory�̖߂�l���擾����֐�ANSI��
    Dim lRet As Long
    '�J�����g�f�B���N�g����(��Q����)���i�[����ϐ�("* 1024"�͕�����̃T�C�Y�ł��B)
    Dim sCurrentDir As String * 1024
    lRet = GetCurrentDirectory(Len(sCurrentDir), sCurrentDir)
    If lRet <> 0 Then
         '�߂�l��"0"�ȊO��������
        '�ϐ�sCurrentDir���̍ŏ���vbNullChar�̈ʒu�����߁ALeft�֐����g�p���ăJ�����g�f�B���N�g���������o���܂��B
        GetCurrentDirectoryEx = Left(sCurrentDir, InStr(sCurrentDir, vbNullChar) - 1)
    Else
        '�߂�l��"0"��������(=�G���[������������)
        GetCurrentDirectoryEx = ""
    End If
End Function
Public Function GetSingleBookAtProjectName(TargetProjectName As String) As Workbook
    Dim sTargetProjectName As String
    sTargetProjectName = TargetProjectName
    Dim wbk As Workbook
    For Each wbk In Workbooks
        If wbk.VBProject.Name = sTargetProjectName Then
            Set GetSingleBookAtProjectName = wbk
            Exit Function
        End If
    Next
End Function

Public Function FindAll(What As Variant, WhereFrom As Range, LookIn As XlFindLookIn, LookAt As XlLookAt, MatchCase As Boolean) As Range
    Dim rc As Range, rngWorker As Range
    If WhereFrom.Areas.Count > 1 Then
        Dim PartsOfWhereFrom As Range
        For Each PartsOfWhereFrom In WhereFrom.Areas
            If rc Is Nothing Then
                Set rc = FindOneArea(What, PartsOfWhereFrom, LookIn, LookAt, MatchCase)
            Else
                Set rngWorker = FindOneArea(What, PartsOfWhereFrom, LookIn, LookAt, MatchCase)
                If rngWorker Is Nothing Then
                    'Nop
                Else
                    Set rc = Application.Union(rc, rngWorker)
                End If
            End If
        Next
    Else
        Set rc = FindOneArea(What, WhereFrom, LookIn, LookAt, MatchCase)
    End If
    Set FindAll = rc
End Function
Private Function FindOneArea(What As Variant, WhereFrom As Range, LookIn As XlFindLookIn, LookAt As XlLookAt, MatchCase As Boolean) As Range
    Dim rngWhole As Range
    Set rngWhole = WhereFrom.Find(What, LookIn:=LookIn, LookAt:=LookAt, MatchCase:=MatchCase)
    If Not (rngWhole Is Nothing) Then
        Dim firstAddress As String
        firstAddress = rngWhole.Address
        Dim rngNext As Range
        Set rngNext = rngWhole
        Do
            Set rngNext = WhereFrom.FindNext(rngNext)
            If rngNext Is Nothing Or rngNext.Address = firstAddress Then Exit Do
            Set rngWhole = Application.Union(rngWhole, rngNext)
        Loop While True
    End If
    Set FindOneArea = rngWhole
End Function
Public Function GetSheetByCodeName(book As Workbook, CodeName As String) As Worksheet
    Dim sh As Worksheet
    
    If book Is Nothing Then Set book = ThisWorkbook
    For Each sh In book.Sheets
        If sh.CodeName = CodeName Then Set GetSheetByCodeName = sh: Exit Function
    Next
End Function
Public Function SuspectPath(FileName, WhereCase) As String
    '�ߗׂ�path�𒲂ׁAFilename������Path���擾����
    Dim oFS As New Scripting.FileSystemObject
    Dim sBasePath As String
    Dim oFolder As Scripting.Folder
    Dim sFileName As String
    Dim WshShell As New IWshRuntimeLibrary.WshShell
    
    sFileName = FileName
    If InStr(sFileName, "\") Then
        sFileName = Mid(sFileName, InStrRev(sFileName, "\") + 1)
    End If
    Select Case LCase(WhereCase)
    Case "current", "currentdir", "currentdirectory", "1", "."
        sBasePath = WshShell.CurrentDirectory & "\"
    Case Else
        sBasePath = ThisWorkbook.Path
        If oFS.FolderExists(WhereCase) Then sBasePath = WhereCase
    End Select
    If oFS.FileExists(oFS.BuildPath(sBasePath, sFileName)) Then
        SuspectPath = oFS.BuildPath(sBasePath, sFileName)
        Exit Function
    End If
    For Each oFolder In oFS.GetFolder(sBasePath).SubFolders
        If oFS.FileExists(oFolder.Path & "\" & sFileName) Then
            SuspectPath = oFolder.Path & "\" & sFileName
            Exit Function
        End If
    Next
    sBasePath = oFS.GetParentFolderName(sBasePath) & "\"
    If oFS.FileExists(oFS.BuildPath(sBasePath, sFileName)) Then
        SuspectPath = oFS.BuildPath(sBasePath, sFileName)
        Exit Function
    End If
    For Each oFolder In oFS.GetFolder(sBasePath).SubFolders
        If oFS.FileExists(oFS.BuildPath(oFolder.Path, sFileName)) Then
            SuspectPath = oFS.BuildPath(oFolder.Path, sFileName)
            Exit Function
        End If
    Next
    sBasePath = oFS.GetParentFolderName(sBasePath) & "\"
    If oFS.FileExists(oFS.BuildPath(sBasePath, sFileName)) Then
        SuspectPath = oFS.BuildPath(sBasePath, sFileName)
        Exit Function
    End If
    For Each oFolder In oFS.GetFolder(sBasePath).SubFolders
        If oFS.FileExists(oFS.BuildPath(oFolder.Path, sFileName)) Then
            SuspectPath = oFS.BuildPath(oFolder.Path, sFileName)
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 53, "SuspectPath", "�\�[�X�t�@�C����������܂���B"
End Function
Public Function TableExists(ConnectionDatabase, sTableName) As Boolean
    Dim axc As Variant 'ADOX.Catalog
    Dim tbl As Variant 'ADOX.Table
    Dim bFind As Boolean
    
    bFind = False
    Set axc = CreateObject("ADOX.Catalog")
    axc.ActiveConnection = ConnectionDatabase.ConnectionString
    For Each tbl In axc.Tables
        bFind = bFind Or tbl.Name = sTableName
        If bFind Then Exit For
    Next
    Set tbl = Nothing
    Set axc = Nothing
    TableExists = bFind
End Function

Public Sub CloseThisWorkbook(Optional SaveChanges As Tristate = TristateFalse)
    Dim bDA As Boolean
    'Stop
    bDA = Application.DisplayAlerts
    If Application.Workbooks.Count = 1 Then
        '�u�b�N��Ȃ�DisplayAlerts�������Ă�OK
        Application.DisplayAlerts = (SaveChanges = TristateMixed)
        'Debug.Print "Saved=" & ThisWorkbook.Saved & ", Alert=" & Application.DisplayAlerts
        'Stop
        If Not ThisWorkbook.Saved And SaveChanges = TristateTrue Then
            Application.DisplayAlerts = False
            ThisWorkbook.Save
        End If
        Application.Quit
        Exit Sub
    End If
    'Close������}�N���I�� Application�I�u�W�F�N�g�͌p��
    'Debug.Print "Saved=" & ThisWorkbook.Saved & ", Alert=" & Application.DisplayAlerts
    'Stop
    If ThisWorkbook.Saved Then
        Application.DisplayAlerts = bDA
        ThisWorkbook.Close False
    ElseIf SaveChanges = TristateMixed Then
        Application.DisplayAlerts = True
        ThisWorkbook.Close 'DisplayAlerts�������Ă���
    ElseIf SaveChanges = TristateTrue Then
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        Application.DisplayAlerts = bDA
        ThisWorkbook.Close False
    Else
        Application.DisplayAlerts = bDA
        ThisWorkbook.Close False
    End If
End Sub

Public Function GetSheetName(BookPath As String, Optional VisibleOnly As Boolean = True) As String()
    Dim rc() As String, sBookName As String, i As Long
    sBookName = Mid(BookPath, InStrRev(BookPath, "\") + 1)
    If BookExists(BookPath, False) Then
        '�u�b�N���J����Ă���
        ReDim rc(Workbooks(sBookName).Sheets.Count - 1)
        Dim vLine As Variant
        i = 0
        For Each vLine In Workbooks(sBookName).Sheets
            If VisibleOnly And vLine.Visible = XlSheetVisibility.xlSheetVisible Then
                rc(i) = vLine.Name
                i = i + 1
            Else
            
            End If
        Next
    Else
        '�u�b�N���J����Ă��Ȃ�(��\���V�[�g�͎擾�ł��Ȃ�)
        If Not VisibleOnly Then
            Err.Raise vbObjectError + 397, "GetSheetName()", "�J���Ă��Ȃ��u�b�N�Ŕ�\���V�[�g���Ώۂɂ���͖̂������ł��B"
        End If
        Dim cn As ADODB.Connection
        Set cn = New ADODB.Connection
        If LCase(Mid(sBookName, InStrRev(sBookName, ".") + 1)) = "xlsx" Then
            cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & BookPath & ";Extended Properties=""Excel 12.0"";"
        ElseIf LCase(Mid(sBookName, InStrRev(sBookName, ".") + 1)) = "xls" Then
            cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & BookPath & ";Extended Properties=""Excel 8.0"";"
        ElseIf LCase(Mid(sBookName, InStrRev(sBookName, ".") + 1)) = "mdb" Then
            cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & BookPath & ";"
        Else
            Err.Raise vbObjectError + 396, "GetSheetName()", "�g���q��xls�܂���xlsx�ł���K�v������܂��B"
        End If
        ReDim rc(256 - 1)
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Set rs = cn.OpenSchema(20) 'adSchemaTables=20
        Dim sLastLetter As String, sPrefix As String
        i = 0
        Do While Not rs.EOF
            sLastLetter = Right$(rs.Fields("TABLE_NAME").Value, 1)
            If LCase(Mid(sBookName, InStrRev(sBookName, ".") + 1)) = "xls" Or LCase(Mid(sBookName, InStrRev(sBookName, ".") + 1)) = "xlsx" Then
                If sLastLetter = "$" Then
                    rc(i) = Left(rs.Fields("TABLE_NAME").Value, Len(rs.Fields("TABLE_NAME").Value) - 1)
                    i = i + 1
                    If UBound(rc) < i Then ReDim Preserve rc((UBound(rc) + 1) * 2 - 1)
                ElseIf sLastLetter = "'" Then
                    rc(i) = Mid(rs.Fields("TABLE_NAME").Value, 2, Len(rs.Fields("TABLE_NAME").Value) - 3)
                    i = i + 1
                    If UBound(rc) < i Then ReDim Preserve rc((UBound(rc) + 1) * 2 - 1)
                Else
                    'Nop
                End If
            ElseIf LCase(Mid(sBookName, InStrRev(sBookName, ".") + 1)) = "mdb" Then
                sPrefix = Left(rs.Fields("TABLE_NAME").Value, 4)
                If sPrefix = "MSys" Then
                    'Nop
                Else
                    rc(i) = rs.Fields("TABLE_NAME").Value
                    i = i + 1
                    If UBound(rc) < i Then ReDim Preserve rc((UBound(rc) + 1) * 2 - 1)
                End If
            Else
                'Nop
            End If
            rs.MoveNext
        Loop
        rs.Close
        cn.Close
        ReDim Preserve rc(i - 1)
    End If
    GetSheetName = rc
End Function
'�J���Ă���u�b�N�̖��O�����݂��邩��Ԃ�
Public Function BookExists(BookName As String, Optional IsLoosePath As Boolean = True) As Boolean
    Dim oBook As Workbook
    Dim sBookName As String
    Dim bHasPath As Boolean
    sBookName = BookName
    If InStr(1, sBookName, "\") > 0 Then
        sBookName = Mid(sBookName, InStrRev(sBookName, "\") + 1)
        bHasPath = True
    End If
    For Each oBook In Application.Workbooks
        If LCase(oBook.Name) = LCase(sBookName) Then
            '�u�b�N���̓��j�[�N�ł��邱�ƂɈˑ�
            If bHasPath And Not IsLoosePath Then
                BookExists = LCase(oBook.FullName) = LCase(BookName)
            Else
                BookExists = True
            End If
            Set oBook = Nothing
            Exit Function
        End If
    Next
    BookExists = False
End Function
Public Function NameExists(Name As String, Optional TargetBook As Workbook = Nothing, Optional AvailedRefference As Boolean = False) As Boolean
    '#REF!���܂܂�Ă����False��Ԃ�
    NameExists = False
    Dim oName As Name
    Dim bIsAvailedRefference As Boolean
    bIsAvailedRefference = AvailedRefference
    Dim bokTarget As Workbook
    If Not (TargetBook Is Nothing) Then Set bokTarget = TargetBook Else Set bokTarget = ThisWorkbook
    For Each oName In bokTarget.Names
        'Debug.Print oName.Name
        If Replace(LCase(oName.Name), "'", "") = Replace(LCase(Name), "'", "") Then
            If Not (bIsAvailedRefference And (InStr(oName.RefersTo, "#REF!") <> 0)) Then
                NameExists = True
                Set oName = Nothing
                Exit Function
            Else
                Set oName = Nothing
                Exit Function
            End If
        End If
    Next
End Function

Public Function OpenBook(SrcPath As String) As Workbook
    'Open���\�b�h�Ăяo�����o���Ȃ��̂ŕۑ����C�x���g(BeforeSave)����Ă΂Ȃ�����
    Dim bok As Workbook
    Dim oFS As Scripting.FileSystemObject
    Dim bokTarget As Workbook
    
    Set bokTarget = Nothing
    Set oFS = New Scripting.FileSystemObject
    For Each bok In Application.Workbooks
        If (LCase(oFS.GetAbsolutePathName(bok.FullName)) = LCase(oFS.GetAbsolutePathName(SrcPath))) Then
            Set bokTarget = bok
            Exit For
        End If
    Next
    If (bokTarget Is Nothing) Then
        If oFS.FileExists(oFS.GetAbsolutePathName(SrcPath)) Then
            Set bokTarget = Application.Workbooks.Open(oFS.GetAbsolutePathName(SrcPath), UpdateLinks:=0)   'Before_Save�C�x���g���g���K�ɍs���Ǝ��s����(1004)���R�s��
        ElseIf oFS.FileExists(SrcPath) Then
            Set bokTarget = Application.Workbooks.Open(SrcPath, UpdateLinks:=0)
        Else
            '���݂��Ȃ��p�X�������̓u�b�N���������w�肷��Ƃ����ɗ��ăG���[
            If BookExists(SrcPath) Then
                Set bokTarget = Workbooks(SrcPath)
            Else
                Set bokTarget = Application.Workbooks.Add()
                bokTarget.SaveAs oFS.GetAbsolutePathName(SrcPath)
            End If
        End If
    End If
    Set OpenBook = bokTarget
End Function
Public Function SheetExists(book As Workbook, TargetName As String) As Boolean
    Dim sh As Worksheet
    Dim bExists As Boolean
    Dim Target As Worksheet
    
    Set Target = Nothing
    If book Is Nothing Then Set book = ThisWorkbook
    For Each sh In book.Sheets
        If sh.Name = TargetName Then
            Set Target = sh
            Exit For
        End If
    Next
    SheetExists = Not (Target Is Nothing)
End Function
Public Function GetNamedSheet(book As Workbook, TargetName As String) As Worksheet
    Dim sh As Worksheet
    Dim bExists As Boolean
    Dim Target As Worksheet
    
    Set Target = Nothing
    For Each sh In book.Sheets
        If sh.Name = TargetName Then
            Set Target = sh
            Exit For
        End If
    Next
    If Target Is Nothing Then
        Set Target = book.Sheets.Add()
        Target.Name = TargetName
    End If
    Set GetNamedSheet = Target
End Function
Public Function GetFilteredRangeFromList(TargetSheet As Worksheet, Criteria As Range) As Range
    Dim rng As Range
    Dim IsNothingList As Boolean
    Dim o As New SheetCondition
    
    o.Stock TargetSheet
    TargetSheet.Activate
    Set rng = Range("A1").CurrentRegion
    rng.AutoFilter
    IsNothingList = (rng.Rows.Count = 1)
    If Not IsNothingList Then
        rng.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Criteria, Unique:=False
    End If
    'Tips-- AutoFilter��CriteriaN�ł͓��t�^�ɑΉ��ł��Ȃ�
    Set rng = Application.Intersect(rng, rng.Offset(1))
    On Error Resume Next
    Set rng = rng.SpecialCells(xlCellTypeVisible)
    If Err.Number = 1004 Then Set rng = Nothing
    If Err.Number = 91 Then Set rng = Nothing
    On Error GoTo 0
    If TargetSheet.FilterMode Then
        TargetSheet.ShowAllData
    End If
    o.Restore
    Set GetFilteredRangeFromList = rng
End Function
Public Sub PasterTraverceWithRCFunction(Appoint As Range, Rowset() As Variant, FunctionColumn As Variant, FunctionSet As Variant)
    Dim i As Long
    Dim t As Long
    Dim vColumn As Variant
    Dim rngAffect As Range
    Dim oCell As Range
    Dim IsListSetEmpty As Boolean
    Dim rng As Range
    
    PasterTraverce Appoint, Rowset
    If Appoint.Cells.Count = 1 Then
        Set rngAffect = Appoint.Resize(UBound(Rowset, 2) + 1, UBound(Rowset) + 1)
    Else
        Set rngAffect = Appoint
    End If
    IsListSetEmpty = True
    For Each rng In rngAffect.Cells
        IsListSetEmpty = IsListSetEmpty And IsEmpty(rng.Value)
        If Not IsListSetEmpty Then Exit For
    Next
    If (rngAffect Is Nothing) Or IsListSetEmpty Then Exit Sub
    If (VarType(FunctionSet) And vbArray) <> vbArray Then FunctionSet = Array(FunctionSet)
    If (VarType(FunctionColumn) And vbArray) <> vbArray Then FunctionColumn = Array(FunctionColumn)
    For i = 0 To UBound(FunctionColumn)
        Set oCell = Application.Intersect(rngAffect.Cells(1, 1).EntireRow, rngAffect).Cells(, FunctionColumn(i) + 1)
        oCell.FormulaR1C1 = FunctionSet(t Mod (UBound(FunctionSet) + 1))
        If rngAffect.Rows.Count > 1 Then
            Application.Intersect(oCell.EntireColumn, rngAffect).FillDown
        End If
        t = t + 1
    Next
End Sub
Public Function GetTraverce(Rowset As Variant) As Variant
    '�s�񔽓]�����z���Ԃ�
    Dim rc() As Variant
    ReDim rc(LBound(Rowset, 2) To UBound(Rowset, 2), LBound(Rowset) To UBound(Rowset))
    Dim i As Long, t As Long
    For i = LBound(Rowset, 2) To UBound(Rowset, 2)
        For t = LBound(Rowset) To UBound(Rowset)
            rc(i, t) = Rowset(t, i)
        Next
    Next
    GetTraverce = rc
End Function
Public Sub PasterTraverce(Appoint As Range, Rowset As Variant)
    '�s�񔽓]����Variant�z����w��͈̔͂ɓ\��t����
    Dim lRowCount As Long
    Dim lColumnCount As Long
    Dim i As Long
    Dim t As Long
    Dim Dest() As Variant
    Dim IsMultiplyRange As Boolean
    Dim o As New SheetCondition
    Dim rngAffect As Range
    
    'If IsEntryEmpty(Rowset) Then Exit Sub
    If IsEmpty(Rowset) Then Exit Sub
    o.Stock Appoint.Worksheet
    Dim vRowset() As Variant
    If (VarType(Rowset) And vbArray) = vbArray Then
        If Not (LBound(Rowset) = 1 And LBound(Rowset, 2) = 1) Then
            Err.Raise vbObjectError + 311, "PasterTraverce", "Rowset�̓Z���͈̓o���A���g�z��ɏ����܂��B"
        End If
        vRowset = Rowset
    Else
        ReDim vRowset(1 To 1, 1 To 1)
        vRowset(1, 1) = Rowset
    End If
    lRowCount = UBound(vRowset, 2)
    lColumnCount = UBound(vRowset)
    If Appoint.Rows.Count <> 1 Then
        IsMultiplyRange = True
        lRowCount = Application.WorksheetFunction.Min(lRowCount, Appoint.Rows.Count)
    End If
    If Appoint.Columns.Count <> 1 Then
        IsMultiplyRange = True
        lColumnCount = Application.WorksheetFunction.Min(lColumnCount, Appoint.Columns.Count)
    End If
    If IsMultiplyRange Then Appoint.ClearContents
    ReDim Dest(lRowCount, lColumnCount)
    For t = 0 To lRowCount
        For i = 0 To lColumnCount
            Dest(t, i) = vRowset(i, t)
        Next
    Next
    Set rngAffect = Appoint.Cells(1, 1).Resize(lRowCount + 1, lColumnCount + 1)
    rngAffect = Dest
    'TemporaryIdea '�b��I
    RestoreFieldAfterPasteDate rngAffect, Dest '�{��
    o.Restore
End Sub
Public Sub RestoreFieldAfterPasteDate(rngDataArea As Range, vSrcData As Variant)
    'Variant�z���\��t���ċN����g���u���΍􃋁[�`��(2000�N���C���p�b�`�����ĂĂ���Ƃ��̖��͔������Ȃ��悤��)
    Dim i As Long
    Dim t As Long
    Dim vDataArea As Variant
    Dim sTepFormattedDate As String
    
    vDataArea = rngDataArea
    For t = 0 To UBound(vSrcData)
        For i = 0 To UBound(vSrcData, 2)
            If VarType(vDataArea(t + 1, i + 1)) = vbString And IsDate(vSrcData(t, i)) Then
                sTepFormattedDate = rngDataArea.Cells(t + 1, i + 1).Value
                rngDataArea.Cells(t + 1, i + 1).Value = Year(vSrcData(t, i)) & "/" & Left(sTepFormattedDate, InStrRev(sTepFormattedDate, "/") - 1)
            End If
        Next
    Next
    
End Sub
Public Sub Range2Name(AppointRange As Range, NamedString As String)
'�w�肵���͈͂�NamedString�Ƃ������O�ɂ���v���V�[�W��
    ActiveWorkbook.Names.Add Name:=NamedString, RefersTo:="=" & AppointRange.Worksheet.Name & "!" & AppointRange.Address
End Sub
Public Sub CreateDeepFolder(Path As String)
    'ver 0.0.1 UNC�Ή�
    Dim oFS As Scripting.FileSystemObject, sTempName As String, i As Long
    Dim saName() As String
    Dim IsUNC As Boolean, Start As Long
    Dim FolderName  As String
    
    IsUNC = False
    Set oFS = CreateObject("Scripting.FileSystemObject")
    FolderName = Path
    If Left(FolderName, 2) <> "\\" Then
        FolderName = oFS.GetAbsolutePathName(FolderName)
    Else
        IsUNC = True
    End If
    If Not oFS.FolderExists(FolderName) Then
        saName = Split(FolderName, "\")
        If IsUNC Then
            sTempName = "\\" & saName(2)
            Start = 3
        Else
            Start = 1
            sTempName = saName(0)
        End If
        For i = Start To UBound(saName)
            sTempName = sTempName & "\" & saName(i)
            If Not oFS.FolderExists(sTempName) Then
                oFS.CreateFolder (sTempName) '71:�f�B�X�N����������Ă��܂���B
            End If
        Next
    End If
End Sub
Public Function GetBoolFormatCondition(Target As Range, Index As Long) As Boolean
    Dim lTempXOffset As Long
    Dim lTempYOffset As Long
    Dim rngTemp As Range
    
    If Target.Cells.Count <> 1 Then Err.Raise vbObjectError + 95, "GetBoolFormatCondition", "�����Ώۂ͂ЂƂ̃Z���݂̂ł��B"
    If Target.FormatConditions.Count = 0 Then Err.Raise vbObjectError + 320, "GetBoolFormatCondition", "�����Ώۂ͏����t�����������Ă��܂���B"
    If Target.FormatConditions.Count < Index Or Index < 0 Then Err.Raise vbObjectError + 1004, "GetBoolFormatCondition", "�����Ƃ��ė^����ꂽ�C���f�b�N�X�͌����͈͂𒴂��Ă��܂��B"
    If Target.FormatConditions(Index).Type <> xlExpression Then Err.Raise vbObjectError + 76, "GetBoolFormatCondition", "�����Ώۂ̏����t�����͐����ł���K�v������܂��B"
    Set rngTemp = Target.Worksheet.Cells(Target.Worksheet.UsedRange.Rows.Count, 1)
    rngTemp.Formula = Target.FormatConditions(Index).Formula1
    GetBoolFormatCondition = rngTemp.Value
    rngTemp.Clear
End Function
Public Function IsFunction(rng As Range) As Boolean
    Dim cel As Range
    Dim bFunc As Boolean
    
    bFunc = True
    If rng.Count = 1 Then
        IsFunction = Left(rng.Formula, 1) = "="
    Else
        For Each cel In rng
            bFunc = bFunc And Left(cel.Formula, 1) = "="
        Next
        IsFunction = bFunc
    End If
End Function
Public Sub CopyPrintSetting(SrcSheet As Worksheet, DestSheet As Worksheet)
    Dim SrcPageSetup As PageSetup
    
    Set SrcPageSetup = SrcSheet.PageSetup
    With DestSheet.PageSetup
        .PrintTitleRows = SrcPageSetup.PrintTitleRows
        .PrintTitleColumns = SrcPageSetup.PrintTitleColumns
        .PrintArea = SrcPageSetup.PrintArea
        .LeftHeader = SrcPageSetup.LeftHeader
        .CenterHeader = SrcPageSetup.CenterHeader
        .RightHeader = SrcPageSetup.RightHeader
        .LeftFooter = SrcPageSetup.LeftFooter
        .CenterFooter = SrcPageSetup.CenterFooter
        .RightFooter = SrcPageSetup.RightFooter
        .LeftMargin = SrcPageSetup.LeftMargin
        .RightMargin = SrcPageSetup.RightMargin
        .TopMargin = SrcPageSetup.TopMargin
        .BottomMargin = SrcPageSetup.BottomMargin
        .HeaderMargin = SrcPageSetup.HeaderMargin
        .FooterMargin = SrcPageSetup.FooterMargin
        .PrintHeadings = SrcPageSetup.PrintHeadings
        .PrintGridlines = SrcPageSetup.PrintGridlines
        .PrintComments = SrcPageSetup.PrintComments
        .CenterHorizontally = SrcPageSetup.CenterHorizontally
        .CenterVertically = SrcPageSetup.CenterVertically
        .Orientation = SrcPageSetup.Orientation
        .Draft = SrcPageSetup.Draft
        .PaperSize = SrcPageSetup.PaperSize
        .FirstPageNumber = SrcPageSetup.FirstPageNumber
        .Order = SrcPageSetup.Order
        .BlackAndWhite = SrcPageSetup.BlackAndWhite
        .Zoom = SrcPageSetup.Zoom
    End With
    SrcSheet.Activate
End Sub
Public Function CJis(ByVal MixedString As String) As String
    '������ �\��:���p���S�p ����:���p�������S�p����
    Dim sWorker As String
    Dim sPocketWorker As String
    Dim i As Long
    
    sWorker = ""
    For i = 1 To Len(MixedString)
        sPocketWorker = Mid(MixedString, i, 1)
        Select Case Asc(sPocketWorker)
        Case Asc("0") To Asc("9")
            sWorker = sWorker & Chr(Asc(sPocketWorker) - Asc("0") + Asc("�O"))
        Case Else
            sWorker = sWorker & sPocketWorker
        End Select
    Next
    CJis = sWorker
End Function

'RangeSet.bas�̃T�u�Z�b�g
'�͈͂𑼃V�[�g�ɓ��e����
Private Function ReflectRangeOverWorksheet(src As Range, Dest As Worksheet, Optional RowOffset As Long = 0, Optional ColumnOffset As Long = 0) As Range
    Dim rng As Range, rngOneArea As Range, rngWorker As Range
    Set rngWorker = Nothing
    For Each rngOneArea In src.Areas
        If rngWorker Is Nothing Then
            Set rngWorker = Dest.Cells(rngOneArea.Row + RowOffset, rngOneArea.Column + ColumnOffset).Resize(rngOneArea.Rows.Count, rngOneArea.Columns.Count)
        Else
            Set rngWorker = Application.Union(rngWorker, Dest.Cells(rngOneArea.Row + RowOffset, rngOneArea.Column + ColumnOffset).Resize(rngOneArea.Rows.Count, rngOneArea.Columns.Count))
        End If
    Next
    Set ReflectRangeOverWorksheet = rngWorker
End Function

