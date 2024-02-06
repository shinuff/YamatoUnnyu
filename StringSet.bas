Attribute VB_Name = "StringSet"
'@Folder("VBAProject")
'CheckIn ���}�g�^�A.xlsm(update 2024/1/8)$64/32���p
'CheckIn PaymentsReceivable.accdb(update 2021/9/9)
'CheckIn ���[�g�ʐ����ʌ�������e���v���[�g.xlsm(update 2020/4/5)
'CheckIn �����U���x�������Z�p.xls(update 2018/4/5)
'CheckIn ���ϒP���ꗗ�\�o��.accdb(update 2014/7/16)
'CheckIn ����X�`�[130212.xls(update 2013/4/28)
'CheckIn ���[�����X�^�[.xls(2012/12/21)
'CheckIn �Q�n.xls(2012/11/17)

Option Explicit
'�Q�Ɛݒ�:
'           : Microsoft Scripting Runtime
'           : Microsoft VBScript Regular Expressions 5.5
' ������̏W����z��ň���StaticClass
' ��W����Split("","")�ŕ\�����遨Ubound(Split("", "")) = 0 �͋�̔��莞���܂��Ȃ�����API�Ŏ擾���遨API�̏ꍇ���ˑ��ŃG���[�ɂȂ�P�[�X����
'
#If VBA7 Then
Private Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32" (ByVal lpSafeArray As Long) As Long
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef ppsaOut() As Any) As Long
Private Declare PtrSafe Sub CopyMemoryFromArray Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef RetPointer As Long, SrcArray() As Any, Optional ByVal Length As Long = 4&)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef src As Any, Optional ByVal Length As Long = 4&)
Private Declare PtrSafe Sub GetMem4 Lib "msvbvm60" (ByVal ptr As Long, ByRef ret As Long)
#Else
Private Declare Function SafeArrayGetDim Lib "oleaut32" (ByVal lpSafeArray As Long) As Long
Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32" (ByVal cDims As Long, ByRef ppsaOut() As Any) As Long
Private Declare Sub CopyMemoryFromArray Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef RetPointer As Long, SrcArray() As Any, Optional ByVal Length As Long = 4&)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef src As Any, Optional ByVal Length As Long = 4&)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal ptr As Long, ByRef ret As Long)

#End If
'
Private Type ARRAYVARIANT
    vt As Integer
    wReserved1 As Integer
    wReserved2 As Integer
    wReserved3 As Integer
    ppArray As Long
End Type

Public Function FindIndex(MotherSet() As String, FindStr As String, FindPartial As Boolean, Optional IgnoreCase As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    FindIndex = -1
    If FindPartial Then
        For i = LBound(MotherSet) To UBound(MotherSet)
            If InStr(1, MotherSet(i), FindStr, IgnoreCase) > 0 Then
                FindIndex = i
                Exit Function
            End If
        Next
    Else
        For i = LBound(MotherSet) To UBound(MotherSet)
            If StrComp(MotherSet(i), FindStr, IgnoreCase) = 0 Then
                FindIndex = i
                Exit Function
            End If
        Next
    End If
End Function

Public Function SplitWithoutQuotedDelimiter(Expression As String, Quoter As String, Delimiter As String) As String()
    If Quoter = Delimiter Then Err.Raise vbObjectError + 496, "SplitWithoutQuotedDelimiter()", "�f���~�^�ƃN�H�[�^�𓯂������ɂ͏o���܂���"""
    If Len(Quoter) <> 1 Or Len(Delimiter) <> 1 Then Err.Raise vbObjectError + 497
    '�Ԓl����z��̏ꍇ�ԋp
    If Expression = "" Then
        SplitWithoutQuotedDelimiter = Split(Expression, Delimiter)
        Exit Function
    End If
    Dim rc() As String
    Dim lIndex As Long, lCounter As Long, sWorker As String, bIsQuoted As Boolean
    ReDim rc(Len(Expression) - Len(Replace(Expression, Delimiter, "")))
    Do While lIndex <= Len(Expression) - 1
        Select Case Mid(Expression, lIndex + 1, 1)
        Case Quoter
            lIndex = lIndex + 1
            Do Until Mid(Expression, lIndex + 1, 1) = Quoter
                rc(lCounter) = rc(lCounter) & Mid(Expression, lIndex + 1, 1)
                lIndex = lIndex + 1
            Loop
        Case Delimiter
            lCounter = lCounter + 1
        Case Else
            rc(lCounter) = rc(lCounter) & Mid(Expression, lIndex + 1, 1)
            
        End Select
        lIndex = lIndex + 1
    Loop
    If UBound(rc) > lCounter Then ReDim Preserve rc(lCounter)
    SplitWithoutQuotedDelimiter = rc
End Function

Public Sub GetSentenceByReg(LogArticles As String, RegularExpressionPattern As String, ByRef Sentences() As String, ByRef Indexies() As Long)
    '�Z���e���X�Ƃ��̃C���f�b�N�X�͎Q�Ƃœn����Ԃ��B
    Dim oReg As RegExp, oMatches As MatchCollection
    Set oReg = New RegExp
    oReg.Pattern = RegularExpressionPattern
    oReg.MultiLine = True
    oReg.Global = True
    Set oMatches = oReg.Execute(LogArticles)
    Dim lFirstSet As Long, lLength As Long
    If oMatches.Count > 0 Then
        ReDim Sentences(oMatches.Count - 1, oMatches(0).SubMatches.Count - 1)
        ReDim Indexies(oMatches.Count - 1, oMatches(0).SubMatches.Count - 1)
        Dim Index As Long, i As Long, t As Long
        For i = 0 To oMatches.Count - 1
            Index = 1
            lFirstSet = oMatches(i).FirstIndex
            For t = 0 To oMatches(i).SubMatches.Count - 1
                Sentences(i, t) = oMatches(i).SubMatches(t)
                'Index�̒T���Ƀo�O����(�\�z�s�\�ȃP�[�X����)ex.(b(a))//bababb�Ȃ�
                Index = InStr(Index, oMatches(i).Value, oMatches(i).SubMatches(t), vbTextCompare)
                If Index > 0 Then
                    Indexies(i, t) = lFirstSet + Index
                End If
            Next
        Next
    Else
        Sentences = NullArrayForStringType
        Indexies = NullArrayForLongIntegerType
    End If
End Sub
Private Function NullArrayForLongIntegerType() As Long()
    Dim laZero() As Long, lZero As Long '�N���C�A���g�R�[�h���ŃG���[�������o����UBound�֐����g����悤�ɖ���`�ł͂Ȃ���̔z��(Long�^)��Ԃ��B
    Dim pSafeArray As Long
    SafeArrayAllocDescriptor 1, laZero
    CopyMemoryFromArray pSafeArray, laZero
    CopyMemory ByVal pSafeArray + 4&, LenB(lZero), 4&
    NullArrayForLongIntegerType = laZero
End Function
Private Function NullArrayForStringType() As String()
    Dim saZero() As String, sZero As String '�N���C�A���g�R�[�h���ŃG���[�������o����UBound�֐����g����悤�ɖ���`�ł͂Ȃ���̔z��(String�^)��Ԃ��B
    Dim pSafeArray As Long
    SafeArrayAllocDescriptor 1, saZero
    CopyMemoryFromArray pSafeArray, saZero
    CopyMemory ByVal pSafeArray + 4&, LenB(sZero), 4&
    NullArrayForStringType = saZero
End Function
Public Function GetOneDimension(SquareArray() As String, Index As Long, Optional AcrossFirstDimension As Boolean = True) As String()
    '2����T[n�~m]��T[n]���擾,index��0���
    If Dimension(SquareArray) = 2 Then
        'Nop
    Else
        Err.Raise vbObjectError + 431, "GetOneDimension()", "SquareArray��2����z�肵�Ă܂��B"
    End If
    If AcrossFirstDimension Then
        If Index < LBound(SquareArray) Or UBound(SquareArray) < Index Then
            Err.Raise vbObjectError + 457, "GetOneDimension()", "index���͈͊O�ł�"
        Else
            'Nop
        End If
    Else
        If Index < LBound(SquareArray, 2) Or UBound(SquareArray, 2) < Index Then
            Err.Raise vbObjectError + 457, "GetOneDimension()", "index���͈͊O�ł�"
        Else
            'Nop
        End If
    End If
    Dim i As Long, rc() As String
    If AcrossFirstDimension Then
        ReDim rc(LBound(SquareArray) To UBound(SquareArray))
        For i = LBound(SquareArray) To UBound(SquareArray)
            rc(i) = SquareArray(i, Index + LBound(SquareArray, 2))
        Next
    Else
        ReDim rc(LBound(SquareArray, 2) To UBound(SquareArray, 2))
        For i = LBound(SquareArray, 2) To UBound(SquareArray, 2)
            rc(i) = SquareArray(Index + LBound(SquareArray), i)
        Next
    End If
    GetOneDimension = rc
End Function

Public Function TransferOrthogonal(SquareArray() As String) As String()
    '[n�~m]�z���[m�~n]�z��ɕϊ�
    If Dimension(SquareArray) = 2 Then
        'Nop
    Else
        Err.Raise vbObjectError + 431, "TransferOrthogonal()", "SquareArray��2����z�肵�Ă܂��B"
    End If
    Dim rc() As String, i As Long, t As Long
    ReDim rc(LBound(SquareArray, 2) To UBound(SquareArray, 2), LBound(SquareArray) To UBound(SquareArray))
    For t = LBound(SquareArray) To UBound(SquareArray)
        For i = LBound(SquareArray, 2) To UBound(SquareArray, 2)
            rc(i, t) = SquareArray(t, i)
        Next
    Next
    TransferOrthogonal = rc
End Function

Public Function Distinct(p() As String) As String()
    '�_�u�������
    Dim bExists() As Boolean, i As Long, t As Long, lCounter As Long
    '�傫���������Ɍ�����
    ReDim bExists(UBound(p))
    For i = UBound(p) To LBound(p) + 1 Step -1
        For t = i - 1 To LBound(p) Step -1
            If p(i) = p(t) Then
                bExists(i) = True
                lCounter = lCounter + 1
                Exit For
            End If
        Next
    Next
    Dim rc() As String
    ReDim rc(UBound(p) - lCounter)
    t = 0
    For i = LBound(p) To UBound(p)
        If Not bExists(i) Then
            rc(t) = p(i)
            t = t + 1
        End If
    Next
    Distinct = rc
End Function

Public Function Exists(MotherSet() As String, Destination As String) As Boolean
    Dim i As Long
    For i = LBound(MotherSet) To UBound(MotherSet)
        If MotherSet(i) = Destination Then
            Exists = True
            Exit Function
        End If
    Next
    Exists = False
End Function

Public Function Concat(ParamArray StrArray()) As String()
    '1���z�����
    '�Y�����̓����z�񓯎m����������B������Ȃ�ΑS�Ă̔z��ɑ΂��Č�������
    Dim i As Long, t As Long, rc() As String, lMax As Long, lMin As Long
    lMin = 999: lMax = -1
    For i = LBound(StrArray) To UBound(StrArray)
        If VarType(StrArray(i)) > vbArray Then
            If lMax < UBound(StrArray(i)) Then lMax = UBound(StrArray(i))
            If lMin > LBound(StrArray(i)) Then lMin = LBound(StrArray(i))
        End If
    Next
    ReDim rc(lMin To lMax)
    Dim sWork As String
    For i = lMin To lMax
        sWork = ""
        For t = 0 To UBound(StrArray)
            If VarType(StrArray(t)) = vbString Then
                sWork = sWork & StrArray(t)
            ElseIf VarType(StrArray(t)) = vbArray + vbString Then
                If LBound(StrArray(t)) <= i And i <= UBound(StrArray(t)) Then
                    sWork = sWork & StrArray(t)(i)
                End If
            Else
                Stop
            End If
        Next
        rc(i) = sWork
    Next
    Concat = rc
End Function

Public Function Contains(MotherSet() As String, SubSet() As String) As Boolean
    'A��B���ӂŐ^
    'v0.02
    Dim sMotherValue As String, i As Long, t As Long
    For t = LBound(MotherSet) To UBound(MotherSet)
        For i = LBound(SubSet) To UBound(SubSet)
            If SubSet(i) = MotherSet(t) Then
                Contains = True
                Exit Function
            End If
        Next
    Next
    Contains = False
End Function

Public Function Involves(MotherSet() As String, SubSet() As String) As Boolean
    'MotherSet:A��SubSet:B=B�Ő^
    Dim i As Long, t As Long, bExists As Boolean
    For i = LBound(SubSet) To UBound(SubSet)
        bExists = False
        For t = LBound(MotherSet) To UBound(MotherSet)
            If SubSet(i) = MotherSet(t) Then
                bExists = True
                Exit For
            End If
        Next
        If Not bExists Then
            Involves = False
            Exit Function
        End If
    Next
    Involves = True
End Function

Private Sub ZeroMassTest()
    Dim saZeroMass() As String
    Debug.Print Sgn(saZeroMass)
    saZeroMass = NullArrayForStringType 'NullArrayForStringType
    Debug.Print Sgn(saZeroMass)
    'Debug.Print UBound(saZeroMass) & ", " & LBound(saZeroMass)
    'Debug.Print UBound(Array()) & ", " & LBound(Array())
    Dim saA() As String, saB() As String, sac() As String
    saA = StringSet.CArrayStr(Array("a", "Aa"))
    saB = StringSet.CArrayStr(Array("s", "b", ""))
    sac = StringSet.Disjunction(saA, saB)
    Debug.Assert Join(sac, ",") = "a,Aa,s,b,"
    sac = StringSet.Disjunction(saZeroMass, saA)
    Debug.Assert Join(sac, ",") = "a,Aa"
    sac = StringSet.Disjunction(saB, saZeroMass)
    Debug.Assert Join(sac, ",") = "s,b,"
    sac = StringSet.Disjunction(saZeroMass, saZeroMass)
    Debug.Assert Join(sac, ",") = ""
    sac = StringSet.MaterialConditional(saZeroMass, saZeroMass)
    Debug.Assert Join(sac, ",") = ""
    sac = StringSet.MaterialConditional(saA, saZeroMass)
    Debug.Assert Join(sac, ",") = ""
    sac = StringSet.MaterialConditional(saZeroMass, saB)
    Debug.Assert Join(sac, ",") = "s,b,"
    sac = StringSet.Conjunction(saZeroMass, saB)
    Debug.Assert Join(sac, ",") = ""
    sac = StringSet.Conjunction(saA, saZeroMass)
    Debug.Assert Join(sac, ",") = ""
    sac = StringSet.Conjunction(saZeroMass, saZeroMass)
    Debug.Assert Join(sac, ",") = ""
    Debug.Assert StringSet.Equals(saZeroMass, saZeroMass)
    Debug.Assert Not StringSet.Equals(saA, saZeroMass)
    Debug.Assert Not StringSet.Equals(saZeroMass, saB)
End Sub

Public Function Equals(MotherSet() As String, CompareSet() As String) As Boolean
    'A=B�Ő^(A��B����B��A�Ő^)
    Dim i As Long, t As Long, bExistsSubset As Boolean
    For i = LBound(MotherSet) To UBound(MotherSet)
        bExistsSubset = False
        For t = LBound(CompareSet) To UBound(CompareSet)
            If MotherSet(i) = CompareSet(t) Then
                bExistsSubset = True
                Exit For
            End If
        Next
        If bExistsSubset = False Then
            Equals = False
            Exit Function
        End If
    Next
    For i = LBound(CompareSet) To UBound(CompareSet)
        bExistsSubset = False
        For t = LBound(MotherSet) To UBound(MotherSet)
            If CompareSet(i) = MotherSet(t) Then
                bExistsSubset = True
                Exit For
            End If
        Next
        If bExistsSubset = False Then
            Equals = False
            Exit Function
        End If
    Next
    Equals = True
End Function

'v0.0.1 ��W���Ή�
Public Function MaterialConditional(A() As String, B() As String) As String()
    'A��B�̊܈ӏW�����擾������A��B��M
    '�Ԓl�̓_�u����������Ȃ��B�_�u���A��B�ł���W����D�悵�A�������Ԓl�ƂȂ�B0���
    If Sgn(A) = 0 Then A = NullArrayForStringType '���I�z�񂪖��ݒ�Ȃ��z����Z�b�g
    If Sgn(B) = 0 Then B = NullArrayForStringType '���I�z�񂪖��ݒ�Ȃ��z����Z�b�g
    If UBound(A) - LBound(B) = -1 And UBound(B) - LBound(B) = -1 Then
        MaterialConditional = NullArrayForStringType
        Exit Function
    ElseIf UBound(A) - LBound(A) = -1 Then
        MaterialConditional = B
        Exit Function
    ElseIf UBound(B) - LBound(B) = -1 Then
        MaterialConditional = NullArrayForStringType
        Exit Function
    End If
    
    Dim bIsHandShaked() As Boolean, bExists() As Boolean, i As Long, t As Long, lCounter As Long
    ReDim bIsHandShaked(LBound(A) To UBound(A)), bExists(LBound(B) To UBound(B))
    For t = LBound(B) To UBound(B)
        For i = LBound(A) To UBound(A)
            If Not bIsHandShaked(i) Then
                If A(i) = B(t) Then
                    bExists(t) = True
                    bIsHandShaked(i) = True
                    Exit For
                End If
            End If
        Next
        If Not bExists(t) Then
            lCounter = lCounter + 1
        End If
    Next
    If lCounter > 0 Then
        Dim rc() As String
        ReDim rc(lCounter - 1)
        i = 0
        For t = LBound(B) To UBound(B)
            If Not bExists(t) Then
                rc(i) = B(t)
                i = i + 1
            End If
        Next
        MaterialConditional = rc
    Else
        MaterialConditional = NullArrayForStringType
    End If
End Function

'v0.0.1 ��W���Ή�
Public Function Disjunction(A() As String, B() As String) As String()
    'A��B�̘a�W�����擾
    '�Ԓl�̓_�u����������Ȃ��B0���
    'A��^A��B�̓����������
    'A��B�����߁AA�̓n���h�V�F�C�N����ꂽ�玟����͑Ώۂɂ��Ȃ��AB�͑��݃`�F�b�N������
    If Sgn(A) = 0 Then A = NullArrayForStringType '���I�z�񂪖��ݒ�Ȃ��z����Z�b�g
    If Sgn(B) = 0 Then B = NullArrayForStringType '���I�z�񂪖��ݒ�Ȃ��z����Z�b�g
    If UBound(A) - LBound(A) = -1 And UBound(B) - LBound(B) = -1 Then
        Disjunction = NullArrayForStringType
        Exit Function
    ElseIf UBound(A) - LBound(A) = -1 Then
        Disjunction = B
        Exit Function
    ElseIf UBound(B) - LBound(B) = -1 Then
        Disjunction = A
        Exit Function
    End If
    Dim bIsHandShaked() As Boolean, bExists() As Boolean
    ReDim bIsHandShaked(LBound(A) To UBound(A)), bExists(LBound(B) To UBound(B))
    Dim i As Long, t As Long
    For t = LBound(B) To UBound(B)
        For i = LBound(A) To UBound(A)
            If Not (bIsHandShaked(i)) Then
                If A(i) = B(t) Then
                    bExists(t) = True
                    bIsHandShaked(i) = True
                    Exit For
                End If
            End If
        Next
    Next
    Dim rc() As String 'RC��0���
    ReDim rc(UBound(A) + UBound(B) - LBound(A) - LBound(B) + 2)
    For i = LBound(A) To UBound(A)
        rc(i) = A(i)
    Next
    t = i
    For i = LBound(B) To UBound(B)
        If Not bExists(i) Then
        rc(t) = B(i)
        t = t + 1
        End If
    Next
    ReDim Preserve rc(t - 1)
    Disjunction = rc
End Function

'v0.0.1 ��W���Ή�
Public Function Conjunction(A() As String, B() As String) As String()
    'A��B�̐ϏW��C���擾
    '�Ԓl��A��C,B��C�𖞂����A�_�u����������Ȃ��B0���
    '��W����NullArrayForStringType��Ԃ��B
    If Sgn(A) = 0 Then A = NullArrayForStringType '���I�z�񂪖��ݒ�Ȃ��z����Z�b�g
    If Sgn(B) = 0 Then B = NullArrayForStringType '���I�z�񂪖��ݒ�Ȃ��z����Z�b�g
    If UBound(A) - LBound(A) = -1 And UBound(B) - LBound(B) = -1 Then
        Conjunction = NullArrayForStringType
        Exit Function
    ElseIf UBound(A) - LBound(A) = -1 Then
        Conjunction = NullArrayForStringType
        Exit Function
    ElseIf UBound(B) - LBound(B) = -1 Then
        Conjunction = NullArrayForStringType
        Exit Function
    End If
    Dim rc() As String
    Dim bHandShaked() As Boolean, bExists() As Boolean
    ReDim bHandShaked(LBound(B) To UBound(B)), bExists(LBound(A) To UBound(A))
    Dim i As Long, t As Long, lCount As Long
    For i = LBound(A) To UBound(A)
        For t = LBound(B) To UBound(B)
            If Not bHandShaked(t) Then
                If B(t) = A(i) Then
                    bExists(i) = True
                    bHandShaked(t) = True
                    lCount = lCount + 1
                    Exit For
                End If
            End If
        Next
    Next
    If lCount > 0 Then
        ReDim rc(lCount - 1)
        t = 0
        For i = LBound(A) To UBound(A)
            If bExists(i) Then
                rc(t) = A(i)
                t = t + 1
            End If
        Next
        Conjunction = rc
    Else
        Conjunction = NullArrayForStringType
    End If
End Function

Public Function CArrayStr(X As Variant) As String()
    '�o���A���g�z��𕶎���z��ɂ���B0���
    Dim rc() As String, i As Long, vLine As Variant
    Dim vX As Variant
    vX = X 'Range��Value Nothing��Null
    If Dimension(vX) = 2 Then
        If UBound(vX, 2) - LBound(vX, 2) = 0 Then
            ReDim rc(UBound(vX) - LBound(vX))
            For i = LBound(vX) To UBound(vX)
                rc(i - LBound(vX)) = vX(i, LBound(vX, 2))
            Next
        ElseIf UBound(vX) - LBound(vX) = 0 Then
            ReDim rc(UBound(vX, 2) - LBound(vX, 2))
            For i = LBound(vX, 2) To UBound(vX, 2)
                rc(i - LBound(vX, 2)) = vX(LBound(vX), i) 'xxx--�G���[(#N/A,#REF!)�͈����܂���B
            Next
        Else
            ReDim rc(UBound(vX) - LBound(vX), UBound(vX, 2) - LBound(vX, 2))
            Dim t As Long
            For i = 0 To UBound(rc, 2)
                For t = 0 To UBound(rc)
                    rc(t, i) = vX(t + LBound(vX), i + LBound(vX))
                Next
            Next
            'Err.Raise vbObjectError + 425, "CArrayStr()", "2���z��͈����܂���B(n,0),(0,m)�ł���K�v������܂��B"
        End If
    Else
        If (VarType(vX) And vbArray) = vbArray Then
            If UBound(vX) - LBound(vX) = -1 Then
                rc = NullArrayForStringType
            Else
            ReDim rc(UBound(vX) - LBound(vX))
                For i = LBound(vX) To UBound(vX)
                    rc(i - LBound(vX)) = vX(i)
                Next
            End If
        Else
            ReDim rc(0)
            rc(0) = vX
        End If
    End If
    CArrayStr = rc
End Function

'vbLf�����s�R�[�h�Ή�
Public Function GetArrayFromCSV(FileName As Variant, ExistsHeader As Boolean, ColumnIndexSet As Variant, Optional StartRow As Long = -1, Optional RowCount As Long = -1) As String()
    'Index ��0���, �z�񂩒P�ϐ��̕����񂩐����݂̂����B�z��͒P��̌^�̂�
    '�t�B�[���h�_�u��̏ꍇ�̃J�����w��̓t�B�[���h����̃_�u�葤��2���̏����𖖔��ɉ�����
    '��w��̃o���f�[�^
    Dim bIsApointsHeader As Boolean, vLine As Variant
    Dim bIsAllSameType As Boolean, lVarType As Long, lColumns As Long
    bIsAllSameType = True
    lColumns = 0
    If IsArray(ColumnIndexSet) Then
        ReDim laColumnIndexSet(UBound(ColumnIndexSet) - LBound(ColumnIndexSet))
        lVarType = VarType(ColumnIndexSet(LBound(ColumnIndexSet)))
        lColumns = UBound(ColumnIndexSet) - LBound(ColumnIndexSet)
        For Each vLine In ColumnIndexSet
            bIsAllSameType = bIsAllSameType And (lVarType = VarType(vLine))
            lVarType = VarType(vLine)
            If Not bIsAllSameType Then Err.Raise vbObjectError + 494, "GetArrayFromCSV()", "�z��̗�w��̏ꍇ�A�P��̌^�݂̂��L���ł��B"
        Next
        bIsApointsHeader = lVarType = vbString
    Else
        ReDim laColumnIndexSet(0)
        lVarType = VarType(ColumnIndexSet)
        bIsApointsHeader = lVarType = vbString
    End If
    If lVarType = vbInteger Or lVarType = vbLong Then
        'Nop
    ElseIf lVarType = vbString Then
        'Nop
    Else
        Err.Raise vbObjectError + 495, "GetArrayFromCSV()", "��w��͕����񂩐����݂̂������܂��B"
    End If
    '��w��ƃw�b�_�L���̃o���f�[�^
    If bIsApointsHeader Then
        If lVarType = vbString Then
            'Nop
        ElseIf lVarType = vbLong Or lVarType = vbInteger Then
            ReDim laColumnIndexSet(0)
        Else
            ' �w�b�_�w��̎��͕�����ł̎w��̂�
            Err.Raise vbObjectError + 496, "GetArrayFromCSV()", "�w�b�_�w��̎��͕�����ł̎w��̂�"
        End If
    Else
        'Nop
    End If
    Dim oFS As New Scripting.FileSystemObject, oTS As TextStream
    Dim strRec As String
    Dim saHeader() As String, saCells() As String
    Dim i As Long, j As Long, k As Long
    Dim lQuote As Long
    Dim strCell As String, sHeader As String
    If FileName = False Then
        Exit Function
    End If
    '���R�[�h���ƃw�b�_�̎擾
    Set oTS = oFS.OpenTextFile(FileName) 'CSV�t�@�C�����I�[�v��
    Dim lRowCount As Long
    Do Until oTS.AtEndOfLine
        lRowCount = lRowCount + 1
        strRec = oTS.ReadLine
        If lRowCount = 1 Then sHeader = strRec
        If strRec = "" Then lRowCount = lRowCount - 1
    Loop
    oTS.Close
    If ExistsHeader Then lRowCount = lRowCount - 1
    Dim lRowLimit As Long
    If RowCount = -1 Then
        lRowLimit = lRowCount
    Else
        If lRowCount < RowCount Then
            lRowLimit = lRowCount
        Else
            lRowLimit = RowCount
        End If
    End If
    If StartRow <> -1 Then
        If lRowCount - RowCount < StartRow Then
            lRowLimit = lRowLimit - StartRow
        End If
    Else
        If ExistsHeader Then
            StartRow = 1
        Else
            StartRow = 0
        End If
    End If
    '��s�ڂ��T���v���Ƀt�B�[���h�����擾�A�w�b�_���g���Ȃ���
    Dim bDoubleField() As Boolean
    ReDim bDoubleField(Len(sHeader) - Len(Replace(sHeader, ",", "")))
    ReDim saHeader(Len(sHeader) - Len(Replace(sHeader, ",", "")))
    j = 0
    lQuote = 0
    strCell = ""
    For k = 1 To Len(sHeader)
        Select Case Mid(sHeader, k, 1)
            Case "," '�u"�v�������Ȃ��؂�A��Ȃ炽���̕���
                If lQuote Mod 2 = 0 Then
                    '�u""�v���u"�v�Œu��
                    strCell = Replace(strCell, """""", """")
                    '�O��́u"�v���폜
                    If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
                        strCell = Mid(strCell, 2, Len(strCell) - 2)
                    End If
                    For i = 0 To j - 1
                        bDoubleField(j) = bDoubleField(j) Or (saHeader(i) = strCell)
                        If bDoubleField(j) Then Exit For
                    Next
                    saHeader(j) = strCell
                    strCell = ""
                    lQuote = 0
                    j = j + 1
                Else
                    strCell = strCell & Mid(sHeader, k, 1)
                End If
            Case """" '�u"�v�̃J�E���g���Ƃ�
                lQuote = lQuote + 1
                strCell = strCell & Mid(sHeader, k, 1)
            Case Else
                strCell = strCell & Mid(sHeader, k, 1)
        End Select
    Next
    '�ŏI��̏���
    '�u""�v���u"�v�Œu��
    strCell = Replace(strCell, """""", """")
    '�O��́u"�v���폜
    If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
        strCell = Mid(strCell, 2, Len(strCell) - 2)
    End If
    saHeader(j) = strCell
    For i = 0 To j - 1
        bDoubleField(j) = bDoubleField(j) Or (saHeader(i) = strCell)
        If bDoubleField(j) Then Exit For
    Next
    strCell = ""
    lQuote = 0
    If j < UBound(saHeader) Then
        ReDim Preserve saHeader(j)
        ReDim Preserve bDoubleField(j)
    End If
    If bIsApointsHeader Then
        Dim lSufix As Long, t As Long, oReg As New RegExp, oMatch As VBScript_RegExp_55.Match, bUniqueField As Boolean
        Dim sBase As String
        oReg.Pattern = "(\D+)(\d+)$"
        For i = 0 To UBound(saHeader)
            If bDoubleField(i) Then
                If oReg.test(saHeader(i)) Then
                    '�������t���ꍇ��������+1����������n�߂�
                    Set oMatch = oReg.Execute(saHeader(i)).matches(0)
                    sBase = oMatch.SubMatches(0)
                    lSufix = CLng(oMatch.SubMatches(1)) + 1
                Else
                    lSufix = 2
                    sBase = saHeader(i)
                End If
                Do While StringSet.Exists(saHeader, sBase & lSufix)
                    lSufix = lSufix + 1
                Loop
                saHeader(i) = sBase & lSufix
            Else
                'Nop
            End If
        Next
        If Not StringSet.Involves(saHeader, StringSet.CArrayStr(ColumnIndexSet)) Then
            Err.Raise "GetArrayFromCSV()", "ColumnIndexSet�̎w���[" & Join(StringSet.MaterialConditional(StringSet.CArrayStr(ColumnIndexSet), saHeader), "],[") & "]���w�b�_�ɂ���܂���"
        End If
        If IsArray(ColumnIndexSet) Then
            For i = 0 To UBound(ColumnIndexSet) - LBound(ColumnIndexSet)
                For t = 0 To UBound(saHeader)
                    If saHeader(t) = ColumnIndexSet(i + LBound(ColumnIndexSet)) Then
                        laColumnIndexSet(i) = t
                        Exit For
                    End If
                Next
            Next
        Else
            For t = 0 To UBound(saHeader)
                If saHeader(t) = ColumnIndexSet Then
                    laColumnIndexSet(0) = t
                    Exit For
                End If
            Next
        End If
    Else
        If IsArray(ColumnIndexSet) Then
            For i = 0 To UBound(ColumnIndexSet) - LBound(ColumnIndexSet)
                If 0 > ColumnIndexSet(i + LBound(ColumnIndexSet)) Or ColumnIndexSet(i + LBound(ColumnIndexSet)) > UBound(saHeader) Then Err.Raise vbObjectError + 495, "GetArrayFromCSV()", ColumnIndexSet(i + LBound(ColumnIndexSet)) & "(" & i + LBound(ColumnIndexSet) & ")�C���f�b�N�X�O�̎w��ł��B"
                laColumnIndexSet(i) = ColumnIndexSet(i + LBound(ColumnIndexSet))
            Next
        Else
            laColumnIndexSet(0) = ColumnIndexSet
        End If
    End If
    Set oTS = oFS.OpenTextFile(FileName)  'CSV�t�@�B�����I�[�v��
    ReDim saCells(lRowLimit - 1, UBound(laColumnIndexSet))
    Dim vLiner() As String
    ReDim vLiner(UBound(bDoubleField))
    i = 0
    If ExistsHeader Then oTS.ReadLine        '�w�b�_�w��Ȃ�1�s�ǂݔ�΂�
    Do While i > StartRow
        strRec = oTS.ReadLine
        i = i + 1
    Loop
    i = 0
    Do Until oTS.AtEndOfLine Or i >= lRowLimit
        strRec = oTS.ReadLine
        j = 0
        lQuote = 0
        strCell = ""
        For k = 1 To Len(strRec)
            Select Case Mid(strRec, k, 1)
                Case "," '�u"�v�������Ȃ��؂�A��Ȃ炽���̕���
                    If lQuote Mod 2 = 0 Then
                        '�u""�v���u"�v�Œu��
                        If Len(strCell) > 2 Then
                            strCell = Replace(strCell, """""", """")
                        End If
                        '�O��́u"�v���폜
                        If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
                            strCell = Mid(strCell, 2, Len(strCell) - 2)
                        End If
                        vLiner(j) = strCell
                        j = j + 1
                        strCell = ""
                        lQuote = 0
                    Else
                        strCell = strCell & Mid(strRec, k, 1)
                    End If
                Case """" '�u"�v�̃J�E���g���Ƃ�
                    lQuote = lQuote + 1
                    strCell = strCell & Mid(strRec, k, 1)
                Case Else
                    strCell = strCell & Mid(strRec, k, 1)
            End Select
        Next
        '�ŏI��̏���
        '�u""�v���u"�v�Œu��
        If Len(strCell) > 2 Then
            strCell = Replace(strCell, """""", """")
        End If
        '�O��́u"�v���폜
        If Left(strCell, 1) = """" And Right(strCell, 1) = """" Then
            strCell = Mid(strCell, 2, Len(strCell) - 2)
        End If
        vLiner(j) = strCell
        j = j + 1
        strCell = ""
        lQuote = 0
        For t = 0 To UBound(laColumnIndexSet)
            saCells(i, t) = vLiner(laColumnIndexSet(t))
        Next
        i = i + 1
    Loop
    oTS.Close
    GetArrayFromCSV = saCells
End Function

'DataObjects����̃T�u�Z�b�g(Private)
Private Function Dimension(VarArray As Variant) As Long
    'varArray��Variant()�̏ꍇ
    Dim av As ARRAYVARIANT
    
    If VarType(VarArray) = (vbArray Or vbString) Then
        'VarArray�������񓮓I�z��(�Ƃ��Đ錾����2���ȏ�ɏ�������������)�̏ꍇ
        Dim iCounter As Long
        Dim lDummy As Long
        Dim lErr As Long
        Do
            iCounter = iCounter + 1
            On Error Resume Next
            lDummy = UBound(VarArray, iCounter)
            lErr = Err.Number
            On Error GoTo 0
            If lErr = 9 Then
                Exit Do
            ElseIf lErr = 0 Then
                'nop
            Else
                Stop
                Err.Raise lErr
            End If
        Loop While True
        Dimension = iCounter - 1
    ElseIf IsArray(VarArray) Then
        'VarArray�����̑��̔z��̏ꍇ
        Dim p As Long
        GetMem4 VarPtr(VarArray) + 8, p
        If p = 0 Then
            Dimension = 0  '����������Ă��Ȃ��z��
        Else
            Dimension = SafeArrayGetDim(p) '������
        End If
    Else
        Dimension = -1     '�z��ł͖���
    End If
    
End Function

