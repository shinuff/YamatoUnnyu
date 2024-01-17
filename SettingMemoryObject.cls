VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingMemoryObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<Written Version="0.1.2" LastUpdate="2023/10/09" LaunchDate="2006/03/09" Author="�V�� �L">
'<Design Summary="���ɂ���ċ��ʂ��Ȃ��f�[�^�̃��[�J���ۑ��N���X">
'<Details>
'<Property Name="Items(Key)" Type="SettingElement" Summary=""/>
'<Interface Name="Concrete"/>
'<Method Name="LoadSetting" Scope="Private"/>
'<Method Name="SaveSetting" Scope="Private"/>
'</Details>
'</Design>
'<RefferenceSetting>
'   <Library Title="Microsoft XML, v3.0"/>
'   <Library Title="Windows Script Host Object Model"/>
'   <Library Title="Microsoft Scripting Runtime"/>
'   <Library Src="SettingElement.cls"/>
'   <Library Src="SettingElementCollection.cls"/>
'   <Library Src="ValueType.bas"/>
'   <Library Src="IClassRule.cls"/>
'</RefferenceSetting>
'</Written>
Option Explicit
Implements IClassRule
Private Const csXmlRootElement As String = "/root"
Private Const csXmlSettingsElement As String = "/Settings"
Private Const csSettingElement As String = "Setting"
Private Const csPurpose As String = "Purpose"
Private Const csRequired As String = "Required"
Private Const csDescription As String = "Description"
Private m_IsForgetting As Boolean
Private m_Items As SettingElementCollection
Private m_sSettingFile As String
Private m_IsDisposed As Boolean
Private m_sComputerName As String
Private m_datLastmodified As Date
Private m_Saved As Boolean
Private m_IsLoadCompleted As Boolean
Private Const mcsDefaultSettingFile As String = "Setting.config"
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

Private Sub Class_Initialize()
    Set m_Items = New SettingElementCollection
    m_sSettingFile = mcsDefaultSettingFile
    m_sComputerName = CreateObject("WScript.Network").ComputerName
    IsForgetting = True
End Sub

Public Property Get SettingFile() As String
    SettingFile = RegulizePath(m_sSettingFile)
End Property
Private Property Let SettingFile(ByVal SettingFilePath As String)
    If Not CreateObject("Scripting.FileSystemObject").FileExists(SettingFilePath) Then Err.Raise vbObjectError + 3300, "SettingFile(w)", "�ݒ�t�@�C����������܂���B(" & SettingFilePath & ")"
    m_sSettingFile = SettingFilePath
End Property
Private Function CastBStr(StringValue As Variant, PurposeForCasting As SettingEnum) As Variant
    Dim sStringValue As String
    sStringValue = CBStr(StringValue)
    Select Case PurposeForCasting
    Case SettingEnum.ForPath
        CastBStr = sStringValue
    Case SettingEnum.ForString
        CastBStr = sStringValue
    Case SettingEnum.None
        CastBStr = vbNull
    Case Else
        Err.Raise vbObjectError + 122, TypeName(Me) & ".CastBStr", "�������̃L���X�g�^�C�v�ł��B"
    End Select
End Function
Private Function CSettingEnum(vSettingEnumString As Variant) As SettingEnum
    Dim sSettingEnumString As String
    sSettingEnumString = LCase(CBStr(vSettingEnumString))
    If sSettingEnumString = "forpath" Then
        CSettingEnum = SettingEnum.ForPath
    ElseIf sSettingEnumString = "forstring" Then
        CSettingEnum = SettingEnum.ForString
    ElseIf sSettingEnumString = "" Or sSettingEnumString = "none" Then
        CSettingEnum = SettingEnum.None
    Else
        Err.Raise vbObjectError + 125, TypeName(Me) & ".CSettingEnum", "�������̐ݒ�񋓌^�ł�"
    End If
End Function
Private Function CStrForSettingEnum(DestSettingEnumValue As SettingEnum) As String
    Select Case DestSettingEnumValue
    Case SettingEnum.None
        CStrForSettingEnum = "None"
    Case SettingEnum.ForPath
        CStrForSettingEnum = "ForPath"
    Case SettingEnum.ForString
        CStrForSettingEnum = "ForString"
    Case SettingEnum.ForNumber
        CStrForSettingEnum = "ForNumber"
    Case Else
        Err.Raise vbObjectError + 126, TypeName(Me) & ".CStrForSettingEnum()", "�������̐ݒ�񋓌^�ł��B"
    End Select
    
End Function
Private Function CBStr(vNullableString As Variant) As String
    If IsNull(vNullableString) Then CBStr = "": Exit Function
    CBStr = CStr(vNullableString)
End Function
Private Function RegulizePath(ByVal SrcPath As String)
    Dim oFS As Scripting.FileSystemObject
    
    RegulizePath = SrcPath
    Set oFS = New Scripting.FileSystemObject
    If LCase(oFS.GetAbsolutePathName(SrcPath)) <> LCase(SrcPath) Then
        SrcPath = GetOneDrivePath(SrcPath)
        SrcPath = oFS.GetAbsolutePathName(SrcPath)
        RegulizePath = SrcPath
    End If
End Function

Private Function GetOneDrivePath(Path As String) As String
     
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
Public Property Get Saved() As Boolean
    Saved = Items.Saved
End Property
Private Property Get ComputerName() As String
    ComputerName = m_sComputerName
End Property

Private Sub RemoveSettingElement(ElementName As String)
    Dim xmlDoc As MSXML2.DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    If Not xmlDoc.Load(SettingFile) Then
        Err.Raise vbObjectError + 1004, "�ݒ�t�@�C�������`���ł͂Ȃ������݂��܂���B" & vbNewLine & "SearchedPath=" & SettingFile
        Exit Sub
    End If
    Dim elmDest As IXMLDOMElement
    Set elmDest = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']/" & csSettingElement & "[@Name='" & ElementName & "']")
    elmDest.ParentNode.RemoveChild elmDest
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
End Sub
Public Sub Forget(ElementName As String)
    If Items.Exists(ElementName) Then
        Items.Remove ElementName
    End If
    RemoveSettingElement ElementName
End Sub
'�ݒ�t�@�C����(��)���[�h����B
'�����̐ݒ�I�u�W�F�N�g�͏㏑������
Public Sub Remember(Optional ByVal ConfigFile As String = "")
    Dim oFS As New Scripting.FileSystemObject
    Dim WshShell As New IWshRuntimeLibrary.WshShell
    Dim sConfigFile As String
    If ConfigFile = "" Then sConfigFile = RegulizePath(SettingFile) Else sConfigFile = RegulizePath(ConfigFile)
    If Not oFS.FileExists(sConfigFile) Then
        m_sSettingFile = sConfigFile
        SaveSetting
    Else
        SettingFile = sConfigFile
    End If
    LoadSetting
    CCast(Me).IsLoadCompleted = True
End Sub
'�ݒ�t�@�C����ǂݍ��݁A�I�u�W�F�N�g�ɑ������
'�����̐ݒ�I�u�W�F�N�g������Ώ㏑������
Private Sub LoadSetting()
    Dim xmlDoc As MSXML2.DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    If Not xmlDoc.Load(SettingFile) Then
        Err.Raise vbObjectError + 1004, "�ݒ�t�@�C�������`���ł͂Ȃ������݂��܂���B" & vbNewLine & "SearchedPath=" & SettingFile
        Exit Sub
    End If
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    m_datLastmodified = oFS.GetFile(SettingFile).DateLastModified
    '�ݒ�t�@�C����ǂݍ��݁A�I�u�W�F�N�g�ɑ������
    Dim elmSetting As IXMLDOMElement
    Dim elmSettings As IXMLDOMElement
    Set elmSettings = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']")
    Dim nodSetting As IXMLDOMNode
    Dim sSettingElementName As String
    Dim seOne As SettingElement
    Dim sKey As String
    If Not elmSettings Is Nothing Then
        For Each nodSetting In elmSettings.childNodes
            If nodSetting.nodeType = NODE_ELEMENT Then
                Set elmSetting = nodSetting
                If elmSetting.nodeName = csSettingElement Then
                    sKey = CBStr(elmSetting.getAttribute("Name"))
                    If sKey <> "" Then
                        If Not Items.Exists(sKey) Then 'xxx--Load�Ȃ̂ɒǋL�^�ŗǂ���?
                            Set seOne = New SettingElement
                            seOne.Concrete sKey, CBStr(elmSetting.getAttribute("Value")), CSettingEnum(elmSetting.getAttribute(csPurpose)), CBool(CBStr(elmSetting.getAttribute(csRequired))), CBStr(elmSetting.getAttribute(csDescription))
                            Items.Add seOne
                        Else
                            Set seOne = New SettingElement
                            seOne.Concrete sKey, CastBStr(elmSetting.getAttribute("Value"), Items(sKey).Purpose), CSettingEnum(elmSetting.getAttribute(csPurpose)), CBool(CBStr(elmSetting.getAttribute(csRequired))), CBStr(elmSetting.getAttribute(csDescription))
                            Set Items(sKey) = seOne
                        End If
                    End If
                End If
            End If
        Next
        IsForgetting = False
    Else
        'INodeList��Item=0�̂Ƃ�Nothing��Ԃ��̂͊��ˑ��H
        
    End If
End Sub
'�ݒ�t�@�C����������΍��
'�K�{�ݒ�I�u�W�F�N�g��Object.Saved=True�ł��ۑ�����
'�C�Ӑݒ�I�u�W�F�N�g��Object.Saved=True�Ȃ�ۑ����Ȃ�
'���̃��\�b�h��ʉ߂������_�őS�Ă�itemObject.Saved=False��True�ɂȂ��Ă��邱��
Private Sub SaveSetting()
    Dim xmlDoc As MSXML2.DOMDocument
    Dim nodSettings As IXMLDOMNodeList
    Dim nodSetting As IXMLDOMNode
    Dim elmSettings As IXMLDOMElement
    Dim elmSetting As IXMLDOMElement ' ���̗v�f�ɑ����l�Ƃ��Đݒ�l������������
    Dim elmOriginal As IXMLDOMElement
    Dim bDirty As Boolean
    Dim elmPi As IXMLDOMProcessingInstruction
    Dim DocumentRoot As IXMLDOMElement
    Dim sRelationalKey As String
    Dim elm As IXMLDOMElement
    Dim oFS As New Scripting.FileSystemObject
    
    bDirty = False
    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    '���`���łȂ���ΐV���ɍ쐬����(�ȑO�̕ۑ����e�͎�����B���C�e���[�V�����ɍl�@�̂���)
    '�ݒ�t�@�C���̐V�K�쐬�̐ӔC������
    If Not xmlDoc.Load(SettingFile) Then
        bDirty = True
        Set elmPi = xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""Shift_JIS"" standalone=""yes""")
        xmlDoc.appendChild elmPi
        Set DocumentRoot = xmlDoc.appendChild(xmlDoc.createElement("root"))
        DocumentRoot.appendChild xmlDoc.createTextNode(vbNewLine)
        Set elmSettings = DocumentRoot.appendChild(xmlDoc.createElement("Settings"))
        DocumentRoot.appendChild xmlDoc.createTextNode(vbNewLine)
        elmSettings.setAttribute "Original", "true"
        elmSettings.appendChild xmlDoc.createTextNode(vbNewLine)
        Set elmSetting = elmSettings.appendChild(xmlDoc.createElement(csSettingElement))
        elmSetting.setAttribute "Name", ""
        elmSetting.setAttribute "Value", ""
        elmSetting.setAttribute csPurpose, ""
        elmSetting.setAttribute csRequired, "false"
        elmSetting.setAttribute "Original", "true"
        Set elmSettings = elmSettings.ParentNode.appendChild(elmSettings.CloneNode(True))
        elmSettings.removeAttribute "Original"
        elmSettings.setAttribute "ComputerName", ComputerName
        elmSettings.ParentNode.appendChild xmlDoc.createTextNode(vbNewLine)
        DocumentRoot.appendChild xmlDoc.createTextNode(vbNewLine)
        IsForgetting = True
    Else
        Set elmSettings = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']")
        '���z���V�K�̏ꍇ�̐ݒ���s��
        If Not elmSettings Is Nothing Then
            'Nop
        Else
            bDirty = True
            Set DocumentRoot = xmlDoc.documentElement
            DocumentRoot.appendChild xmlDoc.createTextNode(vbNewLine)
            Set elmSettings = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@Original='true']").CloneNode(True)
            elmSettings.setAttribute "ComputerName", ComputerName
            elmSettings.removeAttribute "Original"
            Set elmSettings = DocumentRoot.appendChild(elmSettings)
            DocumentRoot.appendChild xmlDoc.createTextNode(vbNewLine)
            elmSettings.appendChild xmlDoc.createTextNode(vbNewLine)
'            Set elmSetting = elmSettings.appendChild(xmlDoc.createElement(csSettingElement))
'            elmSetting.setAttribute "Name", ""
'            elmSetting.setAttribute "Value", ""
'            elmSetting.setAttribute csPurpose, ""
'            elmSetting.setAttribute csRequired, "false"
'            elmSetting.removeAttribute "Original"
'            elmSetting.ParentNode.appendChild xmlDoc.createTextNode(vbNewLine)
            DocumentRoot.appendChild xmlDoc.createTextNode(vbNewLine)

        End If
    End If
    Dim seTarget As SettingElement
    Dim elmOriginalSetting As IXMLDOMElement
    Dim lSettingCounter As Long
    For lSettingCounter = 0 To Items.Count - 1
        Set seTarget = Items.Item(lSettingCounter)
        sRelationalKey = seTarget.Name
        Set elmSetting = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']/" & csSettingElement & "[@Name='" & sRelationalKey & "']")
        If seTarget.Description <> "Proxy" Then
            If seTarget.Required Then
                '�K�{�ݒ�I�u�W�F�N�g��itemObject.Saved=True�ł��ݒ�l���Ⴄ�ꍇ�͕ۑ�����(�čl�̗]�n����?)
                If Not (elmSetting Is Nothing) Then
                    If Not seTarget.Saved Then
                        WriteElement elmSetting, seTarget, bDirty
                    End If
                Else
                    Set elmSettings = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']")
                    If elmSettings Is Nothing Then
                        Err.Raise vbObjectError + 126, TypeName(Me) & ".SaveSetting()", "���ȋL�q��Setting[@ComputerName]�͕K�{�ł�"
                    End If
                    Set elmOriginalSetting = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']/" & csSettingElement & "[@Original='true']")
                    Set elmSetting = elmSettings.appendChild(elmOriginalSetting.CloneNode(True))
                    elmSetting.removeAttribute "Original"
                    WriteElement elmSetting, seTarget, bDirty
                    elmSettings.appendChild xmlDoc.createTextNode(vbNewLine)
                End If
            Else
                '�C�Ӑݒ�I�u�W�F�N�g��Saved=True�̎��ɂ͕ۑ����Ȃ�
                If Not seTarget.Saved Then
                    If Not (elmSetting Is Nothing) Then
                        WriteElement elmSetting, seTarget, bDirty
                    Else
                        Set elmSettings = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']")
                        If elmSettings Is Nothing Then
                            Err.Raise vbObjectError + 126, TypeName(Me) & ".SaveSetting()", "���ȋL�q��Setting[@ComputerName]�͕K�{�ł�"
                        End If
                        Set elmOriginalSetting = xmlDoc.selectSingleNode(csXmlRootElement & csXmlSettingsElement & "[@ComputerName='" & ComputerName & "']/" & csSettingElement & "[@Original='true']")
                        Set elmSetting = elmSettings.appendChild(elmOriginalSetting.CloneNode(True))
                        elmSetting.removeAttribute "Original"
                        WriteElement elmSetting, seTarget, bDirty
                        elmSettings.appendChild xmlDoc.createTextNode(vbNewLine)
                    End If
                End If
            End If
        Else
            'Nop
        End If
    Next
    '�C�Ӑݒ�I�u�W�F�N�g�Őݒ肪����ݒ�l���Ⴂ!Saved�̎��ɕۑ�����
    '�ݒ�I�u�W�F�N�g�ɖ����v�f��ǂݍ��ނ��ǂ����͌�ōl����
'    For Each nodSetting In elmSettings.childNodes
'        If nodSetting.nodeType = NODE_ELEMENT Then
'            Set elmSetting = nodSetting
'            sRelationalKey = elmSetting.getAttribute("Name")
'            If sRelationalKey <> "" Then
'                If Items.Exists(sRelationalKey) Then
'                    Set seTarget = Items.Key(sRelationalKey)
'                    If (Not seTarget.Required) And (Not seTarget.Saved) Then
'                        WriteElement elmSetting, seTarget, bDirty
'                    Else
'                        '�ۑ����Ă���΂����̃��[�v���g�p���Ȃ�
'                        '�K�{�ݒ�I�u�W�F�N�g�͏�L�̃��[�v�ŏ������Ă���͂��Ȃ̂Ŗ�������
'                    End If
'                End If
'            End If
'        Else
'            '�z��OXML�m�[�h�͖���
'        End If
'    Next
    If bDirty Then
        If Not oFS.FolderExists(oFS.GetParentFolderName(SettingFile)) Then
            CreateDeepFolder oFS.GetParentFolderName(oFS.GetAbsolutePathName(SettingFile))
        End If
        xmlDoc.Save SettingFile
        bDirty = False
        m_datLastmodified = oFS.GetFile(SettingFile).DateLastModified
    End If
    
End Sub
Private Sub CreateDeepFolder(Path As String)
    Dim oFS As FileSystemObject
    Dim asSrcPath() As String
    Dim sDestPath As String
    Dim i As Long
    
    Set oFS = New Scripting.FileSystemObject
    Path = oFS.GetAbsolutePathName(Path)
    If Not oFS.FolderExists(Path) Then
        asSrcPath = Split(Path, "\")
        sDestPath = asSrcPath(0) & "\" & asSrcPath(1)
        i = 2
        Do Until oFS.FolderExists(Path)
            If Not oFS.FolderExists(sDestPath) Then
                oFS.CreateFolder sDestPath
            End If
            If UBound(asSrcPath) < i Then Exit Do
            sDestPath = sDestPath & "\" & asSrcPath(i)
            i = i + 1
        Loop
    End If
End Sub
Private Sub WriteElement(ByRef elmSetting As IXMLDOMElement, seTarget As SettingElement, ByRef bDirty As Boolean)
    Dim attr As IXMLDOMAttribute
    For Each attr In elmSetting.Attributes
        If attr.Name = "Original" Then
            elmSetting.removeAttribute "Original"
            Exit For
        End If
    Next
    If seTarget.Purpose <> SettingEnum.None Then
        If IsNull(elmSetting.getAttribute(csPurpose)) Then GoTo WriteOK
        If CastBStr(elmSetting.getAttribute("Value"), CSettingEnum(elmSetting.getAttribute(csPurpose))) <> seTarget.Value Then  'VB�ɕ]���ȗ�������΁A�܂��͂����̕]������Nullable�ɍ�肱�߂Ηǂ������̂���
WriteOK:  'Goto�x�@�ցA�O�s�̕]����And�̑�֕����ށB
            elmSetting.setAttribute "Name", seTarget.Name
            elmSetting.setAttribute "Value", CStr(seTarget.Value)
            elmSetting.setAttribute csRequired, LCase(CStr(seTarget.Required))
            elmSetting.setAttribute csDescription, seTarget.Description
            elmSetting.setAttribute csPurpose, CStrForSettingEnum(seTarget.Purpose)
            'seTarget.Saved = True
            bDirty = True
        End If
    End If
End Sub
Private Function CCast(SettingMemory As SettingMemoryObject) As IClassRule
    Dim io As IClassRule
    Set io = SettingMemory
    Set CCast = io
End Function

Private Sub Class_Terminate()
    If CCast(Me).IsLoadCompleted Then Me.Memorize
    Set m_Items = Nothing
End Sub

Public Function Items() As SettingElementCollection
Attribute Items.VB_UserMemId = 0
    Set Items = m_Items
End Function
'�ݒ�t�@�C���ɕۑ����Đݒ�I�u�W�F�N�g���N���A���Đݒ�����[�h����B
Public Sub Refresh()
    If Not Saved Then SaveSetting
    Items.Clear
    Dim il As IClassRule
    Set il = Me
    il.IsDisposed = True
    LoadSetting
End Sub
Public Sub Memorize(Optional NewSettingFile As String = "")
    If NewSettingFile = "" Then
        If SettingFile = mcsDefaultSettingFile Then
            Err.Raise vbObjectError + 429, "Memorize()", "�ݒ�t�@�C���̏ꏊ���w�肵�ĉ������B"
        Else
            'Nop
        End If
    Else
        SettingFile = RegulizePath(NewSettingFile)
    End If
    Dim lIndexElements As Long
    Dim seTarget As SettingElement
    Dim IsAllSaved As Boolean
    IsAllSaved = True
    For lIndexElements = 0 To Items.Count - 1
        Set seTarget = Items(lIndexElements)
        IsAllSaved = IsAllSaved And seTarget.Saved
        If Not IsAllSaved Then Exit For
    Next
    If Not Saved Or (Not IsAllSaved) Then SaveSetting
End Sub
Private Function PathFinder(NotExistsFile As String) As String
    PathFinder = ""
    Dim oFS As Scripting.FileSystemObject
    Set oFS = New Scripting.FileSystemObject
    Dim sPathFinder As String
    sPathFinder = ""
    On Error Resume Next
    sPathFinder = SuspectPath(oFS.GetFileName(NotExistsFile), ThisWorkbook.Path)
    Err.Clear
    On Error GoTo 0
    If sPathFinder = "" Then
        Dim sExtention As String
        sExtention = LCase(oFS.GetExtensionName(NotExistsFile))
        Dim sExtentionDiscription As String
        Dim sNotExistsFile As String
        sNotExistsFile = oFS.GetAbsolutePathName(NotExistsFile) 'xxx--My Documents�ɍ쐬����͓̂���邩�H
        Dim bIsCreateDummy As Boolean
        If Not oFS.FileExists(sNotExistsFile) Then
            oFS.CreateTextFile(sNotExistsFile, False, False).Write ""
            bIsCreateDummy = True
        End If
        sExtentionDiscription = oFS.GetFile(sNotExistsFile).Type '�w�肳�ꂽ�g���q�̃t�@�C�����(�_�~�[������Ăł�)�^�C�v�������擾����
        If bIsCreateDummy Then oFS.DeleteFile sNotExistsFile
        Dim sDriveLetter As String
        sDriveLetter = Left(sNotExistsFile, 1)
        If oFS.FolderExists(oFS.GetParentFolderName(sNotExistsFile)) Then
            If sDriveLetter <> "\" Then
                If Left(CurDir, 1) <> sDriveLetter Then ChDrive sDriveLetter
            End If
            ChDir oFS.GetParentFolderName(sNotExistsFile)
            Dim WshShell As IWshRuntimeLibrary.WshShell
            Set WshShell = New IWshRuntimeLibrary.WshShell
            WshShell.CurrentDirectory = oFS.GetParentFolderName(sNotExistsFile)
            Set WshShell = Nothing
        End If
        Dim vPathFinder As Variant
        vPathFinder = Application.GetOpenFilename(sExtentionDiscription & "(*." & sExtention & "),(*." & sExtention & ")", , sExtentionDiscription & "��I�����Ă��������B")
        If vPathFinder = False Then vPathFinder = ""
    End If
    PathFinder = CStr(vPathFinder)
End Function
Private Function SuspectPath(FileName, WhereCase) As String
    '�ߗׂ�path�𒲂ׁAFilename������Path���擾����
    Dim oFS As New Scripting.FileSystemObject
    Dim sBasePath As String
    Dim oFolder As Scripting.Folder
    Dim oFolderChild As Scripting.Folder
    Dim sFileName As String
    Dim WshShell As New IWshRuntimeLibrary.WshShell
    Dim sCurrentPath As String
    
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
    sCurrentPath = sBasePath
    '�Z��
    If oFS.FileExists(oFS.BuildPath(sCurrentPath, sFileName)) Then
        SuspectPath = oFS.BuildPath(sCurrentPath, sFileName)
        Exit Function
    End If
    '�q��
    For Each oFolder In oFS.GetFolder(sCurrentPath).SubFolders
        If oFS.FileExists(oFS.BuildPath(oFolder.Path, sFileName)) Then
            SuspectPath = oFS.BuildPath(oFolder.Path, sFileName)
            Exit Function
        End If
    Next
    '�e
    sCurrentPath = oFS.GetParentFolderName(sCurrentPath) & "\"
    If oFS.FileExists(oFS.BuildPath(sCurrentPath, sFileName)) Then
        SuspectPath = oFS.BuildPath(sCurrentPath, sFileName)
        Exit Function
    End If
    '���Ƃ�
    For Each oFolder In oFS.GetFolder(sCurrentPath).SubFolders
        If oFS.FileExists(oFS.BuildPath(oFolder.Path, sFileName)) Then
            SuspectPath = oFS.BuildPath(oFolder.Path, sFileName)
            Exit Function
        End If
    Next
    '�c��
    sCurrentPath = oFS.GetParentFolderName(sCurrentPath) & "\"
    If oFS.FileExists(sCurrentPath & sFileName) Then
        SuspectPath = sCurrentPath & sFileName
        Exit Function
    End If
    '�f��
    For Each oFolder In oFS.GetFolder(sCurrentPath).SubFolders
        If oFS.FileExists(oFS.BuildPath(oFolder.Path, sFileName)) Then
            SuspectPath = oFS.BuildPath(oFolder.Path, sFileName)
            Exit Function
        End If
    Next
    '��
    sCurrentPath = sBasePath
    For Each oFolder In oFS.GetFolder(sCurrentPath).SubFolders
        For Each oFolderChild In oFolder.SubFolders
            If oFS.FileExists(oFS.BuildPath(oFolderChild.Path, sFileName)) Then
                SuspectPath = oFS.BuildPath(oFolderChild.Path, sFileName)
                Exit Function
            End If
        Next
    Next
    Err.Raise vbObjectError + 53, "SuspectPath", "�\�[�X�t�@�C����������܂���B"
End Function
Public Function CreatePathSetting(Name As String, ElementValue As String, Optional IsRequired As Boolean = False, Optional ElementDescription As String = "") As SettingElement
    Dim ReturnValue As SettingElement
    Set ReturnValue = New SettingElement
    ReturnValue.Required = IsRequired
    'IsRequired�Ȃ�ΐݒ肷��ۂɃt�@�C�������݂��Ă���K�v������
    Dim sElementValue As String
    sElementValue = ElementValue
    If IsRequired Then
        Dim oFS As Scripting.FileSystemObject
        Set oFS = New Scripting.FileSystemObject
        If Not oFS.FileExists(sElementValue) Then
            sElementValue = PathFinder(sElementValue)
            If sElementValue = "" Then Err.Raise vbObjectError + 200, TypeName(Me) & ".CreatePathSetting", "�K�{���ڂ̐ݒ���L�����Z������܂����B" & vbNewLine & "�I�����܂��B"
        End If
    End If
    Dim sDescription As String
    If ElementDescription = "" Then
        sDescription = Name
    Else
        sDescription = ElementDescription
    End If
    ReturnValue.Concrete Name, sElementValue, SettingEnum.ForPath, IsRequired, sDescription
    Set CreatePathSetting = ReturnValue
End Function
Public Function CreateStringSetting(Name As String, ElementValue As String, Optional IsRequired As Boolean = False, Optional ElementDescription As String = "") As SettingElement
    Dim ReturnValue As SettingElement
    Set ReturnValue = New SettingElement
    ReturnValue.Required = IsRequired
    Dim sDescription As String
    If ElementDescription = "" Then
        sDescription = Name
    Else
        sDescription = ElementDescription
    End If
    ReturnValue.Concrete Name, ElementValue, SettingEnum.ForString, IsRequired, sDescription
    Set CreateStringSetting = ReturnValue
End Function

Public Property Get IsForgetting() As Boolean
    IsForgetting = m_IsForgetting
End Property

Private Property Let IsForgetting(ByVal vNewValue As Boolean)
    m_IsForgetting = vNewValue
End Property