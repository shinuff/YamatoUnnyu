Attribute VB_Name = "InitSettingHernes"
Option Explicit

'設定ファイルオブジェクト利用のためのハーネス
Const csThisWorkbookKey As String = "ThisWorkbook"
Const csDataKey As String = "Download_Files" '必須設定のうちの一つ。セルに名前を定義しておく。
Const csSettingPath As String = "Setting.config"
Private m_sm As New SettingMemoryObject

Sub TestLoadInit()
    LoadInit
End Sub

Function LoadInit() As Boolean
    'パターン１:設定ファイル優先、常に設定セルに設定値を用意
    '自身のパスをセル値Aとして保存。A=blankなら再設定。
    Dim sSettingFile As String
    '設定ファイルがない場合・必須項目(設定ファイル側)が存在しえない値の時、設定をし直す。
    '設定ファイルに設定値が正しく存在していても環境を移した時・必須項目セル値が設定ファイルのそれと値が違う時には設定ファイルから設定セル値に写す
    Dim oFS As New Scripting.FileSystemObject, bExistsSetting As Boolean
    sSettingFile = oFS.BuildPath(ThisWorkbook.Path, csSettingPath) '同マシンに同様の派生がある場合は配下のフォルダに設定ファイルを入れる
    'sSettingFile = oFS.BuildPath(CreateObject("WshShell").Environments("AppData"), csSettingPath) '同マシンは共有の設定値を持たせたい場合はAppDataフォルダに設定ファイルを入れる
    bExistsSetting = oFS.FileExists(sSettingFile)
    Set m_sm = New SettingMemoryObject
    
    m_sm.Remember sSettingFile
    'Excel側でコントロールする設定リセットトリガ(名前ThisWorkbook値を空白にして保存して再オープン)
    If IsEmpty(ThisWorkbook.Names(csThisWorkbookKey).RefersToRange.Value) Then
        m_sm.Items(csThisWorkbookKey).Value = ""
    End If
    
    'Cが有効
    'Stop
    If ThisWorkbook.Names(csDataKey).RefersToRange.Value <> m_sm(csDataKey).Value Then
        ThisWorkbook.Names(csDataKey).RefersToRange.Value = m_sm(csDataKey).Value
    End If
    Do While Not oFS.FileExists(m_sm.Items(csDataKey).Value)
        SetRequiredItem
        'ダイアログをESCで抜ける
        If m_sm.Items(csDataKey).Value = False Then m_sm.Refresh: Exit Function
    Loop
    If ThisWorkbook.Names(csThisWorkbookKey).RefersToRange.Value <> ThisWorkbook.FullName Then
        ThisWorkbook.Names(csThisWorkbookKey).RefersToRange.Value = ThisWorkbook.FullName
        m_sm.Items(csThisWorkbookKey).Value = ThisWorkbook.FullName
    ElseIf m_sm.Items(csThisWorkbookKey).Value <> ThisWorkbook.FullName Then
        m_sm.Items(csThisWorkbookKey).Value = ThisWorkbook.FullName
    End If
    'D有効？
    If bExistsSetting Then
        'Nop
    Else
        'B存在？

        If Not ThisWorkbook.Names(csDataKey).RefersToRange.Value = m_sm.Items(csDataKey).Value Then ThisWorkbook.Names(csDataKey).RefersToRange.Value = m_sm.Items(csDataKey).Value
        m_sm.Memorize sSettingFile 'Non Proxy
        Exit Function
    End If
    LoadInit = True
ExitInitializing:
    m_sm.Memorize SettingFile
End Function

Sub SetRequiredItem()
    'ダイアログにより設定する。ここで設定する値以外の何の設定値が必須かを知らない。
    '既存の設定セルと比較して異なったら設定セルに値設定
    '設定ファイル優先
    Dim sPaymentDataPath As String, vPath As Variant
    vPath = Misc.GetOpenFilenameOnInitialDir("Microsoft Access Databaseファイル(*.mdb),*.mdb,すべてのファイル (*.*),*.*", 1, "配送台帳を指定してください。", "", False, ThisWorkbook.Path & "\dataStore", True)
    ThisWorkbook.Names(csDataKey).RefersToRange.Value = vPath
    m_sm(csDataKey).Value = vPath
End Sub
Sub ImportBranchCode()
'
    Sheets("支店").Select
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
    Sheets("銀行").Select
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
