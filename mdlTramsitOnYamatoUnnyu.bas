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
    '最後に動かせたファイル(今回のシークエンスではない可能性もある)のタイムスタンプとストッカフォルダの最新とタイムスタンプが一致なら次の評価へ
    Dim oFS As New Scripting.FileSystemObject
    Dim fs As Scripting.File
    If oFS.FileExists(sm.Items("Downloaded_File").Value) Then
        Set fs = oFS.GetFile(sm.Items("Downloaded_File").Value)
        If sm.Items("Downloaded_DateTime").Value <= fs.DateLastModified Then
            
        End If
    End If
    '納品日の内容を
End Sub
Private Function ValidateYamatounnyuCSV(FileName As String) As Boolean
    ValidateYamatounnyuCSV = False
    Dim oFS As New FileSystemObject
    If Not oFS.FileExists(FileName) Then Exit Function
    
    ValidateYamatounnyuCSV = True
End Function
Public Function GetSpanFromYamatounnyuCSV(FileName As String) As DateSpan
    Dim va() As String
    va = StringSet.GetArrayFromCSV(FileName, True, "受付日")
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
    'DownloadedFolderにある(一か月以内の更新かつ決まったフォーマットのファイル名かつ範囲内の年月日)対象があれば一つ所定のフォルダに(移動できれば)移動する
    '動かしたら最新更新ファイルのタイムスタンプ更新し、それより古ければ無視
    Dim sMoveSrc As String, sMoveDest As String
    Dim vLine As Variant
    Dim sm As New SettingMemoryObject
    sm mdlTramsitOnYamatoUnnyu.SettingFile
    For Each vLine In Misc.FilesFinderWithRegExp(SrcFolder, FilterReg)
        
    Next
End Sub
Private Function IsNumericB(Expression As String) As Boolean
    '０-９ と 0-9のみを数字とする
    Dim i As Long
    For i = 1 To Len(Expression)
        Select Case Mid(Expression, i, 1)
        Case "０" To "９", "0" To "9"
        Case Else
            IsNumericB = False: Exit Function
        End Select
    Next
    IsNumericB = True
End Function
'対応フォーマット
'mmdd,mm/dd,m/d,m/dd,mm/d
'/は任意文字列で何文字(ﾊﾞｲﾄ)でも可
Public Function CFormatDateWithoutYear(MonthDateString As String, Format As String, NearType As FixDateEnum, Optional BaseDate As Date = #1/1/1900#) As Date
    Const DefaultDate As Date = #1/1/1900#
    'Formatのバリデート。エラーを許可しない
    'dmは未実装
    If InStr(1, Format, "d", vbTextCompare) < InStr(1, Format, "m", vbTextCompare) Then Err.Raise vbObjectError + 3033, "CFormatDateWithoutYear()", "yet implements."
    'mなし、dなしは定型外
    If InStr(1, Format, "d", vbTextCompare) < 0 Or InStr(1, Format, "m", vbTextCompare) < 0 Then Err.Raise vbObjectError + 3034, "CFormatDateWithoutYear()", "out of range." & vbNewLine & Format
    'm,dが３文字以上は未実装
    If InStr(1, Format, "ddd", vbTextCompare) > 0 Or InStr(1, Format, "mmm", vbTextCompare) > 0 Then Err.Raise vbObjectError + 3035, "CFormatDateWithoutYear()", "yet implements." & vbNewLine & Format
    
    'mの文字数を取得
    Dim lCount_m As Long, lCount_d As Long
    If InStr(1, Format, "mm", vbTextCompare) > 0 Then
        lCount_m = 2
    ElseIf InStr(1, Format, "m", vbTextCompare) > 0 Then
        lCount_m = 1
    End If
    'dの文字数を取得
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
        'セパレータ有りなら見つかった最初のセパレータからlCount_mだけ前から後の文字を取得
        sWork = Mid(MonthDateString, InStr(1, MonthDateString, sSeparator, vbTextCompare) - lCount_m, Len(MonthDateString))
        sMonth = Split(sWork, sSeparator)(0)
        sDay = Split(sWork, sSeparator)(1)
    Else
        'セパレータ無しなら4文字以上の数字の先頭を取得
        Dim i As Long
        For i = 1 To Len(MonthDateString) - lCount_m - lCount_d + 1
            If IsNumericB(Mid(MonthDateString, i, lCount_m + lCount_d)) Then lPointIndex = i: Exit For
        Next
        If lPointIndex = 0 Then Err.Raise vbObjectError + 3036, "CFormatDateWithoutYear()", "dest not have four number_strings."
        sWork = Mid(MonthDateString, lPointIndex, lCount_m + lCount_d)
    End If
    '定型1文字なら対象は1,2文字(2文字の場合1文字目は1)、定型2文字(以上)なら対象は2文字(以上)
    Dim sDay As String
    Dim sMonth As String
    Dim lMonth As Long, lDay As Long
    If lCount_m = 1 Then
        If IsNumericB(Mid(sWork, 1, 2)) Then  '2文字目のみ数字の評価。1文字目空白でも許可
            '2文字
            sMonth = Mid(sWork, 1, 2)
        ElseIf IsNumericB(Mid(sWork, 1, 1)) Then
            '1文字
            sMonth = Mid(sWork, 1, 1)
        End If
        lMonth = Val(sMonth)
    ElseIf lCount_m > 1 Then
        If True Then
            
        Else
        
        End If
    Else
        Err.Raise vbObjectError + 3036, "CFormatDateWithoutYear()", "out of range." & vbNewLine & Format '冗長なバリデート
    End If
    If NearType = FixDateEnum.NearByBase Then
        'mm,ddから数字を取得
        If BaseDate = DefaultDate Then BaseDate = Now()
        
    Else
        Err.Raise vbObjectError + 3044, "CDateWithoutYear()", "yet implements."
    End If
End Function

Public Property Get SettingFile() As String
    '設定ファイルがない場合・必須項目(設定ファイル側)が存在しえない値の時、設定をし直す。
    '設定ファイルに設定値が正しく存在していても環境を移した時・必須項目セル値が設定ファイルのそれと値が違う時には設定ファイルから設定セル値に写す
    Dim oFS As New Scripting.FileSystemObject, bExistsSetting As Boolean
    sSettingFile = oFS.BuildPath(ThisWorkbook.Path, csSettingPath) '同マシンに同様の派生がある場合は配下のフォルダに設定ファイルを入れる
    'sSettingFile = oFS.BuildPath(CreateObject("WshShell").Environments("AppData"), csSettingPath) '同マシンは共有の設定値を持たせたい場合はAppDataフォルダに設定ファイルを入れる
    bExistsSetting = oFS.FileExists(sSettingFile)
    SettingFile = sSettingFile
End Property

