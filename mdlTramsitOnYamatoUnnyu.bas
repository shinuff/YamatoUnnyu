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
    '最後に動かせたファイル(今回のシークエンスではない可能性もある)のタイムスタンプとストッカフォルダの最新とタイムスタンプが一致なら
    Dim oFS As New Scripting.FileSystemObject
    Dim fs As Scripting.File
    If oFS.FileExists(sm.Items("Downloaded_File").Value) Then
        Set fs = oFS.GetFile(sm.Items("Downloaded_File").Value)
        If sm.Items("Downloaded_DateTime").Value <= fs.DateLastModified Then
            
        End If
    End If
End Sub
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
    '設定ファイルがない場合・必須項目(設定ファイル側)が存在しえない値の時、設定をし直す。
    '設定ファイルに設定値が正しく存在していても環境を移した時・必須項目セル値が設定ファイルのそれと値が違う時には設定ファイルから設定セル値に写す
    Dim oFS As New Scripting.FileSystemObject, bExistsSetting As Boolean
    sSettingFile = oFS.BuildPath(ThisWorkbook.Path, csSettingPath) '同マシンに同様の派生がある場合は配下のフォルダに設定ファイルを入れる
    'sSettingFile = oFS.BuildPath(CreateObject("WshShell").Environments("AppData"), csSettingPath) '同マシンは共有の設定値を持たせたい場合はAppDataフォルダに設定ファイルを入れる
    bExistsSetting = oFS.FileExists(sSettingFile)
    SettingFile = sSettingFile
End Property

