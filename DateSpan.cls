VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public First As Date
Public Last As Date

Public Function ConcreteFromArray(DateArray() As Date) As DateSpan
    Dim i As Long, datMin As Date, datMax As Date, lLBound As Long
    datMin = #12/31/9999#: datMax = #12:00:00 AM#
    lLBound = LBound(DateArray)
    For i = 0 To UBound(DateArray) - lLBound
        If datMin > DateArray(i + lLBound) Then datMin = DateArray(i + lLBound)
        If datMax < DateArray(i + lLBound) Then datMax = DateArray(i + lLBound)
    Next
    First = datMin: Last = datMax
    Set ConcreteFromArray = Me
End Function

Public Function ModifyDayOfFiveOrZero(Optional IsExtend As Boolean = True) As DateSpan
    Dim datLastDayAtLastMonth As Date, datIsHasty As Date, datIsLate As Date
    If IsExtend Then
        datIsHasty = First
        datIsLate = Last
    Else
        datIsHasty = Last
        datIsLate = First
        'Err.Raise vbObjectError + 491, "ModifyDayOfFiveOrZero()", "yet implements."
    End If
    
    'If DateSerial(Year(datIsLate), Month(datIsLate), datLastDayAtLastMonth - (datLastDayAtLastMonth Mod 5)) <= datIsLate And datIsLate < datLastDayAtLastMonth + 1 Then '範囲として評価したい
    '五十日単位で最初の日付を戻して後の日付を進める
    '戻す方向は存在する事が保障された日付
    If Day(datIsHasty) > 30 Then
        datIsHasty = datIsHasty - 1
    End If
    datIsHasty = datIsHasty - ((Day(datIsHasty) - 1) Mod 5)
    '進めるのは月に依存する日付の可能性がある
    datLastDayAtLastMonth = DateSerial(Year(datIsLate), Month(datIsLate) + 1, 0)
    If Day(datIsLate) <= 30 Then
        datIsLate = datIsLate + (5 - (Day(datIsLate) Mod 5)) Mod 5
    'ElseIf Month(datIsLate) = 2 And (datIsLate > DateSerial(Year(datIsLate), 2, 25)) Then
    '    datIsLate = datLastDayAtLastMonth
    ElseIf DateSerial(Year(datIsLate), Month(datIsLate), 30) <= datIsLate And datIsLate < datLastDayAtLastMonth + 1 Then
        datIsLate = datLastDayAtLastMonth
    Else
        Err.Raise vbObjectError + 491, "ModifyDayOfFiveOrZero()", "yet implements."
    End If
    If IsExtend Then
        First = datIsHasty
        Last = datIsLate
    Else
        First = datIsLate
        Last = datIsHasty
    End If
    Set ModifyDayOfFiveOrZero = Me
End Function

Public Function Concrete(NewFirst As Date, NewLast As Date) As DateSpan
    First = NewFirst
    Last = NewLast
    Set Concrete = Me
End Function

