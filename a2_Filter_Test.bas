Attribute VB_Name = "a2_Filter_Test"
Option Explicit
' https://www.youtube.com/channel/UCQMbRhaPEFD1NoZLhRzQzSA/videos?view=0&shelf_id=0&sort=dd
' https://inexsu.wordpress.com

'Public Mock As New Mock

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestMethod

Public Sub A2_Row_Empty_TestMethod()

    On Error GoTo TestFail

    Dim a2() As Variant

    Dim iRow As Long

    a2 = Mock.Generator_a2

    iRow = Mock.Rand_Between(LBound(a2), UBound(a2))

    A2_Row_Fill a2, iRow, vbNullString

    Dim varReturn As Boolean

    varReturn = A2_Row_Empty(a2(), iRow)

    If varReturn = False Then Err.Raise 567, "A2_Row_Empty(a2(),iRow)"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub




'@TestMethod

Public Sub A2_Row_Fill_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    Dim iRow As Long

    Dim str As String

    a2 = Mock.Generator_a2

    iRow = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    str = Mock.Generator_String

    A2_Row_Fill a2(), iRow, str

    If a2(iRow, UBound(a2, 2)) <> str Then _
            Err.Raise 567, vbNullString, vbNullString

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub

'@TestMethod

Public Sub A2_Row_Last_NONEmpty_Number_TestMethod()

    On Error GoTo TestFail

    Dim a2() As Variant

    a2 = Mock.Generator_a2(9, 9)

    A2_Row_Fill a2, 9, vbNullString

    Dim varReturn As Long

    varReturn = A2_Row_Last_NONEmpty_Number(a2())

    If varReturn <> 8 Then Err.Raise 567, "A2_Row_Last_NONEmpty_Number(a2())"

    A2_Row_Fill a2, 8, vbNullString

    varReturn = A2_Row_Last_NONEmpty_Number(a2())

    If varReturn <> 7 Then Err.Raise 567, "A2_Row_Last_NONEmpty_Number(a2())"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub


'@TestMethod

Sub A2_Copy_Part_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    Dim row_Max As Long

    Dim col_Max As Long

    a2 = Mock.Generator_a2(9, 9)

    row_Max = Mock.Rand_Between(LBound(a2), UBound(a2))

    col_Max = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    Dim varReturn() As Variant

    varReturn = A2_Copy_Part(a2(), row_Max, col_Max)

    If varReturn(row_Max, col_Max) <> a2(row_Max, col_Max) _
            Then Err.Raise 567, "a2_Copy_Part(a2(),row_Max,col_Max)"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub

'@TestMethod

Public Sub A2_Row_Empty_Cut_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    a2 = Mock.Generator_a2(9, 9)

    Dim iRow As Long
    iRow = UBound(a2)

    A2_Row_Fill a2, iRow, vbNullString

    Dim varReturn() As Variant

    varReturn = A2_Row_Empty_Cut(a2())

    If UBound(varReturn) <> iRow - 1 _
            Then Err.Raise 567, "A2_Row_Empty_Cut(a2())"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub

'@TestMethod

Public Sub A2_Filter_Date_TestMethod()

    On Error GoTo TestFail

    Dim a2() As Variant

    Dim col_Date As Long

    Dim date_Start As Date, date_End As Date, date_Max As Date

    a2 = Mock.Generator_a2(99, 9)

    col_Date = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    A2_Column_Fill_Rand a2, col_Date, "date"

    date_Start = A2_Column_Min(a2, col_Date)

    date_Max = A2_Column_Max(a2, col_Date)

    date_End = date_Start + (date_Max - date_Start) / 2

    Dim varReturn() As Variant

    varReturn = A2_Filter_Date(a2(), col_Date, date_Start, date_End)

    'if varReturn <> 0 Then Err.Raise 567, "a2_Filter_Date(a2(),col_Date,date_Start,date_End)"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub


'@TestMethod

Public Sub A2_Column_Fill_Rand_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    Dim iCol As Long

    Dim sType As String

    a2 = Mock.Generator_a2

    iCol = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    sType = "date"

    A2_Column_Fill_Rand a2(), iCol, sType

    Dim iRow As Long
    iRow = Mock.Rand_Between(LBound(a2), UBound(a2))

    If IsDate( _
            a2(iRow, iCol)) = False _
            Then Err.Raise 567, vbNullString, vbNullString

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub




'@TestMethod

Public Sub A2_Column_2_a1_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    Dim iCol As Long

    a2 = Mock.Generator_a2

    iCol = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    Dim a1 As Variant

    a1 = A2_Column_2_a1(a2(), iCol)

    Dim iRow As Long
    iRow = Mock.Rand_Between(LBound(a2), UBound(a2))

    If a1(iRow) <> a2(iRow, iCol) Then Err.Raise 567, "a2_Column_2_a1(a2(),iCol)"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub


'@TestMethod

Public Sub A2_Column_Min_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    Dim iCol As Long

    a2 = Mock.Generator_a2(3, 3)

    iCol = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    a2(1, iCol) = CDate(Now)
    a2(2, iCol) = CDate(Now) + 1
    a2(3, iCol) = CDate(Now) + 2

    Dim v As Variant

    v = A2_Column_Min(a2(), iCol)

    If v <> a2(1, iCol) _
            Then Err.Raise 567, "A2_Column_Min(a2(),iCol)"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub

Public Sub A2_Column_Max_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    Dim iCol As Long

    a2 = Mock.Generator_a2(3, 3)

    iCol = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    a2(1, iCol) = CDate(Now)
    a2(2, iCol) = CDate(Now) + 1
    a2(3, iCol) = CDate(Now) + 2

    Dim v As Variant

    v = A2_Column_Max(a2(), iCol)

    If v <> a2(3, iCol) _
            Then Err.Raise 567, "A2_Column_Max(a2(),iCol)"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub

'@TestMethod

Public Sub A2_Row_Copy_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2_Sour() As Variant

    Dim a2_Dest() As Variant

    Dim row_Sour As Long

    Dim row_Dest As Long

    a2_Sour = Mock.Generator_a2(9, 9)

    a2_Dest = Mock.Generator_a2(9, 9)

    row_Sour = Mock.Rand_Between(LBound(a2_Sour), UBound(a2_Sour))

    row_Dest = Mock.Rand_Between(LBound(a2_Dest), UBound(a2_Dest))

    A2_Row_Copy a2_Sour(), a2_Dest(), row_Sour, row_Dest

    Dim iCol As Long
    iCol = Mock.Rand_Between(LBound(a2_Dest, 2), UBound(a2_Dest, 2))

    If a2_Dest(row_Dest, iCol) <> a2_Sour(row_Sour, iCol) _
            Then Err.Raise 567, vbNullString, vbNullString

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub



'@TestMethod

Public Sub A2_Empty_TestMethod()

    On Error GoTo TestFail

    Dim a2() As Variant

    a2 = Mock.Generator_a2

    Dim varReturn() As Variant

    varReturn = A2_Empty(a2())

    Dim iRow As Long, iCol As Long

    iRow = Mock.Rand_Between(LBound(a2), UBound(a2))
    iCol = Mock.Rand_Between(LBound(a2, 2), UBound(a2, 2))

    If varReturn(iRow, iCol) <> vbNullString _
            Then Err.Raise 567, "A2_Empty(a2())"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub

'@TestMethod

Public Sub A2_Row_After_Last_NONEmpty_Number_TestMethod()

    On Error GoTo TestFail

    

    

    Dim a2() As Variant

    a2 = Mock.Generator_a2(9, 9)

    Dim iRow As Long

    iRow = UBound(a2)

    A2_Row_Fill a2, iRow, vbNullString

    Dim varReturn As Long

    varReturn = A2_Row_After_Last_NONEmpty_Number(a2())

    If varReturn <> iRow Then Err.Raise 567, "A2_Row_After_Last_NONEmpty_Number(a2())"

    iRow = iRow - 1

    A2_Row_Fill a2, iRow, vbNullString

    varReturn = A2_Row_After_Last_NONEmpty_Number(a2())

    If varReturn <> iRow Then Err.Raise 567, "A2_Row_After_Last_NONEmpty_Number(a2())"

TestExit:

    Mock.wb.Close False

    Exit Sub

TestFail:

    Mock.wb.Close False

    Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description

End Sub
