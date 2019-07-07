Attribute VB_Name = "a2_Filter"
Option Explicit
' https://www.youtube.com/channel/UCQMbRhaPEFD1NoZLhRzQzSA/videos?view=0&shelf_id=0&sort=dd
' https://inexsu.wordpress.com

Public bDebug As Boolean

Function A2_Filter_Date( _
        a2() As Variant, _
        iCol As Long, _
        date_Start As Date, _
        date_End As Date) _
        As Variant()
' Тестом покрыто
' массив фильтровать: копировать в новый пустой массив,
' только нужные строки,
' затем в новом массиве отрезать пустые строки
' Gem

    Dim a2_New() As Variant
    a2_New = A2_Empty(a2)

    Dim iRow As Long

    For iRow = LBound(a2) To UBound(a2)

        If a2(iRow, iCol) >= date_Start And _
                a2(iRow, iCol) <= date_End Then

            A2_Row_Copy a2, a2_New, _
                    iRow, _
                    A2_Row_After_Last_NONEmpty_Number(a2_New)

        End If
    Next

    A2_Filter_Date = A2_Row_Empty_Cut(a2_New)

End Function

Function A2_Row_After_Last_NONEmpty_Number( _
        a2() As Variant) _
        As Long
' Тестом покрыто
' номер строки после последней заполненной
' то есть первая строка пустой части массива

    Dim iRow As Long

    iRow = A2_Row_Last_NONEmpty_Number(a2) + 1

    If iRow > UBound(a2) Then
        iRow = -1
    End If

    A2_Row_After_Last_NONEmpty_Number = iRow

End Function

Function A2_Empty(a2() As Variant) _
        As Variant()
' Тестом покрыто
' вернуть массив пустой такого же размера
' Gem

    Dim a2_New() As Variant
    ReDim a2_New(LBound(a2) To UBound(a2), _
            LBound(a2, 2) To UBound(a2, 2))

    A2_Empty = a2_New

End Function

Sub A2_Row_Copy( _
        a2_Sour() As Variant, _
        a2_Dest() As Variant, _
        row_Sour As Long, _
        row_Dest As Long)
' Тестом покрыто
' массив строку копировать в другой массив
    Dim iCol As Long

    For iCol = LBound(a2_Sour, 2) To UBound(a2_Sour, 2)

        a2_Dest(row_Dest, iCol) = a2_Sour(row_Sour, iCol)

    Next
End Sub

Function A2_Row_Empty_Cut(a2() As Variant) _
        As Variant()
' Тестом покрыто
' массив усечь сжать отрезать пустые строки
' Gem

    Dim row_Max As Long
    row_Max = A2_Row_Last_NONEmpty_Number(a2)

    A2_Row_Empty_Cut = A2_Copy_Part( _
            a2, row_Max, UBound(a2, 2))

End Function

Public Function A2_Copy_Part( _
        a2() As Variant, _
        row_Max As Long, _
        col_Max As Long) _
        As Variant()
' Тестом покрыто
' массив копировать часть, не весь, частично сверху и слева
' Gem

    Dim a2_New() As Variant
    ReDim a2_New( _
            LBound(a2) To row_Max, _
            LBound(a2, 2) To col_Max)

    Dim iRow As Long, iCol As Long

    For iRow = LBound(a2) To row_Max

        For iCol = LBound(a2, 2) To col_Max

            a2_New(iRow, iCol) = a2(iRow, iCol)

        Next iCol
    Next iRow

    A2_Copy_Part = a2_New

End Function

Function A2_Row_Last_NONEmpty_Number(a2() As Variant) _
        As Long
' Тестом покрыто
' вернуть номер первой снизу НЕПустой строки массива
' Gem

    Dim y As Long

    For y = UBound(a2) To LBound(a2) Step -1

        If A2_Row_Empty(a2, y) = False Then

            A2_Row_Last_NONEmpty_Number = y

            Exit Function

        End If
    Next

    If y > UBound(a2) Then _
            A2_Row_Last_NONEmpty_Number = -1

End Function

Public Function A2_Row_Empty(a2() As Variant, iRow As Long) _
        As Boolean
' Тестом покрыто
' пустая ли строка массива?
' Gem

    Dim boo As Boolean
    boo = True

    Dim iCol As Long

    For iCol = LBound(a2, 2) To UBound(a2, 2)

        If a2(iRow, iCol) <> vbNullString Then

            A2_Row_Empty = False

            Exit Function

        End If
    Next

    A2_Row_Empty = boo

End Function

Function Settings(Optional msg As String)
' Тестом НЕ покрыто
'
End Function

Sub A2_Row_Fill(a2() As Variant, _
        iRow As Long, _
        str As String)
' Тестом покрыто
' массив строку заполнить
' Gem, Mock

    Dim iCol As Long

    For iCol = LBound(a2, 2) To UBound(a2, 2)

        a2(iRow, iCol) = str

    Next
End Sub

Sub A2_Column_Fill_Rand( _
        a2() As Variant, _
        iCol As Long, _
        sType As String)
' Тестом покрыто
' массив столбец заполнить типом данных sType
' случайными значениями
' Gem, Mock

    Dim iRow As Long

    For iRow = LBound(a2) To UBound(a2)

        If UCase(sType) = "DATE" Then

            a2(iRow, iCol) = Mock.Generator_Date

        End If
    Next iRow
End Sub

Function A2_Column_Min(a2() As Variant, iCol As Long) _
        As Variant
' Тестом покрыто
' вернуть минимальное значение из столбца массива
' Gem

    Dim vMin As Variant

    vMin = a2(LBound(a2), iCol)

    Dim iRow As Long

    For iRow = LBound(a2) To UBound(a2)
        If vMin > a2(iRow, iCol) Then
            vMin = a2(iRow, iCol)
        End If
    Next

    A2_Column_Min = vMin

End Function

Function A2_Column_Max(a2() As Variant, iCol As Long) _
        As Variant
' Тестом покрыто
' вернуть МАКСимальное значение из столбца массива
' Gem

    Dim vMax As Variant

    vMax = a2(LBound(a2), iCol)

    Dim iRow As Long

    For iRow = LBound(a2) To UBound(a2)
        If vMax < a2(iRow, iCol) Then
            vMax = a2(iRow, iCol)
        End If
    Next

    A2_Column_Max = vMax

End Function

Function A2_Column_2_a1( _
        a2() As Variant, _
        iCol As Long) _
        As Variant
' Тестом  покрыто
' столбец массива двумерного в массив одномерный
' Gem

    Dim a1 As Variant

    ReDim a1(LBound(a2) To UBound(a2))

    Dim iRow As Long

    For iRow = LBound(a2) To UBound(a2)

        a1(iRow) = a2(iRow, iCol)

    Next

    A2_Column_2_a1 = a1

End Function
