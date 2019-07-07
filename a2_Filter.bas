Attribute VB_Name = "a2_Filter"
Option Explicit
'https://www.youtube.com/channel/UCQMbRhaPEFD1NoZLhRzQzSA/videos?view=0&shelf_id=0&sort=dd
'https://inexsu.wordpress.com

Public bDebug As Boolean

Function A2_Filter_Date( _
        a2() As Variant, _
        iCol As Long, _
        date_Start As Date, _
        date_End As Date) _
        As Variant()
'Test covered
'filter array: copy to a new empty array,
'just the right lines
'then cut the empty lines in the new array
'Gem

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
'Test covered
'line number after last filled
'that is, the first line of the empty part of the array

    Dim iRow As Long

    iRow = A2_Row_Last_NONEmpty_Number(a2) + 1

    If iRow > UBound(a2) Then
        iRow = -1
    End If

    A2_Row_After_Last_NONEmpty_Number = iRow

End Function

Function A2_Empty(a2() As Variant) _
        As Variant()
'Test covered
'return an empty array of the same size
'Gem

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
'Test covered
'array string copy to another array
    Dim iCol As Long

    For iCol = LBound(a2_Sour, 2) To UBound(a2_Sour, 2)

        a2_Dest(row_Dest, iCol) = a2_Sour(row_Sour, iCol)

    Next
End Sub

Function A2_Row_Empty_Cut(a2() As Variant) _
        As Variant()
'Test covered
'array truncate cut cut off empty lines
'Gem

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
'Test covered
'array copy part, not all, partly on top and to the left
'Gem

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
'Test covered
'return the first number from the bottom of an empty string array
'Gem

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
'Test covered
' Is the array string empty?
'Gem

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
'NOT covered by Test
'
End Function

Sub A2_Row_Fill(a2() As Variant, _
        iRow As Long, _
        str As String)
'Test covered
'array string fill
'Gem mock

    Dim iCol As Long

    For iCol = LBound(a2, 2) To UBound(a2, 2)

        a2(iRow, iCol) = str

    Next
End Sub

Sub A2_Column_Fill_Rand( _
        a2() As Variant, _
        iCol As Long, _
        sType As String)
'Test covered
'array column fill with data type sType
'random values
'Gem mock

    Dim iRow As Long

    For iRow = LBound(a2) To UBound(a2)

        If UCase(sType) = "DATE" Then

            a2(iRow, iCol) = Mock.Generator_Date

        End If
    Next iRow
End Sub

Function A2_Column_Min(a2() As Variant, iCol As Long) _
        As Variant
'Test covered
'return the minimum value from the array column
'Gem

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
'Test covered
'return the MAX value from an array column
'Gem

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
'Test covered
'two-dimensional array column into one-dimensional array
'Gem

     Dim a1 As Variant

     ReDim a1(LBound(a2) To UBound(a2))

     Dim iRow As Long

     For iRow = LBound(a2) To UBound(a2)

         a1(iRow) = a2(iRow, iCol)

     Next

     A2_Column_2_a1 = a1

End Function

