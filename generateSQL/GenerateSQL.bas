Attribute VB_Name = "generateSQL"
Option Explicit

Const NUMBER_FORMAT As String = "0_ "

Public Function InsertSQL(テーブル名 As Variant, カラム範囲 As Range, データ範囲 As Range) As String

    Dim InsSQL As String
    Dim colPart As String
    Dim valPart As String
    Dim targetCell As Variant

    ' カラム名部分の組み立て
    colPart = ""
    For Each targetCell In カラム範囲
        If colPart <> "" Then
            colPart = colPart & ", "
        End If
        colPart = colPart & targetCell.Value
    Next

    ' データ部分の組み立て
    valPart = ""
    For Each targetCell In データ範囲
        If valPart <> "" Then
            valPart = valPart & ", "
        End If

        If targetCell.Value = "" Then
            valPart = valPart & "null"
        ElseIf targetCell.NumberFormatLocal = NUMBER_FORMAT Then
            valPart = valPart & targetCell.Value
        Else
            valPart = valPart & "'" & targetCell.Value & "'"
        End If
    Next

    ' SQL全体の組み立て
    InsSQL = "insert into " & テーブル名 & " ("
    InsSQL = InsSQL & colPart & ") values ("
    InsSQL = InsSQL & valPart
    InsSQL = InsSQL & ");"

    InsertSQL = InsSQL
End Function

Public Function UpdateSQL(テーブル名 As Variant, 更新カラム範囲 As Range, 更新データ範囲 As Range, 条件カラム範囲 As Range, 条件データ範囲 As Range) As String

    Dim updSql As String
    Dim setPart As String
    Dim wherePart As String
    Dim targetCell As Variant
    Dim i As Long
    Dim cellValue As Variant

    ' 更新カラムとデータ部分の組み立て (SET句)
    setPart = ""
    For i = 1 To 更新カラム範囲.Cells.Count
        If setPart <> "" Then
            setPart = setPart & ", "
        End If
        
        cellValue = 更新データ範囲.Cells(i).Value
        
        If IsEmpty(cellValue) Or cellValue = "" Then
            setPart = setPart & 更新カラム範囲.Cells(i).Value & " = NULL"
        Else
            setPart = setPart & 更新カラム範囲.Cells(i).Value & " = '" & Replace(cellValue, "'", "''") & "'"
        End If
    Next i

    ' 条件カラムとデータ部分の組み立て (WHERE句)
    wherePart = ""
    For i = 1 To 条件カラム範囲.Cells.Count
        If wherePart <> "" Then
            wherePart = wherePart & " and "
        End If
        wherePart = wherePart & 条件カラム範囲.Cells(i).Value & " = '" & Replace(条件データ範囲.Cells(i).Value, "'", "''") & "'"
    Next i

    ' SQL文の組み立て
    updSql = "update " & テーブル名 & " set " & setPart
    If wherePart <> "" Then
        updSql = updSql & " where " & wherePart
    End If
    updSql = updSql & ";"

    ' 関数の戻り値
    UpdateSQL = updSql

End Function

Public Function SelectSQL(テーブル名 As Variant, 取得カラム範囲 As Range, Optional 条件カラム範囲 As Range = Nothing, Optional 条件データ範囲 As Range = Nothing) As String

    Dim slctSql As String
    Dim selectPart As String
    Dim wherePart As String
    Dim i As Long
    Dim val As Variant

    ' SELECT句の作成
    selectPart = ""
    For i = 1 To 取得カラム範囲.Cells.Count
        If selectPart <> "" Then
            selectPart = selectPart & ", "
        End If
        selectPart = selectPart & 取得カラム範囲.Cells(i).Value
    Next i

    ' WHERE句の作成
    wherePart = ""
    If Not 条件カラム範囲 Is Nothing And Not 条件データ範囲 Is Nothing Then
        If 条件カラム範囲.Cells.Count > 0 And 条件データ範囲.Cells.Count > 0 Then
            For i = 1 To 条件カラム範囲.Cells.Count
                If 条件カラム範囲.Cells(i).Value <> "" And 条件データ範囲.Cells(i).Value <> "" Then
                    If wherePart <> "" Then
                        wherePart = wherePart & " and "
                    End If
                    
                    ' セルの値と書式取得
                    val = 条件データ範囲.Cells(i).Value
                    If val <> "" Then
                        wherePart = wherePart & 条件カラム範囲.Cells(i).Value & " = '" & Replace(val, "'", "''") & "'"
                    End If
                End If
            Next i
        End If
    End If

    ' SQL文の組み立て
    slctSql = "select " & selectPart & " from " & テーブル名
    If wherePart <> "" Then
        slctSql = slctSql & " where " & wherePart
    End If
    slctSql = slctSql & ";"
    
    SelectSQL = slctSql

End Function



