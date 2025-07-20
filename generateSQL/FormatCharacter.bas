Attribute VB_Name = "FormatCharacter"
Function MultiConvert(strVal As String) As String
    ' 1. 小文字→大文字
    Dim tmp As String
    tmp = UCase(strVal)
    
    ' 2. 半角→全角（英数字・記号・スペース）
    Dim i As Long
    Dim c As String
    Dim code As Integer
    Dim result As String
    For i = 1 To Len(tmp)
        c = Mid(tmp, i, 1)
        code = AscW(c)
        If code = 32 Then
            result = result & ChrW(&H3000) ' 半角スペース→全角スペース
        ElseIf code >= 33 And code <= 126 Then
            result = result & ChrW(code + &HFEE0) ' 半角英数字・記号→全角
        Else
            result = result & c
        End If
    Next i
    tmp = result
    
    ' 3. 半角カタカナ→全角カタカナ
    tmp = StrConv(tmp, vbWide)
    
    ' 4. ひらがな→カタカナ
    tmp = StrConv(tmp, vbKatakana)
    
    ' 全角コロン2つを半角コロン2つに戻す
    tmp = Replace(tmp, "：：", "::")
    
    MultiConvert = tmp
End Function

