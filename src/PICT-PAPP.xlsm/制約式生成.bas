Attribute VB_Name = "制約式生成"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' 総当たり表の設定した禁則情報からToolの制約式を生成する
Function generateBinaryConstraintExpression(srcWorkbook As Workbook) As Boolean
    generateBinaryConstraintExpression = False
    Dim constraintExpressions As String
    Dim s_expression As String
    Dim roundRobinSheet As Worksheet
    Dim constraintSheet As Worksheet
    If ExistsWorksheet(srcWorkbook, roundRobinSheetName) Then
        Set roundRobinSheet = srcWorkbook.Worksheets(roundRobinSheetName)
    Else
        MsgBox "総当たり表が存在しません。この処理に先立って総当たり表を自動生成し、そのシートに禁則を書き込んで下さい。"
        Exit Function
    End If
    If ExistsWorksheet(srcWorkbook, constraintSheetName) Then
        Set constraintSheet = srcWorkbook.Worksheets(constraintSheetName)
    Else
        MsgBox "制約記述シートが存在しません。この処理には必要です。削除してしまった場合は元のファイルから復元してください。"
        Exit Function
    End If
       
    Dim i As Long
    Dim j As Long
    Dim vN As Long ' 総当たり表の垂直方向の何番目のマスか
    Dim hN As Long ' 総当たり表の水平方向の何番目のマスか
    Dim cellStr As String
    Dim MaxRow As Long
    Dim MaxCol As Long
    
    roundRobinSheet.Unprotect Password:=protectPassword
    Call getMaxRowAndCol(roundRobinSheet, MaxRow, MaxCol)
    
    Dim factorNum As Long
    Dim levelNum As Long
    Dim factorRow As Long
    Dim levelRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    Dim IDs
    factorRow = offsetRows + 1
    levelRow = offsetRows + 2
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    
    For i = levelRow + 1 To MaxRow
        For j = levelCol + 1 To MaxCol
            vN = i - levelRow
            hN = j - levelCol
            If vN < hN Then ' 対角線より右上だけ処理すれば良い
                If roundRobinSheet.Cells(i, j).Value = "×" Then ' 禁則指定されている
                    If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "×" Then
                        MsgBox "禁則のペアが、対角線で線対象になっていません。整合していないセルの値を？に、背景色を赤にしました。" & _
                            i & "行" & j & "列のセルと" & hN + levelRow & "行" & vN + levelCol & "列のセルは対角線で線対称なはずです。"
                        roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "？"
                    End If
                    constraintExpressions = constraintExpressions & "IF [" & roundRobinSheet.Cells(factorRow, j).Value & "] = """ & roundRobinSheet.Cells(levelRow, j).Value & """ THEN [" & _
                                roundRobinSheet.Cells(i, factorCol).Value & "] <> """ & roundRobinSheet.Cells(i, levelCol).Value & """ ;" & vbLf
                    s_expression = s_expression & "(if (== [" & roundRobinSheet.Cells(factorRow, j).Value & "] " & roundRobinSheet.Cells(levelRow, j).Value & ")" & vbLf & _
                                "    (<> [" & roundRobinSheet.Cells(i, factorCol).Value & "] " & roundRobinSheet.Cells(i, levelCol).Value & "))" & vbLf
                End If
            Else ' 対角線より左下については、対角線で線対象になっているかどうかを禁則指定されているペアに限ってチェックする
                If roundRobinSheet.Cells(i, j).Value = "×" Then ' 禁則指定されている
                    If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "×" Then
                        MsgBox "禁則のペアが、対角線で線対象になっていません。整合していないセルの値を？に、背景色を赤にしました。" & _
                            i & "行" & j & "列のセルと" & hN + levelRow & "行" & vN + levelCol & "列のセルは対角線で線対称なはずです。"
                        roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "？"
                    End If
                End If
            End If
        Next j
    Next i
    constraintSheet.Range("自動生成制約").Value = constraintExpressions
    constraintSheet.Cells(constraintSheet.Range("自動生成制約").row, constraintSheet.Range("自動生成制約").Column - 1).Value = roundRobinSheetName
    constraintSheet.Cells(constraintSheet.Range("自動生成制約").row, constraintSheet.Range("自動生成制約").Column + 1).Value = s_expression
    constraintSheet.Activate
    
    roundRobinSheet.Protect Password:=protectPassword

    generateBinaryConstraintExpression = True
End Function

' 多項間禁則マトリクスが定義されている全てのシートについて制約式を生成して制約記述シートに書き込む
Function generateConstraintExpression(srcBook As Workbook) As Boolean
    generateConstraintExpression = False
    Dim i As Long
    Dim srcSheet As Worksheet
    Dim constraintSheet As Worksheet
    Dim idCol As Long
    Dim constrainCol As Long
    Dim binaryConstrainRow As Long
    Dim currentConstrainRow As Long
    Dim expression As String
    Dim s_expression As String
    
    If ExistsWorksheet(srcBook, constraintSheetName) Then
        Set constraintSheet = srcBook.Worksheets(constraintSheetName)
    Else
        MsgBox "制約記述シートが存在しません。この処理には必要です。削除してしまった場合は元のファイルから復元してください。"
        Exit Function
    End If
    If Not generateBinaryConstraintExpression(srcBook) Then
        MsgBox "総当たり表から制約式を生成する際に問題が発生しました。内容を確認してください。"
    End If
    binaryConstrainRow = constraintSheet.Range("自動生成制約").row
    constrainCol = constraintSheet.Range("自動生成制約").Column
    idCol = constrainCol - 1
    
    ' 過去の出力を削除
    Do While True
        If InStr(constraintSheet.Cells(binaryConstrainRow + 1, idCol).Value, kinsokuMatrixSheetBaseName) <> 0 Then
            constraintSheet.Rows(binaryConstrainRow + 1).Delete
        Else
            Exit Do
        End If
    Loop
        
    ' 出力
    currentConstrainRow = binaryConstrainRow
    For i = 1 To srcBook.Sheets.Count ' 全てのシートについて禁則表か判断する
        Set srcSheet = srcBook.Sheets(i)
        If InStr(srcSheet.name, kinsokuMatrixSheetBaseName) <> 0 Then
            If Not kinsokuMatrix2expression(srcSheet, expression, s_expression) Then
                expression = "禁則マトリクスの解析に失敗しました"
            End If
            currentConstrainRow = currentConstrainRow + 1
            constraintSheet.Rows(currentConstrainRow).Insert
            constraintSheet.Cells(currentConstrainRow, idCol).Value = srcSheet.name
            constraintSheet.Cells(currentConstrainRow, constrainCol).Value = expression
            constraintSheet.Cells(currentConstrainRow, constrainCol + 1).Value = s_expression
            constraintSheet.Range("自動生成制約").Copy
            constraintSheet.Cells(currentConstrainRow, idCol).PasteSpecial (xlPasteFormats)
            constraintSheet.Cells(currentConstrainRow, constrainCol).PasteSpecial (xlPasteFormats)
            constraintSheet.Cells(currentConstrainRow, constrainCol + 1).PasteSpecial (xlPasteFormats)
            Application.CutCopyMode = False
        End If
    Next i
    
    constraintSheet.Activate
    
    generateConstraintExpression = True
End Function

' 多項間禁則マトリクスの1枚のシートから制約式を生成する
Function kinsokuMatrix2expression(srcSheet As Worksheet, ByRef expression As String, ByRef s_expression As String) As Boolean
    kinsokuMatrix2expression = False
    
    Dim conditionFactorsRow As Long
    Dim constraintFactorsRow As Long
    Dim firstConditionFactorsCol As Long
    Dim firstConstraintFactorsCol As Long
    Dim conditionNum As Long
    Dim constraintLevelNum As Long
    Dim conditionFactor() As String
    Dim constraintFactor As String
    Dim constraintLevels() As String
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim i As Long
    Dim j As Long
    Dim conditionExpression As String
    Dim conditionSExpression As String

    expression = ""
    s_expression = ""
    srcSheet.Unprotect Password:=protectPassword
    Call getMaxRowAndCol(srcSheet, MaxRow, MaxCol)
    srcSheet.Protect Password:=protectPassword
    
    conditionFactorsRow = offsetRows + 2
    constraintFactorsRow = offsetRows + 1
    firstConditionFactorsCol = offsetColumns + 1
    
    ' 被制約因子の因子名と開始列を見付ける
    For j = firstConditionFactorsCol To MaxCol
        If Not srcSheet.Cells(constraintFactorsRow, j).Value = "" Then
            constraintFactor = srcSheet.Cells(constraintFactorsRow, j).Value
            firstConstraintFactorsCol = j
            Exit For
        End If
    Next j
    If constraintFactor = "" Then
        MsgBox srcSheet.name & "シートの被制約因子が記載されるべき位置において、それを発見できませんでした。"
        Exit Function
    End If
    
    ' 条件因子の数を求める
    conditionNum = firstConstraintFactorsCol - firstConditionFactorsCol
    
    ' 条件因子名を配列に入れる
    ReDim conditionFactors(conditionNum - 1)
    For j = firstConditionFactorsCol To firstConstraintFactorsCol - 1
        conditionFactors(j - firstConditionFactorsCol) = srcSheet.Cells(conditionFactorsRow, j).Value
        If conditionFactors(j - firstConditionFactorsCol) = "" Then
            MsgBox srcSheet.name & "シートの条件因子が記載されるべき位置において、空欄がありました。"
            Exit Function
        End If
    Next j
    
    ' 被制約因子の水準数を求める
    constraintLevelNum = MaxCol - firstConstraintFactorsCol + 1
    
    ' 被制約因子の水準名を配列に入れる
    ReDim constraintLevels(constraintLevelNum - 1)
    For j = firstConstraintFactorsCol To MaxCol
        If Not srcSheet.Cells(constraintFactorsRow, j).Value = constraintFactor Then
            MsgBox srcSheet.name & "シートの被制約因子が記載されるべき位置において、名前が一意になっていません。"
            Exit Function
        End If
        constraintLevels(j - firstConstraintFactorsCol) = srcSheet.Cells(constraintFactorsRow + 1, j).Value
        If constraintLevels(j - firstConstraintFactorsCol) = "" Then
            MsgBox srcSheet.name & "シートの被制約因子の水準が記載されるべき位置において、空欄がありました。"
            Exit Function
        End If
    Next j
    
    For i = conditionFactorsRow + 1 To MaxRow
        ' 条件部
        conditionExpression = ""
        conditionSExpression = ""
        For j = firstConditionFactorsCol To firstConstraintFactorsCol - 1
            If j = firstConditionFactorsCol Then
                conditionExpression = conditionExpression & "IF "
                conditionSExpression = "(if (and "
            Else
                conditionExpression = conditionExpression & " AND "
            End If
            conditionExpression = conditionExpression & "[" & conditionFactors(j - firstConditionFactorsCol) & "] = "
            conditionExpression = conditionExpression & """" & srcSheet.Cells(i, j).Value & """"
            ' S式
            conditionSExpression = conditionSExpression & "(== [" & conditionFactors(j - firstConditionFactorsCol) & "] "
            conditionSExpression = conditionSExpression & srcSheet.Cells(i, j).Value & ") "
        Next j
        ' 制約部
        conditionExpression = conditionExpression & " THEN "
        conditionSExpression = conditionSExpression & ")" & vbLf
        For j = firstConstraintFactorsCol To MaxCol
            If srcSheet.Cells(i, j).Value = "×" Then
                expression = expression & conditionExpression
                expression = expression & "[" & constraintFactor & "] <> "
                expression = expression & """" & constraintLevels(j - firstConstraintFactorsCol) & """"
                expression = expression & ";" & vbLf
                ' S式
                s_expression = s_expression & conditionSExpression
                s_expression = s_expression & "    (<> [" & constraintFactor & "] "
                s_expression = s_expression & constraintLevels(j - firstConstraintFactorsCol) & "))" & vbLf
            End If
        Next j
    Next i

    kinsokuMatrix2expression = True
End Function

' 制約条件の読み取り
Function GetConstraints(ByRef constraints As String) As Boolean
    GetConstraints = False
    
    Dim constraintSheet As Worksheet
    Dim idCol As Long
    Dim constrainCol As Long
    Dim binaryConstrainRow As Long
    Dim i As Long
    
    Set constraintSheet = Worksheets(constraintSheetName)
    binaryConstrainRow = constraintSheet.Range("自動生成制約").row
    constrainCol = constraintSheet.Range("自動生成制約").Column
    idCol = constrainCol - 1
    
    ' まず、総当たり表からの自動生成制約を抽出
    constraints = constraintSheet.Range("自動生成制約").Value
    
    ' 全ての自動生成制約を連結
    For i = binaryConstrainRow + 1 To constraintSheet.Range("Alloyによる検証用表現").row - 2
        If InStr(constraintSheet.Cells(i, idCol).Value, kinsokuMatrixSheetBaseName) <> 0 Then
            constraints = constraints & constraintSheet.Cells(i, constrainCol).Value
        Else
            Exit For
        End If
    Next i
    
    constraints = constraints & constraintSheet.Range("自由記述制約").Value
    
    GetConstraints = True
End Function

' S式(CIT-BACH用)制約条件の読み取り
Function GetSConstraints(ByRef constraints As String) As Boolean
    GetSConstraints = False
    
    Dim constraintSheet As Worksheet
    Dim idCol As Long
    Dim sConstrainCol As Long
    Dim binaryConstrainRow As Long
    Dim i As Long
    
    Set constraintSheet = Worksheets(constraintSheetName)
    binaryConstrainRow = constraintSheet.Range("自動生成制約").row
    sConstrainCol = constraintSheet.Range("自動生成制約").Column + 1
    idCol = sConstrainCol - 2
    
    ' まず、総当たり表からの自動生成制約を抽出
    constraints = constraintSheet.Cells(binaryConstrainRow, sConstrainCol).Value
    
    ' 全ての自動生成制約を連結
    For i = binaryConstrainRow + 1 To constraintSheet.Range("Alloyによる検証用表現").row - 2
        If InStr(constraintSheet.Cells(i, idCol).Value, kinsokuMatrixSheetBaseName) <> 0 Then
            constraints = constraints & constraintSheet.Cells(i, sConstrainCol).Value
        Else
            Exit For
        End If
    Next i
    
    constraints = constraints & constraintSheet.Cells(constraintSheet.Range("自由記述制約").row, sConstrainCol).Value
    
    GetSConstraints = True
End Function


