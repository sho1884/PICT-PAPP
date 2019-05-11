Attribute VB_Name = "Alloyソース生成"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' Toolが組み合わせを出力しなかった全てのPairについて、それが(間接)禁則であることを確認するAlloyのソースを生成する
Function createAlloySrc(srcWorkbook As Workbook, ByRef alloySrc As String) As Boolean
    createAlloySrc = False
    Dim factorNum As Long
    Dim levelNum As Long
    Dim flSet As String
    Dim defSystem As String
    Dim pict As String
    Dim alloy As String
    Dim factorNames() As String
    Dim levelLists()
    Dim pairs() As String
    Dim dicDuplication
    Dim predicate As String
    Dim n As Long
    Dim levelName As String
    If FLTable2array(srcWorkbook.Worksheets(FLtblSheetName), factorNames, levelLists) Then
        ' 水準名の衝突を避けるため、水準名別に何因子の中で使われているか数える
        Call generateDicDuplication(factorNames, levelLists, dicDuplication)
        ' 各因子毎の取り得る水準の集合を因子名に同じ名前の集合として定義する
        defSystem = "sig システム {" & vbLf
        For factorNum = LBound(factorNames) To UBound(factorNames)
            flSet = flSet & "enum " & factorNames(factorNum) & " {"
            defSystem = defSystem & vbTab & factorNames(factorNum) & alloyLevelSuffix & ":one " & factorNames(factorNum) & "," & vbLf
            For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
                If levelNum > 0 Then flSet = flSet & ", "
                levelName = levelLists(factorNum)(levelNum)
                If dicDuplication.Item(levelName) > 1 Then ' 水準名が衝突している場合は"_[因子名]"を連結する
                    levelName = levelName & "_" & factorNames(factorNum)
                End If
                flSet = flSet & levelName
            Next levelNum
            flSet = flSet & "}" & vbLf
        Next factorNum
        defSystem = Left(defSystem, Len(defSystem) - 2)
        defSystem = defSystem & vbLf & "}" & vbLf
    Else
        MsgBox "因子・水準の取得に失敗しました"
    End If
    If GetConstraints(pict) Then
        If pict2alloy(pict, dicDuplication, alloy) Then
        Else
            MsgBox "PICT制約式からalloy形式への変換に失敗しました"
        End If
    Else
        MsgBox "PICT制約式の取得に失敗しました"
    End If
    If Not ExistsWorksheet(srcWorkbook, mappedRoundRobinSheetName) Then
        MsgBox "IDマッピング済み総当たり表が存在しません。Toolを実行すると自動生成されるので、先に実行してください。"
        Exit Function
    End If
    If pairsWithoutTestcase(srcWorkbook.Worksheets(mappedRoundRobinSheetName), dicDuplication, pairs) Then
        predicate = "pred 組合せ状態が存在する(s:システム) {" & vbLf
        For n = LBound(pairs) To UBound(pairs)
            If n > 0 Then
                predicate = predicate & " ||" & vbLf
            End If
            predicate = predicate & vbTab & pairs(n)
        Next n
        predicate = predicate & vbLf & "}" & vbLf
    Else
        MsgBox "テストケースの存在しないPair集合の取得に失敗しました"
    End If
    alloySrc = alloySrc & flSet
    alloySrc = alloySrc & defSystem
    alloySrc = alloySrc & alloy
    alloySrc = alloySrc & predicate
    alloySrc = alloySrc & vbLf & alloyExec
    createAlloySrc = True
End Function

' Toolが組み合わせに出力しなかった全てのPairについての情報を収集する。
' それが全て禁則であることをAlloyに確認させるため、Alloyに都合の良い書式で結果が得られる。
Function pairsWithoutTestcase(roundRobinSheet As Worksheet, dicDuplication, ByRef pair() As String) As Boolean
    pairsWithoutTestcase = False
    Dim i As Long
    Dim j As Long
    Dim vN As Long ' 総当たり表の垂直方向の何番目のマスか
    Dim hN As Long ' 総当たり表の水平方向の何番目のマスか
    Dim cellStr As String
    Dim counter As Long
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim factorNum As Long
    Dim levelNum As Long
    Dim factorRow As Long
    Dim levelRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    Dim factor1Name As String
    Dim level1Name As String
    Dim factor2Name As String
    Dim level2Name As String
    
    Call getMaxRowAndCol(roundRobinSheet, MaxRow, MaxCol)
    
    factorRow = offsetRows + 1
    levelRow = offsetRows + 2
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    counter = 0
    
    For i = levelRow + 1 To MaxRow
        For j = levelCol + 1 To MaxCol
            vN = i - levelRow
            hN = j - levelCol
            If vN < hN Then ' 基本的に対角線より右上だけ処理すれば良いはずだが、、、
                Select Case roundRobinSheet.Cells(i, j).Value
                    Case "―" ' 対角線上の自身との組み合わせで、意味が無いので無視
                        If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "―" Then
                            MsgBox "無効エリア―について、対角線で線対象になっていません。整合していない双方のセルの背景色を赤にしました。"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' 赤色
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' 赤色
                        End If
                    Case "×", "？", "?", "" ' テストケースが存在していない場合
                        cellStr = roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value
                        If Not (cellStr = "×" Or cellStr = "?" Or cellStr = "") Then
                            MsgBox "禁則またはテストケース無しのペアが、対角線で線対象になっていません。整合していない双方のセルの背景色を赤にしました。"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' 赤色
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' 赤色
                        End If
                        counter = counter + 1
                        ReDim Preserve pair(counter - 1)
                        factor1Name = roundRobinSheet.Cells(factorRow, j).Value
                        level1Name = roundRobinSheet.Cells(levelRow, j).Value
                        factor2Name = roundRobinSheet.Cells(i, factorCol).Value
                        level2Name = roundRobinSheet.Cells(i, levelCol).Value
                        If dicDuplication.Item(level1Name) > 1 Then ' 水準名が衝突している場合は"_[因子名]"を連結する
                            level1Name = level1Name & "_" & factor1Name
                        End If
                        If dicDuplication.Item(level2Name) > 1 Then ' 水準名が衝突している場合は"_[因子名]"を連結する
                            level2Name = level2Name & "_" & factor2Name
                        End If
                        pair(counter - 1) = "s." & factor1Name & alloyLevelSuffix & " = " & level1Name & " && s." & _
                                    factor2Name & alloyLevelSuffix & " = " & level2Name
                End Select
            Else ' 対角線より左下もテスト項目が無いと思われるペアについてだけ線対象になっているかどうか、チェックだけする
                Select Case roundRobinSheet.Cells(i, j).Value
                    Case "―" ' 対角線上の自身との組み合わせで、意味が無いので無視
                        If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "―" Then
                            MsgBox "無効エリア―について、対角線で線対象になっていません。整合していない双方のセルの背景色を赤にしました。"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' 赤色
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' 赤色
                        End If
                    Case "×", "？", "?", "" ' テストケースが存在していない場合
                        cellStr = roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value
                        If Not (cellStr = "×" Or cellStr = "?" Or cellStr = "") Then
                            MsgBox "禁則またはテストケース無しのペアが、対角線で線対象になっていません。整合していない双方のセルの背景色を赤にしました。"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' 赤色
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' 赤色
                        End If
                End Select
            End If
        Next j
    Next i
       
    pairsWithoutTestcase = True
End Function

' セル内の制約表現式をalloyで検証できる形式に変換する
Function pict2alloy(pict As String, dicDuplication, ByRef alloy As String) As Boolean
    pict2alloy = False
    Dim oReg As Object
    Dim match As Object
    Dim matches As Object
    Dim match2 As Object
    Dim matches2 As Object
    Dim condition As Variant
    Dim conditionStr As String
    Dim i As Long
    Dim factorName As String
    Dim levelName As String
    
    Set oReg = CreateObject("VBScript.Regexp")
    oReg.Pattern = "IF *(\[.+)THEN(.+)<>([^;]+);"
    oReg.Pattern = "IF *(\[.+\] *= *"".+"" *)THEN(.+)<>([^;]+);"
    oReg.Pattern = "IF *(\[.+)THEN *\[(.+)\] *<> *""(.+)"" *;"
    oReg.Global = True
    Set match = oReg.Execute(pict)

    alloy = "{" & vbLf
    For Each matches In match
        oReg.Pattern = " *\[(.+)\] *= *""(.+)"" *"
        condition = Split(matches.Submatches(0), " AND ")
        conditionStr = ""
        For i = LBound(condition) To UBound(condition)
            Set match2 = oReg.Execute(condition(i))
            If i > LBound(condition) Then conditionStr = conditionStr & " and "
            factorName = match2(0).Submatches(0)
            levelName = match2(0).Submatches(1)
            If dicDuplication.Item(levelName) > 1 Then ' 水準名が衝突している場合は"_[因子名]"を連結する
                levelName = levelName & "_" & factorName
            End If
            conditionStr = conditionStr & factorName & alloyLevelSuffix & "=" & levelName
        Next i
        factorName = matches.Submatches(1)
        levelName = matches.Submatches(2)
        If dicDuplication.Item(levelName) > 1 Then ' 水準名が衝突している場合は"_[因子名]"を連結する
            levelName = levelName & "_" & factorName
        End If
        alloy = alloy & vbTab & conditionStr & "=>"
        alloy = alloy & factorName & alloyLevelSuffix & "!="
        alloy = alloy & levelName & vbLf
    Next
    alloy = alloy & vbLf & "}" & vbLf
    
    pict2alloy = True
End Function

