Attribute VB_Name = "その他・Utility"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' 因子・水準表の各因子全てに、MASK状態を表すシンボルの水準を追加する
Function insertMaskSymbol(srcBook As Workbook) As Boolean
    insertMaskSymbol = False
                
    Dim flSheet As Worksheet
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim titleRow As Long
    Dim factorCol As Long
    Dim i As Long
    Dim j As Long
    Dim obj As Object
    Dim content As String
    Dim fEmptyFlg As Boolean
    Dim lEmptyFlg As Boolean

    ' 因子・水準のシートを特定する
    If ExistsWorksheet(srcBook, FLtblSheetName) Then
        Set flSheet = srcBook.Sheets(FLtblSheetName)
    End If
    If flSheet Is Nothing Then ' 結局因子・水準表シートが見つからない
        MsgBox "因子・水準表の名前として「" & FLtblSheetName & "」が指定されていますが、その名前のシートが存在しません。"
        Exit Function
    End If
    
    Call getMaxRowAndCol(flSheet, MaxRow, MaxCol)
    
    Set obj = flSheet.Cells.Find("因子", LookAt:=xlWhole) 'まずは「因子」のセルを探すことでFL表であるかどうかの手掛かりとする
    If obj Is Nothing Then
        Exit Function
    End If
    
    titleRow = obj.row
    factorCol = obj.Column
    If Not flSheet.Cells(titleRow, factorCol + 1).Value Like "水準*" Then
        Exit Function
    End If
    
    fEmptyFlg = False
    
    For i = titleRow + 1 To MaxRow
        content = flSheet.Cells(i, factorCol).Value
        If content = "" Then
            fEmptyFlg = True
        Else
            If fEmptyFlg Then
                MsgBox "因子列の途中に空のセルがあります。空のセル以下の行を無視します。"
                Exit For
            End If
            lEmptyFlg = False
            For j = factorCol + 1 To MaxCol
                content = flSheet.Cells(i, j).Value
                If content = "" Then ' 左から順にみて最初に空であったセルにMASKの状態を表す水準を挿入する
                    If Not lEmptyFlg Then
                        flSheet.Cells(i, j).Value = maskSymbol
                        lEmptyFlg = True
                    End If
                Else
                    If lEmptyFlg Then
                        MsgBox "水準列の途中に空のセルがあり、MASK用の水準を挿入しました。しかし、それより右の列に値が入っています。この状態は問題を起こすので確認してください。"
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    
    insertMaskSymbol = True
End Function

' 因子水準の組み合わせからテストIDを引ける連想配列を作る。同時にペア・リストを出力する
Sub makeFLIDDictionary(paramNames() As String, tuples(), ByRef dicFL As Object, Optional pairListSheet As Worksheet = Nothing)
    Dim testCaseNum As Long
    Dim paramNum As Long
    Dim flKey As String
    
    Set dicFL = CreateObject("Scripting.Dictionary")
    dicFL.CompareMode = vbBinaryCompare
        
    If Not Not tuples Then
        testCaseNum = UBound(tuples) - LBound(tuples) + 1 ' 試験の個数
    Else
        testCaseNum = 0 ' 動的配列の割り当てが未設定
    End If
    
    If Not Not paramNames Then
        paramNum = UBound(paramNames) - LBound(paramNames) + 1 ' 条件項目数
    Else
        paramNum = 0 ' 動的配列の割り当てが未設定
    End If
    
    Dim pairNum As Long
    Dim i As Long
    Dim j As Long
    Dim j2 As Long
    Dim paramL As Long
    Dim paramU As Long
    Dim tuple() As String
    Dim newIDs As String
    
    paramL = LBound(paramNames)
    paramU = UBound(paramNames)
    If Not (pairListSheet Is Nothing) Then ' ペア・リストを出力するシートが指定されている
        pairNum = 0
        Worksheets(controlSheetName).Range("項目タイトル書式").Copy
        pairListSheet.Cells(pairNum + offsetRows + 1, 1 + offsetColumns).Value = "No."
        pairListSheet.Cells(pairNum + offsetRows + 1, 1 + offsetColumns).PasteSpecial (xlPasteFormats)
        pairListSheet.Cells(pairNum + offsetRows + 1, 2 + offsetColumns).Value = "第1因子:第2因子"
        pairListSheet.Cells(pairNum + offsetRows + 1, 2 + offsetColumns).PasteSpecial (xlPasteFormats)
        pairListSheet.Cells(pairNum + offsetRows + 1, 3 + offsetColumns).Value = "第1因子の水準値:第2因子の水準値"
        pairListSheet.Cells(pairNum + offsetRows + 1, 3 + offsetColumns).PasteSpecial (xlPasteFormats)
        pairListSheet.Cells(pairNum + offsetRows + 1, 4 + offsetColumns).Value = "PairwiseID"
        pairListSheet.Cells(pairNum + offsetRows + 1, 4 + offsetColumns).PasteSpecial (xlPasteFormats)
        Worksheets(controlSheetName).Range("値書式").Copy
    End If
    
    Debug.Print Time & " - 辞書作成開始"
    
    For i = 0 To testCaseNum - 1
        tuple = tuples(i)
        For j = paramL To paramU
            For j2 = j + 1 To paramU
                If Not (pairListSheet Is Nothing) Then ' ペア・リストを出力するシートが指定されている
                    pairNum = pairNum + 1
                    pairListSheet.Cells(pairNum + offsetRows + 1, 1 + offsetColumns).Value = pairNum
                    pairListSheet.Cells(pairNum + offsetRows + 1, 1 + offsetColumns).PasteSpecial (xlPasteFormats)
                    pairListSheet.Cells(pairNum + offsetRows + 1, 2 + offsetColumns).Value = paramNames(j) & ":" & paramNames(j2)
                    pairListSheet.Cells(pairNum + offsetRows + 1, 2 + offsetColumns).PasteSpecial (xlPasteFormats)
                    pairListSheet.Cells(pairNum + offsetRows + 1, 3 + offsetColumns).Value = tuple(j) & ":" & tuple(j2)
                    pairListSheet.Cells(pairNum + offsetRows + 1, 3 + offsetColumns).PasteSpecial (xlPasteFormats)
                    pairListSheet.Cells(pairNum + offsetRows + 1, 4 + offsetColumns).Value = "#" & i + 1
                    pairListSheet.Cells(pairNum + offsetRows + 1, 4 + offsetColumns).PasteSpecial (xlPasteFormats)
                End If
                flKey = paramNames(j) & ":" & tuple(j) & "*" & paramNames(j2) & ":" & tuple(j2)
                If dicFL.exists(flKey) Then
                    newIDs = dicFL.Item(flKey) & ", #" & i + 1
                    dicFL.Remove (flKey)
                    dicFL.Add flKey, newIDs
                Else
                    dicFL.Add flKey, "#" & i + 1
                End If
            Next j2
        Next j
    Next i
    
    Debug.Print Time & " - 辞書作成終了"
    Debug.Print dicFL.Count
    
End Sub

' 因子・水準表を特定して、因子・水準を配列に読み込む処理を呼び出す
Function FLTableSheet2array(srcBook As Workbook, ByRef factorNames() As String, ByRef levelLists()) As Boolean
    FLTableSheet2array = False

    Dim FLtblSheet As Worksheet

    ' 因子・水準のシートを特定する
    If Not FLtblSheetName = "" Then
        If ExistsWorksheet(srcBook, FLtblSheetName) Then
            Set FLtblSheet = srcBook.Sheets(FLtblSheetName)
        End If
    End If
    If FLtblSheet Is Nothing Then ' 結局因子・水準表シートが見つからない
        MsgBox "因子・水準表の名前として「" & FLtblSheetName & "」が指定されていますが、その名前のシートが存在しません。"
        Exit Function
    End If

    ' 因子・水準を配列に読み込む
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "因子・水準の解析に失敗しました。処理を中止します。"
        Exit Function
    End If

    FLTableSheet2array = True
End Function

' 特定された因子・水準表の因子・水準を配列に読み込む
Function FLTable2array(flSheet As Worksheet, ByRef factorNames() As String, ByRef levelLists()) As Boolean
    FLTable2array = False ' FL表形式だと判断したら必ずTrueを返す
    
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim titleRow As Long
    Dim factorCol As Long
    Dim factorNum As Long
    Dim levelNum As Long
    Dim i As Long
    Dim j As Long
    Dim obj As Object
    Dim content As String
    Dim emptyFlg As Boolean
    Dim levelNames() As String

    Call getMaxRowAndCol(flSheet, MaxRow, MaxCol)
    
    Set obj = flSheet.Cells.Find("因子", LookAt:=xlWhole) 'まずは「因子」のセルを探すことでFL表であるかどうかの手掛かりとする
    If obj Is Nothing Then
        Exit Function
    End If
    
    titleRow = obj.row
    factorCol = obj.Column
    If Not flSheet.Cells(titleRow, factorCol + 1).Value Like "水準*" Then
        Exit Function
    End If
    
    factorNum = 0
    emptyFlg = False
    
    For i = titleRow + 1 To MaxRow
        content = flSheet.Cells(i, factorCol).Value
        If content = "" Then
            emptyFlg = True
        Else
            If emptyFlg Then
                MsgBox "因子列の途中に空のセルがあります。空のセル以下の行を無視します。"
                Exit For
            End If
            factorNum = factorNum + 1
            ReDim Preserve factorNames(factorNum - 1)
            factorNames(factorNum - 1) = content
        End If
    Next i
    
    ReDim levelLists(factorNum - 1)
    For i = 0 To factorNum - 1
        levelNum = 0
        emptyFlg = False
        ReDim levelNames(0)
        For j = factorCol + 1 To MaxCol
            content = flSheet.Cells(i + titleRow + 1, j).Value
            If content = "" Then
                emptyFlg = True
            Else
                If emptyFlg Then
                    MsgBox "水準列の途中に空のセルがあります。空のセルより右の列を無視します。"
                    Exit For
                End If
                levelNum = levelNum + 1
                ReDim Preserve levelNames(levelNum - 1)
                levelNames(levelNum - 1) = content
                levelLists(i) = levelNames
            End If
        Next j
    Next i

    FLTable2array = True
End Function

' Toolの実行結果の入っている文字列を配列に変換する
Function textTable2array(tupleStr As String, delimiter As String, ByRef paramNames() As String, ByRef tuples()) As Boolean
    textTable2array = False
    
    Dim lines() As String
    Dim tuple() As String
    Dim i As Long
    Dim j As Long
    Dim testCaseNum As Long
    
    tupleStr = Replace(tupleStr, vbCrLf, vbLf)
    lines = Split(tupleStr, vbLf)
    
    paramNames = Split(lines(LBound(lines)), delimiter) ' 1行目は因子名
    
    testCaseNum = 0
    For i = LBound(lines) + 1 To UBound(lines) ' 1行目は因子名として既に取り込んだので読み飛ばす
        If Not lines(i) = "" Then
            tuple = Split(lines(i), delimiter)
            testCaseNum = testCaseNum + 1
            ReDim Preserve tuples(testCaseNum - 1)
            tuples(testCaseNum - 1) = tuple
        End If
    Next i
            
    textTable2array = True
End Function

' テストケースの入っているシートの情報を配列に変換する
Function testCaseSheet2array(testCaseSheet As Worksheet, ByRef paramNames() As String, ByRef tuples()) As Boolean
    testCaseSheet2array = False
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim titleRow As Long
    Dim idCol As Long
    Dim factorNum As Long
    Dim i As Long
    Dim j As Long
    Dim obj As Object
    Dim emptyFlg As Boolean
    Dim content As String
    
    Call getMaxRowAndCol(testCaseSheet, MaxRow, MaxCol)
    
    Set obj = testCaseSheet.Cells.Find("ID", LookAt:=xlWhole) 'まずは「ID」のセルを探すことで形式が変更されていないかどうかの手掛かりとする
    If obj Is Nothing Then
        Exit Function
    End If
    
    titleRow = obj.row
    idCol = obj.Column
    
    Dim tuple() As String
    Dim testCaseNum As Long
    
    factorNum = 0
    emptyFlg = False
    ReDim paramNames(MaxCol - idCol - 1)
    For j = idCol + 1 To MaxCol
        factorNum = factorNum + 1
        paramNames(factorNum - 1) = testCaseSheet.Cells(titleRow, j).Value ' 1行目は因子名
    Next j
    
    testCaseNum = 0
    For i = titleRow + 1 To MaxRow
        content = testCaseSheet.Cells(i, idCol).Value
        If content = "" Then
            emptyFlg = True
        Else
            If emptyFlg Then
                MsgBox "ID列の途中に空のセルがあります。空のセル以下の行を無視します。"
                Exit For
            End If
            factorNum = 0
            ReDim tuple(MaxCol - idCol - 1)
            For j = idCol + 1 To MaxCol
                factorNum = factorNum + 1
                tuple(factorNum - 1) = testCaseSheet.Cells(i, j).Value ' 1行目は因子名
            Next j
            testCaseNum = testCaseNum + 1
            ReDim Preserve tuples(testCaseNum - 1)
            tuples(testCaseNum - 1) = tuple
        End If
    Next i
            
    testCaseSheet2array = True
End Function

' 定義における水準名の重複を調べるために水準名の出現回数の辞書を作る
' 一般に因子が違えば水準名は重複して良いが、Alloyでは不都合が起きる
Sub generateDicDuplication(factorNames, levelLists, ByRef dicDuplication)
    Dim factorNum As Long
    Dim levelNum As Long
    Dim levelName As String
    Dim newVal As Long
    Set dicDuplication = CreateObject("Scripting.Dictionary")
        
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            levelName = levelLists(factorNum)(levelNum)
            If dicDuplication.exists(levelName) Then
                newVal = dicDuplication.Item(levelName) + 1
                dicDuplication.Remove (levelName)
                dicDuplication.Add levelName, newVal
            Else
                dicDuplication.Add levelName, 1
            End If

        Next levelNum
    Next factorNum
    
End Sub

' シートで使われている最大の行とカラムを求める
Sub getMaxRowAndCol(wkSheet As Worksheet, ByRef MaxRow As Long, ByRef MaxCol As Long)
    Dim lDummy As Long: lDummy = wkSheet.UsedRange.row ' 一度 UsedRange を使うと最終セルが補正されるようだ
    Dim i As Long
    Dim j As Long
    MaxRow = wkSheet.Cells.SpecialCells(xlLastCell).row
    MaxCol = wkSheet.Cells.SpecialCells(xlLastCell).Column
    
    If wkSheet.Cells.SpecialCells(xlLastCell).MergeCells Then ' セル結合がある場合に対応して最終セルの位置を修正する
        i = MaxRow
        j = MaxCol
        MaxRow = MaxRow + wkSheet.Cells(i, j).MergeArea.Rows.Count - 1
        MaxCol = MaxCol + wkSheet.Cells(i, j).MergeArea.Columns.Count - 1
    End If
End Sub

' 指定した名前のシートが存在するか確認します。
Function ExistsWorksheet(wb As Workbook, name As String)
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        If ws.name = name Then
            ' 存在する
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    
    ' 存在しない
    ExistsWorksheet = False
End Function

' EXCELの名前定義のリストをシートに書き出す。
' 本ツールでは使用していない。デバッグ用
Sub nameList()
    Dim nm As name
    Dim i As Long
    Dim ws As Worksheet
    Dim shortName As String
    Worksheets.Add
    'Set ws = ActiveWorkbook.Worksheets(controlSheetName)
    i = 1
    Cells(i, 1) = "name"
    Cells(i, 2) = "Value"
    Cells(i, 3) = "Row"
    Cells(i, 4) = "Column"
    Cells(i, 5) = "シート名"
    Cells(i, 6) = "シート内name"
    Cells(i, 7) = "Parent.name"
    For Each nm In ActiveWorkbook.Names
    'For Each nm In ws.Names　必ずしもシート内限定で名前を付けてくれているとは限らないので
        i = i + 1
        shortName = Replace(nm.name, "'" & Range(nm).Worksheet.name & "'!", "")
        shortName = Replace(shortName, Range(nm).Worksheet.name & "!", "")
        Cells(i, 1) = nm.name
        Cells(i, 2) = "'" & nm.Value
        Cells(i, 3) = Range(nm).row
        Cells(i, 4) = Range(nm).Column
        Cells(i, 5) = Range(nm).Worksheet.name
        Cells(i, 6) = shortName
        Cells(i, 7) = nm.Parent.name
    Next
End Sub

