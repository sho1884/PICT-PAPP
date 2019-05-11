Attribute VB_Name = "全組み合わせ・多項間禁則表"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' 因子・水準の全組み合わせ生成
Sub outputCombination(ws As Worksheet, row As Long, testCase() As String, factorN As Integer, levelLists())
    Dim levelN As Integer

    If factorN > UBound(levelLists) - LBound(levelLists) + 1 - 1 Then
        Call outputTuple(ws, row, testCase)
        row = row + 1
        Exit Sub
    End If
    For levelN = LBound(levelLists(factorN)) To UBound(levelLists(factorN))
        testCase(factorN) = levelLists(factorN)(levelN)
        '再帰呼出し
        Call outputCombination(ws, row, testCase, factorN + 1, levelLists)
    Next
End Sub

' 生成済みの因子・水準の全組み合わせをシートに書き込む
Sub outputTuple(ws As Worksheet, row As Long, testCase() As String)
    Dim j As Integer
    ' オフセット分＋タイトルの1行を空ける
    Worksheets(controlSheetName).Range("値書式").Copy
    ws.Cells(row + offsetRows + 1, offsetColumns + 1).Value = "#" & row ' IDとしてシーケンシャル番号
    ws.Cells(row + offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    ' オフセット分＋IDの1列を空ける
    For j = LBound(testCase) To UBound(testCase)
        ws.Cells(row + offsetRows + 1, offsetColumns + j + 2).Value = testCase(j)
        ws.Cells(row + offsetRows + 1, offsetColumns + j + 2).PasteSpecial (xlPasteFormats)
    Next
End Sub

' 因子・水準表から全組合せのシートを生成する
Function fillInTupleSheets(srcBook As Workbook) As Boolean
    fillInTupleSheets = False

    Dim tuplelSheet As Worksheet
    Dim j As Long

    Dim factorNames() As String
    Dim levelLists()
    Dim testCase() As String

    ' 新たなシートを生成する
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = tuplelSheetName
    Set tuplelSheet = srcBook.Sheets(tuplelSheetName)
    If tuplelSheet Is Nothing Then ' 結局組合せ生成するシートが特定できない
        MsgBox "全組合せ生成するシートの名前として「" & tuplelSheetName & "」が指定されていますが、その名前のシートが生成できません。"
        Exit Function
    End If

    ' 因子・水準のシートを特定して因子・水準を配列に読み込む
    If Not FLTableSheet2array(srcBook, factorNames, levelLists) Then
        Exit Function
    End If
    
    ' タイトル行となる因子名の行を書き込み
    tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 1).Value = "ID"
    Worksheets(controlSheetName).Range("項目タイトル書式").Copy
    tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 1).PasteSpecial (xlPasteFormats)
    For j = LBound(factorNames) To UBound(factorNames)
        tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 2).Value = factorNames(j)
        tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 2).PasteSpecial (xlPasteFormats)
    Next j
    
    ' 水準の全組み合わせの出力
    ReDim testCase(UBound(levelLists) - LBound(levelLists) + 1 - 1)
    Call outputCombination(tuplelSheet, 1, testCase, 0, levelLists)

    fillInTupleSheets = True
End Function

' 多項間禁則を定義するための禁則マトリクスシートを生成する
Function createKinsokuMatrix(srcBook As Workbook) As Boolean
    createKinsokuMatrix = False
                
    Dim kinsokuMatrixSheetName As String
    Dim kinsokuMatrix As Worksheet
    Dim FLtblSheet As Worksheet
    Dim rng As Range
    Dim totalLevelNum As Long: totalLevelNum = 0
    Dim i As Long
    Dim j As Long
    Dim conditionNum As Long
    Dim constraintLevelsNum As Long
    
    Dim constraintLevels
    Dim levelLists()
    Dim allLevelLists()
    Dim conditionComb() As String
    Dim conditionCombNum As Long

    ' 因子・水準のシートを特定して因子・水準を配列に読み込む
    If Not FLTableSheet2array(srcBook, publicFactorNames, allLevelLists) Then
        Exit Function
    End If
    
    ' 条件因子と被制約因子を選択するリストダイアログを表示
    Call SelectFactors.doModal
    Unload SelectFactors 'ユーザーフォームはここで閉じる
    
    If constraintFactors = "" Then
        MsgBox "被制約因子が選択されなかったので、処理を中止します。"
        Exit Function
    End If
    
    If Not Not conditionFactors Then
        conditionNum = UBound(conditionFactors) - LBound(conditionFactors) + 1 ' 条件因子の個数
    Else
        conditionNum = 0 ' 動的配列の割り当てが未設定
    End If
    
    If conditionNum < 2 Then
        MsgBox "条件因子が2つ以上選択されなかったので、処理を中止します。１つで良いのであれば総当たり表がその目的に使えます。"
        Exit Function
    End If
    
    ' 条件因子と被制約因子の重複チェック
    For i = LBound(conditionFactors) To UBound(conditionFactors)
        If conditionFactors(i) = constraintFactors Then
            MsgBox "被制約因子として選択した[" & constraintFactors & "]が条件因子にも含まれていたので、処理を中止します。"
            Exit Function
        End If
    Next i
    
    ' 選択された条件因子のみの水準列を抽出
    ReDim levelLists(conditionNum - 1)
    conditionCombNum = 1
    For i = LBound(conditionFactors) To UBound(conditionFactors)
        For j = LBound(publicFactorNames) To UBound(publicFactorNames)
            If publicFactorNames(j) = conditionFactors(i) Then
                levelLists(i) = allLevelLists(j)
                conditionCombNum = conditionCombNum * (UBound(levelLists(i)) - LBound(levelLists(i)) + 1)
            ElseIf publicFactorNames(j) = constraintFactors Then
                constraintLevels = allLevelLists(j)
            End If
        Next j
    Next i
    
    ' 多項間禁則表のまだ使われていない名前を作る（空いている番号をつける）
    For i = 1 To kinsokuMatrixSheetMax
        kinsokuMatrixSheetName = kinsokuMatrixSheetBaseName & "(" & i & ")"
        If Not ExistsWorksheet(srcBook, kinsokuMatrixSheetName) Then
            Exit For
        End If
    Next i
    
    ' 現実にはありそうにないことなので、禁則表が増え過ぎたらやめる。未整理で削除忘れのためと思われるので。
    If kinsokuMatrixSheetName = "" Then
        MsgBox "多項の禁則関係定義用シートの数が最大数を超えました。処理を中止します。"
        Exit Function
    Else
        ' 新たなシートを生成する
        srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
        ActiveSheet.name = kinsokuMatrixSheetName
        Set kinsokuMatrix = srcBook.Sheets(kinsokuMatrixSheetName)
    End If
    If kinsokuMatrix Is Nothing Then ' 結局禁則定義用シートが特定できない
        MsgBox "多項の禁則関係定義用シートの名前として「" & kinsokuMatrix & "」が指定されていますが、その名前のシートが存在しません。"
        Exit Function
    End If

    ' 条件因子の行を出力
    kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + 1).Value = "ID"
    Worksheets(controlSheetName).Range("因子書式").Copy
    kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    For j = LBound(conditionFactors) To UBound(conditionFactors)
        kinsokuMatrix.Cells(offsetRows + 1, j + offsetColumns + 2).Value = conditionFactors(j)
        kinsokuMatrix.Cells(offsetRows + 1, j + offsetColumns + 2).PasteSpecial (xlPasteFormats)
    Next j
    
    ' 水準の全組み合わせの出力
    ReDim conditionComb(UBound(levelLists) - LBound(levelLists) + 1 - 1)
    Call outputCombination(kinsokuMatrix, 1, conditionComb, 0, levelLists) ' 要するに全組み合わせ生成処理の流用
    
    ' 列の幅を自動調整
    For j = LBound(conditionFactors) To UBound(conditionFactors)
        kinsokuMatrix.Columns(j + offsetColumns + 2).EntireColumn.AutoFit
    Next j

    ' 全組み合わせ生成処理の流用による不都合を修正
    ' 被制約因子名を記述する行を１行挿入する
    kinsokuMatrix.Rows(offsetRows).Insert
    ' IDの列を削除する
    kinsokuMatrix.Columns(offsetColumns + 1).Delete
    
    ' 被制約因子側の情報を追記
    For j = LBound(constraintLevels) To UBound(constraintLevels)
        kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + conditionNum + j + 1).Value = constraintFactors
        Worksheets(controlSheetName).Range("因子書式").Copy
        kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + conditionNum + j + 1).PasteSpecial (xlPasteFormats)

        kinsokuMatrix.Cells(offsetRows + 2, offsetColumns + conditionNum + j + 1).Value = constraintLevels(j)
        Worksheets(controlSheetName).Range("水準書式").Copy
        kinsokuMatrix.Cells(offsetRows + 2, offsetColumns + conditionNum + j + 1).PasteSpecial (xlPasteFormats)
        kinsokuMatrix.Columns(offsetColumns + conditionNum + j + 1).ColumnWidth = 3
    Next j

    ' 禁則設定で編集可能な範囲やセル色が変わる設定をする
    constraintLevelsNum = UBound(constraintLevels) - LBound(constraintLevels) + 1
    
    kinsokuMatrix.Range(Cells(offsetRows + 1, offsetColumns + conditionNum + 1), Cells(offsetRows + 2, offsetColumns + conditionNum + constraintLevelsNum)).Orientation = xlDownward
    Set rng = kinsokuMatrix.Range(Cells(offsetRows + 3, offsetColumns + conditionNum + 1), Cells(offsetRows + conditionCombNum + 2, offsetColumns + conditionNum + constraintLevelsNum))
    Worksheets(controlSheetName).Range("値書式").Copy
    rng.PasteSpecial (xlPasteFormats)
    Call kinsokuFormatCollectionsAdd(rng)

    kinsokuMatrix.Cells.Locked = True ' 全セルをロック
    On Error Resume Next
    rng.SpecialCells(Type:=xlCellTypeBlanks).Locked = False ' ロックを解除
    On Error GoTo 0
    kinsokuMatrix.Protect Password:=protectPassword 'シートの保護
    
    createKinsokuMatrix = True
End Function

