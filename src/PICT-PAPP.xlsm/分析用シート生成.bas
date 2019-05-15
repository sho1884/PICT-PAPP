Attribute VB_Name = "分析用シート生成"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' いくつかの分析用シートの生成
Function analysis(srcBook As Workbook, paramNames() As String, tuples()) As Boolean
    analysis = False
                
    Dim pairListSheet As Worksheet
    Dim dicFL As Object
    
    If pairListFlg Then
        ' 2因子間の組み合わせ出現数分析用にペア・リスト生成用のシートを用意する
        srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
        ActiveSheet.name = pairListSheetName
        Set pairListSheet = srcBook.Sheets(pairListSheetName)
        
        ' 因子水準の組み合わせからテストIDを引ける連想配列を作る
        ' 同時にペア・リスト生成用のシートに情報を書き込む
        Call makeFLIDDictionary(paramNames(), tuples, dicFL, pairListSheet)
    Else
        ' 因子水準の組み合わせからテストIDを引ける連想配列を作る
        ' ペア・リスト生成用のシートは作成しない
        Call makeFLIDDictionary(paramNames(), tuples, dicFL)
    End If
    
    ' Toolの出力結果を複製した総当たり表にマップし、同時にカバレッジ状態を示すシートを生成する
    If Not fillInRoundRobinTable(srcBook, paramNames(), tuples, dicFL) Then
        MsgBox "総当たり表へのテストIDマッピング処理に失敗しました。"
    End If

    analysis = True
End Function

' 総当たり表を複製した"IDマッピング済み総当たり表"シートを生成し、そこにToolの結果として得たテストIDや数を追記する。
' （未だ総当たり表が存在しない場合は、勝手に前処理として生成する。）
Function fillInRoundRobinTable(srcWorkbook As Workbook, paramNames() As String, tuples(), dicFL As Object) As Boolean
    fillInRoundRobinTable = False
    
    Dim destSheet As Worksheet
    Dim testCaseNum As Long
    Dim paramNum As Long
    Dim flKey As String
    Dim pairCount As Long ' 禁則が無い場合の全Pairの数
    Dim someCount As Long ' Pairwiseに少なくとも1回は出現したPairの数
    Dim kinsokuCount As Long ' 総当たり表において禁則と明示された全Pairの数
    Dim uncertainCount As Long ' 総当たり表において禁則と明示されていないが、Toolが出力しなかったPairの数
    
    pairCount = 0
    someCount = 0
    kinsokuCount = 0
    uncertainCount = 0
    
    ' 総当たり表が存在しない場合は、勝手に前処理として生成する
    If Not ExistsWorksheet(srcWorkbook, roundRobinSheetName) Then
        If createRoundRobinTable(ThisWorkbook) Then
            MsgBox "総当たり表を生成処理しました。"
        Else
            MsgBox "総当たり表の生成処理に失敗しました。処理を中止します。"
            Exit Function
        End If
    End If
    
    ' 総当たり表を複製した"IDマッピング済み総当たり表"シートを生成
    srcWorkbook.Worksheets(roundRobinSheetName).Copy Before:=srcWorkbook.Worksheets(roundRobinSheetName)
    ActiveSheet.name = mappedRoundRobinSheetName
    Set destSheet = srcWorkbook.Worksheets(mappedRoundRobinSheetName)
    destSheet.Unprotect Password:=protectPassword ' 特に必要ないので保護しない

    ' シートに書き込み
    Dim i As Long
    Dim j As Long
    Dim MaxRow As Long
    Dim MaxCol As Long
    
    Call getMaxRowAndCol(destSheet, MaxRow, MaxCol)
    
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
    
    If Not (MaxRow - levelRow = MaxCol - levelCol) Then
        MsgBox "総当たり表の縦横のサイズがあっていません。ゴミデータ等が入っていると思われるので、処理を中止します。"
        Exit Function
    End If
    
    For i = levelRow + 1 To MaxRow
        For j = levelCol + 1 To MaxCol
            If i - levelRow > j - levelCol Then ' 対角線より左下
                flKey = destSheet.Cells(factorRow, j).Value & ":" & destSheet.Cells(levelRow, j).Value & "*" & destSheet.Cells(i, factorCol).Value & ":" & destSheet.Cells(i, levelCol).Value
            Else ' 対角線より右上
                flKey = destSheet.Cells(i, factorCol).Value & ":" & destSheet.Cells(i, levelCol).Value & "*" & destSheet.Cells(factorRow, j).Value & ":" & destSheet.Cells(levelRow, j).Value
            End If
            If dicFL.exists(flKey) Then
                pairCount = pairCount + 1
                someCount = someCount + 1
                If destSheet.Cells(i, j).Value = "×" Then
                    kinsokuCount = kinsokuCount + 1
                    MsgBox "禁則指定した組み合わせがテストケースの中に出現しています。制約式の自動生成などの正しい手順を踏んだか確認してください。（シートの" & i & "行目" & j & "列）"
                Else
                    If i - levelRow > j - levelCol Then ' 対角線より左下
                        destSheet.Cells(i, j).Value = dicFL.Item(flKey)
                    Else ' 対角線より右上
                        IDs = Split(dicFL.Item(flKey), ",")
                        destSheet.Cells(i, j).Value = UBound(IDs) - LBound(IDs) + 1
                    End If
                End If
            Else
                Select Case destSheet.Cells(i, j).Value
                    Case "―" ' 対角線上の自身との組み合わせで、意味が無いので無視
                    Case "×" ' 禁則指定されているならば問題ない
                        pairCount = pairCount + 1
                        kinsokuCount = kinsokuCount + 1
                    Case "" ' 禁則指定されていないならば問題なので、出現していない原因を究明する必要がある。
                        pairCount = pairCount + 1
                        uncertainCount = uncertainCount + 1
                        destSheet.Cells(i, j).Value = "?"
'                        destSheet.Cells(i, j).Interior.Color = RGB(255, 255, 0) ' 黄色
                    Case Else
                        pairCount = pairCount + 1
                        someCount = someCount + 1
                        MsgBox "総当たり表に意味不明の入力値が入っています。正しい手順を踏んだか確認してください。（シートの" & i & "行目" & j & "列）"
                End Select
            End If
        Next j
    Next i
    
    ' 数の整合性チェック
    someCount = someCount / 2
    kinsokuCount = kinsokuCount / 2
    uncertainCount = uncertainCount / 2
    pairCount = pairCount / 2
    If Not (someCount + kinsokuCount + uncertainCount = pairCount) Then
        MsgBox "総当たり表の禁則や無効の設定が不適切なため、各Pair数の合計数について不整合が起こっています。確認してください。"
    End If
    
    ' 網羅率シート生成
    If Not fillInCoverageSheet(srcWorkbook, someCount, kinsokuCount, uncertainCount, pairCount, UBound(tuples) - LBound(tuples) + 1) Then
        MsgBox "網羅率のシート生成に失敗しました。"
    End If
    
    destSheet.Activate
    
    fillInRoundRobinTable = True
End Function
    
' 網羅率シート生成と書き込み
Function fillInCoverageSheet(srcWorkbook As Workbook, someCount, kinsokuCount, uncertainCount, pairCount, testcaseCount) As Boolean
    fillInCoverageSheet = False
    
    ' 網羅率の出力シートを用意する
    Dim coverageSheet As Worksheet
    srcWorkbook.Worksheets.Add Before:=Worksheets(mappedRoundRobinSheetName)
    ActiveSheet.name = coverageSheetName
    Set coverageSheet = srcWorkbook.Sheets(coverageSheetName)
    
    Worksheets(controlSheetName).Range("項目タイトル書式").Copy
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 1).Value = "A"
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 1).Value = "B"
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 1).Value = "B'"
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 1).Value = "C"
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 1).Value = "D"
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 1).Value = "E"
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 1).Value = "F"
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 1).Value = "G"
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    coverageSheet.Columns(offsetColumns + 1).EntireColumn.AutoFit
    
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 2).Value = "Tool出力結果のテスト項目数"
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 2).Value = "組合せ網羅率(％) ( B = C / E × 100)"
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 2).Value = "Dの組み合わせも全て禁則である場合の組合せ網羅率(％) ( B' = C / (E-D) × 100)"
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 2).Value = "Tool出力結が網羅する2因子間組合せ数"
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 2).Value = "Tool出力結果が網羅せず、しかも禁則と明示設定されてもいない2因子間組合せ数"
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 2).Value = "禁則と明示設定されているものを除いた2因子間組合せ数 ( E = G - F )"
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 2).Value = "禁則と明示設定されている2因子間組み合せ数"
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 2).Value = "禁則が無いと仮定した場合の2因子間組み合せ数"
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Columns(offsetColumns + 2).EntireColumn.AutoFit
    
    Worksheets(controlSheetName).Range("値書式").Copy
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 3).Value = testcaseCount ' Tool出力結の項目数
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 3).Value = someCount / (pairCount - kinsokuCount) * 100 ' 組合せ網羅率(％) ( B = C / E × 100)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 3).Value = someCount / (pairCount - kinsokuCount - uncertainCount) * 100 ' Dの組み合わせも全て禁則である場合の組合せ網羅率(％) ( B' = C / (E-D) × 100)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 3).Interior.Color = RGB(255, 255, 0) ' 黄色
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 3).Value = someCount ' Tool出力結が網羅する2因子間組合せ数
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 3).Value = uncertainCount ' Tool出力結が網羅せず、しかも禁則と明示設定されてもいない2因子間組合せ数
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 3).Interior.Color = RGB(255, 255, 0) ' 黄色
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 3).Value = pairCount - kinsokuCount ' 禁則と明示設定されているものを除いた2因子間組合せ数 ( E = G - F )
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 3).Value = kinsokuCount ' 禁則と明示設定されている2因子間組み合せ数
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 3).Value = pairCount ' 禁則が無いと仮定した場合の2因子間組み合せ数
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Columns(offsetColumns + 3).EntireColumn.AutoFit
    
    fillInCoverageSheet = True
End Function

