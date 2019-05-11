Attribute VB_Name = "総当たり表"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' 因子・水準表から総当たり表を生成する
Function createRoundRobinTable(srcBook As Workbook) As Boolean
    createRoundRobinTable = False
                
    Dim roundRobinSheet As Worksheet
    Dim FLtblSheet As Worksheet
    Dim rng As Range
    Dim totalLevelNum As Long: totalLevelNum = 0
    Dim i As Long
    Dim j As Long
    
    Dim factorNames() As String
    Dim levelLists()
    Dim testCase() As String

    ' 総当たり表シートを新たに生成する
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = roundRobinSheetName
    Set roundRobinSheet = srcBook.Sheets(roundRobinSheetName)
    If roundRobinSheet Is Nothing Then ' 結局総当たり表シートが特定できない
        MsgBox "総当たり表シートの名前として「" & roundRobinSheetName & "」が指定されていますが、その名前のシートが存在しません。"
        Exit Function
    End If
    
    ' 因子・水準のシートを特定する
    If ExistsWorksheet(srcBook, FLtblSheetName) Then
        Set FLtblSheet = srcBook.Sheets(FLtblSheetName)
    End If
    If FLtblSheet Is Nothing Then ' 結局因子・水準表シートが見つからない
        MsgBox "因子・水準表の名前として「" & FLtblSheetName & "」が指定されていますが、その名前のシートが存在しません。"
        Exit Function
    End If
    
    ' 因子・水準を読み込む
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "因子・水準の解析に失敗しました。処理を中止します。"
        Exit Function
    End If

    ' 総当たり表を出力
    Dim factorNum As Long
    Dim levelNum As Long
    Dim factorRow As Long
    Dim levelRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    
    factorRow = offsetRows + 1
    levelRow = offsetRows + 2
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    
    ' roundRobinSheet.Unprotect Password:=protectPassword ' 生成したところなので保護されていない
    
    Worksheets(controlSheetName).Range("項目タイトル書式").Copy
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            totalLevelNum = totalLevelNum + 1
            j = totalLevelNum + levelCol
            roundRobinSheet.Cells(factorRow, j).Value = factorNames(factorNum)
            Worksheets(controlSheetName).Range("因子書式").Copy
            roundRobinSheet.Cells(factorRow, j).PasteSpecial (xlPasteFormats)

            roundRobinSheet.Cells(levelRow, j).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("水準書式").Copy
            roundRobinSheet.Cells(levelRow, j).PasteSpecial (xlPasteFormats)
            roundRobinSheet.Columns(j).ColumnWidth = 3
        Next levelNum
    Next factorNum
    roundRobinSheet.Range(Cells(factorRow, levelCol + 1), Cells(levelRow, j)).Orientation = xlDownward
    Set rng = roundRobinSheet.Range(Cells(levelRow + 1, levelCol + 1), Cells(levelRow + totalLevelNum, levelCol + totalLevelNum))
    Worksheets(controlSheetName).Range("値書式").Copy
    rng.PasteSpecial (xlPasteFormats)
    Call kinsokuFormatCollectionsAdd(rng)
    
    roundRobinSheet.Cells.Locked = True ' 全セルをロック
    On Error Resume Next
    rng.SpecialCells(Type:=xlCellTypeBlanks).Locked = False ' ロックを解除
    On Error GoTo 0
    
    i = levelRow
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            i = i + 1
            roundRobinSheet.Cells(i, factorCol).Value = factorNames(factorNum)
            Worksheets(controlSheetName).Range("因子書式").Copy
            roundRobinSheet.Cells(i, factorCol).PasteSpecial (xlPasteFormats)
            
            roundRobinSheet.Cells(i, levelCol).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("水準書式").Copy
            roundRobinSheet.Cells(i, levelCol).PasteSpecial (xlPasteFormats)
        Next levelNum
    Next factorNum
    roundRobinSheet.Columns(factorCol).EntireColumn.AutoFit
    roundRobinSheet.Columns(levelCol).EntireColumn.AutoFit
    
    Dim r As Long
    Dim c As Long
    r = levelRow
    c = levelCol
    For factorNum = LBound(factorNames) To UBound(factorNames)
        levelNum = UBound(levelLists(factorNum)) - LBound(levelLists(factorNum)) + 1
        For i = 1 To levelNum
            For j = 1 To levelNum
                roundRobinSheet.Cells(i + r, j + c).Value = "―"
                Worksheets(controlSheetName).Range("無効書式").Copy
                roundRobinSheet.Cells(i + r, j + c).PasteSpecial (xlPasteFormats)
            Next j
        Next i
        r = r + levelNum
        c = c + levelNum
    Next factorNum
    
    roundRobinSheet.Protect Password:=protectPassword 'シートの保護
    
    createRoundRobinTable = True
End Function

' 総当たり表のシートに、「禁則を表す×が入力されたときに、そのセルの背景色などを自動変更する」ルールを設定する
Sub kinsokuFormatCollectionsAdd(r As Range)
    Dim f   As FormatCondition
    
    '// 条件付き書式の追加（セルに×が入力された場合）
    Set f = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="×")
    '// フォント太字、文字色、背景色
    f.Font.Bold = Worksheets(controlSheetName).Range("禁則書式").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("禁則書式").Font.Color
    f.Interior.Color = Worksheets(controlSheetName).Range("禁則書式").Interior.Color
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("禁則書式").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("禁則書式").Borders(xlEdgeTop).Weight
    
    '// 条件付き書式の追加（セルに？が入力された場合）
    Set f = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="？")
    '// フォント太字、文字色、背景色
    f.Font.Bold = Worksheets(controlSheetName).Range("値書式").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("値書式").Font.Color
    f.Interior.Color = RGB(255, 0, 0)  ' 赤色
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("値書式").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("値書式").Borders(xlEdgeTop).Weight
    
    '// 条件付き書式の追加（セルに?が入力された場合）
    Set f = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="?")
    '// フォント太字、文字色、背景色
    f.Font.Bold = Worksheets(controlSheetName).Range("値書式").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("値書式").Font.Color
    f.Interior.Color = RGB(255, 255, 0) ' 黄色
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("値書式").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("値書式").Borders(xlEdgeTop).Weight
End Sub

