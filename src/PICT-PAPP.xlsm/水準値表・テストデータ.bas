Attribute VB_Name = "水準値表・テストデータ"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' 因子・水準表から因子・水準・水準値表を生成する
Function createFLLVSheet(srcBook As Workbook) As Boolean
    createFLLVSheet = False
                
    Dim FLLVSheet As Worksheet
    Dim FLtblSheet As Worksheet
    Dim rng As Range
    Dim i As Long
    
    Dim factorNames() As String
    Dim levelLists()

    ' 新たな因子・水準・水準値設定表シートを生成する
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = FLLVSheetName
    Set FLLVSheet = srcBook.Sheets(FLLVSheetName)
    
    If FLLVSheet Is Nothing Then ' 結局総当たり表シートが特定できない
        MsgBox "因子・水準・水準値設定表シートの名前として「" & FLLVSheetName & "」が指定されていますが、その名前のシートが存在しません。"
        Exit Function
    End If
    
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
    
    ' 因子・水準を読み込む
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "因子・水準の解析に失敗しました。処理を中止します。"
        Exit Function
    End If

    ' 表を出力
    Dim factorNum As Long
    Dim levelNum As Long
    Dim startRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    
    startRow = offsetRows + 1
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    
    i = startRow
    Worksheets(controlSheetName).Range("項目タイトル書式").Copy
    FLLVSheet.Cells(i, factorCol).Value = "因子"
    FLLVSheet.Cells(i, factorCol).PasteSpecial (xlPasteFormats)
    FLLVSheet.Cells(i, levelCol).Value = "水準"
    FLLVSheet.Cells(i, levelCol).PasteSpecial (xlPasteFormats)
    FLLVSheet.Cells(i, levelCol + 1).Value = "水準値"
    FLLVSheet.Cells(i, levelCol + 1).PasteSpecial (xlPasteFormats)
    FLLVSheet.Cells(i, levelCol + 2).Value = "備考"
    FLLVSheet.Cells(i, levelCol + 2).PasteSpecial (xlPasteFormats)
    
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            i = i + 1
            FLLVSheet.Cells(i, factorCol).Value = factorNames(factorNum)
            Worksheets(controlSheetName).Range("因子書式").Copy
            FLLVSheet.Cells(i, factorCol).PasteSpecial (xlPasteFormats)
            
            FLLVSheet.Cells(i, levelCol).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("水準書式").Copy
            FLLVSheet.Cells(i, levelCol).PasteSpecial (xlPasteFormats)
        
            FLLVSheet.Cells(i, levelCol + 1).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("値書式").Copy
            FLLVSheet.Cells(i, levelCol + 1).PasteSpecial (xlPasteFormats)
        
            FLLVSheet.Cells(i, levelCol + 2).Value = ""
            Worksheets(controlSheetName).Range("値書式").Copy
            FLLVSheet.Cells(i, levelCol + 2).PasteSpecial (xlPasteFormats)
        Next levelNum
    Next factorNum
    FLLVSheet.Columns(factorCol).EntireColumn.AutoFit
    FLLVSheet.Columns(levelCol).EntireColumn.AutoFit
    FLLVSheet.Columns(levelCol + 1).EntireColumn.AutoFit
    FLLVSheet.Columns(levelCol + 2).ColumnWidth = 80
    
    Set rng = FLLVSheet.Range(Cells(startRow + 1, levelCol + 1), Cells(i, levelCol + 2))
    FLLVSheet.Cells.Locked = True ' 全セルをロック
    On Error Resume Next
    rng.Locked = False ' ロックを解除
    Set rng = FLLVSheet.Range(Cells(startRow + 1, levelCol + 1), Cells(i, levelCol + 1))
    Call levelValFormatCollectionsAdd(rng)
    On Error GoTo 0
    FLLVSheet.Protect Password:=protectPassword 'シートの保護
    
    createFLLVSheet = True
End Function

' 因子・水準・水準値表のシートで、水準値が空になっている場合に、そのセルの背景色を赤にするルールを設定する
Sub levelValFormatCollectionsAdd(r As Range)
    Dim f   As FormatCondition
    
    '// 条件付き書式の追加（セルに？が入力された場合）
    Set f = r.FormatConditions.Add(Type:=xlBlanksCondition)
    '// フォント太字、文字色、背景色
    f.Font.Bold = Worksheets(controlSheetName).Range("値書式").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("値書式").Font.Color
    f.Interior.Color = RGB(255, 0, 0)  ' 赤色
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("値書式").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("値書式").Borders(xlEdgeTop).Weight
End Sub

' 因子・水準・水準値設定表シートの情報を使って、因子・水準の対の名前から水準値を引ける連想配列を作る。
' セルの値を持っていく場合のコードがコメントアウトされている。現状、セルのアドレスを持っていく仕様になっている。
Function makeFLLVDictionary(FLLVSheet As Worksheet, ByRef dicFLLV As Object) As Boolean
    makeFLLVDictionary = False
    
    Dim flKey As String
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim i As Long
    Dim j As Long
    Dim titleRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    Dim factorName As String
    Dim levelName As String
    Dim levelVal As String
    Dim levelValAddress As String
    
    FLLVSheet.Unprotect Password:=protectPassword
    Call getMaxRowAndCol(FLLVSheet, MaxRow, MaxCol)
    FLLVSheet.Protect Password:=protectPassword
    
    titleRow = offsetRows + 1
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    
    If Not FLLVSheet.Cells(titleRow, factorCol).Value = "因子" Then
        MsgBox "因子・水準・水準値設定表シートの形式が想定外で認識できません。処理を中止します。"
        Exit Function
    End If
    If Not FLLVSheet.Cells(titleRow, levelCol).Value = "水準" Then
        MsgBox "因子・水準・水準値設定表シートの形式が想定外で認識できません。処理を中止します。"
        Exit Function
    End If
    If Not FLLVSheet.Cells(titleRow, levelCol + 1).Value = "水準値" Then
        MsgBox "因子・水準・水準値設定表シートの形式が想定外で認識できません。処理を中止します。"
        Exit Function
    End If
    
    Set dicFLLV = CreateObject("Scripting.Dictionary")

    For i = titleRow + 1 To MaxRow
        factorName = FLLVSheet.Cells(i, factorCol).Value
        levelName = FLLVSheet.Cells(i, levelCol).Value
'        levelVal = FLLVSheet.Cells(i, levelCol + 1).Value
        levelValAddress = "=" & FLLVSheetName & "!" & FLLVSheet.Cells(i, levelCol + 1).Address
        If factorName = "" Then
            MsgBox "因子・水準・水準値設定表シートの形式が想定外で認識できません。処理を中止します。"
            Exit Function
        End If
        If levelName = "" Then
            MsgBox "因子・水準・水準値設定表シートの形式が想定外で認識できません。処理を中止します。"
            Exit Function
        End If
'        If levelVal = "" Then
'            MsgBox "因子:" & factorName & "水準:" & levelName & "の水準値設定が空になっています。「[水準名]?」に置換します。"
'            FLLVSheet.Unprotect Password:=protectPassword
'            FLLVSheet.Cells(i, levelCol + 1).Interior.Color = RGB(255, 0, 0)  ' 赤色
'            FLLVSheet.Protect Password:=protectPassword
'            levelVal = levelName & "?"
'        End If
        flKey = factorName & ":" & levelName
        If dicFLLV.exists(flKey) Then
            MsgBox "因子・水準・水準値設定表シートにおいて、因子名・水準名の組み合わせに重複があります。処理を中止します。"
            Exit Function
        Else
'            dicFLLV.Add flKey, levelVal
            dicFLLV.Add flKey, levelValAddress
        End If
    Next i
    
    makeFLLVDictionary = True
End Function

' テストデータシートに水準値を書き込む
Function fillInTestDataSheet(dicFLLV As Object, testDataSheet As Worksheet) As Boolean
    fillInTestDataSheet = False
    
    Dim flKey As String
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim i As Long
    Dim j As Long
    Dim titleRow As Long
    Dim factorCol As Long
    Dim idCol As Long
    Dim factorName As String
    Dim levelName As String
    Dim levelVal As String
    
    Call getMaxRowAndCol(testDataSheet, MaxRow, MaxCol)
    titleRow = offsetRows + 1
    idCol = offsetColumns + 1
    
    If Not testDataSheet.Cells(titleRow, idCol).Value = "ID" Then
        MsgBox "テストデータシートの形式が想定外で認識できません。処理を中止します。"
        Exit Function
    End If

    For i = titleRow + 1 To MaxRow
        For j = idCol + 1 To MaxCol
            factorName = testDataSheet.Cells(titleRow, j).Value
            levelName = testDataSheet.Cells(i, j).Value
            flKey = factorName & ":" & levelName
            If dicFLLV.exists(flKey) Then
                ' testDataSheet.Cells(i, j).Value = dicFLLV.Item(flKey)
                testDataSheet.Cells(i, j).Formula = dicFLLV.Item(flKey)
            Else
                MsgBox "因子名・水準名の組み合わせに因子・水準・水準値設定表シートにおいて定義されていないもの「" & flKey & "」があります。処理を中止します。"
                testDataSheet.Cells(i, j).Value = "？"
                testDataSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0)  ' 赤色
            End If
        Next j
    Next i
    
    fillInTestDataSheet = True
End Function

