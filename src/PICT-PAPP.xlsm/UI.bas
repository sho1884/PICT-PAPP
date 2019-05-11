Attribute VB_Name = "UI"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

Public Const offsetRows As Integer = 1 ' 表を出力する際、最上部に何行を空けるか
Public Const offsetColumns As Integer = 1 ' 表を出力する際、最左部に何列を空けるか
Public Const alloyLevelSuffix As String = "の水準" ' alloy用で因子に対応する要素を表すために使用する名前として使うsuffix
Public Const alloyExec As String = "run 組合せ状態が存在する for 1 but exactly 1 システム" ' alloyでの実行指示

Public Const alloySrcFileName As String = "ペアが存在しなくて良いことを検証するalloyソース.als"
Public Const pictInFileName As String = "PICTin.txt"
Public Const pictOutFileName As String = "PICTout.txt"
Public Const citBachInFileName As String = "CitBachIn.txt"
Public Const citBachOutFileName As String = "CitBachOut.txt"

Public Const controlSheetName As String = "〔処理の指示＆設定〕"
Public Const tuplelSheetName As String = "全組み合わせ"
Public Const coverageSheetName As String = "網羅率"
Public Const roundRobinSheetName As String = "総当たり表"
Public Const mappedRoundRobinSheetName As String = "IDマッピング済み総当たり表"
Public Const pairListSheetName As String = "ペア・リスト"
Public Const pairListFlg As Boolean = False 'ペア・リストを生成する
Public Const toolOutSheetName As String = "ツールの生成結果"
Public Const testCaseSheetName As String = "テストケース"
Public Const testDataSheetName As String = "テストデータ"
Public Const FLtblSheetName As String = "因子・水準"
Public Const FLLVSheetName As String = "因子・水準・水準値"
Public Const constraintSheetName As String = "制約記述"
Public Const kinsokuMatrixSheetBaseName As String = "多項間禁則表"
Public Const kinsokuMatrixSheetMax As Integer = 100 ' 多項間禁則表シートの上限数

Public maskSymbol As String ' MASK状態を表すシンボル
Public protectPassword As String ' シート保護に使うパスワード
Public toolName As String ' 実行Tool名
Public pictCmdOption As String ' PICTコマンドオプション
Public citBachCmdOption As String ' CIT-BACHコマンドオプション

Public conditionFactors() As String
Public constraintFactors As String
Public publicFactorNames() As String

' Step0)
Sub MASK水準の自動挿入()
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If insertMaskSymbol(ThisWorkbook) Then
        MsgBox "処理しました。"
    Else
        MsgBox "処理に失敗しました。"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step1)
Sub 総当たり()
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    ' この処理で生成される総当たり表シートが既存ではないことを確認する
    If ExistsWorksheet(ThisWorkbook, roundRobinSheetName) Then ' 総当たり表シートの名前が既に使われているか？
        MsgBox "総当たり表シートの名前として指定されている「" & roundRobinSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If createRoundRobinTable(ThisWorkbook) Then
        MsgBox "処理しました。"
    Else
        MsgBox "処理に失敗しました。"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step1')
Sub 多項禁則マトリクス生成()
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If createKinsokuMatrix(ThisWorkbook) Then
        MsgBox "処理しました。"
    Else
        MsgBox "処理に失敗しました。"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step2)
Sub 制約自動生成()
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If generateConstraintExpression(ThisWorkbook) Then
        MsgBox "制約自動生成処理しました。"
    Else
        MsgBox "制約自動生成処理に失敗しました。"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step3) Tool実行の準備、実行、結果の取り込み、いくつかの分析用シートの生成までの一連の流れを駆動する
Sub Tool実行()
    Dim paramNames() As String
    Dim tuples()
    
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    ' この処理で生成されるシート名のシートが既存ではないことを確認する
    If ExistsWorksheet(ThisWorkbook, mappedRoundRobinSheetName) Then ' IDマッピング済み総当たり表シートの名前が既に使われているか？
        MsgBox "IDマッピング済み総当たり表シートの名前として指定されている「" & mappedRoundRobinSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, pairListSheetName) Then ' ペア・リストシートの名前が既に使われているか？
        MsgBox "ペア・リストシートの名前として指定されている「" & pairListSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' ツールの生成結果シートの名前が既に使われているか？
        MsgBox "ツールの生成結果シートの名前として指定されている「" & toolOutSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, coverageSheetName) Then ' 網羅率シートの名前が既に使われているか？
        MsgBox "網羅率シートの名前として指定されている「" & coverageSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If

    If toolName = "PICT" Then
        If Not PictFEP(paramNames, tuples) Then
            Exit Sub
        End If
    ElseIf toolName = "CIT-BACH" Then
        If Not CitBachFEP(paramNames, tuples) Then
            Exit Sub
        End If
    Else
        MsgBox "実行Tool名が正しく設定されていないので、処理を中止します。"
        Exit Sub
    End If
    ThisWorkbook.Sheets(toolOutSheetName).Activate
' 以下の処理も継続して実施してしまうと便利であるが、一方でデータサイズによっては処理時間がかかることがある。
' プログレスバーなどを付けて、中断もできるようにすると良いかもしれない。(要検討)
'    On Error GoTo ErrLabel
'    Application.ScreenUpdating = False
'    If Not analysis(ThisWorkbook, paramNames, tuples) Then
'        Application.ScreenUpdating = True
'        Exit Sub
'    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' ToolにPICTが選択されている場合、Step3でPICT実行の準備、実行、結果の取り込み
Function PictFEP(ByRef paramNames() As String, ByRef tuples()) As Boolean
    PictFEP = False
    Dim pairwiseStr As String
    
    If Not getToolInputFile(ThisWorkbook, pictInFileName, citBachInFileName) Then
        MsgBox "Tool入力用ファイルの生成に失敗しましたので。処理を中止します。"
        Exit Function
    End If
    If Not execPict(pictInFileName, pictOutFileName) Then
        MsgBox "PICTの実行処理に失敗しましたので、処理を中止します。"
        Exit Function
    End If
    If Not inputUtf8(pictOutFileName, pairwiseStr) Then
        MsgBox "PICTの実行結果ファイルを読み取ることに失敗したので、処理を中止します。"
        Exit Function
    End If
    If Not textTable2array(pairwiseStr, vbTab, paramNames(), tuples) Then
        MsgBox "PICTの実行結果を解析することに失敗したので、処理を中止します。"
        Exit Function
    End If
    ' PICT出力結果用シートに結果を書き込む
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "PICTの実行結果をシートに出力することに失敗したので、処理を中止します。"
        Exit Function
    End If
    
    PictFEP = True
    
ErrLabel:
    Application.ScreenUpdating = True
End Function

' ToolにCIT-BACHが選択されている場合、Step3でCIT-BACH実行の準備、実行、結果の取り込み
Function CitBachFEP(ByRef paramNames() As String, ByRef tuples()) As Boolean
    CitBachFEP = False
    Dim pairwiseStr As String
    Dim lengthLine1 As Long
    Dim line1 As String
    
    If Not getToolInputFile(ThisWorkbook, pictInFileName, citBachInFileName) Then
        MsgBox "Tool入力用ファイルの生成に失敗しましたので。処理を中止します。"
        Exit Function
    End If
    If Not execCitBach(citBachInFileName, citBachOutFileName) Then
        MsgBox "CIT-BACHの実行処理に失敗しましたので、処理を中止します。"
        Exit Function
    End If
    ' Pairwiseの生成結果ファイルの読み込み
    If Not inputFile(citBachOutFileName, pairwiseStr) Then
        MsgBox "Pairwiseの生成結果ファイルを読み取ることに失敗したので、処理を中止します。"
        Exit Function
    End If
    
    lengthLine1 = InStr(pairwiseStr, vbLf) - 1
    line1 = Left(pairwiseStr, lengthLine1)
    
    If Left(line1, 1) = "#" And InStr(line1, vbTab) = 0 Then ' CIT-BACHの出力と思われる
        pairwiseStr = Mid(pairwiseStr, lengthLine1 + 2)
    Else
        MsgBox "CIT-BACHの生成結果ファイルを読み取ることに失敗したので、処理を中止します。"
        Exit Function
    End If
    If Not textTable2array(pairwiseStr, ",", paramNames(), tuples) Then
        MsgBox "CIT-BACHの実行結果を解析することに失敗したので、処理を中止します。"
        Exit Function
    End If
    ' CIT-BACH出力結果用シートに結果を書き込む
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "CIT-BACHの実行結果をシートに出力することに失敗したので、処理を中止します。"
        Exit Function
    End If
    
    CitBachFEP = True

ErrLabel:
    Application.ScreenUpdating = True
End Function

' Step3') Pairwiseのファイル生成は済んでいることを前提に、その実行済みの結果ファイルを指定してシートに取り込み
Sub Tool実行済み結果ファイルのシートへの読み込み処理()
    Dim pairwiseStr As String
    Dim lengthLine1 As Long
    Dim line1 As String
    Dim pwFileFullName As String
    Dim paramNames() As String
    Dim tuples()
    Dim delimiter As String
    
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    ' この処理で生成されるシート名のシートが既存ではないことを確認する
    If ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' ツールの生成結果シートの名前が既に使われているか？
        MsgBox "ツールの生成結果シートの名前として指定されている「" & toolOutSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If

    ' Pairwiseの生成結果ファイルの読み込み
    If Not inputFile(pwFileFullName, pairwiseStr) Then
        MsgBox "Pairwiseの生成結果ファイルを読み取ることに失敗したので、処理を中止します。"
        Exit Sub
    End If
    
    lengthLine1 = InStr(pairwiseStr, vbLf) - 1
    line1 = Left(pairwiseStr, lengthLine1)
    
    If Left(line1, 1) = "#" And InStr(line1, vbTab) = 0 Then ' CIT-BACHの出力と思われる
        pairwiseStr = Mid(pairwiseStr, lengthLine1 + 2)
        delimiter = ","
    Else ' PICTの出力と思われる（もう一度読み直し）
        If Not inputUtf8(pwFileFullName, pairwiseStr) Then ' ファイル名は既に指定されている同じもの
            MsgBox "PICTの実行結果ファイルを読み取ることに失敗したので、処理を中止します。"
            Exit Sub
        End If
        delimiter = vbTab
    End If
    
    If Not textTable2array(pairwiseStr, delimiter, paramNames(), tuples) Then
        MsgBox "Toolの実行結果を解析することに失敗したので、処理を中止します。"
        Exit Sub
    End If
    ' Tool出力結果用シートに結果を書き込む
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "Toolの実行結果をシートに出力することに失敗したので、処理を中止します。"
        Exit Sub
    End If
    ThisWorkbook.Sheets(toolOutSheetName).Activate

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step4) 生成済み結果のシートまたはそれを複製した編集済みシートの情報から、総当たり表の複製シートにIDをマップした分析シートを作成する
Sub Tool結果シートまたは編集済みシートから分析までの処理()
    Dim paramNames() As String
    Dim tuples()
    Dim srcSheet As Worksheet
    Dim pairListSheet As Worksheet
    Dim toolOutStr As String
    
    Debug.Print Time & " - Tool結果シートまたは編集済みシートから分析までの処理開始"
    
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    ' この処理で生成されるシート名のシートが既存ではないことを確認する
    If ExistsWorksheet(ThisWorkbook, mappedRoundRobinSheetName) Then ' IDマッピング済み総当たり表シートの名前が既に使われているか？
        MsgBox "IDマッピング済み総当たり表シートの名前として指定されている「" & mappedRoundRobinSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, pairListSheetName) Then ' ペア・リストシートの名前が既に使われているか？
        MsgBox "ペア・リストシートの名前として指定されている「" & pairListSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, coverageSheetName) Then ' 網羅率シートの名前が既に使われているか？
        MsgBox "網羅率シートの名前として指定されている「" & coverageSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If

    If ExistsWorksheet(ThisWorkbook, testCaseSheetName) Then ' テストケースのシートを見付ける
        Set srcSheet = ThisWorkbook.Sheets(testCaseSheetName)
    ElseIf ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' ツールの生成結果のシートを見付ける
        Set srcSheet = ThisWorkbook.Sheets(toolOutSheetName)
    Else
        MsgBox "「" & testCaseSheetName & "」シートまたは「" & toolOutSheetName & "」シートが存在することが必要です。存在しなかったので処理を中止します。"
        Exit Sub
    End If

    ' テストケースのシートまたはツールの生成結果のシートから実行結果の読み込み
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not testCaseSheet2array(srcSheet, paramNames(), tuples) Then
        MsgBox "テストケースの解析に失敗したので、処理を中止します。"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If Not analysis(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Debug.Print Time & " - Tool結果シートまたは編集済みシートから分析までの処理終了"
    MsgBox "処理しました。"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step4') 外部ツールによる生成結果のファイルが既に存在していることを前提に、その実行済みの結果ファイルを指定して取り込み、総当たり表の複製シートにIDをマップした分析シートを作成する
Sub Tool生成結果のファイルから分析までの処理()
    Dim pairwiseStr As String
    Dim lengthLine1 As Long
    Dim line1 As String
    Dim pwFileFullName As String
    Dim paramNames() As String
    Dim tuples()
    Dim delimiter As String
    
    Debug.Print Time & " - Tool生成結果のファイルから分析までの処理開始"
    
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    ' この処理で生成されるシート名のシートが既存ではないことを確認する
    If ExistsWorksheet(ThisWorkbook, mappedRoundRobinSheetName) Then ' IDマッピング済み総当たり表シートの名前が既に使われているか？
        MsgBox "IDマッピング済み総当たり表シートの名前として指定されている「" & mappedRoundRobinSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, pairListSheetName) Then ' ペア・リストシートの名前が既に使われているか？
        MsgBox "ペア・リストシートの名前として指定されている「" & pairListSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' ツールの生成結果シートの名前が既に使われているか？
        MsgBox "ツールの生成結果シートの名前として指定されている「" & toolOutSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, coverageSheetName) Then ' 網羅率シートの名前が既に使われているか？
        MsgBox "網羅率シートの名前として指定されている「" & coverageSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If

    ' Pairwiseの生成結果ファイルの読み込み
    If Not inputFile(pwFileFullName, pairwiseStr) Then
        MsgBox "Pairwiseの生成結果ファイルを読み取ることに失敗したので、処理を中止します。"
        Exit Sub
    End If
    
    lengthLine1 = InStr(pairwiseStr, vbLf) - 1
    line1 = Left(pairwiseStr, lengthLine1)
    
    If Left(line1, 1) = "#" And InStr(line1, vbTab) = 0 Then ' CIT-BACHの出力と思われる
        pairwiseStr = Mid(pairwiseStr, lengthLine1 + 2)
        delimiter = ","
    Else ' PICTの出力と思われる（もう一度読み直し）
        If Not inputUtf8(pwFileFullName, pairwiseStr) Then ' ファイル名は既に指定されている同じもの
            MsgBox "PICTの実行結果ファイルを読み取ることに失敗したので、処理を中止します。"
            Exit Sub
        End If
        delimiter = vbTab
    End If
    
    If Not textTable2array(pairwiseStr, delimiter, paramNames(), tuples) Then
        MsgBox "Toolの実行結果を解析することに失敗したので、処理を中止します。"
        Exit Sub
    End If
    ' Tool出力結果用シートに結果を書き込む
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "Toolの実行結果をシートに出力することに失敗したので、処理を中止します。"
        Exit Sub
    End If
    
    If Not analysis(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Debug.Print Time & " - Tool生成結果のファイルから分析までの処理終了"
    MsgBox "処理しました。"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step5)
Sub ペアが存在しなくて良いことを検証するalloyソースの生成()
    Dim alloySrc As String
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    If createAlloySrc(ThisWorkbook, alloySrc) Then
        ThisWorkbook.Worksheets(constraintSheetName).Range("Alloyによる検証用表現").Value = alloySrc
        ThisWorkbook.Worksheets(constraintSheetName).Activate
        ThisWorkbook.Worksheets(constraintSheetName).Range("Alloyによる検証用表現").Select
    Else
        MsgBox "ペアが存在しなくて良いことを検証するalloyソースの生成に失敗しました。"
        Exit Sub
    End If
    If outputUtf8(alloySrc, alloySrcFileName) Then
        If Not execAlloy(alloySrcFileName) Then
            MsgBox "Alloyの起動に失敗しました。"
        End If
    Else
        MsgBox "ペアが存在しなくて良いことを検証するalloyソースファイルの出力に失敗しました。"
    End If
End Sub

' Step6)
Sub 因子_水準_水準値設定表の生成()
    Dim alloySrc As String
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    If ExistsWorksheet(ThisWorkbook, FLLVSheetName) Then ' 因子・水準・水準値設定表シートの名前が既に使われているか？
        MsgBox "因子・水準・水準値設定表シートの名前として指定されている「" & FLLVSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not createFLLVSheet(ThisWorkbook) Then
        MsgBox "因子・水準・水準値設定表シートの生成に失敗しました。"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    MsgBox "処理しました。"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step6) 因子・水準・水準値設定表に従って生成済みPairwise結果のシートの水準を水準値に置換する
Sub 水準を水準値に置換()
    Dim paramNames() As String
    Dim tuples()
    Dim FLLVSheet As Worksheet
    Dim srcSheet As Worksheet
    Dim testDataSheet As Worksheet
    Dim toolOutStr As String
    Dim dicFLLV As Object
    
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    
    If ExistsWorksheet(ThisWorkbook, testDataSheetName) Then ' テストデータシートの名前が既に使われているか？
        MsgBox "テストデータシートの名前として指定されている「" & testDataSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    
    ' 因子・水準・水準値設定表シートを見付ける
    If ExistsWorksheet(ThisWorkbook, FLLVSheetName) Then
        Set FLLVSheet = ThisWorkbook.Sheets(FLLVSheetName)
    Else
        MsgBox "「" & FLLVSheetName & "」シートが存在することが必要です。存在しなかったので処理を中止します。"
        Exit Sub
    End If

    If ExistsWorksheet(ThisWorkbook, testCaseSheetName) Then ' テストケースのシートを見付ける
        Set srcSheet = ThisWorkbook.Sheets(testCaseSheetName)
    ElseIf ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' ツールの生成結果のシートを見付ける
        Set srcSheet = ThisWorkbook.Sheets(toolOutSheetName)
    Else
        MsgBox "「" & testCaseSheetName & "」シートまたは「" & toolOutSheetName & "」シートが存在することが必要です。存在しなかったので処理を中止します。"
        Exit Sub
    End If
    ThisWorkbook.Worksheets(srcSheet.name).Copy Before:=ThisWorkbook.Worksheets(srcSheet.name)
    ActiveSheet.name = testDataSheetName
    Set testDataSheet = ThisWorkbook.Worksheets(testDataSheetName)

    ' 因子・水準の対の名前から水準値への辞書を作成する
    If Not makeFLLVDictionary(FLLVSheet, dicFLLV) Then
        Exit Sub
    End If
    
    ' 辞書の情報に基づいてテストデータのシートの水準名を水準値に置換
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInTestDataSheet(dicFLLV, testDataSheet) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    MsgBox "処理しました。"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' おまけ
Sub 全組み合わせ()
    If GetSetValues = False Then
        MsgBox "設定値の読み取りに失敗したので終了します。"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, tuplelSheetName) Then ' 全組み合わせシートの名前が既に使われているか？
        MsgBox "全組み合わせシートの名前として指定されている「" & tuplelSheetName & "」が既に存在しています。削除するか改名してください。"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If fillInTupleSheets(ThisWorkbook) Then
        MsgBox "処理しました。"
    Else
        MsgBox "処理に失敗しました。"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' 作業フォルダをExplorerで開く
Sub 作業フォルダをExplorerで開く()
    
    Call Shell("explorer.exe " & Worksheets(controlSheetName).Range("作業パス").Value, vbNormalFocus)

End Sub

Sub getWorkingPath()
    Worksheets(controlSheetName).Range("作業パス").Value = getPath("処理対象ファイルが格納されているフォルダを選択")
End Sub

Sub getPictPath()
    Worksheets(controlSheetName).Range("PICTフルパス").Value = getPictExePath("pict.exeファイルを選択")
End Sub

Sub getCitBachPath()
    Worksheets(controlSheetName).Range("CIT_BACHフルパス").Value = getCitBachJarPath("cit-bach.jarファイルを選択")
End Sub

Sub getAlloyPath()
    Worksheets(controlSheetName).Range("Alloyフルパス").Value = getAlloyJarPath("alloy.jarファイルを選択")
End Sub

' ユーザにフォルダを選択させて、そのパスを得る
Function getPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    If fileDialog.Show = -1 Then
        getPath = fileDialog.SelectedItems(1)
    End If
End Function

' ユーザにpict.exeファイルを選択させて、そのフルパスを得る
Function getPictExePath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "pict.exeファイル", "*.exe"
    If fileDialog.Show = -1 Then
        getPictExePath = fileDialog.SelectedItems(1)
    End If
End Function

' ユーザにcit-bach.jarファイルを選択させて、そのフルパスを得る
Function getCitBachJarPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "cit-bach.jarファイル", "*.jar"
    If fileDialog.Show = -1 Then
        getCitBachJarPath = fileDialog.SelectedItems(1)
    End If
End Function

' ユーザにalloy.jarファイルを選択させて、そのフルパスを得る
Function getAlloyJarPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "alloy.jarファイル", "*.jar"
    If fileDialog.Show = -1 Then
        getAlloyJarPath = fileDialog.SelectedItems(1)
    End If
End Function

' ユーザにファイルを選択させて、そのフルパスを得る
Function getFilePath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.InitialFileName = Worksheets(controlSheetName).Range("作業パス").Value & "\" & pictOutFileName
    fileDialog.Filters.Add title, "*.*"
    If fileDialog.Show = -1 Then
        getFilePath = fileDialog.SelectedItems(1)
    End If
End Function

' 設定値の読み取り
Function GetSetValues() As Boolean
    GetSetValues = False
    
    maskSymbol = Worksheets(controlSheetName).Range("MASK状態を表すシンボル").Value
    protectPassword = Worksheets(controlSheetName).Range("シート保護に使うパスワード").Value
    toolName = Worksheets(controlSheetName).Range("実行Tool名").Value
    If maskSymbol = "" Then
        maskSymbol = "mask"
        Worksheets(controlSheetName).Range("MASK状態を表すシンボル").Value = maskSymbol
        MsgBox "MASK状態を表すシンボルが設定されていないので、maskとしました。"
    End If
    If protectPassword = "" Then
        protectPassword = "password"
        Worksheets(controlSheetName).Range("シート保護に使うパスワード").Value = protectPassword
        MsgBox "シート保護に使うパスワードが設定されていないので、passwordとしました。"
    End If
    If Not (toolName = "PICT" Or toolName = "CIT-BACH") Then
        toolName = "CIT-BACH"
        Worksheets(controlSheetName).Range("実行Tool名").Value = toolName
        MsgBox "実行Toolの選択が正しく設定されていないので、CIT-BACHとしました。"
    End If
    pictCmdOption = Worksheets(controlSheetName).Range("PICTオプション").Value
    citBachCmdOption = Worksheets(controlSheetName).Range("CIT_BACHオプション").Value
    
    GetSetValues = True
End Function

