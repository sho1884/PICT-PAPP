Attribute VB_Name = "ツール実行・入出力"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

' Tool入力用のソースをファイルに出力する
Function getToolInputFile(srcBook As Workbook, inFileName As String, inFileNameS As String) As Boolean
    getToolInputFile = False
                
    Dim pictSrc As String
    Dim citBachSrc As String

    ' TOOL入力用のソースを生成する処理を呼び出す
    If Not createToolInputSrc(srcBook, pictSrc, citBachSrc) Then
        MsgBox "Toolのソースを生成することに失敗したので、処理を中止します。"
        Exit Function
    End If
    
    ' PICT入力用のソースをファイルに出力する処理を呼び出す
    If Not outputUtf8(pictSrc, inFileName) Then
        MsgBox "PICTのソースをファイル出力することに失敗したので、処理を中止します。"
        Exit Function
    End If

    ' CIT-BACH入力用のソースをファイルに出力する処理を呼び出す
    If Not outputFile(citBachSrc, inFileNameS) Then
        MsgBox "CIT-BACHのソースをファイル出力することに失敗したので、処理を中止します。"
        Exit Function
    End If

    getToolInputFile = True
End Function

' PICT実行
Function execPict(inFileName As String, outFileName As String) As Boolean
    execPict = False
    
    Dim objShell As Object
    Dim oExec
    Dim rc As Long
    Dim comStr As String

    ' PICT実行
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = Worksheets(controlSheetName).Range("作業パス").Value
    
    comStr = """" & Worksheets(controlSheetName).Range("PICTフルパス").Value & """ """ & inFileName & """ " & pictCmdOption & " > """ & outFileName & """"
    comStr = "%ComSpec% /c """ & comStr & """"
        
    rc = objShell.Run(comStr, 1, True)
    If Not rc = 0 Then
        MsgBox "PICTの実行結果の戻り値が" & rc & "なので、実行に失敗している可能性があります。"
    End If
    execPict = True
End Function

' CIT-BACH実行
Function execCitBach(inFileName As String, outFileName As String) As Boolean
    execCitBach = False
    
    Dim objShell As Object
    Dim oExec
    Dim rc As Long
    Dim comStr As String

    ' CIT-BACH実行
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = Worksheets(controlSheetName).Range("作業パス").Value
    
    comStr = "java -jar " & """" & Worksheets(controlSheetName).Range("CIT_BACHフルパス").Value & """ " & citBachCmdOption & " -i """ & inFileName & """ -o """ & outFileName & """"
    comStr = "%ComSpec% /c """ & comStr & """"
        
    rc = objShell.Run(comStr, 1, True)
    If Not rc = 0 Then
        MsgBox "CIT-BACHの実行結果の戻り値が" & rc & "なので、実行に失敗している可能性があります。"
    End If
    
    execCitBach = True
End Function

' Alloy実行
Function execAlloy(inFileName As String) As Boolean
    execAlloy = False
    
    Dim objShell As Object
    Dim oExec
    Dim rc As Long
    Dim comStr As String

    ' Alloy実行
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = Worksheets(controlSheetName).Range("作業パス").Value
    
    comStr = "java -jar " & """" & Worksheets(controlSheetName).Range("Alloyフルパス").Value & """ """ & inFileName ' & """ -o """ & outFileName & """"
    comStr = "%ComSpec% /c """ & comStr & """"
        
    rc = objShell.Run(comStr, 0, False)
    If Not rc = 0 Then
        MsgBox "Alloyの実行結果の戻り値が" & rc & "なので、実行に失敗している可能性があります。"
    End If
    
    execAlloy = True
End Function

' Toolへの入力ファイルを因子・水準の情報と制約(総当たり表から自動生成した制約も含め)記述の情報から生成する
Function createToolInputSrc(srcBook As Workbook, ByRef pictSrc As String, ByRef citBachSrc As String) As Boolean
    createToolInputSrc = False
    Dim paramNames() As String
    Dim tuples()
    Dim FLtblSheet As Worksheet
    Dim factorNames() As String
    Dim levelLists()
    Dim factorNum As Long
    Dim levelNum As Long
    
    ' 因子・水準のシートを特定する
    If Not FLtblSheetName = "" Then
        If ExistsWorksheet(srcBook, FLtblSheetName) Then
            Set FLtblSheet = srcBook.Sheets(FLtblSheetName)
        End If
    End If
    If FLtblSheet Is Nothing Then ' 結局因子・水準表シートが見つからない
        MsgBox "因子・水準表の名前として「" & FLtblSheetName & "」が指定されていますが、その名前のシートが存在しません。処理を中止します。"
        Exit Function
    End If
    
    ' 因子・水準を読み込む
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "因子・水準の解析に失敗しました。処理を中止します。"
        Exit Function
    End If

    ' 因子・水準情報
    For factorNum = LBound(factorNames) To UBound(factorNames)
        pictSrc = pictSrc & factorNames(factorNum) & ": "
        citBachSrc = citBachSrc & factorNames(factorNum) & " ("
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            If levelNum > 0 Then
                pictSrc = pictSrc & ", "
                citBachSrc = citBachSrc & " "
            End If
            pictSrc = pictSrc & levelLists(factorNum)(levelNum)
            citBachSrc = citBachSrc & levelLists(factorNum)(levelNum)
        Next levelNum
        pictSrc = pictSrc & vbLf
        citBachSrc = citBachSrc & ")" & vbLf
    Next factorNum
        
    ' 制約情報の追加
    Dim constraints As String
    If Not GetConstraints(constraints) Then
        MsgBox "制約記述の取得に失敗しました。処理を中止します。"
        Exit Function
    End If
    
    ' S式側制約情報の追加
    Dim s_constraints As String
    If Not GetSConstraints(s_constraints) Then
        MsgBox "CIT-BACH制約記述の取得に失敗しました。処理を中止します。"
        Exit Function
    End If
    
    pictSrc = pictSrc & constraints
    citBachSrc = citBachSrc & s_constraints

    createToolInputSrc = True
End Function

' Toolの出力結果をシートに書き込む
Function fillInToolOutSheets(srcBook As Workbook, paramNames() As String, tuples()) As Boolean
    fillInToolOutSheets = False
    
    Dim testCaseNum As Long
    Dim paramNum As Long
    Dim toolOutSheet As Worksheet
    Dim j As Long
    Dim paramL As Long
    Dim paramU As Long
     
    ' Tool出力結果用のシートを用意する
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = toolOutSheetName
    Set toolOutSheet = srcBook.Sheets(toolOutSheetName)
    
    ' データ数を確認してシートに書き込む
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
    
    paramL = LBound(paramNames)
    paramU = UBound(paramNames)
    Worksheets(controlSheetName).Range("項目タイトル書式").Copy
    toolOutSheet.Cells(offsetRows + 1, offsetColumns + 1).Value = "ID"
    toolOutSheet.Cells(offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    For j = paramL To paramU
        toolOutSheet.Cells(offsetRows + 1, offsetColumns + j - paramL + 2).Value = paramNames(j)
        toolOutSheet.Cells(offsetRows + 1, offsetColumns + j - paramL + 2).PasteSpecial (xlPasteFormats)
    Next j
    Dim i As Long
    Dim tuple() As String
    Worksheets(controlSheetName).Range("値書式").Copy
    For i = 0 To testCaseNum - 1
        tuple = tuples(i)
        toolOutSheet.Cells(offsetRows + 2 + i, offsetColumns + 1).Value = "#" & i + 1 ' IDとしてシーケンシャル番号
        toolOutSheet.Cells(offsetRows + 2 + i, offsetColumns + 1).PasteSpecial (xlPasteFormats)
        For j = paramL To paramU
            toolOutSheet.Cells(offsetRows + 2 + i, offsetColumns + j - paramL + 2).Value = tuple(j)
            toolOutSheet.Cells(offsetRows + 2 + i, offsetColumns + j - paramL + 2).PasteSpecial (xlPasteFormats)
        Next j
    Next i
    
    fillInToolOutSheets = True
End Function

Function inputUtf8(srcFilename As String, ByRef str As String) As Boolean
    inputUtf8 = False
    'ADODB.Streamオブジェクトを生成
    Dim srcFileFullName As String
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8への変換のために使う
    
    If srcFilename = "" Then
        srcFileFullName = getFilePath("PICT実行結果のファイルを選択")
        If srcFileFullName = "" Then
            MsgBox "PICT実行結果のファイルが選択されなかったので、処理を中止します。"
            Exit Function
        End If
    ElseIf InStr(srcFilename, "\") > 0 Then ' 元々フルパスで指定されていると判断する
        srcFileFullName = srcFilename
    Else
        srcFileFullName = Worksheets(controlSheetName).Range("作業パス").Value & "\" & srcFilename
    End If
    
    With adoSt
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adCRLF
        .Open
        .LoadFromFile (srcFileFullName)
        str = .ReadText
        .Close
    End With
    
    inputUtf8 = True
End Function

Function inputFile(ByRef srcFilename As String, ByRef str As String) As Boolean
    inputFile = False
    
    Dim srcFileFullName As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As TextStream
    
    If srcFilename = "" Then
        srcFileFullName = getFilePath("PICTまたはCIT-BACHの実行結果のファイルを選択")
        If srcFileFullName = "" Then
            MsgBox "PICTまたはCIT-BACHの実行結果のファイルが選択されなかったので、処理を中止します。"
            Exit Function
        End If
    ElseIf InStr(srcFilename, "\") > 0 Then ' 元々フルパスで指定されていると判断する
        srcFileFullName = srcFilename
    Else
        srcFileFullName = Worksheets(controlSheetName).Range("作業パス").Value & "\" & srcFilename
    End If
    
    Set ts = fso.OpenTextFile(srcFileFullName, Format:=TristateFalse) ' ファイルを Shift-JIS で開く

    ' 全てのデータを読み込み
    str = ts.ReadAll
   
    srcFilename = srcFileFullName ' 実際に読み込んだファイルのフルパス名を返す
    
    inputFile = True
End Function

Function outputUtf8(src As String, destFilename As String) As Boolean
    outputUtf8 = False
    
    Dim destFileFullName As String
    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8への変換のために使う
    
    destFileFullName = Worksheets(controlSheetName).Range("作業パス").Value & "\" & destFilename
    
    With adoSt
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText src
        ' BOMを削除する
        Dim byteData() As Byte
        .Position = 0
        .Type = adTypeBinary
        .Position = 3
        byteData = adoSt.Read
        .Close
        .Open
        .Write byteData
    
        .SaveToFile destFileFullName, adSaveCreateOverWrite
        .Close
    End With
    
    outputUtf8 = True
End Function

Function outputFile(src As String, destFilename As String) As Boolean
    
    Dim destFileFullName As String
    
    destFileFullName = Worksheets(controlSheetName).Range("作業パス").Value & "\" & destFilename
    
    outputFile = False
    Open destFileFullName For Output As #1
    Print #1, src
    Close #1
    
    outputFile = True
End Function

