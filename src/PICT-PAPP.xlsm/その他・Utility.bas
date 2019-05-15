Attribute VB_Name = "���̑��EUtility"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' ���q�E�����\�̊e���q�S�ĂɁAMASK��Ԃ�\���V���{���̐�����ǉ�����
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

    ' ���q�E�����̃V�[�g����肷��
    If ExistsWorksheet(srcBook, FLtblSheetName) Then
        Set flSheet = srcBook.Sheets(FLtblSheetName)
    End If
    If flSheet Is Nothing Then ' ���ǈ��q�E�����\�V�[�g��������Ȃ�
        MsgBox "���q�E�����\�̖��O�Ƃ��āu" & FLtblSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B"
        Exit Function
    End If
    
    Call getMaxRowAndCol(flSheet, MaxRow, MaxCol)
    
    Set obj = flSheet.Cells.Find("���q", LookAt:=xlWhole) '�܂��́u���q�v�̃Z����T�����Ƃ�FL�\�ł��邩�ǂ����̎�|����Ƃ���
    If obj Is Nothing Then
        Exit Function
    End If
    
    titleRow = obj.row
    factorCol = obj.Column
    If Not flSheet.Cells(titleRow, factorCol + 1).Value Like "����*" Then
        Exit Function
    End If
    
    fEmptyFlg = False
    
    For i = titleRow + 1 To MaxRow
        content = flSheet.Cells(i, factorCol).Value
        If content = "" Then
            fEmptyFlg = True
        Else
            If fEmptyFlg Then
                MsgBox "���q��̓r���ɋ�̃Z��������܂��B��̃Z���ȉ��̍s�𖳎����܂��B"
                Exit For
            End If
            lEmptyFlg = False
            For j = factorCol + 1 To MaxCol
                content = flSheet.Cells(i, j).Value
                If content = "" Then ' �����珇�ɂ݂čŏ��ɋ�ł������Z����MASK�̏�Ԃ�\��������}������
                    If Not lEmptyFlg Then
                        flSheet.Cells(i, j).Value = maskSymbol
                        lEmptyFlg = True
                    End If
                Else
                    If lEmptyFlg Then
                        MsgBox "������̓r���ɋ�̃Z��������AMASK�p�̐�����}�����܂����B�������A������E�̗�ɒl�������Ă��܂��B���̏�Ԃ͖����N�����̂Ŋm�F���Ă��������B"
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    
    insertMaskSymbol = True
End Function

' ���q�����̑g�ݍ��킹����e�X�gID��������A�z�z������B�����Ƀy�A�E���X�g���o�͂���
Sub makeFLIDDictionary(paramNames() As String, tuples(), ByRef dicFL As Object, Optional pairListSheet As Worksheet = Nothing)
    Dim testCaseNum As Long
    Dim paramNum As Long
    Dim flKey As String
    
    Set dicFL = CreateObject("Scripting.Dictionary")
    dicFL.CompareMode = vbBinaryCompare
        
    If Not Not tuples Then
        testCaseNum = UBound(tuples) - LBound(tuples) + 1 ' �����̌�
    Else
        testCaseNum = 0 ' ���I�z��̊��蓖�Ă����ݒ�
    End If
    
    If Not Not paramNames Then
        paramNum = UBound(paramNames) - LBound(paramNames) + 1 ' �������ڐ�
    Else
        paramNum = 0 ' ���I�z��̊��蓖�Ă����ݒ�
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
    If Not (pairListSheet Is Nothing) Then ' �y�A�E���X�g���o�͂���V�[�g���w�肳��Ă���
        pairNum = 0
        Worksheets(controlSheetName).Range("���ڃ^�C�g������").Copy
        pairListSheet.Cells(pairNum + offsetRows + 1, 1 + offsetColumns).Value = "No."
        pairListSheet.Cells(pairNum + offsetRows + 1, 1 + offsetColumns).PasteSpecial (xlPasteFormats)
        pairListSheet.Cells(pairNum + offsetRows + 1, 2 + offsetColumns).Value = "��1���q:��2���q"
        pairListSheet.Cells(pairNum + offsetRows + 1, 2 + offsetColumns).PasteSpecial (xlPasteFormats)
        pairListSheet.Cells(pairNum + offsetRows + 1, 3 + offsetColumns).Value = "��1���q�̐����l:��2���q�̐����l"
        pairListSheet.Cells(pairNum + offsetRows + 1, 3 + offsetColumns).PasteSpecial (xlPasteFormats)
        pairListSheet.Cells(pairNum + offsetRows + 1, 4 + offsetColumns).Value = "PairwiseID"
        pairListSheet.Cells(pairNum + offsetRows + 1, 4 + offsetColumns).PasteSpecial (xlPasteFormats)
        Worksheets(controlSheetName).Range("�l����").Copy
    End If
    
    Debug.Print Time & " - �����쐬�J�n"
    
    For i = 0 To testCaseNum - 1
        tuple = tuples(i)
        For j = paramL To paramU
            For j2 = j + 1 To paramU
                If Not (pairListSheet Is Nothing) Then ' �y�A�E���X�g���o�͂���V�[�g���w�肳��Ă���
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
    
    Debug.Print Time & " - �����쐬�I��"
    Debug.Print dicFL.Count
    
End Sub

' ���q�E�����\����肵�āA���q�E������z��ɓǂݍ��ޏ������Ăяo��
Function FLTableSheet2array(srcBook As Workbook, ByRef factorNames() As String, ByRef levelLists()) As Boolean
    FLTableSheet2array = False

    Dim FLtblSheet As Worksheet

    ' ���q�E�����̃V�[�g����肷��
    If Not FLtblSheetName = "" Then
        If ExistsWorksheet(srcBook, FLtblSheetName) Then
            Set FLtblSheet = srcBook.Sheets(FLtblSheetName)
        End If
    End If
    If FLtblSheet Is Nothing Then ' ���ǈ��q�E�����\�V�[�g��������Ȃ�
        MsgBox "���q�E�����\�̖��O�Ƃ��āu" & FLtblSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B"
        Exit Function
    End If

    ' ���q�E������z��ɓǂݍ���
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "���q�E�����̉�͂Ɏ��s���܂����B�����𒆎~���܂��B"
        Exit Function
    End If

    FLTableSheet2array = True
End Function

' ���肳�ꂽ���q�E�����\�̈��q�E������z��ɓǂݍ���
Function FLTable2array(flSheet As Worksheet, ByRef factorNames() As String, ByRef levelLists()) As Boolean
    FLTable2array = False ' FL�\�`�����Ɣ��f������K��True��Ԃ�
    
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
    
    Set obj = flSheet.Cells.Find("���q", LookAt:=xlWhole) '�܂��́u���q�v�̃Z����T�����Ƃ�FL�\�ł��邩�ǂ����̎�|����Ƃ���
    If obj Is Nothing Then
        Exit Function
    End If
    
    titleRow = obj.row
    factorCol = obj.Column
    If Not flSheet.Cells(titleRow, factorCol + 1).Value Like "����*" Then
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
                MsgBox "���q��̓r���ɋ�̃Z��������܂��B��̃Z���ȉ��̍s�𖳎����܂��B"
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
                    MsgBox "������̓r���ɋ�̃Z��������܂��B��̃Z�����E�̗�𖳎����܂��B"
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

' Tool�̎��s���ʂ̓����Ă��镶�����z��ɕϊ�����
Function textTable2array(tupleStr As String, delimiter As String, ByRef paramNames() As String, ByRef tuples()) As Boolean
    textTable2array = False
    
    Dim lines() As String
    Dim tuple() As String
    Dim i As Long
    Dim j As Long
    Dim testCaseNum As Long
    
    tupleStr = Replace(tupleStr, vbCrLf, vbLf)
    lines = Split(tupleStr, vbLf)
    
    paramNames = Split(lines(LBound(lines)), delimiter) ' 1�s�ڂ͈��q��
    
    testCaseNum = 0
    For i = LBound(lines) + 1 To UBound(lines) ' 1�s�ڂ͈��q���Ƃ��Ċ��Ɏ�荞�񂾂̂œǂݔ�΂�
        If Not lines(i) = "" Then
            tuple = Split(lines(i), delimiter)
            testCaseNum = testCaseNum + 1
            ReDim Preserve tuples(testCaseNum - 1)
            tuples(testCaseNum - 1) = tuple
        End If
    Next i
            
    textTable2array = True
End Function

' �e�X�g�P�[�X�̓����Ă���V�[�g�̏���z��ɕϊ�����
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
    
    Set obj = testCaseSheet.Cells.Find("ID", LookAt:=xlWhole) '�܂��́uID�v�̃Z����T�����ƂŌ`�����ύX����Ă��Ȃ����ǂ����̎�|����Ƃ���
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
        paramNames(factorNum - 1) = testCaseSheet.Cells(titleRow, j).Value ' 1�s�ڂ͈��q��
    Next j
    
    testCaseNum = 0
    For i = titleRow + 1 To MaxRow
        content = testCaseSheet.Cells(i, idCol).Value
        If content = "" Then
            emptyFlg = True
        Else
            If emptyFlg Then
                MsgBox "ID��̓r���ɋ�̃Z��������܂��B��̃Z���ȉ��̍s�𖳎����܂��B"
                Exit For
            End If
            factorNum = 0
            ReDim tuple(MaxCol - idCol - 1)
            For j = idCol + 1 To MaxCol
                factorNum = factorNum + 1
                tuple(factorNum - 1) = testCaseSheet.Cells(i, j).Value ' 1�s�ڂ͈��q��
            Next j
            testCaseNum = testCaseNum + 1
            ReDim Preserve tuples(testCaseNum - 1)
            tuples(testCaseNum - 1) = tuple
        End If
    Next i
            
    testCaseSheet2array = True
End Function

' ��`�ɂ����鐅�����̏d���𒲂ׂ邽�߂ɐ������̏o���񐔂̎��������
' ��ʂɈ��q���Ⴆ�ΐ������͏d�����ėǂ����AAlloy�ł͕s�s�����N����
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

' �V�[�g�Ŏg���Ă���ő�̍s�ƃJ���������߂�
Sub getMaxRowAndCol(wkSheet As Worksheet, ByRef MaxRow As Long, ByRef MaxCol As Long)
    Dim lDummy As Long: lDummy = wkSheet.UsedRange.row ' ��x UsedRange ���g���ƍŏI�Z�����␳�����悤��
    Dim i As Long
    Dim j As Long
    MaxRow = wkSheet.Cells.SpecialCells(xlLastCell).row
    MaxCol = wkSheet.Cells.SpecialCells(xlLastCell).Column
    
    If wkSheet.Cells.SpecialCells(xlLastCell).MergeCells Then ' �Z������������ꍇ�ɑΉ����čŏI�Z���̈ʒu���C������
        i = MaxRow
        j = MaxCol
        MaxRow = MaxRow + wkSheet.Cells(i, j).MergeArea.Rows.Count - 1
        MaxCol = MaxCol + wkSheet.Cells(i, j).MergeArea.Columns.Count - 1
    End If
End Sub

' �w�肵�����O�̃V�[�g�����݂��邩�m�F���܂��B
Function ExistsWorksheet(wb As Workbook, name As String)
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        If ws.name = name Then
            ' ���݂���
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    
    ' ���݂��Ȃ�
    ExistsWorksheet = False
End Function

' EXCEL�̖��O��`�̃��X�g���V�[�g�ɏ����o���B
' �{�c�[���ł͎g�p���Ă��Ȃ��B�f�o�b�O�p
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
    Cells(i, 5) = "�V�[�g��"
    Cells(i, 6) = "�V�[�g��name"
    Cells(i, 7) = "Parent.name"
    For Each nm In ActiveWorkbook.Names
    'For Each nm In ws.Names�@�K�������V�[�g������Ŗ��O��t���Ă���Ă���Ƃ͌���Ȃ��̂�
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

