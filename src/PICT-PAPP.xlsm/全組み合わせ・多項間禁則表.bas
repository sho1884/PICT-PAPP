Attribute VB_Name = "�S�g�ݍ��킹�E�����ԋ֑��\"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' ���q�E�����̑S�g�ݍ��킹����
Sub outputCombination(ws As Worksheet, row As Long, testCase() As String, factorN As Integer, levelLists())
    Dim levelN As Integer

    If factorN > UBound(levelLists) - LBound(levelLists) + 1 - 1 Then
        Call outputTuple(ws, row, testCase)
        row = row + 1
        Exit Sub
    End If
    For levelN = LBound(levelLists(factorN)) To UBound(levelLists(factorN))
        testCase(factorN) = levelLists(factorN)(levelN)
        '�ċA�ďo��
        Call outputCombination(ws, row, testCase, factorN + 1, levelLists)
    Next
End Sub

' �����ς݂̈��q�E�����̑S�g�ݍ��킹���V�[�g�ɏ�������
Sub outputTuple(ws As Worksheet, row As Long, testCase() As String)
    Dim j As Integer
    ' �I�t�Z�b�g���{�^�C�g����1�s���󂯂�
    Worksheets(controlSheetName).Range("�l����").Copy
    ws.Cells(row + offsetRows + 1, offsetColumns + 1).Value = "#" & row ' ID�Ƃ��ăV�[�P���V�����ԍ�
    ws.Cells(row + offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    ' �I�t�Z�b�g���{ID��1����󂯂�
    For j = LBound(testCase) To UBound(testCase)
        ws.Cells(row + offsetRows + 1, offsetColumns + j + 2).Value = testCase(j)
        ws.Cells(row + offsetRows + 1, offsetColumns + j + 2).PasteSpecial (xlPasteFormats)
    Next
End Sub

' ���q�E�����\����S�g�����̃V�[�g�𐶐�����
Function fillInTupleSheets(srcBook As Workbook) As Boolean
    fillInTupleSheets = False

    Dim tuplelSheet As Worksheet
    Dim j As Long

    Dim factorNames() As String
    Dim levelLists()
    Dim testCase() As String

    ' �V���ȃV�[�g�𐶐�����
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = tuplelSheetName
    Set tuplelSheet = srcBook.Sheets(tuplelSheetName)
    If tuplelSheet Is Nothing Then ' ���Ǒg������������V�[�g������ł��Ȃ�
        MsgBox "�S�g������������V�[�g�̖��O�Ƃ��āu" & tuplelSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�������ł��܂���B"
        Exit Function
    End If

    ' ���q�E�����̃V�[�g����肵�Ĉ��q�E������z��ɓǂݍ���
    If Not FLTableSheet2array(srcBook, factorNames, levelLists) Then
        Exit Function
    End If
    
    ' �^�C�g���s�ƂȂ���q���̍s����������
    tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 1).Value = "ID"
    Worksheets(controlSheetName).Range("���ڃ^�C�g������").Copy
    tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 1).PasteSpecial (xlPasteFormats)
    For j = LBound(factorNames) To UBound(factorNames)
        tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 2).Value = factorNames(j)
        tuplelSheet.Cells(offsetRows + 1, j + offsetColumns + 2).PasteSpecial (xlPasteFormats)
    Next j
    
    ' �����̑S�g�ݍ��킹�̏o��
    ReDim testCase(UBound(levelLists) - LBound(levelLists) + 1 - 1)
    Call outputCombination(tuplelSheet, 1, testCase, 0, levelLists)

    fillInTupleSheets = True
End Function

' �����ԋ֑����`���邽�߂֑̋��}�g���N�X�V�[�g�𐶐�����
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

    ' ���q�E�����̃V�[�g����肵�Ĉ��q�E������z��ɓǂݍ���
    If Not FLTableSheet2array(srcBook, publicFactorNames, allLevelLists) Then
        Exit Function
    End If
    
    ' �������q�Ɣ퐧����q��I�����郊�X�g�_�C�A���O��\��
    Call SelectFactors.doModal
    Unload SelectFactors '���[�U�[�t�H�[���͂����ŕ���
    
    If constraintFactors = "" Then
        MsgBox "�퐧����q���I������Ȃ������̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    
    If Not Not conditionFactors Then
        conditionNum = UBound(conditionFactors) - LBound(conditionFactors) + 1 ' �������q�̌�
    Else
        conditionNum = 0 ' ���I�z��̊��蓖�Ă����ݒ�
    End If
    
    If conditionNum < 2 Then
        MsgBox "�������q��2�ȏ�I������Ȃ������̂ŁA�����𒆎~���܂��B�P�ŗǂ��̂ł���Α�������\�����̖ړI�Ɏg���܂��B"
        Exit Function
    End If
    
    ' �������q�Ɣ퐧����q�̏d���`�F�b�N
    For i = LBound(conditionFactors) To UBound(conditionFactors)
        If conditionFactors(i) = constraintFactors Then
            MsgBox "�퐧����q�Ƃ��đI������[" & constraintFactors & "]���������q�ɂ��܂܂�Ă����̂ŁA�����𒆎~���܂��B"
            Exit Function
        End If
    Next i
    
    ' �I�����ꂽ�������q�݂̂̐�����𒊏o
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
    
    ' �����ԋ֑��\�̂܂��g���Ă��Ȃ����O�����i�󂢂Ă���ԍ�������j
    For i = 1 To kinsokuMatrixSheetMax
        kinsokuMatrixSheetName = kinsokuMatrixSheetBaseName & "(" & i & ")"
        If Not ExistsWorksheet(srcBook, kinsokuMatrixSheetName) Then
            Exit For
        End If
    Next i
    
    ' �����ɂ͂��肻���ɂȂ����ƂȂ̂ŁA�֑��\�������߂������߂�B�������ō폜�Y��̂��߂Ǝv����̂ŁB
    If kinsokuMatrixSheetName = "" Then
        MsgBox "�����֑̋��֌W��`�p�V�[�g�̐����ő吔�𒴂��܂����B�����𒆎~���܂��B"
        Exit Function
    Else
        ' �V���ȃV�[�g�𐶐�����
        srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
        ActiveSheet.name = kinsokuMatrixSheetName
        Set kinsokuMatrix = srcBook.Sheets(kinsokuMatrixSheetName)
    End If
    If kinsokuMatrix Is Nothing Then ' ���ǋ֑���`�p�V�[�g������ł��Ȃ�
        MsgBox "�����֑̋��֌W��`�p�V�[�g�̖��O�Ƃ��āu" & kinsokuMatrix & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B"
        Exit Function
    End If

    ' �������q�̍s���o��
    kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + 1).Value = "ID"
    Worksheets(controlSheetName).Range("���q����").Copy
    kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    For j = LBound(conditionFactors) To UBound(conditionFactors)
        kinsokuMatrix.Cells(offsetRows + 1, j + offsetColumns + 2).Value = conditionFactors(j)
        kinsokuMatrix.Cells(offsetRows + 1, j + offsetColumns + 2).PasteSpecial (xlPasteFormats)
    Next j
    
    ' �����̑S�g�ݍ��킹�̏o��
    ReDim conditionComb(UBound(levelLists) - LBound(levelLists) + 1 - 1)
    Call outputCombination(kinsokuMatrix, 1, conditionComb, 0, levelLists) ' �v����ɑS�g�ݍ��킹���������̗��p
    
    ' ��̕�����������
    For j = LBound(conditionFactors) To UBound(conditionFactors)
        kinsokuMatrix.Columns(j + offsetColumns + 2).EntireColumn.AutoFit
    Next j

    ' �S�g�ݍ��킹���������̗��p�ɂ��s�s�����C��
    ' �퐧����q�����L�q����s���P�s�}������
    kinsokuMatrix.Rows(offsetRows).Insert
    ' ID�̗���폜����
    kinsokuMatrix.Columns(offsetColumns + 1).Delete
    
    ' �퐧����q���̏���ǋL
    For j = LBound(constraintLevels) To UBound(constraintLevels)
        kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + conditionNum + j + 1).Value = constraintFactors
        Worksheets(controlSheetName).Range("���q����").Copy
        kinsokuMatrix.Cells(offsetRows + 1, offsetColumns + conditionNum + j + 1).PasteSpecial (xlPasteFormats)

        kinsokuMatrix.Cells(offsetRows + 2, offsetColumns + conditionNum + j + 1).Value = constraintLevels(j)
        Worksheets(controlSheetName).Range("��������").Copy
        kinsokuMatrix.Cells(offsetRows + 2, offsetColumns + conditionNum + j + 1).PasteSpecial (xlPasteFormats)
        kinsokuMatrix.Columns(offsetColumns + conditionNum + j + 1).ColumnWidth = 3
    Next j

    ' �֑��ݒ�ŕҏW�\�Ȕ͈͂�Z���F���ς��ݒ������
    constraintLevelsNum = UBound(constraintLevels) - LBound(constraintLevels) + 1
    
    kinsokuMatrix.Range(Cells(offsetRows + 1, offsetColumns + conditionNum + 1), Cells(offsetRows + 2, offsetColumns + conditionNum + constraintLevelsNum)).Orientation = xlDownward
    Set rng = kinsokuMatrix.Range(Cells(offsetRows + 3, offsetColumns + conditionNum + 1), Cells(offsetRows + conditionCombNum + 2, offsetColumns + conditionNum + constraintLevelsNum))
    Worksheets(controlSheetName).Range("�l����").Copy
    rng.PasteSpecial (xlPasteFormats)
    Call kinsokuFormatCollectionsAdd(rng)

    kinsokuMatrix.Cells.Locked = True ' �S�Z�������b�N
    On Error Resume Next
    rng.SpecialCells(Type:=xlCellTypeBlanks).Locked = False ' ���b�N������
    On Error GoTo 0
    kinsokuMatrix.Protect Password:=protectPassword '�V�[�g�̕ی�
    
    createKinsokuMatrix = True
End Function

