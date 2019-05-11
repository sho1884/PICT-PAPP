Attribute VB_Name = "��������\"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' ���q�E�����\���瑍������\�𐶐�����
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

    ' ��������\�V�[�g��V���ɐ�������
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = roundRobinSheetName
    Set roundRobinSheet = srcBook.Sheets(roundRobinSheetName)
    If roundRobinSheet Is Nothing Then ' ���Ǒ�������\�V�[�g������ł��Ȃ�
        MsgBox "��������\�V�[�g�̖��O�Ƃ��āu" & roundRobinSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B"
        Exit Function
    End If
    
    ' ���q�E�����̃V�[�g����肷��
    If ExistsWorksheet(srcBook, FLtblSheetName) Then
        Set FLtblSheet = srcBook.Sheets(FLtblSheetName)
    End If
    If FLtblSheet Is Nothing Then ' ���ǈ��q�E�����\�V�[�g��������Ȃ�
        MsgBox "���q�E�����\�̖��O�Ƃ��āu" & FLtblSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B"
        Exit Function
    End If
    
    ' ���q�E������ǂݍ���
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "���q�E�����̉�͂Ɏ��s���܂����B�����𒆎~���܂��B"
        Exit Function
    End If

    ' ��������\���o��
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
    
    ' roundRobinSheet.Unprotect Password:=protectPassword ' ���������Ƃ���Ȃ̂ŕی삳��Ă��Ȃ�
    
    Worksheets(controlSheetName).Range("���ڃ^�C�g������").Copy
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            totalLevelNum = totalLevelNum + 1
            j = totalLevelNum + levelCol
            roundRobinSheet.Cells(factorRow, j).Value = factorNames(factorNum)
            Worksheets(controlSheetName).Range("���q����").Copy
            roundRobinSheet.Cells(factorRow, j).PasteSpecial (xlPasteFormats)

            roundRobinSheet.Cells(levelRow, j).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("��������").Copy
            roundRobinSheet.Cells(levelRow, j).PasteSpecial (xlPasteFormats)
            roundRobinSheet.Columns(j).ColumnWidth = 3
        Next levelNum
    Next factorNum
    roundRobinSheet.Range(Cells(factorRow, levelCol + 1), Cells(levelRow, j)).Orientation = xlDownward
    Set rng = roundRobinSheet.Range(Cells(levelRow + 1, levelCol + 1), Cells(levelRow + totalLevelNum, levelCol + totalLevelNum))
    Worksheets(controlSheetName).Range("�l����").Copy
    rng.PasteSpecial (xlPasteFormats)
    Call kinsokuFormatCollectionsAdd(rng)
    
    roundRobinSheet.Cells.Locked = True ' �S�Z�������b�N
    On Error Resume Next
    rng.SpecialCells(Type:=xlCellTypeBlanks).Locked = False ' ���b�N������
    On Error GoTo 0
    
    i = levelRow
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            i = i + 1
            roundRobinSheet.Cells(i, factorCol).Value = factorNames(factorNum)
            Worksheets(controlSheetName).Range("���q����").Copy
            roundRobinSheet.Cells(i, factorCol).PasteSpecial (xlPasteFormats)
            
            roundRobinSheet.Cells(i, levelCol).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("��������").Copy
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
                roundRobinSheet.Cells(i + r, j + c).Value = "�\"
                Worksheets(controlSheetName).Range("��������").Copy
                roundRobinSheet.Cells(i + r, j + c).PasteSpecial (xlPasteFormats)
            Next j
        Next i
        r = r + levelNum
        c = c + levelNum
    Next factorNum
    
    roundRobinSheet.Protect Password:=protectPassword '�V�[�g�̕ی�
    
    createRoundRobinTable = True
End Function

' ��������\�̃V�[�g�ɁA�u�֑���\���~�����͂��ꂽ�Ƃ��ɁA���̃Z���̔w�i�F�Ȃǂ������ύX����v���[����ݒ肷��
Sub kinsokuFormatCollectionsAdd(r As Range)
    Dim f   As FormatCondition
    
    '// �����t�������̒ǉ��i�Z���Ɂ~�����͂��ꂽ�ꍇ�j
    Set f = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="�~")
    '// �t�H���g�����A�����F�A�w�i�F
    f.Font.Bold = Worksheets(controlSheetName).Range("�֑�����").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("�֑�����").Font.Color
    f.Interior.Color = Worksheets(controlSheetName).Range("�֑�����").Interior.Color
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("�֑�����").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("�֑�����").Borders(xlEdgeTop).Weight
    
    '// �����t�������̒ǉ��i�Z���ɁH�����͂��ꂽ�ꍇ�j
    Set f = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="�H")
    '// �t�H���g�����A�����F�A�w�i�F
    f.Font.Bold = Worksheets(controlSheetName).Range("�l����").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("�l����").Font.Color
    f.Interior.Color = RGB(255, 0, 0)  ' �ԐF
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("�l����").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("�l����").Borders(xlEdgeTop).Weight
    
    '// �����t�������̒ǉ��i�Z����?�����͂��ꂽ�ꍇ�j
    Set f = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="?")
    '// �t�H���g�����A�����F�A�w�i�F
    f.Font.Bold = Worksheets(controlSheetName).Range("�l����").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("�l����").Font.Color
    f.Interior.Color = RGB(255, 255, 0) ' ���F
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("�l����").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("�l����").Borders(xlEdgeTop).Weight
End Sub

