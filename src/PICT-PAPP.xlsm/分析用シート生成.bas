Attribute VB_Name = "���͗p�V�[�g����"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' �������̕��͗p�V�[�g�̐���
Function analysis(srcBook As Workbook, paramNames() As String, tuples()) As Boolean
    analysis = False
                
    Dim pairListSheet As Worksheet
    Dim dicFL As Object
    
    If pairListFlg Then
        ' 2���q�Ԃ̑g�ݍ��킹�o�������͗p�Ƀy�A�E���X�g�����p�̃V�[�g��p�ӂ���
        srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
        ActiveSheet.name = pairListSheetName
        Set pairListSheet = srcBook.Sheets(pairListSheetName)
        
        ' ���q�����̑g�ݍ��킹����e�X�gID��������A�z�z������
        ' �����Ƀy�A�E���X�g�����p�̃V�[�g�ɏ�����������
        Call makeFLIDDictionary(paramNames(), tuples, dicFL, pairListSheet)
    Else
        ' ���q�����̑g�ݍ��킹����e�X�gID��������A�z�z������
        ' �y�A�E���X�g�����p�̃V�[�g�͍쐬���Ȃ�
        Call makeFLIDDictionary(paramNames(), tuples, dicFL)
    End If
    
    ' Tool�̏o�͌��ʂ𕡐�������������\�Ƀ}�b�v���A�����ɃJ�o���b�W��Ԃ������V�[�g�𐶐�����
    If Not fillInRoundRobinTable(srcBook, paramNames(), tuples, dicFL) Then
        MsgBox "��������\�ւ̃e�X�gID�}�b�s���O�����Ɏ��s���܂����B"
    End If

    analysis = True
End Function

' ��������\�𕡐�����"ID�}�b�s���O�ςݑ�������\"�V�[�g�𐶐����A������Tool�̌��ʂƂ��ē����e�X�gID�␔��ǋL����B
' �i������������\�����݂��Ȃ��ꍇ�́A����ɑO�����Ƃ��Đ�������B�j
Function fillInRoundRobinTable(srcWorkbook As Workbook, paramNames() As String, tuples(), dicFL As Object) As Boolean
    fillInRoundRobinTable = False
    
    Dim destSheet As Worksheet
    Dim testCaseNum As Long
    Dim paramNum As Long
    Dim flKey As String
    Dim pairCount As Long ' �֑��������ꍇ�̑SPair�̐�
    Dim someCount As Long ' Pairwise�ɏ��Ȃ��Ƃ�1��͏o������Pair�̐�
    Dim kinsokuCount As Long ' ��������\�ɂ����ċ֑��Ɩ������ꂽ�SPair�̐�
    Dim uncertainCount As Long ' ��������\�ɂ����ċ֑��Ɩ�������Ă��Ȃ����ATool���o�͂��Ȃ�����Pair�̐�
    
    pairCount = 0
    someCount = 0
    kinsokuCount = 0
    uncertainCount = 0
    
    ' ��������\�����݂��Ȃ��ꍇ�́A����ɑO�����Ƃ��Đ�������
    If Not ExistsWorksheet(srcWorkbook, roundRobinSheetName) Then
        If createRoundRobinTable(ThisWorkbook) Then
            MsgBox "��������\�𐶐��������܂����B"
        Else
            MsgBox "��������\�̐��������Ɏ��s���܂����B�����𒆎~���܂��B"
            Exit Function
        End If
    End If
    
    ' ��������\�𕡐�����"ID�}�b�s���O�ςݑ�������\"�V�[�g�𐶐�
    srcWorkbook.Worksheets(roundRobinSheetName).Copy Before:=srcWorkbook.Worksheets(roundRobinSheetName)
    ActiveSheet.name = mappedRoundRobinSheetName
    Set destSheet = srcWorkbook.Worksheets(mappedRoundRobinSheetName)
    destSheet.Unprotect Password:=protectPassword ' ���ɕK�v�Ȃ��̂ŕی삵�Ȃ�

    ' �V�[�g�ɏ�������
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
        MsgBox "��������\�̏c���̃T�C�Y�������Ă��܂���B�S�~�f�[�^���������Ă���Ǝv����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    
    For i = levelRow + 1 To MaxRow
        For j = levelCol + 1 To MaxCol
            If i - levelRow > j - levelCol Then ' �Ίp����荶��
                flKey = destSheet.Cells(factorRow, j).Value & ":" & destSheet.Cells(levelRow, j).Value & "*" & destSheet.Cells(i, factorCol).Value & ":" & destSheet.Cells(i, levelCol).Value
            Else ' �Ίp�����E��
                flKey = destSheet.Cells(i, factorCol).Value & ":" & destSheet.Cells(i, levelCol).Value & "*" & destSheet.Cells(factorRow, j).Value & ":" & destSheet.Cells(levelRow, j).Value
            End If
            If dicFL.exists(flKey) Then
                pairCount = pairCount + 1
                someCount = someCount + 1
                If destSheet.Cells(i, j).Value = "�~" Then
                    kinsokuCount = kinsokuCount + 1
                    MsgBox "�֑��w�肵���g�ݍ��킹���e�X�g�P�[�X�̒��ɏo�����Ă��܂��B���񎮂̎��������Ȃǂ̐������菇�𓥂񂾂��m�F���Ă��������B�i�V�[�g��" & i & "�s��" & j & "��j"
                Else
                    If i - levelRow > j - levelCol Then ' �Ίp����荶��
                        destSheet.Cells(i, j).Value = dicFL.Item(flKey)
                    Else ' �Ίp�����E��
                        IDs = Split(dicFL.Item(flKey), ",")
                        destSheet.Cells(i, j).Value = UBound(IDs) - LBound(IDs) + 1
                    End If
                End If
            Else
                Select Case destSheet.Cells(i, j).Value
                    Case "�\" ' �Ίp����̎��g�Ƃ̑g�ݍ��킹�ŁA�Ӗ��������̂Ŗ���
                    Case "�~" ' �֑��w�肳��Ă���Ȃ�Ζ��Ȃ�
                        pairCount = pairCount + 1
                        kinsokuCount = kinsokuCount + 1
                    Case "" ' �֑��w�肳��Ă��Ȃ��Ȃ�Ζ��Ȃ̂ŁA�o�����Ă��Ȃ���������������K�v������B
                        pairCount = pairCount + 1
                        uncertainCount = uncertainCount + 1
                        destSheet.Cells(i, j).Value = "?"
'                        destSheet.Cells(i, j).Interior.Color = RGB(255, 255, 0) ' ���F
                    Case Else
                        pairCount = pairCount + 1
                        someCount = someCount + 1
                        MsgBox "��������\�ɈӖ��s���̓��͒l�������Ă��܂��B�������菇�𓥂񂾂��m�F���Ă��������B�i�V�[�g��" & i & "�s��" & j & "��j"
                End Select
            End If
        Next j
    Next i
    
    ' ���̐������`�F�b�N
    someCount = someCount / 2
    kinsokuCount = kinsokuCount / 2
    uncertainCount = uncertainCount / 2
    pairCount = pairCount / 2
    If Not (someCount + kinsokuCount + uncertainCount = pairCount) Then
        MsgBox "��������\�֑̋��△���̐ݒ肪�s�K�؂Ȃ��߁A�ePair���̍��v���ɂ��ĕs�������N�����Ă��܂��B�m�F���Ă��������B"
    End If
    
    ' �ԗ����V�[�g����
    If Not fillInCoverageSheet(srcWorkbook, someCount, kinsokuCount, uncertainCount, pairCount, UBound(tuples) - LBound(tuples) + 1) Then
        MsgBox "�ԗ����̃V�[�g�����Ɏ��s���܂����B"
    End If
    
    destSheet.Activate
    
    fillInRoundRobinTable = True
End Function
    
' �ԗ����V�[�g�����Ə�������
Function fillInCoverageSheet(srcWorkbook As Workbook, someCount, kinsokuCount, uncertainCount, pairCount, testcaseCount) As Boolean
    fillInCoverageSheet = False
    
    ' �ԗ����̏o�̓V�[�g��p�ӂ���
    Dim coverageSheet As Worksheet
    srcWorkbook.Worksheets.Add Before:=Worksheets(mappedRoundRobinSheetName)
    ActiveSheet.name = coverageSheetName
    Set coverageSheet = srcWorkbook.Sheets(coverageSheetName)
    
    Worksheets(controlSheetName).Range("���ڃ^�C�g������").Copy
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
    
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 2).Value = "Tool�o�͌��ʂ̃e�X�g���ڐ�"
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 2).Value = "�g�����ԗ���(��) ( B = C / E �~ 100)"
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 2).Value = "D�̑g�ݍ��킹���S�ċ֑��ł���ꍇ�̑g�����ԗ���(��) ( B' = C / (E-D) �~ 100)"
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 2).Value = "Tool�o�͌����ԗ�����2���q�ԑg������"
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 2).Value = "Tool�o�͌��ʂ��ԗ������A�������֑��Ɩ����ݒ肳��Ă����Ȃ�2���q�ԑg������"
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 2).Value = "�֑��Ɩ����ݒ肳��Ă�����̂�������2���q�ԑg������ ( E = G - F )"
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 2).Value = "�֑��Ɩ����ݒ肳��Ă���2���q�ԑg�ݍ�����"
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 2).Value = "�֑��������Ɖ��肵���ꍇ��2���q�ԑg�ݍ�����"
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 2).PasteSpecial (xlPasteFormats)
    coverageSheet.Columns(offsetColumns + 2).EntireColumn.AutoFit
    
    Worksheets(controlSheetName).Range("�l����").Copy
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 3).Value = testcaseCount ' Tool�o�͌��̍��ڐ�
    coverageSheet.Cells(offsetRows + 1, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 3).Value = someCount / (pairCount - kinsokuCount) * 100 ' �g�����ԗ���(��) ( B = C / E �~ 100)
    coverageSheet.Cells(offsetRows + 2, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 3).Value = someCount / (pairCount - kinsokuCount - uncertainCount) * 100 ' D�̑g�ݍ��킹���S�ċ֑��ł���ꍇ�̑g�����ԗ���(��) ( B' = C / (E-D) �~ 100)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 3, offsetColumns + 3).Interior.Color = RGB(255, 255, 0) ' ���F
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 3).Value = someCount ' Tool�o�͌����ԗ�����2���q�ԑg������
    coverageSheet.Cells(offsetRows + 4, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 3).Value = uncertainCount ' Tool�o�͌����ԗ������A�������֑��Ɩ����ݒ肳��Ă����Ȃ�2���q�ԑg������
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 5, offsetColumns + 3).Interior.Color = RGB(255, 255, 0) ' ���F
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 3).Value = pairCount - kinsokuCount ' �֑��Ɩ����ݒ肳��Ă�����̂�������2���q�ԑg������ ( E = G - F )
    coverageSheet.Cells(offsetRows + 6, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 3).Value = kinsokuCount ' �֑��Ɩ����ݒ肳��Ă���2���q�ԑg�ݍ�����
    coverageSheet.Cells(offsetRows + 7, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 3).Value = pairCount ' �֑��������Ɖ��肵���ꍇ��2���q�ԑg�ݍ�����
    coverageSheet.Cells(offsetRows + 8, offsetColumns + 3).PasteSpecial (xlPasteFormats)
    coverageSheet.Columns(offsetColumns + 3).EntireColumn.AutoFit
    
    fillInCoverageSheet = True
End Function

