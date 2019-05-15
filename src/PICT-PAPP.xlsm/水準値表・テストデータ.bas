Attribute VB_Name = "�����l�\�E�e�X�g�f�[�^"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' ���q�E�����\������q�E�����E�����l�\�𐶐�����
Function createFLLVSheet(srcBook As Workbook) As Boolean
    createFLLVSheet = False
                
    Dim FLLVSheet As Worksheet
    Dim FLtblSheet As Worksheet
    Dim rng As Range
    Dim i As Long
    
    Dim factorNames() As String
    Dim levelLists()

    ' �V���Ȉ��q�E�����E�����l�ݒ�\�V�[�g�𐶐�����
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = FLLVSheetName
    Set FLLVSheet = srcBook.Sheets(FLLVSheetName)
    
    If FLLVSheet Is Nothing Then ' ���Ǒ�������\�V�[�g������ł��Ȃ�
        MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̖��O�Ƃ��āu" & FLLVSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B"
        Exit Function
    End If
    
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
    
    ' ���q�E������ǂݍ���
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "���q�E�����̉�͂Ɏ��s���܂����B�����𒆎~���܂��B"
        Exit Function
    End If

    ' �\���o��
    Dim factorNum As Long
    Dim levelNum As Long
    Dim startRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    
    startRow = offsetRows + 1
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    
    i = startRow
    Worksheets(controlSheetName).Range("���ڃ^�C�g������").Copy
    FLLVSheet.Cells(i, factorCol).Value = "���q"
    FLLVSheet.Cells(i, factorCol).PasteSpecial (xlPasteFormats)
    FLLVSheet.Cells(i, levelCol).Value = "����"
    FLLVSheet.Cells(i, levelCol).PasteSpecial (xlPasteFormats)
    FLLVSheet.Cells(i, levelCol + 1).Value = "�����l"
    FLLVSheet.Cells(i, levelCol + 1).PasteSpecial (xlPasteFormats)
    FLLVSheet.Cells(i, levelCol + 2).Value = "���l"
    FLLVSheet.Cells(i, levelCol + 2).PasteSpecial (xlPasteFormats)
    
    For factorNum = LBound(factorNames) To UBound(factorNames)
        For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
            i = i + 1
            FLLVSheet.Cells(i, factorCol).Value = factorNames(factorNum)
            Worksheets(controlSheetName).Range("���q����").Copy
            FLLVSheet.Cells(i, factorCol).PasteSpecial (xlPasteFormats)
            
            FLLVSheet.Cells(i, levelCol).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("��������").Copy
            FLLVSheet.Cells(i, levelCol).PasteSpecial (xlPasteFormats)
        
            FLLVSheet.Cells(i, levelCol + 1).Value = levelLists(factorNum)(levelNum)
            Worksheets(controlSheetName).Range("�l����").Copy
            FLLVSheet.Cells(i, levelCol + 1).PasteSpecial (xlPasteFormats)
        
            FLLVSheet.Cells(i, levelCol + 2).Value = ""
            Worksheets(controlSheetName).Range("�l����").Copy
            FLLVSheet.Cells(i, levelCol + 2).PasteSpecial (xlPasteFormats)
        Next levelNum
    Next factorNum
    FLLVSheet.Columns(factorCol).EntireColumn.AutoFit
    FLLVSheet.Columns(levelCol).EntireColumn.AutoFit
    FLLVSheet.Columns(levelCol + 1).EntireColumn.AutoFit
    FLLVSheet.Columns(levelCol + 2).ColumnWidth = 80
    
    Set rng = FLLVSheet.Range(Cells(startRow + 1, levelCol + 1), Cells(i, levelCol + 2))
    FLLVSheet.Cells.Locked = True ' �S�Z�������b�N
    On Error Resume Next
    rng.Locked = False ' ���b�N������
    Set rng = FLLVSheet.Range(Cells(startRow + 1, levelCol + 1), Cells(i, levelCol + 1))
    Call levelValFormatCollectionsAdd(rng)
    On Error GoTo 0
    FLLVSheet.Protect Password:=protectPassword '�V�[�g�̕ی�
    
    createFLLVSheet = True
End Function

' ���q�E�����E�����l�\�̃V�[�g�ŁA�����l����ɂȂ��Ă���ꍇ�ɁA���̃Z���̔w�i�F��Ԃɂ��郋�[����ݒ肷��
Sub levelValFormatCollectionsAdd(r As Range)
    Dim f   As FormatCondition
    
    '// �����t�������̒ǉ��i�Z���ɁH�����͂��ꂽ�ꍇ�j
    Set f = r.FormatConditions.Add(Type:=xlBlanksCondition)
    '// �t�H���g�����A�����F�A�w�i�F
    f.Font.Bold = Worksheets(controlSheetName).Range("�l����").Font.Bold
    f.Font.Color = Worksheets(controlSheetName).Range("�l����").Font.Color
    f.Interior.Color = RGB(255, 0, 0)  ' �ԐF
    f.Borders.LineStyle = Worksheets(controlSheetName).Range("�l����").Borders(xlEdgeTop).LineStyle
    f.Borders.Weight = Worksheets(controlSheetName).Range("�l����").Borders(xlEdgeTop).Weight
End Sub

' ���q�E�����E�����l�ݒ�\�V�[�g�̏����g���āA���q�E�����̑΂̖��O���琅���l��������A�z�z������B
' �Z���̒l�������Ă����ꍇ�̃R�[�h���R�����g�A�E�g����Ă���B����A�Z���̃A�h���X�������Ă����d�l�ɂȂ��Ă���B
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
    
    If Not FLLVSheet.Cells(titleRow, factorCol).Value = "���q" Then
        MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̌`�����z��O�ŔF���ł��܂���B�����𒆎~���܂��B"
        Exit Function
    End If
    If Not FLLVSheet.Cells(titleRow, levelCol).Value = "����" Then
        MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̌`�����z��O�ŔF���ł��܂���B�����𒆎~���܂��B"
        Exit Function
    End If
    If Not FLLVSheet.Cells(titleRow, levelCol + 1).Value = "�����l" Then
        MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̌`�����z��O�ŔF���ł��܂���B�����𒆎~���܂��B"
        Exit Function
    End If
    
    Set dicFLLV = CreateObject("Scripting.Dictionary")

    For i = titleRow + 1 To MaxRow
        factorName = FLLVSheet.Cells(i, factorCol).Value
        levelName = FLLVSheet.Cells(i, levelCol).Value
'        levelVal = FLLVSheet.Cells(i, levelCol + 1).Value
        levelValAddress = "=" & FLLVSheetName & "!" & FLLVSheet.Cells(i, levelCol + 1).Address
        If factorName = "" Then
            MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̌`�����z��O�ŔF���ł��܂���B�����𒆎~���܂��B"
            Exit Function
        End If
        If levelName = "" Then
            MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̌`�����z��O�ŔF���ł��܂���B�����𒆎~���܂��B"
            Exit Function
        End If
'        If levelVal = "" Then
'            MsgBox "���q:" & factorName & "����:" & levelName & "�̐����l�ݒ肪��ɂȂ��Ă��܂��B�u[������]?�v�ɒu�����܂��B"
'            FLLVSheet.Unprotect Password:=protectPassword
'            FLLVSheet.Cells(i, levelCol + 1).Interior.Color = RGB(255, 0, 0)  ' �ԐF
'            FLLVSheet.Protect Password:=protectPassword
'            levelVal = levelName & "?"
'        End If
        flKey = factorName & ":" & levelName
        If dicFLLV.exists(flKey) Then
            MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�ɂ����āA���q���E�������̑g�ݍ��킹�ɏd��������܂��B�����𒆎~���܂��B"
            Exit Function
        Else
'            dicFLLV.Add flKey, levelVal
            dicFLLV.Add flKey, levelValAddress
        End If
    Next i
    
    makeFLLVDictionary = True
End Function

' �e�X�g�f�[�^�V�[�g�ɐ����l����������
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
        MsgBox "�e�X�g�f�[�^�V�[�g�̌`�����z��O�ŔF���ł��܂���B�����𒆎~���܂��B"
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
                MsgBox "���q���E�������̑g�ݍ��킹�Ɉ��q�E�����E�����l�ݒ�\�V�[�g�ɂ����Ē�`����Ă��Ȃ����́u" & flKey & "�v������܂��B�����𒆎~���܂��B"
                testDataSheet.Cells(i, j).Value = "�H"
                testDataSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0)  ' �ԐF
            End If
        Next j
    Next i
    
    fillInTestDataSheet = True
End Function

