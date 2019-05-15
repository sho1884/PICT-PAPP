Attribute VB_Name = "���񎮐���"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' ��������\�̐ݒ肵���֑���񂩂�Tool�̐��񎮂𐶐�����
Function generateBinaryConstraintExpression(srcWorkbook As Workbook) As Boolean
    generateBinaryConstraintExpression = False
    Dim constraintExpressions As String
    Dim s_expression As String
    Dim roundRobinSheet As Worksheet
    Dim constraintSheet As Worksheet
    If ExistsWorksheet(srcWorkbook, roundRobinSheetName) Then
        Set roundRobinSheet = srcWorkbook.Worksheets(roundRobinSheetName)
    Else
        MsgBox "��������\�����݂��܂���B���̏����ɐ旧���đ�������\�������������A���̃V�[�g�ɋ֑�����������ŉ������B"
        Exit Function
    End If
    If ExistsWorksheet(srcWorkbook, constraintSheetName) Then
        Set constraintSheet = srcWorkbook.Worksheets(constraintSheetName)
    Else
        MsgBox "����L�q�V�[�g�����݂��܂���B���̏����ɂ͕K�v�ł��B�폜���Ă��܂����ꍇ�͌��̃t�@�C�����畜�����Ă��������B"
        Exit Function
    End If
       
    Dim i As Long
    Dim j As Long
    Dim vN As Long ' ��������\�̐��������̉��Ԗڂ̃}�X��
    Dim hN As Long ' ��������\�̐��������̉��Ԗڂ̃}�X��
    Dim cellStr As String
    Dim MaxRow As Long
    Dim MaxCol As Long
    
    roundRobinSheet.Unprotect Password:=protectPassword
    Call getMaxRowAndCol(roundRobinSheet, MaxRow, MaxCol)
    
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
    
    For i = levelRow + 1 To MaxRow
        For j = levelCol + 1 To MaxCol
            vN = i - levelRow
            hN = j - levelCol
            If vN < hN Then ' �Ίp�����E�ゾ����������Ηǂ�
                If roundRobinSheet.Cells(i, j).Value = "�~" Then ' �֑��w�肳��Ă���
                    If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "�~" Then
                        MsgBox "�֑��̃y�A���A�Ίp���Ő��ΏۂɂȂ��Ă��܂���B�������Ă��Ȃ��Z���̒l���H�ɁA�w�i�F��Ԃɂ��܂����B" & _
                            i & "�s" & j & "��̃Z����" & hN + levelRow & "�s" & vN + levelCol & "��̃Z���͑Ίp���Ő��Ώ̂Ȃ͂��ł��B"
                        roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "�H"
                    End If
                    constraintExpressions = constraintExpressions & "IF [" & roundRobinSheet.Cells(factorRow, j).Value & "] = """ & roundRobinSheet.Cells(levelRow, j).Value & """ THEN [" & _
                                roundRobinSheet.Cells(i, factorCol).Value & "] <> """ & roundRobinSheet.Cells(i, levelCol).Value & """ ;" & vbLf
                    s_expression = s_expression & "(if (== [" & roundRobinSheet.Cells(factorRow, j).Value & "] " & roundRobinSheet.Cells(levelRow, j).Value & ")" & vbLf & _
                                "    (<> [" & roundRobinSheet.Cells(i, factorCol).Value & "] " & roundRobinSheet.Cells(i, levelCol).Value & "))" & vbLf
                End If
            Else ' �Ίp����荶���ɂ��ẮA�Ίp���Ő��ΏۂɂȂ��Ă��邩�ǂ������֑��w�肳��Ă���y�A�Ɍ����ă`�F�b�N����
                If roundRobinSheet.Cells(i, j).Value = "�~" Then ' �֑��w�肳��Ă���
                    If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "�~" Then
                        MsgBox "�֑��̃y�A���A�Ίp���Ő��ΏۂɂȂ��Ă��܂���B�������Ă��Ȃ��Z���̒l���H�ɁA�w�i�F��Ԃɂ��܂����B" & _
                            i & "�s" & j & "��̃Z����" & hN + levelRow & "�s" & vN + levelCol & "��̃Z���͑Ίp���Ő��Ώ̂Ȃ͂��ł��B"
                        roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "�H"
                    End If
                End If
            End If
        Next j
    Next i
    constraintSheet.Range("������������").Value = constraintExpressions
    constraintSheet.Cells(constraintSheet.Range("������������").row, constraintSheet.Range("������������").Column - 1).Value = roundRobinSheetName
    constraintSheet.Cells(constraintSheet.Range("������������").row, constraintSheet.Range("������������").Column + 1).Value = s_expression
    constraintSheet.Activate
    
    roundRobinSheet.Protect Password:=protectPassword

    generateBinaryConstraintExpression = True
End Function

' �����ԋ֑��}�g���N�X����`����Ă���S�ẴV�[�g�ɂ��Đ��񎮂𐶐����Đ���L�q�V�[�g�ɏ�������
Function generateConstraintExpression(srcBook As Workbook) As Boolean
    generateConstraintExpression = False
    Dim i As Long
    Dim srcSheet As Worksheet
    Dim constraintSheet As Worksheet
    Dim idCol As Long
    Dim constrainCol As Long
    Dim binaryConstrainRow As Long
    Dim currentConstrainRow As Long
    Dim expression As String
    Dim s_expression As String
    
    If ExistsWorksheet(srcBook, constraintSheetName) Then
        Set constraintSheet = srcBook.Worksheets(constraintSheetName)
    Else
        MsgBox "����L�q�V�[�g�����݂��܂���B���̏����ɂ͕K�v�ł��B�폜���Ă��܂����ꍇ�͌��̃t�@�C�����畜�����Ă��������B"
        Exit Function
    End If
    If Not generateBinaryConstraintExpression(srcBook) Then
        MsgBox "��������\���琧�񎮂𐶐�����ۂɖ�肪�������܂����B���e���m�F���Ă��������B"
    End If
    binaryConstrainRow = constraintSheet.Range("������������").row
    constrainCol = constraintSheet.Range("������������").Column
    idCol = constrainCol - 1
    
    ' �ߋ��̏o�͂��폜
    Do While True
        If InStr(constraintSheet.Cells(binaryConstrainRow + 1, idCol).Value, kinsokuMatrixSheetBaseName) <> 0 Then
            constraintSheet.Rows(binaryConstrainRow + 1).Delete
        Else
            Exit Do
        End If
    Loop
        
    ' �o��
    currentConstrainRow = binaryConstrainRow
    For i = 1 To srcBook.Sheets.Count ' �S�ẴV�[�g�ɂ��ċ֑��\�����f����
        Set srcSheet = srcBook.Sheets(i)
        If InStr(srcSheet.name, kinsokuMatrixSheetBaseName) <> 0 Then
            If Not kinsokuMatrix2expression(srcSheet, expression, s_expression) Then
                expression = "�֑��}�g���N�X�̉�͂Ɏ��s���܂���"
            End If
            currentConstrainRow = currentConstrainRow + 1
            constraintSheet.Rows(currentConstrainRow).Insert
            constraintSheet.Cells(currentConstrainRow, idCol).Value = srcSheet.name
            constraintSheet.Cells(currentConstrainRow, constrainCol).Value = expression
            constraintSheet.Cells(currentConstrainRow, constrainCol + 1).Value = s_expression
            constraintSheet.Range("������������").Copy
            constraintSheet.Cells(currentConstrainRow, idCol).PasteSpecial (xlPasteFormats)
            constraintSheet.Cells(currentConstrainRow, constrainCol).PasteSpecial (xlPasteFormats)
            constraintSheet.Cells(currentConstrainRow, constrainCol + 1).PasteSpecial (xlPasteFormats)
            Application.CutCopyMode = False
        End If
    Next i
    
    constraintSheet.Activate
    
    generateConstraintExpression = True
End Function

' �����ԋ֑��}�g���N�X��1���̃V�[�g���琧�񎮂𐶐�����
Function kinsokuMatrix2expression(srcSheet As Worksheet, ByRef expression As String, ByRef s_expression As String) As Boolean
    kinsokuMatrix2expression = False
    
    Dim conditionFactorsRow As Long
    Dim constraintFactorsRow As Long
    Dim firstConditionFactorsCol As Long
    Dim firstConstraintFactorsCol As Long
    Dim conditionNum As Long
    Dim constraintLevelNum As Long
    Dim conditionFactor() As String
    Dim constraintFactor As String
    Dim constraintLevels() As String
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim i As Long
    Dim j As Long
    Dim conditionExpression As String
    Dim conditionSExpression As String

    expression = ""
    s_expression = ""
    srcSheet.Unprotect Password:=protectPassword
    Call getMaxRowAndCol(srcSheet, MaxRow, MaxCol)
    srcSheet.Protect Password:=protectPassword
    
    conditionFactorsRow = offsetRows + 2
    constraintFactorsRow = offsetRows + 1
    firstConditionFactorsCol = offsetColumns + 1
    
    ' �퐧����q�̈��q���ƊJ�n������t����
    For j = firstConditionFactorsCol To MaxCol
        If Not srcSheet.Cells(constraintFactorsRow, j).Value = "" Then
            constraintFactor = srcSheet.Cells(constraintFactorsRow, j).Value
            firstConstraintFactorsCol = j
            Exit For
        End If
    Next j
    If constraintFactor = "" Then
        MsgBox srcSheet.name & "�V�[�g�̔퐧����q���L�ڂ����ׂ��ʒu�ɂ����āA����𔭌��ł��܂���ł����B"
        Exit Function
    End If
    
    ' �������q�̐������߂�
    conditionNum = firstConstraintFactorsCol - firstConditionFactorsCol
    
    ' �������q����z��ɓ����
    ReDim conditionFactors(conditionNum - 1)
    For j = firstConditionFactorsCol To firstConstraintFactorsCol - 1
        conditionFactors(j - firstConditionFactorsCol) = srcSheet.Cells(conditionFactorsRow, j).Value
        If conditionFactors(j - firstConditionFactorsCol) = "" Then
            MsgBox srcSheet.name & "�V�[�g�̏������q���L�ڂ����ׂ��ʒu�ɂ����āA�󗓂�����܂����B"
            Exit Function
        End If
    Next j
    
    ' �퐧����q�̐����������߂�
    constraintLevelNum = MaxCol - firstConstraintFactorsCol + 1
    
    ' �퐧����q�̐�������z��ɓ����
    ReDim constraintLevels(constraintLevelNum - 1)
    For j = firstConstraintFactorsCol To MaxCol
        If Not srcSheet.Cells(constraintFactorsRow, j).Value = constraintFactor Then
            MsgBox srcSheet.name & "�V�[�g�̔퐧����q���L�ڂ����ׂ��ʒu�ɂ����āA���O����ӂɂȂ��Ă��܂���B"
            Exit Function
        End If
        constraintLevels(j - firstConstraintFactorsCol) = srcSheet.Cells(constraintFactorsRow + 1, j).Value
        If constraintLevels(j - firstConstraintFactorsCol) = "" Then
            MsgBox srcSheet.name & "�V�[�g�̔퐧����q�̐������L�ڂ����ׂ��ʒu�ɂ����āA�󗓂�����܂����B"
            Exit Function
        End If
    Next j
    
    For i = conditionFactorsRow + 1 To MaxRow
        ' ������
        conditionExpression = ""
        conditionSExpression = ""
        For j = firstConditionFactorsCol To firstConstraintFactorsCol - 1
            If j = firstConditionFactorsCol Then
                conditionExpression = conditionExpression & "IF "
                conditionSExpression = "(if (and "
            Else
                conditionExpression = conditionExpression & " AND "
            End If
            conditionExpression = conditionExpression & "[" & conditionFactors(j - firstConditionFactorsCol) & "] = "
            conditionExpression = conditionExpression & """" & srcSheet.Cells(i, j).Value & """"
            ' S��
            conditionSExpression = conditionSExpression & "(== [" & conditionFactors(j - firstConditionFactorsCol) & "] "
            conditionSExpression = conditionSExpression & srcSheet.Cells(i, j).Value & ") "
        Next j
        ' ����
        conditionExpression = conditionExpression & " THEN "
        conditionSExpression = conditionSExpression & ")" & vbLf
        For j = firstConstraintFactorsCol To MaxCol
            If srcSheet.Cells(i, j).Value = "�~" Then
                expression = expression & conditionExpression
                expression = expression & "[" & constraintFactor & "] <> "
                expression = expression & """" & constraintLevels(j - firstConstraintFactorsCol) & """"
                expression = expression & ";" & vbLf
                ' S��
                s_expression = s_expression & conditionSExpression
                s_expression = s_expression & "    (<> [" & constraintFactor & "] "
                s_expression = s_expression & constraintLevels(j - firstConstraintFactorsCol) & "))" & vbLf
            End If
        Next j
    Next i

    kinsokuMatrix2expression = True
End Function

' ��������̓ǂݎ��
Function GetConstraints(ByRef constraints As String) As Boolean
    GetConstraints = False
    
    Dim constraintSheet As Worksheet
    Dim idCol As Long
    Dim constrainCol As Long
    Dim binaryConstrainRow As Long
    Dim i As Long
    
    Set constraintSheet = Worksheets(constraintSheetName)
    binaryConstrainRow = constraintSheet.Range("������������").row
    constrainCol = constraintSheet.Range("������������").Column
    idCol = constrainCol - 1
    
    ' �܂��A��������\����̎�����������𒊏o
    constraints = constraintSheet.Range("������������").Value
    
    ' �S�Ă̎������������A��
    For i = binaryConstrainRow + 1 To constraintSheet.Range("Alloy�ɂ�錟�ؗp�\��").row - 2
        If InStr(constraintSheet.Cells(i, idCol).Value, kinsokuMatrixSheetBaseName) <> 0 Then
            constraints = constraints & constraintSheet.Cells(i, constrainCol).Value
        Else
            Exit For
        End If
    Next i
    
    constraints = constraints & constraintSheet.Range("���R�L�q����").Value
    
    GetConstraints = True
End Function

' S��(CIT-BACH�p)��������̓ǂݎ��
Function GetSConstraints(ByRef constraints As String) As Boolean
    GetSConstraints = False
    
    Dim constraintSheet As Worksheet
    Dim idCol As Long
    Dim sConstrainCol As Long
    Dim binaryConstrainRow As Long
    Dim i As Long
    
    Set constraintSheet = Worksheets(constraintSheetName)
    binaryConstrainRow = constraintSheet.Range("������������").row
    sConstrainCol = constraintSheet.Range("������������").Column + 1
    idCol = sConstrainCol - 2
    
    ' �܂��A��������\����̎�����������𒊏o
    constraints = constraintSheet.Cells(binaryConstrainRow, sConstrainCol).Value
    
    ' �S�Ă̎������������A��
    For i = binaryConstrainRow + 1 To constraintSheet.Range("Alloy�ɂ�錟�ؗp�\��").row - 2
        If InStr(constraintSheet.Cells(i, idCol).Value, kinsokuMatrixSheetBaseName) <> 0 Then
            constraints = constraints & constraintSheet.Cells(i, sConstrainCol).Value
        Else
            Exit For
        End If
    Next i
    
    constraints = constraints & constraintSheet.Cells(constraintSheet.Range("���R�L�q����").row, sConstrainCol).Value
    
    GetSConstraints = True
End Function


