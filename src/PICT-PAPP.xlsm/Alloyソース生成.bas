Attribute VB_Name = "Alloy�\�[�X����"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' Tool���g�ݍ��킹���o�͂��Ȃ������S�Ă�Pair�ɂ��āA���ꂪ(�Ԑ�)�֑��ł��邱�Ƃ��m�F����Alloy�̃\�[�X�𐶐�����
Function createAlloySrc(srcWorkbook As Workbook, ByRef alloySrc As String) As Boolean
    createAlloySrc = False
    Dim factorNum As Long
    Dim levelNum As Long
    Dim flSet As String
    Dim defSystem As String
    Dim pict As String
    Dim alloy As String
    Dim factorNames() As String
    Dim levelLists()
    Dim pairs() As String
    Dim dicDuplication
    Dim predicate As String
    Dim n As Long
    Dim levelName As String
    If FLTable2array(srcWorkbook.Worksheets(FLtblSheetName), factorNames, levelLists) Then
        ' �������̏Փ˂�����邽�߁A�������ʂɉ����q�̒��Ŏg���Ă��邩������
        Call generateDicDuplication(factorNames, levelLists, dicDuplication)
        ' �e���q���̎�蓾�鐅���̏W�������q���ɓ������O�̏W���Ƃ��Ē�`����
        defSystem = "sig �V�X�e�� {" & vbLf
        For factorNum = LBound(factorNames) To UBound(factorNames)
            flSet = flSet & "enum " & factorNames(factorNum) & " {"
            defSystem = defSystem & vbTab & factorNames(factorNum) & alloyLevelSuffix & ":one " & factorNames(factorNum) & "," & vbLf
            For levelNum = LBound(levelLists(factorNum)) To UBound(levelLists(factorNum))
                If levelNum > 0 Then flSet = flSet & ", "
                levelName = levelLists(factorNum)(levelNum)
                If dicDuplication.Item(levelName) > 1 Then ' ���������Փ˂��Ă���ꍇ��"_[���q��]"��A������
                    levelName = levelName & "_" & factorNames(factorNum)
                End If
                flSet = flSet & levelName
            Next levelNum
            flSet = flSet & "}" & vbLf
        Next factorNum
        defSystem = Left(defSystem, Len(defSystem) - 2)
        defSystem = defSystem & vbLf & "}" & vbLf
    Else
        MsgBox "���q�E�����̎擾�Ɏ��s���܂���"
    End If
    If GetConstraints(pict) Then
        If pict2alloy(pict, dicDuplication, alloy) Then
        Else
            MsgBox "PICT���񎮂���alloy�`���ւ̕ϊ��Ɏ��s���܂���"
        End If
    Else
        MsgBox "PICT���񎮂̎擾�Ɏ��s���܂���"
    End If
    If Not ExistsWorksheet(srcWorkbook, mappedRoundRobinSheetName) Then
        MsgBox "ID�}�b�s���O�ςݑ�������\�����݂��܂���BTool�����s����Ǝ������������̂ŁA��Ɏ��s���Ă��������B"
        Exit Function
    End If
    If pairsWithoutTestcase(srcWorkbook.Worksheets(mappedRoundRobinSheetName), dicDuplication, pairs) Then
        predicate = "pred �g������Ԃ����݂���(s:�V�X�e��) {" & vbLf
        For n = LBound(pairs) To UBound(pairs)
            If n > 0 Then
                predicate = predicate & " ||" & vbLf
            End If
            predicate = predicate & vbTab & pairs(n)
        Next n
        predicate = predicate & vbLf & "}" & vbLf
    Else
        MsgBox "�e�X�g�P�[�X�̑��݂��Ȃ�Pair�W���̎擾�Ɏ��s���܂���"
    End If
    alloySrc = alloySrc & flSet
    alloySrc = alloySrc & defSystem
    alloySrc = alloySrc & alloy
    alloySrc = alloySrc & predicate
    alloySrc = alloySrc & vbLf & alloyExec
    createAlloySrc = True
End Function

' Tool���g�ݍ��킹�ɏo�͂��Ȃ������S�Ă�Pair�ɂ��Ă̏������W����B
' ���ꂪ�S�ċ֑��ł��邱�Ƃ�Alloy�Ɋm�F�����邽�߁AAlloy�ɓs���̗ǂ������Ō��ʂ�������B
Function pairsWithoutTestcase(roundRobinSheet As Worksheet, dicDuplication, ByRef pair() As String) As Boolean
    pairsWithoutTestcase = False
    Dim i As Long
    Dim j As Long
    Dim vN As Long ' ��������\�̐��������̉��Ԗڂ̃}�X��
    Dim hN As Long ' ��������\�̐��������̉��Ԗڂ̃}�X��
    Dim cellStr As String
    Dim counter As Long
    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim factorNum As Long
    Dim levelNum As Long
    Dim factorRow As Long
    Dim levelRow As Long
    Dim factorCol As Long
    Dim levelCol As Long
    Dim factor1Name As String
    Dim level1Name As String
    Dim factor2Name As String
    Dim level2Name As String
    
    Call getMaxRowAndCol(roundRobinSheet, MaxRow, MaxCol)
    
    factorRow = offsetRows + 1
    levelRow = offsetRows + 2
    factorCol = offsetColumns + 1
    levelCol = offsetColumns + 2
    counter = 0
    
    For i = levelRow + 1 To MaxRow
        For j = levelCol + 1 To MaxCol
            vN = i - levelRow
            hN = j - levelCol
            If vN < hN Then ' ��{�I�ɑΊp�����E�ゾ����������Ηǂ��͂������A�A�A
                Select Case roundRobinSheet.Cells(i, j).Value
                    Case "�\" ' �Ίp����̎��g�Ƃ̑g�ݍ��킹�ŁA�Ӗ��������̂Ŗ���
                        If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "�\" Then
                            MsgBox "�����G���A�\�ɂ��āA�Ίp���Ő��ΏۂɂȂ��Ă��܂���B�������Ă��Ȃ��o���̃Z���̔w�i�F��Ԃɂ��܂����B"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' �ԐF
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' �ԐF
                        End If
                    Case "�~", "�H", "?", "" ' �e�X�g�P�[�X�����݂��Ă��Ȃ��ꍇ
                        cellStr = roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value
                        If Not (cellStr = "�~" Or cellStr = "?" Or cellStr = "") Then
                            MsgBox "�֑��܂��̓e�X�g�P�[�X�����̃y�A���A�Ίp���Ő��ΏۂɂȂ��Ă��܂���B�������Ă��Ȃ��o���̃Z���̔w�i�F��Ԃɂ��܂����B"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' �ԐF
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' �ԐF
                        End If
                        counter = counter + 1
                        ReDim Preserve pair(counter - 1)
                        factor1Name = roundRobinSheet.Cells(factorRow, j).Value
                        level1Name = roundRobinSheet.Cells(levelRow, j).Value
                        factor2Name = roundRobinSheet.Cells(i, factorCol).Value
                        level2Name = roundRobinSheet.Cells(i, levelCol).Value
                        If dicDuplication.Item(level1Name) > 1 Then ' ���������Փ˂��Ă���ꍇ��"_[���q��]"��A������
                            level1Name = level1Name & "_" & factor1Name
                        End If
                        If dicDuplication.Item(level2Name) > 1 Then ' ���������Փ˂��Ă���ꍇ��"_[���q��]"��A������
                            level2Name = level2Name & "_" & factor2Name
                        End If
                        pair(counter - 1) = "s." & factor1Name & alloyLevelSuffix & " = " & level1Name & " && s." & _
                                    factor2Name & alloyLevelSuffix & " = " & level2Name
                End Select
            Else ' �Ίp����荶�����e�X�g���ڂ������Ǝv����y�A�ɂ��Ă������ΏۂɂȂ��Ă��邩�ǂ����A�`�F�b�N��������
                Select Case roundRobinSheet.Cells(i, j).Value
                    Case "�\" ' �Ίp����̎��g�Ƃ̑g�ݍ��킹�ŁA�Ӗ��������̂Ŗ���
                        If Not roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value = "�\" Then
                            MsgBox "�����G���A�\�ɂ��āA�Ίp���Ő��ΏۂɂȂ��Ă��܂���B�������Ă��Ȃ��o���̃Z���̔w�i�F��Ԃɂ��܂����B"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' �ԐF
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' �ԐF
                        End If
                    Case "�~", "�H", "?", "" ' �e�X�g�P�[�X�����݂��Ă��Ȃ��ꍇ
                        cellStr = roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Value
                        If Not (cellStr = "�~" Or cellStr = "?" Or cellStr = "") Then
                            MsgBox "�֑��܂��̓e�X�g�P�[�X�����̃y�A���A�Ίp���Ő��ΏۂɂȂ��Ă��܂���B�������Ă��Ȃ��o���̃Z���̔w�i�F��Ԃɂ��܂����B"
                            roundRobinSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) ' �ԐF
                            roundRobinSheet.Cells(hN + levelRow, vN + levelCol).Interior.Color = RGB(255, 0, 0) ' �ԐF
                        End If
                End Select
            End If
        Next j
    Next i
       
    pairsWithoutTestcase = True
End Function

' �Z�����̐���\������alloy�Ō��؂ł���`���ɕϊ�����
Function pict2alloy(pict As String, dicDuplication, ByRef alloy As String) As Boolean
    pict2alloy = False
    Dim oReg As Object
    Dim match As Object
    Dim matches As Object
    Dim match2 As Object
    Dim matches2 As Object
    Dim condition As Variant
    Dim conditionStr As String
    Dim i As Long
    Dim factorName As String
    Dim levelName As String
    
    Set oReg = CreateObject("VBScript.Regexp")
    oReg.Pattern = "IF *(\[.+)THEN(.+)<>([^;]+);"
    oReg.Pattern = "IF *(\[.+\] *= *"".+"" *)THEN(.+)<>([^;]+);"
    oReg.Pattern = "IF *(\[.+)THEN *\[(.+)\] *<> *""(.+)"" *;"
    oReg.Global = True
    Set match = oReg.Execute(pict)

    alloy = "{" & vbLf
    For Each matches In match
        oReg.Pattern = " *\[(.+)\] *= *""(.+)"" *"
        condition = Split(matches.Submatches(0), " AND ")
        conditionStr = ""
        For i = LBound(condition) To UBound(condition)
            Set match2 = oReg.Execute(condition(i))
            If i > LBound(condition) Then conditionStr = conditionStr & " and "
            factorName = match2(0).Submatches(0)
            levelName = match2(0).Submatches(1)
            If dicDuplication.Item(levelName) > 1 Then ' ���������Փ˂��Ă���ꍇ��"_[���q��]"��A������
                levelName = levelName & "_" & factorName
            End If
            conditionStr = conditionStr & factorName & alloyLevelSuffix & "=" & levelName
        Next i
        factorName = matches.Submatches(1)
        levelName = matches.Submatches(2)
        If dicDuplication.Item(levelName) > 1 Then ' ���������Փ˂��Ă���ꍇ��"_[���q��]"��A������
            levelName = levelName & "_" & factorName
        End If
        alloy = alloy & vbTab & conditionStr & "=>"
        alloy = alloy & factorName & alloyLevelSuffix & "!="
        alloy = alloy & levelName & vbLf
    Next
    alloy = alloy & vbLf & "}" & vbLf
    
    pict2alloy = True
End Function

