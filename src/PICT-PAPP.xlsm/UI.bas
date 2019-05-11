Attribute VB_Name = "UI"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

Public Const offsetRows As Integer = 1 ' �\���o�͂���ہA�ŏ㕔�ɉ��s���󂯂邩
Public Const offsetColumns As Integer = 1 ' �\���o�͂���ہA�ō����ɉ�����󂯂邩
Public Const alloyLevelSuffix As String = "�̐���" ' alloy�p�ň��q�ɑΉ�����v�f��\�����߂Ɏg�p���閼�O�Ƃ��Ďg��suffix
Public Const alloyExec As String = "run �g������Ԃ����݂��� for 1 but exactly 1 �V�X�e��" ' alloy�ł̎��s�w��

Public Const alloySrcFileName As String = "�y�A�����݂��Ȃ��ėǂ����Ƃ����؂���alloy�\�[�X.als"
Public Const pictInFileName As String = "PICTin.txt"
Public Const pictOutFileName As String = "PICTout.txt"
Public Const citBachInFileName As String = "CitBachIn.txt"
Public Const citBachOutFileName As String = "CitBachOut.txt"

Public Const controlSheetName As String = "�k�����̎w�����ݒ�l"
Public Const tuplelSheetName As String = "�S�g�ݍ��킹"
Public Const coverageSheetName As String = "�ԗ���"
Public Const roundRobinSheetName As String = "��������\"
Public Const mappedRoundRobinSheetName As String = "ID�}�b�s���O�ςݑ�������\"
Public Const pairListSheetName As String = "�y�A�E���X�g"
Public Const pairListFlg As Boolean = False '�y�A�E���X�g�𐶐�����
Public Const toolOutSheetName As String = "�c�[���̐�������"
Public Const testCaseSheetName As String = "�e�X�g�P�[�X"
Public Const testDataSheetName As String = "�e�X�g�f�[�^"
Public Const FLtblSheetName As String = "���q�E����"
Public Const FLLVSheetName As String = "���q�E�����E�����l"
Public Const constraintSheetName As String = "����L�q"
Public Const kinsokuMatrixSheetBaseName As String = "�����ԋ֑��\"
Public Const kinsokuMatrixSheetMax As Integer = 100 ' �����ԋ֑��\�V�[�g�̏����

Public maskSymbol As String ' MASK��Ԃ�\���V���{��
Public protectPassword As String ' �V�[�g�ی�Ɏg���p�X���[�h
Public toolName As String ' ���sTool��
Public pictCmdOption As String ' PICT�R�}���h�I�v�V����
Public citBachCmdOption As String ' CIT-BACH�R�}���h�I�v�V����

Public conditionFactors() As String
Public constraintFactors As String
Public publicFactorNames() As String

' Step0)
Sub MASK�����̎����}��()
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If insertMaskSymbol(ThisWorkbook) Then
        MsgBox "�������܂����B"
    Else
        MsgBox "�����Ɏ��s���܂����B"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step1)
Sub ��������()
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    ' ���̏����Ő�������鑍������\�V�[�g�������ł͂Ȃ����Ƃ��m�F����
    If ExistsWorksheet(ThisWorkbook, roundRobinSheetName) Then ' ��������\�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "��������\�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & roundRobinSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If createRoundRobinTable(ThisWorkbook) Then
        MsgBox "�������܂����B"
    Else
        MsgBox "�����Ɏ��s���܂����B"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step1')
Sub �����֑��}�g���N�X����()
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If createKinsokuMatrix(ThisWorkbook) Then
        MsgBox "�������܂����B"
    Else
        MsgBox "�����Ɏ��s���܂����B"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step2)
Sub ���񎩓�����()
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If generateConstraintExpression(ThisWorkbook) Then
        MsgBox "���񎩓������������܂����B"
    Else
        MsgBox "���񎩓����������Ɏ��s���܂����B"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step3) Tool���s�̏����A���s�A���ʂ̎�荞�݁A�������̕��͗p�V�[�g�̐����܂ł̈�A�̗�����쓮����
Sub Tool���s()
    Dim paramNames() As String
    Dim tuples()
    
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    ' ���̏����Ő��������V�[�g���̃V�[�g�������ł͂Ȃ����Ƃ��m�F����
    If ExistsWorksheet(ThisWorkbook, mappedRoundRobinSheetName) Then ' ID�}�b�s���O�ςݑ�������\�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "ID�}�b�s���O�ςݑ�������\�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & mappedRoundRobinSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, pairListSheetName) Then ' �y�A�E���X�g�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�y�A�E���X�g�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & pairListSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' �c�[���̐������ʃV�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�c�[���̐������ʃV�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & toolOutSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, coverageSheetName) Then ' �ԗ����V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�ԗ����V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & coverageSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
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
        MsgBox "���sTool�����������ݒ肳��Ă��Ȃ��̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    ThisWorkbook.Sheets(toolOutSheetName).Activate
' �ȉ��̏������p�����Ď��{���Ă��܂��ƕ֗��ł��邪�A����Ńf�[�^�T�C�Y�ɂ���Ă͏������Ԃ������邱�Ƃ�����B
' �v���O���X�o�[�Ȃǂ�t���āA���f���ł���悤�ɂ���Ɨǂ���������Ȃ��B(�v����)
'    On Error GoTo ErrLabel
'    Application.ScreenUpdating = False
'    If Not analysis(ThisWorkbook, paramNames, tuples) Then
'        Application.ScreenUpdating = True
'        Exit Sub
'    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Tool��PICT���I������Ă���ꍇ�AStep3��PICT���s�̏����A���s�A���ʂ̎�荞��
Function PictFEP(ByRef paramNames() As String, ByRef tuples()) As Boolean
    PictFEP = False
    Dim pairwiseStr As String
    
    If Not getToolInputFile(ThisWorkbook, pictInFileName, citBachInFileName) Then
        MsgBox "Tool���͗p�t�@�C���̐����Ɏ��s���܂����̂ŁB�����𒆎~���܂��B"
        Exit Function
    End If
    If Not execPict(pictInFileName, pictOutFileName) Then
        MsgBox "PICT�̎��s�����Ɏ��s���܂����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    If Not inputUtf8(pictOutFileName, pairwiseStr) Then
        MsgBox "PICT�̎��s���ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    If Not textTable2array(pairwiseStr, vbTab, paramNames(), tuples) Then
        MsgBox "PICT�̎��s���ʂ���͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    ' PICT�o�͌��ʗp�V�[�g�Ɍ��ʂ���������
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "PICT�̎��s���ʂ��V�[�g�ɏo�͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    
    PictFEP = True
    
ErrLabel:
    Application.ScreenUpdating = True
End Function

' Tool��CIT-BACH���I������Ă���ꍇ�AStep3��CIT-BACH���s�̏����A���s�A���ʂ̎�荞��
Function CitBachFEP(ByRef paramNames() As String, ByRef tuples()) As Boolean
    CitBachFEP = False
    Dim pairwiseStr As String
    Dim lengthLine1 As Long
    Dim line1 As String
    
    If Not getToolInputFile(ThisWorkbook, pictInFileName, citBachInFileName) Then
        MsgBox "Tool���͗p�t�@�C���̐����Ɏ��s���܂����̂ŁB�����𒆎~���܂��B"
        Exit Function
    End If
    If Not execCitBach(citBachInFileName, citBachOutFileName) Then
        MsgBox "CIT-BACH�̎��s�����Ɏ��s���܂����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    ' Pairwise�̐������ʃt�@�C���̓ǂݍ���
    If Not inputFile(citBachOutFileName, pairwiseStr) Then
        MsgBox "Pairwise�̐������ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    
    lengthLine1 = InStr(pairwiseStr, vbLf) - 1
    line1 = Left(pairwiseStr, lengthLine1)
    
    If Left(line1, 1) = "#" And InStr(line1, vbTab) = 0 Then ' CIT-BACH�̏o�͂Ǝv����
        pairwiseStr = Mid(pairwiseStr, lengthLine1 + 2)
    Else
        MsgBox "CIT-BACH�̐������ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    If Not textTable2array(pairwiseStr, ",", paramNames(), tuples) Then
        MsgBox "CIT-BACH�̎��s���ʂ���͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    ' CIT-BACH�o�͌��ʗp�V�[�g�Ɍ��ʂ���������
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "CIT-BACH�̎��s���ʂ��V�[�g�ɏo�͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    
    CitBachFEP = True

ErrLabel:
    Application.ScreenUpdating = True
End Function

' Step3') Pairwise�̃t�@�C�������͍ς�ł��邱�Ƃ�O��ɁA���̎��s�ς݂̌��ʃt�@�C�����w�肵�ăV�[�g�Ɏ�荞��
Sub Tool���s�ς݌��ʃt�@�C���̃V�[�g�ւ̓ǂݍ��ݏ���()
    Dim pairwiseStr As String
    Dim lengthLine1 As Long
    Dim line1 As String
    Dim pwFileFullName As String
    Dim paramNames() As String
    Dim tuples()
    Dim delimiter As String
    
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    ' ���̏����Ő��������V�[�g���̃V�[�g�������ł͂Ȃ����Ƃ��m�F����
    If ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' �c�[���̐������ʃV�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�c�[���̐������ʃV�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & toolOutSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If

    ' Pairwise�̐������ʃt�@�C���̓ǂݍ���
    If Not inputFile(pwFileFullName, pairwiseStr) Then
        MsgBox "Pairwise�̐������ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    
    lengthLine1 = InStr(pairwiseStr, vbLf) - 1
    line1 = Left(pairwiseStr, lengthLine1)
    
    If Left(line1, 1) = "#" And InStr(line1, vbTab) = 0 Then ' CIT-BACH�̏o�͂Ǝv����
        pairwiseStr = Mid(pairwiseStr, lengthLine1 + 2)
        delimiter = ","
    Else ' PICT�̏o�͂Ǝv����i������x�ǂݒ����j
        If Not inputUtf8(pwFileFullName, pairwiseStr) Then ' �t�@�C�����͊��Ɏw�肳��Ă��铯������
            MsgBox "PICT�̎��s���ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
            Exit Sub
        End If
        delimiter = vbTab
    End If
    
    If Not textTable2array(pairwiseStr, delimiter, paramNames(), tuples) Then
        MsgBox "Tool�̎��s���ʂ���͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    ' Tool�o�͌��ʗp�V�[�g�Ɍ��ʂ���������
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "Tool�̎��s���ʂ��V�[�g�ɏo�͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    ThisWorkbook.Sheets(toolOutSheetName).Activate

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step4) �����ς݌��ʂ̃V�[�g�܂��͂���𕡐������ҏW�ς݃V�[�g�̏�񂩂�A��������\�̕����V�[�g��ID���}�b�v�������̓V�[�g���쐬����
Sub Tool���ʃV�[�g�܂��͕ҏW�ς݃V�[�g���番�͂܂ł̏���()
    Dim paramNames() As String
    Dim tuples()
    Dim srcSheet As Worksheet
    Dim pairListSheet As Worksheet
    Dim toolOutStr As String
    
    Debug.Print Time & " - Tool���ʃV�[�g�܂��͕ҏW�ς݃V�[�g���番�͂܂ł̏����J�n"
    
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    ' ���̏����Ő��������V�[�g���̃V�[�g�������ł͂Ȃ����Ƃ��m�F����
    If ExistsWorksheet(ThisWorkbook, mappedRoundRobinSheetName) Then ' ID�}�b�s���O�ςݑ�������\�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "ID�}�b�s���O�ςݑ�������\�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & mappedRoundRobinSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, pairListSheetName) Then ' �y�A�E���X�g�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�y�A�E���X�g�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & pairListSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, coverageSheetName) Then ' �ԗ����V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�ԗ����V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & coverageSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If

    If ExistsWorksheet(ThisWorkbook, testCaseSheetName) Then ' �e�X�g�P�[�X�̃V�[�g�����t����
        Set srcSheet = ThisWorkbook.Sheets(testCaseSheetName)
    ElseIf ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' �c�[���̐������ʂ̃V�[�g�����t����
        Set srcSheet = ThisWorkbook.Sheets(toolOutSheetName)
    Else
        MsgBox "�u" & testCaseSheetName & "�v�V�[�g�܂��́u" & toolOutSheetName & "�v�V�[�g�����݂��邱�Ƃ��K�v�ł��B���݂��Ȃ������̂ŏ����𒆎~���܂��B"
        Exit Sub
    End If

    ' �e�X�g�P�[�X�̃V�[�g�܂��̓c�[���̐������ʂ̃V�[�g������s���ʂ̓ǂݍ���
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not testCaseSheet2array(srcSheet, paramNames(), tuples) Then
        MsgBox "�e�X�g�P�[�X�̉�͂Ɏ��s�����̂ŁA�����𒆎~���܂��B"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    If Not analysis(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Debug.Print Time & " - Tool���ʃV�[�g�܂��͕ҏW�ς݃V�[�g���番�͂܂ł̏����I��"
    MsgBox "�������܂����B"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step4') �O���c�[���ɂ�鐶�����ʂ̃t�@�C�������ɑ��݂��Ă��邱�Ƃ�O��ɁA���̎��s�ς݂̌��ʃt�@�C�����w�肵�Ď�荞�݁A��������\�̕����V�[�g��ID���}�b�v�������̓V�[�g���쐬����
Sub Tool�������ʂ̃t�@�C�����番�͂܂ł̏���()
    Dim pairwiseStr As String
    Dim lengthLine1 As Long
    Dim line1 As String
    Dim pwFileFullName As String
    Dim paramNames() As String
    Dim tuples()
    Dim delimiter As String
    
    Debug.Print Time & " - Tool�������ʂ̃t�@�C�����番�͂܂ł̏����J�n"
    
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    ' ���̏����Ő��������V�[�g���̃V�[�g�������ł͂Ȃ����Ƃ��m�F����
    If ExistsWorksheet(ThisWorkbook, mappedRoundRobinSheetName) Then ' ID�}�b�s���O�ςݑ�������\�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "ID�}�b�s���O�ςݑ�������\�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & mappedRoundRobinSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, pairListSheetName) Then ' �y�A�E���X�g�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�y�A�E���X�g�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & pairListSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' �c�[���̐������ʃV�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�c�[���̐������ʃV�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & toolOutSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, coverageSheetName) Then ' �ԗ����V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�ԗ����V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & coverageSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If

    ' Pairwise�̐������ʃt�@�C���̓ǂݍ���
    If Not inputFile(pwFileFullName, pairwiseStr) Then
        MsgBox "Pairwise�̐������ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    
    lengthLine1 = InStr(pairwiseStr, vbLf) - 1
    line1 = Left(pairwiseStr, lengthLine1)
    
    If Left(line1, 1) = "#" And InStr(line1, vbTab) = 0 Then ' CIT-BACH�̏o�͂Ǝv����
        pairwiseStr = Mid(pairwiseStr, lengthLine1 + 2)
        delimiter = ","
    Else ' PICT�̏o�͂Ǝv����i������x�ǂݒ����j
        If Not inputUtf8(pwFileFullName, pairwiseStr) Then ' �t�@�C�����͊��Ɏw�肳��Ă��铯������
            MsgBox "PICT�̎��s���ʃt�@�C����ǂݎ�邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
            Exit Sub
        End If
        delimiter = vbTab
    End If
    
    If Not textTable2array(pairwiseStr, delimiter, paramNames(), tuples) Then
        MsgBox "Tool�̎��s���ʂ���͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    ' Tool�o�͌��ʗp�V�[�g�Ɍ��ʂ���������
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInToolOutSheets(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        MsgBox "Tool�̎��s���ʂ��V�[�g�ɏo�͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Sub
    End If
    
    If Not analysis(ThisWorkbook, paramNames, tuples) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Debug.Print Time & " - Tool�������ʂ̃t�@�C�����番�͂܂ł̏����I��"
    MsgBox "�������܂����B"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step5)
Sub �y�A�����݂��Ȃ��ėǂ����Ƃ����؂���alloy�\�[�X�̐���()
    Dim alloySrc As String
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    If createAlloySrc(ThisWorkbook, alloySrc) Then
        ThisWorkbook.Worksheets(constraintSheetName).Range("Alloy�ɂ�錟�ؗp�\��").Value = alloySrc
        ThisWorkbook.Worksheets(constraintSheetName).Activate
        ThisWorkbook.Worksheets(constraintSheetName).Range("Alloy�ɂ�錟�ؗp�\��").Select
    Else
        MsgBox "�y�A�����݂��Ȃ��ėǂ����Ƃ����؂���alloy�\�[�X�̐����Ɏ��s���܂����B"
        Exit Sub
    End If
    If outputUtf8(alloySrc, alloySrcFileName) Then
        If Not execAlloy(alloySrcFileName) Then
            MsgBox "Alloy�̋N���Ɏ��s���܂����B"
        End If
    Else
        MsgBox "�y�A�����݂��Ȃ��ėǂ����Ƃ����؂���alloy�\�[�X�t�@�C���̏o�͂Ɏ��s���܂����B"
    End If
End Sub

' Step6)
Sub ���q_����_�����l�ݒ�\�̐���()
    Dim alloySrc As String
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    If ExistsWorksheet(ThisWorkbook, FLLVSheetName) Then ' ���q�E�����E�����l�ݒ�\�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & FLLVSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not createFLLVSheet(ThisWorkbook) Then
        MsgBox "���q�E�����E�����l�ݒ�\�V�[�g�̐����Ɏ��s���܂����B"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    MsgBox "�������܂����B"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' Step6) ���q�E�����E�����l�ݒ�\�ɏ]���Đ����ς�Pairwise���ʂ̃V�[�g�̐����𐅏��l�ɒu������
Sub �����𐅏��l�ɒu��()
    Dim paramNames() As String
    Dim tuples()
    Dim FLLVSheet As Worksheet
    Dim srcSheet As Worksheet
    Dim testDataSheet As Worksheet
    Dim toolOutStr As String
    Dim dicFLLV As Object
    
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    
    If ExistsWorksheet(ThisWorkbook, testDataSheetName) Then ' �e�X�g�f�[�^�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�e�X�g�f�[�^�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & testDataSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    
    ' ���q�E�����E�����l�ݒ�\�V�[�g�����t����
    If ExistsWorksheet(ThisWorkbook, FLLVSheetName) Then
        Set FLLVSheet = ThisWorkbook.Sheets(FLLVSheetName)
    Else
        MsgBox "�u" & FLLVSheetName & "�v�V�[�g�����݂��邱�Ƃ��K�v�ł��B���݂��Ȃ������̂ŏ����𒆎~���܂��B"
        Exit Sub
    End If

    If ExistsWorksheet(ThisWorkbook, testCaseSheetName) Then ' �e�X�g�P�[�X�̃V�[�g�����t����
        Set srcSheet = ThisWorkbook.Sheets(testCaseSheetName)
    ElseIf ExistsWorksheet(ThisWorkbook, toolOutSheetName) Then ' �c�[���̐������ʂ̃V�[�g�����t����
        Set srcSheet = ThisWorkbook.Sheets(toolOutSheetName)
    Else
        MsgBox "�u" & testCaseSheetName & "�v�V�[�g�܂��́u" & toolOutSheetName & "�v�V�[�g�����݂��邱�Ƃ��K�v�ł��B���݂��Ȃ������̂ŏ����𒆎~���܂��B"
        Exit Sub
    End If
    ThisWorkbook.Worksheets(srcSheet.name).Copy Before:=ThisWorkbook.Worksheets(srcSheet.name)
    ActiveSheet.name = testDataSheetName
    Set testDataSheet = ThisWorkbook.Worksheets(testDataSheetName)

    ' ���q�E�����̑΂̖��O���琅���l�ւ̎������쐬����
    If Not makeFLLVDictionary(FLLVSheet, dicFLLV) Then
        Exit Sub
    End If
    
    ' �����̏��Ɋ�Â��ăe�X�g�f�[�^�̃V�[�g�̐������𐅏��l�ɒu��
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If Not fillInTestDataSheet(dicFLLV, testDataSheet) Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    MsgBox "�������܂����B"

ErrLabel:
    Application.ScreenUpdating = True
End Sub

' ���܂�
Sub �S�g�ݍ��킹()
    If GetSetValues = False Then
        MsgBox "�ݒ�l�̓ǂݎ��Ɏ��s�����̂ŏI�����܂��B"
        Exit Sub
    End If
    If ExistsWorksheet(ThisWorkbook, tuplelSheetName) Then ' �S�g�ݍ��킹�V�[�g�̖��O�����Ɏg���Ă��邩�H
        MsgBox "�S�g�ݍ��킹�V�[�g�̖��O�Ƃ��Ďw�肳��Ă���u" & tuplelSheetName & "�v�����ɑ��݂��Ă��܂��B�폜���邩�������Ă��������B"
        Exit Sub
    End If
    On Error GoTo ErrLabel
    Application.ScreenUpdating = False
    If fillInTupleSheets(ThisWorkbook) Then
        MsgBox "�������܂����B"
    Else
        MsgBox "�����Ɏ��s���܂����B"
    End If
ErrLabel:
    Application.ScreenUpdating = True
End Sub

' ��ƃt�H���_��Explorer�ŊJ��
Sub ��ƃt�H���_��Explorer�ŊJ��()
    
    Call Shell("explorer.exe " & Worksheets(controlSheetName).Range("��ƃp�X").Value, vbNormalFocus)

End Sub

Sub getWorkingPath()
    Worksheets(controlSheetName).Range("��ƃp�X").Value = getPath("�����Ώۃt�@�C�����i�[����Ă���t�H���_��I��")
End Sub

Sub getPictPath()
    Worksheets(controlSheetName).Range("PICT�t���p�X").Value = getPictExePath("pict.exe�t�@�C����I��")
End Sub

Sub getCitBachPath()
    Worksheets(controlSheetName).Range("CIT_BACH�t���p�X").Value = getCitBachJarPath("cit-bach.jar�t�@�C����I��")
End Sub

Sub getAlloyPath()
    Worksheets(controlSheetName).Range("Alloy�t���p�X").Value = getAlloyJarPath("alloy.jar�t�@�C����I��")
End Sub

' ���[�U�Ƀt�H���_��I�������āA���̃p�X�𓾂�
Function getPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    If fileDialog.Show = -1 Then
        getPath = fileDialog.SelectedItems(1)
    End If
End Function

' ���[�U��pict.exe�t�@�C����I�������āA���̃t���p�X�𓾂�
Function getPictExePath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "pict.exe�t�@�C��", "*.exe"
    If fileDialog.Show = -1 Then
        getPictExePath = fileDialog.SelectedItems(1)
    End If
End Function

' ���[�U��cit-bach.jar�t�@�C����I�������āA���̃t���p�X�𓾂�
Function getCitBachJarPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "cit-bach.jar�t�@�C��", "*.jar"
    If fileDialog.Show = -1 Then
        getCitBachJarPath = fileDialog.SelectedItems(1)
    End If
End Function

' ���[�U��alloy.jar�t�@�C����I�������āA���̃t���p�X�𓾂�
Function getAlloyJarPath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "alloy.jar�t�@�C��", "*.jar"
    If fileDialog.Show = -1 Then
        getAlloyJarPath = fileDialog.SelectedItems(1)
    End If
End Function

' ���[�U�Ƀt�@�C����I�������āA���̃t���p�X�𓾂�
Function getFilePath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.InitialFileName = Worksheets(controlSheetName).Range("��ƃp�X").Value & "\" & pictOutFileName
    fileDialog.Filters.Add title, "*.*"
    If fileDialog.Show = -1 Then
        getFilePath = fileDialog.SelectedItems(1)
    End If
End Function

' �ݒ�l�̓ǂݎ��
Function GetSetValues() As Boolean
    GetSetValues = False
    
    maskSymbol = Worksheets(controlSheetName).Range("MASK��Ԃ�\���V���{��").Value
    protectPassword = Worksheets(controlSheetName).Range("�V�[�g�ی�Ɏg���p�X���[�h").Value
    toolName = Worksheets(controlSheetName).Range("���sTool��").Value
    If maskSymbol = "" Then
        maskSymbol = "mask"
        Worksheets(controlSheetName).Range("MASK��Ԃ�\���V���{��").Value = maskSymbol
        MsgBox "MASK��Ԃ�\���V���{�����ݒ肳��Ă��Ȃ��̂ŁAmask�Ƃ��܂����B"
    End If
    If protectPassword = "" Then
        protectPassword = "password"
        Worksheets(controlSheetName).Range("�V�[�g�ی�Ɏg���p�X���[�h").Value = protectPassword
        MsgBox "�V�[�g�ی�Ɏg���p�X���[�h���ݒ肳��Ă��Ȃ��̂ŁApassword�Ƃ��܂����B"
    End If
    If Not (toolName = "PICT" Or toolName = "CIT-BACH") Then
        toolName = "CIT-BACH"
        Worksheets(controlSheetName).Range("���sTool��").Value = toolName
        MsgBox "���sTool�̑I�����������ݒ肳��Ă��Ȃ��̂ŁACIT-BACH�Ƃ��܂����B"
    End If
    pictCmdOption = Worksheets(controlSheetName).Range("PICT�I�v�V����").Value
    citBachCmdOption = Worksheets(controlSheetName).Range("CIT_BACH�I�v�V����").Value
    
    GetSetValues = True
End Function

