Attribute VB_Name = "�c�[�����s�E���o��"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

' Tool���͗p�̃\�[�X���t�@�C���ɏo�͂���
Function getToolInputFile(srcBook As Workbook, inFileName As String, inFileNameS As String) As Boolean
    getToolInputFile = False
                
    Dim pictSrc As String
    Dim citBachSrc As String

    ' TOOL���͗p�̃\�[�X�𐶐����鏈�����Ăяo��
    If Not createToolInputSrc(srcBook, pictSrc, citBachSrc) Then
        MsgBox "Tool�̃\�[�X�𐶐����邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If
    
    ' PICT���͗p�̃\�[�X���t�@�C���ɏo�͂��鏈�����Ăяo��
    If Not outputUtf8(pictSrc, inFileName) Then
        MsgBox "PICT�̃\�[�X���t�@�C���o�͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If

    ' CIT-BACH���͗p�̃\�[�X���t�@�C���ɏo�͂��鏈�����Ăяo��
    If Not outputFile(citBachSrc, inFileNameS) Then
        MsgBox "CIT-BACH�̃\�[�X���t�@�C���o�͂��邱�ƂɎ��s�����̂ŁA�����𒆎~���܂��B"
        Exit Function
    End If

    getToolInputFile = True
End Function

' PICT���s
Function execPict(inFileName As String, outFileName As String) As Boolean
    execPict = False
    
    Dim objShell As Object
    Dim oExec
    Dim rc As Long
    Dim comStr As String

    ' PICT���s
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = Worksheets(controlSheetName).Range("��ƃp�X").Value
    
    comStr = """" & Worksheets(controlSheetName).Range("PICT�t���p�X").Value & """ """ & inFileName & """ " & pictCmdOption & " > """ & outFileName & """"
    comStr = "%ComSpec% /c """ & comStr & """"
        
    rc = objShell.Run(comStr, 1, True)
    If Not rc = 0 Then
        MsgBox "PICT�̎��s���ʂ̖߂�l��" & rc & "�Ȃ̂ŁA���s�Ɏ��s���Ă���\��������܂��B"
    End If
    execPict = True
End Function

' CIT-BACH���s
Function execCitBach(inFileName As String, outFileName As String) As Boolean
    execCitBach = False
    
    Dim objShell As Object
    Dim oExec
    Dim rc As Long
    Dim comStr As String

    ' CIT-BACH���s
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = Worksheets(controlSheetName).Range("��ƃp�X").Value
    
    comStr = "java -jar " & """" & Worksheets(controlSheetName).Range("CIT_BACH�t���p�X").Value & """ " & citBachCmdOption & " -i """ & inFileName & """ -o """ & outFileName & """"
    comStr = "%ComSpec% /c """ & comStr & """"
        
    rc = objShell.Run(comStr, 1, True)
    If Not rc = 0 Then
        MsgBox "CIT-BACH�̎��s���ʂ̖߂�l��" & rc & "�Ȃ̂ŁA���s�Ɏ��s���Ă���\��������܂��B"
    End If
    
    execCitBach = True
End Function

' Alloy���s
Function execAlloy(inFileName As String) As Boolean
    execAlloy = False
    
    Dim objShell As Object
    Dim oExec
    Dim rc As Long
    Dim comStr As String

    ' Alloy���s
    Set objShell = CreateObject("WScript.Shell")
    objShell.CurrentDirectory = Worksheets(controlSheetName).Range("��ƃp�X").Value
    
    comStr = "java -jar " & """" & Worksheets(controlSheetName).Range("Alloy�t���p�X").Value & """ """ & inFileName ' & """ -o """ & outFileName & """"
    comStr = "%ComSpec% /c """ & comStr & """"
        
    rc = objShell.Run(comStr, 0, False)
    If Not rc = 0 Then
        MsgBox "Alloy�̎��s���ʂ̖߂�l��" & rc & "�Ȃ̂ŁA���s�Ɏ��s���Ă���\��������܂��B"
    End If
    
    execAlloy = True
End Function

' Tool�ւ̓��̓t�@�C�������q�E�����̏��Ɛ���(��������\���玩����������������܂�)�L�q�̏�񂩂琶������
Function createToolInputSrc(srcBook As Workbook, ByRef pictSrc As String, ByRef citBachSrc As String) As Boolean
    createToolInputSrc = False
    Dim paramNames() As String
    Dim tuples()
    Dim FLtblSheet As Worksheet
    Dim factorNames() As String
    Dim levelLists()
    Dim factorNum As Long
    Dim levelNum As Long
    
    ' ���q�E�����̃V�[�g����肷��
    If Not FLtblSheetName = "" Then
        If ExistsWorksheet(srcBook, FLtblSheetName) Then
            Set FLtblSheet = srcBook.Sheets(FLtblSheetName)
        End If
    End If
    If FLtblSheet Is Nothing Then ' ���ǈ��q�E�����\�V�[�g��������Ȃ�
        MsgBox "���q�E�����\�̖��O�Ƃ��āu" & FLtblSheetName & "�v���w�肳��Ă��܂����A���̖��O�̃V�[�g�����݂��܂���B�����𒆎~���܂��B"
        Exit Function
    End If
    
    ' ���q�E������ǂݍ���
    If Not FLTable2array(FLtblSheet, factorNames, levelLists) Then
        MsgBox "���q�E�����̉�͂Ɏ��s���܂����B�����𒆎~���܂��B"
        Exit Function
    End If

    ' ���q�E�������
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
        
    ' ������̒ǉ�
    Dim constraints As String
    If Not GetConstraints(constraints) Then
        MsgBox "����L�q�̎擾�Ɏ��s���܂����B�����𒆎~���܂��B"
        Exit Function
    End If
    
    ' S����������̒ǉ�
    Dim s_constraints As String
    If Not GetSConstraints(s_constraints) Then
        MsgBox "CIT-BACH����L�q�̎擾�Ɏ��s���܂����B�����𒆎~���܂��B"
        Exit Function
    End If
    
    pictSrc = pictSrc & constraints
    citBachSrc = citBachSrc & s_constraints

    createToolInputSrc = True
End Function

' Tool�̏o�͌��ʂ��V�[�g�ɏ�������
Function fillInToolOutSheets(srcBook As Workbook, paramNames() As String, tuples()) As Boolean
    fillInToolOutSheets = False
    
    Dim testCaseNum As Long
    Dim paramNum As Long
    Dim toolOutSheet As Worksheet
    Dim j As Long
    Dim paramL As Long
    Dim paramU As Long
     
    ' Tool�o�͌��ʗp�̃V�[�g��p�ӂ���
    srcBook.Worksheets.Add Before:=Worksheets(FLtblSheetName)
    ActiveSheet.name = toolOutSheetName
    Set toolOutSheet = srcBook.Sheets(toolOutSheetName)
    
    ' �f�[�^�����m�F���ăV�[�g�ɏ�������
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
    
    paramL = LBound(paramNames)
    paramU = UBound(paramNames)
    Worksheets(controlSheetName).Range("���ڃ^�C�g������").Copy
    toolOutSheet.Cells(offsetRows + 1, offsetColumns + 1).Value = "ID"
    toolOutSheet.Cells(offsetRows + 1, offsetColumns + 1).PasteSpecial (xlPasteFormats)
    For j = paramL To paramU
        toolOutSheet.Cells(offsetRows + 1, offsetColumns + j - paramL + 2).Value = paramNames(j)
        toolOutSheet.Cells(offsetRows + 1, offsetColumns + j - paramL + 2).PasteSpecial (xlPasteFormats)
    Next j
    Dim i As Long
    Dim tuple() As String
    Worksheets(controlSheetName).Range("�l����").Copy
    For i = 0 To testCaseNum - 1
        tuple = tuples(i)
        toolOutSheet.Cells(offsetRows + 2 + i, offsetColumns + 1).Value = "#" & i + 1 ' ID�Ƃ��ăV�[�P���V�����ԍ�
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
    'ADODB.Stream�I�u�W�F�N�g�𐶐�
    Dim srcFileFullName As String
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8�ւ̕ϊ��̂��߂Ɏg��
    
    If srcFilename = "" Then
        srcFileFullName = getFilePath("PICT���s���ʂ̃t�@�C����I��")
        If srcFileFullName = "" Then
            MsgBox "PICT���s���ʂ̃t�@�C�����I������Ȃ������̂ŁA�����𒆎~���܂��B"
            Exit Function
        End If
    ElseIf InStr(srcFilename, "\") > 0 Then ' ���X�t���p�X�Ŏw�肳��Ă���Ɣ��f����
        srcFileFullName = srcFilename
    Else
        srcFileFullName = Worksheets(controlSheetName).Range("��ƃp�X").Value & "\" & srcFilename
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
        srcFileFullName = getFilePath("PICT�܂���CIT-BACH�̎��s���ʂ̃t�@�C����I��")
        If srcFileFullName = "" Then
            MsgBox "PICT�܂���CIT-BACH�̎��s���ʂ̃t�@�C�����I������Ȃ������̂ŁA�����𒆎~���܂��B"
            Exit Function
        End If
    ElseIf InStr(srcFilename, "\") > 0 Then ' ���X�t���p�X�Ŏw�肳��Ă���Ɣ��f����
        srcFileFullName = srcFilename
    Else
        srcFileFullName = Worksheets(controlSheetName).Range("��ƃp�X").Value & "\" & srcFilename
    End If
    
    Set ts = fso.OpenTextFile(srcFileFullName, Format:=TristateFalse) ' �t�@�C���� Shift-JIS �ŊJ��

    ' �S�Ẵf�[�^��ǂݍ���
    str = ts.ReadAll
   
    srcFilename = srcFileFullName ' ���ۂɓǂݍ��񂾃t�@�C���̃t���p�X����Ԃ�
    
    inputFile = True
End Function

Function outputUtf8(src As String, destFilename As String) As Boolean
    outputUtf8 = False
    
    Dim destFileFullName As String
    'ADODB.Stream�I�u�W�F�N�g�𐶐�
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8�ւ̕ϊ��̂��߂Ɏg��
    
    destFileFullName = Worksheets(controlSheetName).Range("��ƃp�X").Value & "\" & destFilename
    
    With adoSt
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText src
        ' BOM���폜����
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
    
    destFileFullName = Worksheets(controlSheetName).Range("��ƃp�X").Value & "\" & destFilename
    
    outputFile = False
    Open destFileFullName For Output As #1
    Print #1, src
    Close #1
    
    outputFile = True
End Function

