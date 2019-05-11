VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectFactors 
   Caption         =   "�����Ԑ���(�֑�)�֌W���q�I��"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   OleObjectBlob   =   "SelectFactors.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SelectFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
'    (http://www.gnu.org/copyleft/gpl.html)

'�{�^����������
Private Sub CreateButton_Click()
    Call GetFactors
    Me.Hide '���[�U�[�t�H�[�����\���ɂ���
End Sub

Private Sub CancelButton_Click()
    constraintFactors = ""
    Me.Hide '���[�U�[�t�H�[�����\���ɂ���
End Sub

Private Sub UserForm_Initialize()
'    MsgBox "initialized!"
    Dim i As Long
    'ListBox�̏�����
    With Me.ListBox1
        For i = LBound(publicFactorNames) To UBound(publicFactorNames)
            .AddItem publicFactorNames(i)
        Next i
        .ListStyle = fmListStyleOption      '�`�F�b�N�{�b�N�X�ɂ���
        .MultiSelect = fmMultiSelectMulti   '�����I���ɂ���
    End With
    With Me.ListBox2
        For i = LBound(publicFactorNames) To UBound(publicFactorNames)
            .AddItem publicFactorNames(i)
        Next i
        .ListStyle = fmListStyleOption      '�`�F�b�N�{�b�N�X�ɂ���
'        .MultiSelect = fmMultiSelectMulti   '�����I���ɂ���
    End With
End Sub

'�`�F�b�N���ꂽ���X�g��z��ɓ����
Private Sub GetFactors()
    Dim i As Long, j As Long
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) Then
            ReDim Preserve conditionFactors(j)
            conditionFactors(j) = Me.ListBox1.List(i)
            j = j + 1
        End If
    Next
    constraintFactors = ""
    For i = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(i) Then
            If constraintFactors = "" Then
                constraintFactors = Me.ListBox2.List(i)
            Else
                MsgBox "�����I������Ă���BGUI�̐����łł��Ȃ��͂������B"
            End If
            j = j + 1
        End If
    Next
End Sub

'�W�����W���[������Ăяo���֐�
'ListBox�Ń`�F�b�N���ꂽ���ږ��Q��Ԃ�
Public Function doModal() As String()
    Me.Show
    doModal = conditionFactors
End Function

