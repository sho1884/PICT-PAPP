VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectFactors 
   Caption         =   "多項間制約(禁則)関係因子選択"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   OleObjectBlob   =   "SelectFactors.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SelectFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

'ボタン押した時
Private Sub CreateButton_Click()
    Call GetFactors
    Me.Hide 'ユーザーフォームを非表示にする
End Sub

Private Sub CancelButton_Click()
    constraintFactors = ""
    Me.Hide 'ユーザーフォームを非表示にする
End Sub

Private Sub UserForm_Initialize()
'    MsgBox "initialized!"
    Dim i As Long
    'ListBoxの初期化
    With Me.ListBox1
        For i = LBound(publicFactorNames) To UBound(publicFactorNames)
            .AddItem publicFactorNames(i)
        Next i
        .ListStyle = fmListStyleOption      'チェックボックスにする
        .MultiSelect = fmMultiSelectMulti   '複数選択可にする
    End With
    With Me.ListBox2
        For i = LBound(publicFactorNames) To UBound(publicFactorNames)
            .AddItem publicFactorNames(i)
        Next i
        .ListStyle = fmListStyleOption      'チェックボックスにする
'        .MultiSelect = fmMultiSelectMulti   '複数選択可にする
    End With
End Sub

'チェックされたリストを配列に入れる
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
                MsgBox "複数選択されている。GUIの制限でできないはずだが。"
            End If
            j = j + 1
        End If
    Next
End Sub

'標準モジュールから呼び出す関数
'ListBoxでチェックされた項目名群を返す
Public Function doModal() As String()
    Me.Show
    doModal = conditionFactors
End Function

