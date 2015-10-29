VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKekka 
   Caption         =   "�����X�|�[�c���ߑO����"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   OleObjectBlob   =   "frmKekka.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmKekka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTeamResult() As TeamData
Private mShowResult As Boolean

Sub ���ʕ\��(ByRef TeamResult() As TeamData)
    mTeamResult = TeamResult
    
    txtKekka.text = "����,�`�[����,���_,�����_��,�����_,���s" & vbCrLf
    Dim i As Integer
    i = 0
    For i = 0 To UBound(TeamResult) - 1
        txtKekka.text = txtKekka.text & TeamResult(i).Rank & "," & _
            TeamResult(i).teamName & "," & TeamResult(i).Tokuten & _
            "," & TeamResult(i).Tokushittensa & "," & TeamResult(i).Soutokuten & _
            "," & TeamResult(i).Shourisuu & "��" & TeamResult(i).Haisensuu & "�s" & TeamResult(i).Hikiwakesuu & "����" & vbCrLf
    Next i
    
    ' ��ԏ�Ɉړ��B
    txtKekka.SelStart = 0
End Sub

Private Sub btnCopy_Click()
    txtKekka.SelStart = 0
    txtKekka.SelLength = Len(txtKekka.text)
    txtKekka.SetFocus
    txtKekka.Copy
End Sub

Private Sub btnGenerateCSV_Click()
    ' CSV�t�@�C���𐶐�����B
    Dim CSVFile As String
    CSVFile = ActivePresentation.Path & "\���ʁi" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Minute(Now) & "-" & Second(Now) & "�j.csv"
    
    Dim fileNumber As Integer
    fileNumber = FreeFile()
    
    Open CSVFile For Output As fileNumber
    
    Print #fileNumber, txtKekka.text
    Close #fileNumber
    
    MsgBox "���ʂ��ȉ��̃t�@�C���ɏ����o���܂����B" & vbCrLf & CSVFile, vbOKOnly + vbInformation, "�����X�|�[�c���"
End Sub

Private Sub btnKekkaHanei_Click()
    If MsgBox("�ߌ�̕��̕\�ɕύX���������܂��B��낵���ł����H", vbQuestion + vbYesNo, "�����X�|�[�c���") = vbNo Then
        Exit Sub
    End If
    
    mShowResult = True
    
    Unload Me
End Sub

Property Get ShowResult() As Boolean
    ShowResult = mShowResult
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    Me.Hide
End Sub
