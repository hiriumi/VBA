VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKekka 
   Caption         =   "中高スポーツ大会午前結果"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   OleObjectBlob   =   "frmKekka.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKekka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTeamResult() As TeamData
Private mShowResult As Boolean

Sub 結果表示(ByRef TeamResult() As TeamData)
    mTeamResult = TeamResult
    
    txtKekka.text = "順位,チーム名,得点,得失点差,総得点,勝敗" & vbCrLf
    Dim i As Integer
    i = 0
    For i = 0 To UBound(TeamResult) - 1
        txtKekka.text = txtKekka.text & TeamResult(i).Rank & "," & _
            TeamResult(i).teamName & "," & TeamResult(i).Tokuten & _
            "," & TeamResult(i).Tokushittensa & "," & TeamResult(i).Soutokuten & _
            "," & TeamResult(i).Shourisuu & "勝" & TeamResult(i).Haisensuu & "敗" & TeamResult(i).Hikiwakesuu & "分け" & vbCrLf
    Next i
    
    ' 一番上に移動。
    txtKekka.SelStart = 0
End Sub

Private Sub btnCopy_Click()
    txtKekka.SelStart = 0
    txtKekka.SelLength = Len(txtKekka.text)
    txtKekka.SetFocus
    txtKekka.Copy
End Sub

Private Sub btnGenerateCSV_Click()
    ' CSVファイルを生成する。
    Dim CSVFile As String
    CSVFile = ActivePresentation.Path & "\結果（" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Minute(Now) & "-" & Second(Now) & "）.csv"
    
    Dim fileNumber As Integer
    fileNumber = FreeFile()
    
    Open CSVFile For Output As fileNumber
    
    Print #fileNumber, txtKekka.text
    Close #fileNumber
    
    MsgBox "結果を以下のファイルに書き出しました。" & vbCrLf & CSVFile, vbOKOnly + vbInformation, "中高スポーツ大会"
End Sub

Private Sub btnKekkaHanei_Click()
    If MsgBox("午後の部の表に変更が加えられます。よろしいですか？", vbQuestion + vbYesNo, "中高スポーツ大会") = vbNo Then
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
