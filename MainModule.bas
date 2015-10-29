Attribute VB_Name = "MainModule"
Option Explicit

Private Const GozenSlideStart As Integer = 1
Private Const GozenSlideEnd As Integer = 5

Private Const GogoSlideStart As Integer = 6
Private Const GogoSlideEnd As Integer = 9

Private Const TournamentSlide As Integer = 10

Private Const StartingRowNumber As Integer = 2

Private Const GozenTeam1ColumnNumber As Integer = 3
Private Const GozenTeam2ColumnNumber As Integer = 5

Private Const GogoTeam1ColumnNumber As Integer = 4
Private Const GogoTeam2ColumnNumber As Integer = 6


Private Const GozenShiaiIDColumnNumber As Integer = 2
Private Const GogoShiaiIDColumnNumber As Integer = 3

Private TeamResult As Dictionary

Public Sub テストデータを入力()
   
    Dim Slide1 As Slide
    Dim i As Integer
    Dim j As Integer
    Set TeamResult = New Dictionary
    
    Dim team1Name As String
    Dim team2Name As String

    For i = GozenSlideStart To GozenSlideEnd
        Set Slide1 = ActivePresentation.Slides(i)
        Dim s As Shape
        Dim RandomNumber As Integer
        
        ' スライドの中のShapeをループしてそのタイプがTableだったらそれを取り扱う。
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    
                    RandomNumber = Int((15 - 0 + 1) * Rnd + 0)
                    
                    If InStr(Team1Text, "(") > 0 Or InStr(Team1Text, "（") > 0 Then
                        team1Name = チーム名抽出(Team1Text)
                        s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team1Name
                    Else
                        team1Name = Trim(Team1Text)
                    End If
                    
                    ' ここで適当な数字を入れる。
                    s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team1Name & " (" & CStr(RandomNumber) & ")"
                    
                    If InStr(Team2Text, "(") > 0 Or InStr(Team2Text, "（") > 0 Then
                        team2Name = チーム名抽出(Team2Text)
                        s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team2Name
                    Else
                        team2Name = Trim(Team2Text)
                    End If
                    
                    ' 別の乱数を生成。
                    RandomNumber = Int((15 - 0 + 1) * Rnd + 0)
                    s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team2Name & " (" & CStr(RandomNumber) & ")"
                    
                Next j
            End If
        Next s
    Next i
    
    MsgBox "テストデータを生成しました。", vbInformation + vbOKOnly, "中高スポーツ大会"

End Sub

Public Sub 午前のデータをクリア()
    If MsgBox("午前のスライドに入力されているデータをすべて消去します。よろしいですか？", vbYesNo + vbQuestion, "中高スポーツ大会") = vbNo Then
        Exit Sub
    End If
    
    Dim Slide1 As Slide
    Dim i As Integer
    Dim j As Integer
    Set TeamResult = New Dictionary
    
    Dim team1Name As String
    Dim team2Name As String

    For i = GozenSlideStart To GozenSlideEnd
        Set Slide1 = ActivePresentation.Slides(i)
        Dim s As Shape
        
        ' スライドの中のShapeをループしてそのタイプがTableだったらそれを取り扱う。
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    
                    If InStr(Team1Text, "(") > 0 Or InStr(Team1Text, "（") > 0 Then
                        team1Name = チーム名抽出(Team1Text)
                        s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team1Name
                    End If
                    
                    If InStr(Team2Text, "(") > 0 Or InStr(Team2Text, "（") > 0 Then
                        team2Name = チーム名抽出(Team2Text)
                        s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team2Name
                    End If
                    
                Next j
            End If
        Next s
    Next i
    
    MsgBox "データの削除が完了しました！(｀･ω･´)ゞ", vbInformation + vbOKOnly, "中高スポーツ大会"

End Sub

Public Sub 午前の得点計算()
    Dim Slide1 As Slide
    Dim i As Integer
    Dim j As Integer
    Set TeamResult = New Dictionary
    
    Dim team1Name As String
    Dim team2Name As String
    Dim shiaiID As String
        
    ' まずは午前の部のスライドをループして表から得点を抽出する。
    For i = GozenSlideStart To GozenSlideEnd
        Set Slide1 = ActivePresentation.Slides(i)
        Dim s As Shape
        
        ' スライドの中のShapeをループしてそのタイプがTableだったらそれを取り扱う。
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    shiaiID = s.Table.Cell(j, GozenShiaiIDColumnNumber).Shape.TextFrame.TextRange.text
                    
                    ' まずはチーム名を抽出
                    team1Name = チーム名抽出(Team1Text)
                    team2Name = チーム名抽出(Team2Text)
                    
                    Dim Team1Data As TeamData
                    Dim Team2Data As TeamData
                    If Not ExcludedShiai(shiaiID) Then
                        If Not TeamResult.Exists(team1Name) Then
                            Set Team1Data = New TeamData
                            Team1Data.teamName = team1Name
                            Team1Data.Tokuten = 0
                            Team1Data.Tokushittensa = 0
                            
                            TeamResult.Add team1Name, Team1Data
                        End If
                        
                        If Not TeamResult.Exists(team2Name) Then
                            Set Team2Data = New TeamData
                            Team2Data.teamName = team2Name
                            Team2Data.Tokuten = 0
                            Team2Data.Tokushittensa = 0
                            
                            TeamResult.Add team2Name, Team2Data
                        End If
                        
                        Dim team1Tokuten As Integer
                        Dim team2Tokuten As Integer
                        
                        team1Tokuten = 得点抽出(Team1Text)
                        team2Tokuten = 得点抽出(Team2Text)
                        
                         ' 得点を計算して記録する。
                         得点計算 team1Name, team1Tokuten, team2Name, team2Tokuten
                    End If
                Next j
            End If
        Next s
    Next i
    
    ' データを並び替え出来るように配列に入れる。
    Dim SortData() As TeamData
    ReDim SortData(TeamResult.Count)
    
    Dim teamName As Variant
    i = 0
    
    For Each teamName In TeamResult.Keys
        Set SortData(i) = TeamResult(teamName)
        i = i + 1
    Next
    
    ' 得点を基準に降順に並び替える。
    QuickSort SortData, 0, TeamResult.Count - 1, "得点"
    
    ' Debug Code
    Dim k
    For k = 0 To UBound(SortData) - 1
        Debug.Print SortData(k).teamName & "," & SortData(k).Tokuten & ", " & SortData(k).Tokushittensa
    Next k
    
    ' 得失点差並び替え
    得失点差並べ替え SortData
    
    総得点並べ替え SortData
    
    ' 最終的な順位を決定
    Dim 同じ順位が存在 As Boolean
    同じ順位が存在 = 順位決定(SortData)
    
    If 同じ順位が存在 Then
        MsgBox "同じ順位が存在しています。( ´ﾟдﾟ｀)ｱﾁｬｰ", vbInformation + vbOKOnly, "中高スポーツ大会"
    End If
    
    ' 結果を表示。
    Dim fKekka As frmKekka
    Set fKekka = New frmKekka
    fKekka.結果表示 SortData
    fKekka.Show vbModal
    
    If Not fKekka Is Nothing Then
        If fKekka.ShowResult Then
            Unload fKekka
            Set fKekka = Nothing
            ' まずはバックアップ
            Dim BackupFilePath As String
            BackupFilePath = ActivePresentation.Path & "\" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Minute(Now) & "-" & Second(Now) & " " & ActivePresentation.Name
            
            ActivePresentation.SaveCopyAs BackupFilePath
            
            午後の表に書きこむ SortData
            
            トーナメント表に書きこむ SortData
            
            MsgBox "午後の表への書き込みを終了しました。やったぁ！(∩´∀｀)∩ﾜｰｲ", vbOKOnly + vbInformation, "中高スポーツ大会"
        End If
    End If

    
    ' いらないオブジェクトをメモリから排除。
    Set TeamResult = Nothing
    Set fKekka = Nothing
End Sub

Private Sub 午後の表に書きこむ(ByRef SortData() As TeamData)
    Dim Slide1 As Slide
    Dim i As Integer
    Dim j As Integer
    
    Dim team1Name As String
    Dim team2Name As String
    
    Dim jyuni As Integer
    
    Dim team As TeamData
    
    For i = GogoSlideStart To GogoSlideEnd
        Set Slide1 = ActivePresentation.Slides(i)
        Dim s As Shape
        
        ' スライドの中のShapeをループしてそのタイプがTableだったらそれを取り扱う。
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GogoTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GogoTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    
                    If InStr(Team1Text, "位") > 0 Then
                            jyuni = CInt(Mid(Team1Text, 1, InStr(Team1Text, "位") - 1))
                            Set team = 順位でチームデータを取得(SortData, jyuni)
                            
                            If Not team Is Nothing Then
                                s.Table.Cell(j, GogoTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team.teamName
                            End If
                    End If
                    
                    If InStr(Team2Text, "位") > 0 Then
                        jyuni = CInt(Mid(Team2Text, 1, InStr(Team2Text, "位") - 1))
                        Set team = 順位でチームデータを取得(SortData, jyuni)
                            
                        If Not team Is Nothing Then
                            s.Table.Cell(j, GogoTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team.teamName
                        End If
                    End If
                 Next j
            End If
        Next s
    Next i

End Sub

Private Function 順位でチームデータを取得(ByRef SortData() As TeamData, ByVal jyuni As Integer) As TeamData
    Dim i As Integer
    Dim result As TeamData
    Set result = Nothing
    
    For i = 0 To UBound(SortData) - 1
        If SortData(i).Rank = jyuni Then
            Set result = SortData(i)
            Exit For
        End If
    Next i
    
    Set 順位でチームデータを取得 = result
End Function

Private Sub トーナメント表に書きこむ(ByRef SortData() As TeamData)
' Public Sub トーナメント表に書きこむ()
    Dim TounamentSlide As Slide
    Dim team As TeamData
    
    Set TounamentSlide = ActivePresentation.Slides(TournamentSlide)
    Dim s As Shape
    Dim jyuni As Integer
    ' スライドの中のShapeをループしてそのタイプがTextBoxだったらそれを取り扱う。
    For Each s In TounamentSlide.Shapes
        If s.Type = msoAutoShape And s.HasTextFrame Then
            If InStr(s.TextEffect.text, "位") > 0 And InStr(s.TextEffect.text, "決定戦") = 0 Then
                Debug.Print s.TextEffect.text
                jyuni = Replace(Trim(s.TextEffect.text), "位", vbNullString)
                Set team = 順位でチームデータを取得(SortData, jyuni)
                If Not team Is Nothing Then
                    s.TextEffect.text = チーム名をトーナメント表用に変換(team.teamName)
                End If
            End If
        End If
    Next s
    
    
End Sub

Private Function チーム名をトーナメント表用に変換(ByVal teamName As String) As String
    Dim newString As String
    Dim re As New RegExp
    re.Pattern = "[A-Za-z]"
    
    If re.Test(teamName) Then
        Dim result As MatchCollection
        Set result = re.Execute(teamName)
        newString = re.Replace(teamName, vbCrLf & Mid(teamName, result(0).FirstIndex + 1))
    Else
        newString = teamName
    End If
    
    チーム名をトーナメント表用に変換 = newString
End Function

'Private Function 全角数字を半角数字に変換(ByVal text As String) As Integer
'    Dim i As Integer
'    Dim suuji As Variant
'
'    For i = 1 To Len(text)
'        If Mid(text, i, 1) Like "[０-９]" Then
'            suuji = suuji & StrConv(Mid(text, i, 1), vbNarrow)
'        Else
'            suuji = suuji & Mid(text, i, 1)
'        End If
'    Next i
'
'    全角数字を半角数字に変換 = CInt(suuji)
'End Function


Private Function 順位決定(ByRef SortData() As TeamData) As Boolean
    Dim i As Integer
    Dim jyuni As Integer
    Dim SameRankExists As Boolean
    
    SameRankExists = False
    i = 0
    jyuni = 1
    
    Dim OnajiJyuniSonzai As Boolean
    OnajiJyuniSonzai = False
        
    Do While i < UBound(SortData)
        If SortData(i + 1) Is Nothing Then
            If SortData(i).Rank = 0 Then
                SortData(i).Rank = UBound(SortData)
            End If
            Exit Do
        End If
        
        If SortData(i).Tokuten = SortData(i + 1).Tokuten And SortData(i).Tokushittensa = SortData(i + 1).Tokushittensa And SortData(i).Soutokuten = SortData(i + 1).Soutokuten Then
            SameRankExists = True
            SortData(i).Rank = jyuni
            SortData(i + 1).Rank = jyuni
            
            OnajiJyuniSonzai = True
        Else
            If OnajiJyuniSonzai Then
                jyuni = i + 2
            Else
                SortData(i).Rank = jyuni
                jyuni = jyuni + 1
            End If
            
            OnajiJyuniSonzai = False
        End If
        i = i + 1
    Loop
    
    順位決定 = SameRankExists
End Function

Private Sub 総得点並べ替え(ByRef SortData() As TeamData)
    ' 得点がも得失点差も同点の場合に総得点で同点のチームのみを対象に並び替えをする。
    Dim SaiNarabekaeStartIndex As Integer
    Dim SaiNarabekaeEndIndex As Integer
    Dim i As Integer
    
    i = 0
    While i < UBound(SortData) - 1
        If SortData(i).Tokuten = SortData(i + 1).Tokuten And SortData(i).Tokushittensa = SortData(i + 1).Tokushittensa Then
            SaiNarabekaeStartIndex = i
            Do While SortData(i).Tokuten = SortData(i + 1).Tokuten And SortData(i).Tokushittensa = SortData(i + 1).Tokushittensa
                SaiNarabekaeEndIndex = i + 1
                i = i + 1
                If i = UBound(SortData) - 1 Then
                    Exit Do
                End If
            Loop
            
            If SaiNarabekaeStartIndex <> SaiNarabekaeEndIndex Then
                QuickSort SortData, SaiNarabekaeStartIndex, SaiNarabekaeEndIndex, "総得点"
                SaiNarabekaeStartIndex = 0
                SaiNarabekaeEndIndex = 0
            End If
        End If
        i = i + 1
    Wend

End Sub

Private Sub 得失点差並べ替え(ByRef SortData() As TeamData)
    ' 最後に得点が同点の場合に得失点差で同点のチームのみを対象に並び替えをする。
    Dim SaiNarabekaeStartIndex As Integer
    Dim SaiNarabekaeEndIndex As Integer
    Dim i As Integer
    
    i = 0
    While i < UBound(SortData) - 1
        If SortData(i).Tokuten = SortData(i + 1).Tokuten Then
            SaiNarabekaeStartIndex = i
            Do While SortData(i).Tokuten = SortData(i + 1).Tokuten
                SaiNarabekaeEndIndex = i + 1
                i = i + 1
                If i = UBound(SortData) - 1 Then
                    Exit Do
                End If
            Loop
            
            If SaiNarabekaeStartIndex <> SaiNarabekaeEndIndex Then
                QuickSort SortData, SaiNarabekaeStartIndex, SaiNarabekaeEndIndex, "得失点差"
                SaiNarabekaeStartIndex = 0
                SaiNarabekaeEndIndex = 0
            End If
        End If
        i = i + 1
    Wend
End Sub

Private Sub 得点計算(ByVal team1Name As String, ByVal team1Tokuten As Integer, ByVal team2Name As String, ByVal team2Tokuten As Integer)
    ' もしどちらも０点で入ってきた場合は、まだ試合が行われていないのでこのプロセスを省略。
    If team1Tokuten = 0 And team2Tokuten = 0 Then
        Exit Sub
    End If
    
    ' チーム１とチーム２の既存の得点をメモリーから呼び出す。
    Dim team1KizonData As TeamData
    Dim team2KizonData As TeamData
    Set team1KizonData = TeamResult(team1Name)
    Set team2KizonData = TeamResult(team2Name)

    team1KizonData.Tokushittensa = team1KizonData.Tokushittensa + (team1Tokuten - team2Tokuten)
    team2KizonData.Tokushittensa = team2KizonData.Tokushittensa + (team2Tokuten - team1Tokuten)
    
    team1KizonData.Soutokuten = team1KizonData.Soutokuten + team1Tokuten
    team2KizonData.Soutokuten = team2KizonData.Soutokuten + team2Tokuten
    
    ' 試合数を追加。
    team1KizonData.Shiaisuu = team1KizonData.Shiaisuu + 1
    team2KizonData.Shiaisuu = team2KizonData.Shiaisuu + 1
    
    ' まず両者同点の場合各チームに１点ずつ加算。
    If team1Tokuten = team2Tokuten Then
        team1KizonData.Tokuten = team1KizonData.Tokuten + 1
        team2KizonData.Tokuten = team2KizonData.Tokuten + 1
        team1KizonData.Hikiwakesuu = team1KizonData.Hikiwakesuu + 1
        team2KizonData.Hikiwakesuu = team2KizonData.Hikiwakesuu + 1
    End If
    
    ' チーム１が勝った場合。
    If team1Tokuten > team2Tokuten Then
        team1KizonData.Tokuten = team1KizonData.Tokuten + 2
        team1KizonData.Shourisuu = team1KizonData.Shourisuu + 1
        team2KizonData.Haisensuu = team2KizonData.Haisensuu + 1
    End If
    
    ' チーム２が勝った場合。
    If team2Tokuten > team1Tokuten Then
        team2KizonData.Tokuten = team2KizonData.Tokuten + 2
        team2KizonData.Shourisuu = team2KizonData.Shourisuu + 1
        team1KizonData.Haisensuu = team1KizonData.Haisensuu + 1
    End If
    
    ' 再びメモリーに保存。
    Set TeamResult(team1Name) = team1KizonData
    Set TeamResult(team2Name) = team2KizonData
    
End Sub

Private Function チーム名抽出(ByVal text As String) As String
    Dim teamName As String
    teamName = Trim(text)
    
    If InStr(text, "(") > 0 Then
        teamName = Trim(Mid(text, 1, InStr(text, "(") - 1))
    End If
    
    If InStr(text, "（") > 0 Then
        teamName = Trim(Mid(text, 1, InStr(text, "（") - 1))
    End If
            
    チーム名抽出 = teamName
End Function

Private Function 得点抽出(ByVal text As String) As Integer
    Dim Tokuten As Integer
    Tokuten = 0
    
    Dim HajimenoKakkoIndex As Long
    Dim OwarinoKakkoIndex As Long
    
    HajimenoKakkoIndex = InStr(text, "(")
    OwarinoKakkoIndex = InStr(text, ")")
    
    If HajimenoKakkoIndex = 0 And OwarinoKakkoIndex = 0 Then
            HajimenoKakkoIndex = InStr(text, "（")
            HajimenoKakkoIndex = InStr(text, "）")
    End If
    
    If HajimenoKakkoIndex = 0 And OwarinoKakkoIndex = 0 Then
        Exit Function
    End If
    
    Tokuten = CInt(Mid(text, HajimenoKakkoIndex + 1, OwarinoKakkoIndex - HajimenoKakkoIndex - 1))
        
    得点抽出 = Tokuten
    
End Function

Private Function ExcludedShiai(ByVal shiaiID As String) As Boolean
    Dim ShiaiToExclude As Variant
    ShiaiToExclude = Array("91", "92", "93", "94")
    
    Dim t As Variant
    For Each t In ShiaiToExclude
        If shiaiID = t Then
            ExcludedShiai = True
            Exit Function
        End If
    Next t
    
    ExcludedShiai = False
End Function



Sub QuickSort(ByRef SortData() As TeamData, ByVal Min As Long, ByVal Max As Long, Optional ByVal SortBy As String)
    '----------------------------------------------------------------------------------
    'クイックソート(配列データ, 配列データ最小インデックス, 配列データ最大インデックス)
    '----------------------------------------------------------------------------------
        Dim lngIdxL        As Long
        Dim lngIdxR        As Long
        Dim Kijyunchi  As TeamData
        Dim vntWk          As Variant
        
        '中央付近のインデックス内容値を基準値とします。
        Set Kijyunchi = SortData((Min + Max) \ 2)
        '最小インデックスをセット。
        lngIdxL = Min
        '最大インデックスをセット。
        lngIdxR = Max
        Do
            '配列のインデックスの小さい方から基準値に向かって、インデックス内容値が基準値より大きい値かイコールな値を探します。
            For lngIdxL = lngIdxL To Max Step 1
                Select Case SortBy
                    Case "得点"
                        If SortData(lngIdxL).Tokuten <= Kijyunchi.Tokuten Then  '＊１、降順は『>=』を『<=』 にする。
                            Exit For
                        End If
                    Case "得失点差"
                        If SortData(lngIdxL).Tokushittensa <= Kijyunchi.Tokushittensa Then  '＊１、降順は『>=』を『<=』 にする。
                            Exit For
                        End If
                    Case "総得点"
                        If SortData(lngIdxL).Soutokuten <= Kijyunchi.Soutokuten Then  '＊１、降順は『>=』を『<=』 にする。
                            Exit For
                        End If
                End Select
                

            Next
            '配列のインデックスの大きい方から基準値に向かって、インデックス内容値が基準値より小さい値かイコールな値を探します。
            For lngIdxR = lngIdxR To Min Step -1
                Select Case SortBy
                    Case "得点"
                        If SortData(lngIdxR).Tokuten >= Kijyunchi.Tokuten Then  '＊１、降順は『>=』を『<=』 にする。
                            Exit For
                        End If
                    Case "得失点差"
                        If SortData(lngIdxR).Tokushittensa >= Kijyunchi.Tokushittensa Then '＊１、降順は『>=』を『<=』 にする。
                            Exit For
                        End If
                    Case "総得点"
                        If SortData(lngIdxR).Soutokuten >= Kijyunchi.Soutokuten Then  '＊１、降順は『>=』を『<=』 にする。
                            Exit For
                        End If
                End Select
            Next
            
            '最小インデックスと最大インデックスが同じか大きくになったらブレイクDOループ。
            If lngIdxL >= lngIdxR Then
               'この段階で、基準値より小さな値が基準値より左に、大きな値が基準値より右にきています。
               Exit Do
            End If
            '双方の値を交換します。
            Set vntWk = SortData(lngIdxL)
            Set SortData(lngIdxL) = SortData(lngIdxR)
            Set SortData(lngIdxR) = vntWk
            '配列のインデックスの小さい方から基準値に向かってのインデックスを更新する。
            lngIdxL = lngIdxL + 1
            '配列のインデックスの大きい方から基準値に向かってのインデックスを更新する。
            lngIdxR = lngIdxR - 1
        Loop
        '左の配列の処理が必要か
        If (Min < lngIdxL - 1) Then
              '基準値の左の配列に対してのソート処理を行う。
              QuickSort SortData(), Min, lngIdxL - 1, SortBy
        End If
        '右の配列の処理が必要か
        If (Max > lngIdxR + 1) Then
              '基準値の右の配列に対してのソート処理を行う。
              QuickSort SortData(), lngIdxR + 1, Max, SortBy
        End If
End Sub



