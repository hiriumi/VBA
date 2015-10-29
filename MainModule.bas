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

Public Sub �e�X�g�f�[�^�����()
   
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
        
        ' �X���C�h�̒���Shape�����[�v���Ă��̃^�C�v��Table�������炻�����舵���B
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    
                    RandomNumber = Int((15 - 0 + 1) * Rnd + 0)
                    
                    If InStr(Team1Text, "(") > 0 Or InStr(Team1Text, "�i") > 0 Then
                        team1Name = �`�[�������o(Team1Text)
                        s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team1Name
                    Else
                        team1Name = Trim(Team1Text)
                    End If
                    
                    ' �����œK���Ȑ���������B
                    s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team1Name & " (" & CStr(RandomNumber) & ")"
                    
                    If InStr(Team2Text, "(") > 0 Or InStr(Team2Text, "�i") > 0 Then
                        team2Name = �`�[�������o(Team2Text)
                        s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team2Name
                    Else
                        team2Name = Trim(Team2Text)
                    End If
                    
                    ' �ʂ̗����𐶐��B
                    RandomNumber = Int((15 - 0 + 1) * Rnd + 0)
                    s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team2Name & " (" & CStr(RandomNumber) & ")"
                    
                Next j
            End If
        Next s
    Next i
    
    MsgBox "�e�X�g�f�[�^�𐶐����܂����B", vbInformation + vbOKOnly, "�����X�|�[�c���"

End Sub

Public Sub �ߑO�̃f�[�^���N���A()
    If MsgBox("�ߑO�̃X���C�h�ɓ��͂���Ă���f�[�^�����ׂď������܂��B��낵���ł����H", vbYesNo + vbQuestion, "�����X�|�[�c���") = vbNo Then
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
        
        ' �X���C�h�̒���Shape�����[�v���Ă��̃^�C�v��Table�������炻�����舵���B
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    
                    If InStr(Team1Text, "(") > 0 Or InStr(Team1Text, "�i") > 0 Then
                        team1Name = �`�[�������o(Team1Text)
                        s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team1Name
                    End If
                    
                    If InStr(Team2Text, "(") > 0 Or InStr(Team2Text, "�i") > 0 Then
                        team2Name = �`�[�������o(Team2Text)
                        s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team2Name
                    End If
                    
                Next j
            End If
        Next s
    Next i
    
    MsgBox "�f�[�^�̍폜���������܂����I(�M��֥�L)�U", vbInformation + vbOKOnly, "�����X�|�[�c���"

End Sub

Public Sub �ߑO�̓��_�v�Z()
    Dim Slide1 As Slide
    Dim i As Integer
    Dim j As Integer
    Set TeamResult = New Dictionary
    
    Dim team1Name As String
    Dim team2Name As String
    Dim shiaiID As String
        
    ' �܂��͌ߑO�̕��̃X���C�h�����[�v���ĕ\���瓾�_�𒊏o����B
    For i = GozenSlideStart To GozenSlideEnd
        Set Slide1 = ActivePresentation.Slides(i)
        Dim s As Shape
        
        ' �X���C�h�̒���Shape�����[�v���Ă��̃^�C�v��Table�������炻�����舵���B
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GozenTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GozenTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    shiaiID = s.Table.Cell(j, GozenShiaiIDColumnNumber).Shape.TextFrame.TextRange.text
                    
                    ' �܂��̓`�[�����𒊏o
                    team1Name = �`�[�������o(Team1Text)
                    team2Name = �`�[�������o(Team2Text)
                    
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
                        
                        team1Tokuten = ���_���o(Team1Text)
                        team2Tokuten = ���_���o(Team2Text)
                        
                         ' ���_���v�Z���ċL�^����B
                         ���_�v�Z team1Name, team1Tokuten, team2Name, team2Tokuten
                    End If
                Next j
            End If
        Next s
    Next i
    
    ' �f�[�^����ёւ��o����悤�ɔz��ɓ����B
    Dim SortData() As TeamData
    ReDim SortData(TeamResult.Count)
    
    Dim teamName As Variant
    i = 0
    
    For Each teamName In TeamResult.Keys
        Set SortData(i) = TeamResult(teamName)
        i = i + 1
    Next
    
    ' ���_����ɍ~���ɕ��ёւ���B
    QuickSort SortData, 0, TeamResult.Count - 1, "���_"
    
    ' Debug Code
    Dim k
    For k = 0 To UBound(SortData) - 1
        Debug.Print SortData(k).teamName & "," & SortData(k).Tokuten & ", " & SortData(k).Tokushittensa
    Next k
    
    ' �����_�����ёւ�
    �����_�����בւ� SortData
    
    �����_���בւ� SortData
    
    ' �ŏI�I�ȏ��ʂ�����
    Dim �������ʂ����� As Boolean
    �������ʂ����� = ���ʌ���(SortData)
    
    If �������ʂ����� Then
        MsgBox "�������ʂ����݂��Ă��܂��B( �L߄t߁M)����", vbInformation + vbOKOnly, "�����X�|�[�c���"
    End If
    
    ' ���ʂ�\���B
    Dim fKekka As frmKekka
    Set fKekka = New frmKekka
    fKekka.���ʕ\�� SortData
    fKekka.Show vbModal
    
    If Not fKekka Is Nothing Then
        If fKekka.ShowResult Then
            Unload fKekka
            Set fKekka = Nothing
            ' �܂��̓o�b�N�A�b�v
            Dim BackupFilePath As String
            BackupFilePath = ActivePresentation.Path & "\" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Minute(Now) & "-" & Second(Now) & " " & ActivePresentation.Name
            
            ActivePresentation.SaveCopyAs BackupFilePath
            
            �ߌ�̕\�ɏ������� SortData
            
            �g�[�i�����g�\�ɏ������� SortData
            
            MsgBox "�ߌ�̕\�ւ̏������݂��I�����܂����B��������I(���L�́M)��ܰ�", vbOKOnly + vbInformation, "�����X�|�[�c���"
        End If
    End If

    
    ' ����Ȃ��I�u�W�F�N�g������������r���B
    Set TeamResult = Nothing
    Set fKekka = Nothing
End Sub

Private Sub �ߌ�̕\�ɏ�������(ByRef SortData() As TeamData)
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
        
        ' �X���C�h�̒���Shape�����[�v���Ă��̃^�C�v��Table�������炻�����舵���B
        For Each s In Slide1.Shapes
            If s.Type = msoTable Then
                For j = StartingRowNumber To s.Table.Rows.Count
                    Dim Team1Text As String, Team2Text As String
                    Team1Text = s.Table.Cell(j, GogoTeam1ColumnNumber).Shape.TextFrame.TextRange.text
                    Team2Text = s.Table.Cell(j, GogoTeam2ColumnNumber).Shape.TextFrame.TextRange.text
                    
                    If InStr(Team1Text, "��") > 0 Then
                            jyuni = CInt(Mid(Team1Text, 1, InStr(Team1Text, "��") - 1))
                            Set team = ���ʂŃ`�[���f�[�^���擾(SortData, jyuni)
                            
                            If Not team Is Nothing Then
                                s.Table.Cell(j, GogoTeam1ColumnNumber).Shape.TextFrame.TextRange.text = team.teamName
                            End If
                    End If
                    
                    If InStr(Team2Text, "��") > 0 Then
                        jyuni = CInt(Mid(Team2Text, 1, InStr(Team2Text, "��") - 1))
                        Set team = ���ʂŃ`�[���f�[�^���擾(SortData, jyuni)
                            
                        If Not team Is Nothing Then
                            s.Table.Cell(j, GogoTeam2ColumnNumber).Shape.TextFrame.TextRange.text = team.teamName
                        End If
                    End If
                 Next j
            End If
        Next s
    Next i

End Sub

Private Function ���ʂŃ`�[���f�[�^���擾(ByRef SortData() As TeamData, ByVal jyuni As Integer) As TeamData
    Dim i As Integer
    Dim result As TeamData
    Set result = Nothing
    
    For i = 0 To UBound(SortData) - 1
        If SortData(i).Rank = jyuni Then
            Set result = SortData(i)
            Exit For
        End If
    Next i
    
    Set ���ʂŃ`�[���f�[�^���擾 = result
End Function

Private Sub �g�[�i�����g�\�ɏ�������(ByRef SortData() As TeamData)
' Public Sub �g�[�i�����g�\�ɏ�������()
    Dim TounamentSlide As Slide
    Dim team As TeamData
    
    Set TounamentSlide = ActivePresentation.Slides(TournamentSlide)
    Dim s As Shape
    Dim jyuni As Integer
    ' �X���C�h�̒���Shape�����[�v���Ă��̃^�C�v��TextBox�������炻�����舵���B
    For Each s In TounamentSlide.Shapes
        If s.Type = msoAutoShape And s.HasTextFrame Then
            If InStr(s.TextEffect.text, "��") > 0 And InStr(s.TextEffect.text, "�����") = 0 Then
                Debug.Print s.TextEffect.text
                jyuni = Replace(Trim(s.TextEffect.text), "��", vbNullString)
                Set team = ���ʂŃ`�[���f�[�^���擾(SortData, jyuni)
                If Not team Is Nothing Then
                    s.TextEffect.text = �`�[�������g�[�i�����g�\�p�ɕϊ�(team.teamName)
                End If
            End If
        End If
    Next s
    
    
End Sub

Private Function �`�[�������g�[�i�����g�\�p�ɕϊ�(ByVal teamName As String) As String
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
    
    �`�[�������g�[�i�����g�\�p�ɕϊ� = newString
End Function

'Private Function �S�p�����𔼊p�����ɕϊ�(ByVal text As String) As Integer
'    Dim i As Integer
'    Dim suuji As Variant
'
'    For i = 1 To Len(text)
'        If Mid(text, i, 1) Like "[�O-�X]" Then
'            suuji = suuji & StrConv(Mid(text, i, 1), vbNarrow)
'        Else
'            suuji = suuji & Mid(text, i, 1)
'        End If
'    Next i
'
'    �S�p�����𔼊p�����ɕϊ� = CInt(suuji)
'End Function


Private Function ���ʌ���(ByRef SortData() As TeamData) As Boolean
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
    
    ���ʌ��� = SameRankExists
End Function

Private Sub �����_���בւ�(ByRef SortData() As TeamData)
    ' ���_���������_�������_�̏ꍇ�ɑ����_�œ��_�̃`�[���݂̂�Ώۂɕ��ёւ�������B
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
                QuickSort SortData, SaiNarabekaeStartIndex, SaiNarabekaeEndIndex, "�����_"
                SaiNarabekaeStartIndex = 0
                SaiNarabekaeEndIndex = 0
            End If
        End If
        i = i + 1
    Wend

End Sub

Private Sub �����_�����בւ�(ByRef SortData() As TeamData)
    ' �Ō�ɓ��_�����_�̏ꍇ�ɓ����_���œ��_�̃`�[���݂̂�Ώۂɕ��ёւ�������B
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
                QuickSort SortData, SaiNarabekaeStartIndex, SaiNarabekaeEndIndex, "�����_��"
                SaiNarabekaeStartIndex = 0
                SaiNarabekaeEndIndex = 0
            End If
        End If
        i = i + 1
    Wend
End Sub

Private Sub ���_�v�Z(ByVal team1Name As String, ByVal team1Tokuten As Integer, ByVal team2Name As String, ByVal team2Tokuten As Integer)
    ' �����ǂ�����O�_�œ����Ă����ꍇ�́A�܂��������s���Ă��Ȃ��̂ł��̃v���Z�X���ȗ��B
    If team1Tokuten = 0 And team2Tokuten = 0 Then
        Exit Sub
    End If
    
    ' �`�[���P�ƃ`�[���Q�̊����̓��_���������[����Ăяo���B
    Dim team1KizonData As TeamData
    Dim team2KizonData As TeamData
    Set team1KizonData = TeamResult(team1Name)
    Set team2KizonData = TeamResult(team2Name)

    team1KizonData.Tokushittensa = team1KizonData.Tokushittensa + (team1Tokuten - team2Tokuten)
    team2KizonData.Tokushittensa = team2KizonData.Tokushittensa + (team2Tokuten - team1Tokuten)
    
    team1KizonData.Soutokuten = team1KizonData.Soutokuten + team1Tokuten
    team2KizonData.Soutokuten = team2KizonData.Soutokuten + team2Tokuten
    
    ' ��������ǉ��B
    team1KizonData.Shiaisuu = team1KizonData.Shiaisuu + 1
    team2KizonData.Shiaisuu = team2KizonData.Shiaisuu + 1
    
    ' �܂����ғ��_�̏ꍇ�e�`�[���ɂP�_�����Z�B
    If team1Tokuten = team2Tokuten Then
        team1KizonData.Tokuten = team1KizonData.Tokuten + 1
        team2KizonData.Tokuten = team2KizonData.Tokuten + 1
        team1KizonData.Hikiwakesuu = team1KizonData.Hikiwakesuu + 1
        team2KizonData.Hikiwakesuu = team2KizonData.Hikiwakesuu + 1
    End If
    
    ' �`�[���P���������ꍇ�B
    If team1Tokuten > team2Tokuten Then
        team1KizonData.Tokuten = team1KizonData.Tokuten + 2
        team1KizonData.Shourisuu = team1KizonData.Shourisuu + 1
        team2KizonData.Haisensuu = team2KizonData.Haisensuu + 1
    End If
    
    ' �`�[���Q���������ꍇ�B
    If team2Tokuten > team1Tokuten Then
        team2KizonData.Tokuten = team2KizonData.Tokuten + 2
        team2KizonData.Shourisuu = team2KizonData.Shourisuu + 1
        team1KizonData.Haisensuu = team1KizonData.Haisensuu + 1
    End If
    
    ' �Ăу������[�ɕۑ��B
    Set TeamResult(team1Name) = team1KizonData
    Set TeamResult(team2Name) = team2KizonData
    
End Sub

Private Function �`�[�������o(ByVal text As String) As String
    Dim teamName As String
    teamName = Trim(text)
    
    If InStr(text, "(") > 0 Then
        teamName = Trim(Mid(text, 1, InStr(text, "(") - 1))
    End If
    
    If InStr(text, "�i") > 0 Then
        teamName = Trim(Mid(text, 1, InStr(text, "�i") - 1))
    End If
            
    �`�[�������o = teamName
End Function

Private Function ���_���o(ByVal text As String) As Integer
    Dim Tokuten As Integer
    Tokuten = 0
    
    Dim HajimenoKakkoIndex As Long
    Dim OwarinoKakkoIndex As Long
    
    HajimenoKakkoIndex = InStr(text, "(")
    OwarinoKakkoIndex = InStr(text, ")")
    
    If HajimenoKakkoIndex = 0 And OwarinoKakkoIndex = 0 Then
            HajimenoKakkoIndex = InStr(text, "�i")
            HajimenoKakkoIndex = InStr(text, "�j")
    End If
    
    If HajimenoKakkoIndex = 0 And OwarinoKakkoIndex = 0 Then
        Exit Function
    End If
    
    Tokuten = CInt(Mid(text, HajimenoKakkoIndex + 1, OwarinoKakkoIndex - HajimenoKakkoIndex - 1))
        
    ���_���o = Tokuten
    
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
    '�N�C�b�N�\�[�g(�z��f�[�^, �z��f�[�^�ŏ��C���f�b�N�X, �z��f�[�^�ő�C���f�b�N�X)
    '----------------------------------------------------------------------------------
        Dim lngIdxL        As Long
        Dim lngIdxR        As Long
        Dim Kijyunchi  As TeamData
        Dim vntWk          As Variant
        
        '�����t�߂̃C���f�b�N�X���e�l����l�Ƃ��܂��B
        Set Kijyunchi = SortData((Min + Max) \ 2)
        '�ŏ��C���f�b�N�X���Z�b�g�B
        lngIdxL = Min
        '�ő�C���f�b�N�X���Z�b�g�B
        lngIdxR = Max
        Do
            '�z��̃C���f�b�N�X�̏������������l�Ɍ������āA�C���f�b�N�X���e�l����l���傫���l���C�R�[���Ȓl��T���܂��B
            For lngIdxL = lngIdxL To Max Step 1
                Select Case SortBy
                    Case "���_"
                        If SortData(lngIdxL).Tokuten <= Kijyunchi.Tokuten Then  '���P�A�~���́w>=�x���w<=�x �ɂ���B
                            Exit For
                        End If
                    Case "�����_��"
                        If SortData(lngIdxL).Tokushittensa <= Kijyunchi.Tokushittensa Then  '���P�A�~���́w>=�x���w<=�x �ɂ���B
                            Exit For
                        End If
                    Case "�����_"
                        If SortData(lngIdxL).Soutokuten <= Kijyunchi.Soutokuten Then  '���P�A�~���́w>=�x���w<=�x �ɂ���B
                            Exit For
                        End If
                End Select
                

            Next
            '�z��̃C���f�b�N�X�̑傫���������l�Ɍ������āA�C���f�b�N�X���e�l����l��菬�����l���C�R�[���Ȓl��T���܂��B
            For lngIdxR = lngIdxR To Min Step -1
                Select Case SortBy
                    Case "���_"
                        If SortData(lngIdxR).Tokuten >= Kijyunchi.Tokuten Then  '���P�A�~���́w>=�x���w<=�x �ɂ���B
                            Exit For
                        End If
                    Case "�����_��"
                        If SortData(lngIdxR).Tokushittensa >= Kijyunchi.Tokushittensa Then '���P�A�~���́w>=�x���w<=�x �ɂ���B
                            Exit For
                        End If
                    Case "�����_"
                        If SortData(lngIdxR).Soutokuten >= Kijyunchi.Soutokuten Then  '���P�A�~���́w>=�x���w<=�x �ɂ���B
                            Exit For
                        End If
                End Select
            Next
            
            '�ŏ��C���f�b�N�X�ƍő�C���f�b�N�X���������傫���ɂȂ�����u���C�NDO���[�v�B
            If lngIdxL >= lngIdxR Then
               '���̒i�K�ŁA��l��菬���Ȓl����l��荶�ɁA�傫�Ȓl����l���E�ɂ��Ă��܂��B
               Exit Do
            End If
            '�o���̒l���������܂��B
            Set vntWk = SortData(lngIdxL)
            Set SortData(lngIdxL) = SortData(lngIdxR)
            Set SortData(lngIdxR) = vntWk
            '�z��̃C���f�b�N�X�̏������������l�Ɍ������ẴC���f�b�N�X���X�V����B
            lngIdxL = lngIdxL + 1
            '�z��̃C���f�b�N�X�̑傫���������l�Ɍ������ẴC���f�b�N�X���X�V����B
            lngIdxR = lngIdxR - 1
        Loop
        '���̔z��̏������K�v��
        If (Min < lngIdxL - 1) Then
              '��l�̍��̔z��ɑ΂��Ẵ\�[�g�������s���B
              QuickSort SortData(), Min, lngIdxL - 1, SortBy
        End If
        '�E�̔z��̏������K�v��
        If (Max > lngIdxR + 1) Then
              '��l�̉E�̔z��ɑ΂��Ẵ\�[�g�������s���B
              QuickSort SortData(), lngIdxR + 1, Max, SortBy
        End If
End Sub



