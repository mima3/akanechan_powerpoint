Attribute VB_Name = "main"
Option Explicit



'*
'* ノートに書いた内容をスライドの切り替え時にsofttalkを用いてしゃべるようにします。
'* 「画像切り替え」タブのサウンドにそのファイルは指定されています。
'* このリストに追加したファイルはPowerPointerを再起動したときに使用されていなければリストから消えます
'*
Public Sub AddSoftTalk()
    Dim sld As Slide
    Dim note As Slide
    Dim msg As String
    Dim wavPath As String
    ' Visit each slide
    For Each sld In ActivePresentation.Slides
        Call AddSoftTalkToSlide(sld)
    Next
End Sub
'*
'* 選択中のスライドに対して音声を追加する
'*
Public Sub AddSoftTalkToSelectedSlide()
    Dim sld As Slide
    Set sld = ActivePresentation.Slides.Item(ActiveWindow.Selection.SlideRange.SlideIndex)
    Call AddSoftTalkToSlide(sld)
End Sub

Private Sub AddSoftTalkToSlide(ByRef sld As Slide)
    Dim note As Slide
    Dim msg As String
    Dim wavPath As String
    Dim line As Variant
    
    Dim vr As VoiceRoid
    Set vr = New VoiceRoid
    
    Dim i As Long
    
    For Each note In sld.NotesPage
        msg = note.Shapes.Item(2).TextEffect.text
        Debug.Print msg
        If msg <> "" Then
            i = 0
            line = Split(msg, vbCr)
            For i = LBound(line) To UBound(line)
                If line(i) <> "" Then
                    wavPath = ActivePresentation.Path & "\" & sld.name & "_" & i & ".wav"
                    ' Wavファイル作成
                    Call vr.CreateWavFile(line(i), wavPath)
            
                    Call ApeendWavFile(sld, wavPath)
                End If
            Next i
        End If
    Next

End Sub


'**
'* スライドにファイルを追加します。
'* この際全てのShapesをチェックしてすでに追加されていないか確認します。
'* @param[in,out] sld 対象のスライド
'* @param[in] wavPath 作成するwavファイルのパス
'**
Private Sub ApeendWavFile(ByRef sld As Slide, ByVal wavPath As String)
    ' 重複チェック & 削除
    Dim shp As Shape
    Dim rmIndex As Long
    rmIndex = 0
    Dim i As Long
    i = 1
    For Each shp In sld.Shapes
        If shp.Type = msoMedia Then
            If shp.MediaType = ppMediaTypeSound Then
                If Dir(wavPath) = shp.name Then
                    rmIndex = i
                    Exit For
                End If
            End If
        End If
        i = i + 1
    Next
    If rmIndex <> 0 Then
        sld.Shapes.Item(rmIndex).Delete
    End If
    
    Set shp = sld.Shapes.AddMediaObject2(wavPath)
    shp.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue
    shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoTrue

    
End Sub


