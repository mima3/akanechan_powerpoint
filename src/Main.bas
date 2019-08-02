Attribute VB_Name = "main"
Option Explicit



'*
'* �m�[�g�ɏ��������e���X���C�h�̐؂�ւ�����softtalk��p���Ă���ׂ�悤�ɂ��܂��B
'* �u�摜�؂�ւ��v�^�u�̃T�E���h�ɂ��̃t�@�C���͎w�肳��Ă��܂��B
'* ���̃��X�g�ɒǉ������t�@�C����PowerPointer���ċN�������Ƃ��Ɏg�p����Ă��Ȃ���΃��X�g��������܂�
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
'* �I�𒆂̃X���C�h�ɑ΂��ĉ�����ǉ�����
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
                    ' Wav�t�@�C���쐬
                    Call vr.CreateWavFile(line(i), wavPath)
            
                    Call ApeendWavFile(sld, wavPath)
                End If
            Next i
        End If
    Next

End Sub


'**
'* �X���C�h�Ƀt�@�C����ǉ����܂��B
'* ���̍ۑS�Ă�Shapes���`�F�b�N���Ă��łɒǉ�����Ă��Ȃ����m�F���܂��B
'* @param[in,out] sld �Ώۂ̃X���C�h
'* @param[in] wavPath �쐬����wav�t�@�C���̃p�X
'**
Private Sub ApeendWavFile(ByRef sld As Slide, ByVal wavPath As String)
    ' �d���`�F�b�N & �폜
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


