VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VoiceRoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private vua As VbaUiAuto
Private mainForm As IUIAutomationElement
'*
'* ������
'*
Private Sub Class_Initialize()
    Set vua = New VbaUiAuto
    Set mainForm = vua.GetMainWindowByTitle(vua.GetRoot(), "VOICEROID2")
    If (mainForm Is Nothing) Then
        Set mainForm = vua.GetMainWindowByTitle(vua.GetRoot(), "VOICEROID2*")
        If (mainForm Is Nothing) Then
            Err.Raise 999, "VoiceRoid.Init", "VOICEROIDE2���N�����Ă��Ȃ�"
            Exit Sub
        End If
    End If
End Sub

'**
'* VOICEROID2�ɂ��Wav�t�@�C���̍쐬
'* @param[in] msg ����ׂ���e
'* @param[in] wavPath �쐬����wav�t�@�C���̃p�X
'**
Public Sub CreateWavFile(ByVal msg As String, ByVal wavPath As String)
    ' �������̃Z���t�ݒ�
    Call vua.SetText(mainForm, 0, msg)
    
    ' �����ۑ�
    Call vua.pushButton(mainForm, 4)

    ' 5�b�ȓ��ɉ����ۑ���ʂ��\�����ꂽ��ۑ��{�^��������
    Dim saveWvForm As IUIAutomationElement
    Set saveWvForm = vua.WaitMainWindowByTitle(mainForm, "�����ۑ�", 5)
    Call vua.pushButton(saveWvForm, 0)

    ' ���O��t���ĕۑ��ɓ��t�̃t�@�C���������
    Dim saveFileForm As IUIAutomationElement
    Set saveFileForm = vua.WaitMainWindowByTitle(saveWvForm, "���O��t���ĕۑ�", 5)
    Call vua.SetTextById(saveFileForm, "1001", wavPath)
    SendKeys "{ENTER}"

    ' ���|�b�v�A�b�v��OK������
    Dim infoForm As IUIAutomationElement
    Set infoForm = vua.WaitMainWindowByTitle(saveWvForm, "���", 60)
    Call vua.pushButton(infoForm, 0)

End Sub
