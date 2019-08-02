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
'* 初期化
'*
Private Sub Class_Initialize()
    Set vua = New VbaUiAuto
    Set mainForm = vua.GetMainWindowByTitle(vua.GetRoot(), "VOICEROID2")
    If (mainForm Is Nothing) Then
        Set mainForm = vua.GetMainWindowByTitle(vua.GetRoot(), "VOICEROID2*")
        If (mainForm Is Nothing) Then
            Err.Raise 999, "VoiceRoid.Init", "VOICEROIDE2が起動していない"
            Exit Sub
        End If
    End If
End Sub

'**
'* VOICEROID2によるWavファイルの作成
'* @param[in] msg しゃべる内容
'* @param[in] wavPath 作成するwavファイルのパス
'**
Public Sub CreateWavFile(ByVal msg As String, ByVal wavPath As String)
    ' 茜ちゃんのセリフ設定
    Call vua.SetText(mainForm, 0, msg)
    
    ' 音声保存
    Call vua.pushButton(mainForm, 4)

    ' 5秒以内に音声保存画面が表示されたら保存ボタンを押す
    Dim saveWvForm As IUIAutomationElement
    Set saveWvForm = vua.WaitMainWindowByTitle(mainForm, "音声保存", 5)
    Call vua.pushButton(saveWvForm, 0)

    ' 名前を付けて保存に日付のファイル名を作る
    Dim saveFileForm As IUIAutomationElement
    Set saveFileForm = vua.WaitMainWindowByTitle(saveWvForm, "名前を付けて保存", 5)
    Call vua.SetTextById(saveFileForm, "1001", wavPath)
    SendKeys "{ENTER}"

    ' 情報ポップアップのOKを押下
    Dim infoForm As IUIAutomationElement
    Set infoForm = vua.WaitMainWindowByTitle(saveWvForm, "情報", 60)
    Call vua.pushButton(infoForm, 0)

End Sub

