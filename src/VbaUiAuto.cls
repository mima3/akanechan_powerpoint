VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaUiAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'* UIAutomationClientを参照設定してください。
Private uia As UIAutomationClient.CUIAutomation
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()
    Set uia = New UIAutomationClient.CUIAutomation
End Sub
Private Sub Class_Terminate()
    Set uia = Nothing
End Sub
Public Sub SleepMilli(ByVal millisec As Long)
    Call Sleep(millisec)
End Sub
' ルートのディスクトップ要素を取得
Public Function GetRoot() As IUIAutomationElement
    Dim ret As IUIAutomationElement
    Set ret = uia.GetRootElement
    Set GetRoot = ret
End Function

' 指定の子ウィンドウをタイトルから取得する
Public Function GetMainWindowByTitle(ByRef form As IUIAutomationElement, ByVal name As String) As IUIAutomationElement
    Dim cnd As IUIAutomationCondition
    Dim ret As IUIAutomationElement
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, name)

    Set ret = form.FindFirst(TreeScope_Element Or TreeScope_Children, cnd)

    Set GetMainWindowByTitle = ret
End Function

' 指定の子ウィンドウをタイトルから取得できまで待機
Public Function WaitMainWindowByTitle(ByRef form As IUIAutomationElement, ByVal name As String, ByVal timeOutSec As Double) As IUIAutomationElement
    Dim start As Variant
    start = Timer()
    Dim ret As IUIAutomationElement

    Set ret = GetMainWindowByTitle(form, name)
    Do While ret Is Nothing
        If Timer() - start > timeOutSec Then
            Exit Function
        End If
        Set ret = GetMainWindowByTitle(form, name)
        Call SleepMilli(100)
    Loop
    Set WaitMainWindowByTitle = ret
End Function

' ボタンを指定Indexを押下する
Public Sub pushButton(ByRef form As IUIAutomationElement, ByVal ix As Long)
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ClassNamePropertyId, "Button")

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)

    Dim ptn As IUIAutomationInvokePattern
    Set ptn = list.GetElement(ix).GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId)
    Call ptn.Invoke

End Sub


' 指定のClassNameがTextBoxに値を設定する
Public Sub SetText(ByRef form As IUIAutomationElement, ByVal ix As Long, ByVal text As String)
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ClassNamePropertyId, "TextBox")

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)


    Dim editValue As IUIAutomationValuePattern
    Set editValue = list.GetElement(ix).GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId)
    Call editValue.SetValue(text)

End Sub

' 指定のAutomationIdでTextBoxに値を設定する
Public Sub SetTextById(ByRef form As IUIAutomationElement, ByVal id As String, ByVal text As String)
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_AutomationIdPropertyId, id)

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)


    Dim editValue As IUIAutomationValuePattern
    Set editValue = list.GetElement(0).GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId)
    Call editValue.SetValue(text)

End Sub

' 指定のClassNameがTextBoxの値を取得する
Public Function GetStatusBarItemText(ByRef form As IUIAutomationElement, ByVal ix As Long) As String
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ClassNamePropertyId, "StatusBarItem")

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)

    GetStatusBarItemText = list.GetElement(ix).CurrentName

End Function
