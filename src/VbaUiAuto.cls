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
'* UIAutomationClient���Q�Ɛݒ肵�Ă��������B
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
' ���[�g�̃f�B�X�N�g�b�v�v�f���擾
Public Function GetRoot() As IUIAutomationElement
    Dim ret As IUIAutomationElement
    Set ret = uia.GetRootElement
    Set GetRoot = ret
End Function

' �w��̎q�E�B���h�E���^�C�g������擾����
Public Function GetMainWindowByTitle(ByRef form As IUIAutomationElement, ByVal name As String) As IUIAutomationElement
    Dim cnd As IUIAutomationCondition
    Dim ret As IUIAutomationElement
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, name)

    Set ret = form.FindFirst(TreeScope_Element Or TreeScope_Children, cnd)

    Set GetMainWindowByTitle = ret
End Function

' �w��̎q�E�B���h�E���^�C�g������擾�ł��܂őҋ@
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

' �{�^�����w��Index����������
Public Sub pushButton(ByRef form As IUIAutomationElement, ByVal ix As Long)
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ClassNamePropertyId, "Button")

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)

    Dim ptn As IUIAutomationInvokePattern
    Set ptn = list.GetElement(ix).GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId)
    Call ptn.Invoke

End Sub


' �w���ClassName��TextBox�ɒl��ݒ肷��
Public Sub SetText(ByRef form As IUIAutomationElement, ByVal ix As Long, ByVal text As String)
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ClassNamePropertyId, "TextBox")

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)


    Dim editValue As IUIAutomationValuePattern
    Set editValue = list.GetElement(ix).GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId)
    Call editValue.SetValue(text)

End Sub

' �w���AutomationId��TextBox�ɒl��ݒ肷��
Public Sub SetTextById(ByRef form As IUIAutomationElement, ByVal id As String, ByVal text As String)
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_AutomationIdPropertyId, id)

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)


    Dim editValue As IUIAutomationValuePattern
    Set editValue = list.GetElement(0).GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId)
    Call editValue.SetValue(text)

End Sub

' �w���ClassName��TextBox�̒l���擾����
Public Function GetStatusBarItemText(ByRef form As IUIAutomationElement, ByVal ix As Long) As String
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ClassNamePropertyId, "StatusBarItem")

    Dim list As IUIAutomationElementArray
    Set list = form.FindAll(TreeScope_Element Or TreeScope_Descendants, cnd)

    GetStatusBarItemText = list.GetElement(ix).CurrentName

End Function