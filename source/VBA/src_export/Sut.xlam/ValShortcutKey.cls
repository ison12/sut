VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValShortcutKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CommandBarControlオブジェクト
Private m_commandBarControl As commandBarControl

' ショートカットキーラベル
Private m_shortcutKeyLabel As String
' ショートカットキーコード Application#OnKeyに設定するコード
Private m_shortcutKeyCode  As String

Public Property Get commandName() As String

    commandName = m_commandBarControl.DescriptionText
End Property

Public Property Get commandBarControl() As commandBarControl

    Set commandBarControl = m_commandBarControl
End Property

Public Property Let commandBarControl(ByVal vNewValue As commandBarControl)

    Set m_commandBarControl = vNewValue
End Property

Public Property Get shortcutKeyCode() As String

    shortcutKeyCode = m_shortcutKeyCode
End Property

Public Property Let shortcutKeyCode(ByVal vNewValue As String)

    m_shortcutKeyCode = vNewValue
End Property

Public Property Get shortcutKeyLabel() As String

    shortcutKeyLabel = m_shortcutKeyLabel
End Property

Public Property Let shortcutKeyLabel(ByVal vNewValue As String)

    m_shortcutKeyLabel = vNewValue
End Property

Public Function Clone() As ValShortcutKey

    Dim ret As ValShortcutKey
    
    Set ret = New ValShortcutKey
    ret.commandBarControl = m_commandBarControl
    ret.shortcutKeyCode = m_shortcutKeyCode
    ret.shortcutKeyLabel = m_shortcutKeyLabel
    
    Set Clone = ret

End Function
