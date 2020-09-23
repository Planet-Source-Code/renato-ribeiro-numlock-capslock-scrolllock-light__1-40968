Attribute VB_Name = "Module1"
'luzes
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Global vk_numlock
Global vk_scrolllock
Global vk_capslock
Global light As Integer

Sub lightOff()
    If GetKeyState(vk_numlock) = 1 Then
        keybd_event vk_numlock, 0, 1, 0
        keybd_event vk_numlock, 0, 2, 0
    End If
    If GetKeyState(vk_capslock) = 1 Then
        keybd_event vk_capslock, 0, 1, 0
        keybd_event vk_capslock, 0, 2, 0
    End If
    If GetKeyState(vk_scrolllock) = 1 Then
        keybd_event vk_scrolllock, 0, 1, 0
        keybd_event vk_scrolllock, 0, 2, 0
    End If
End Sub

Sub lightOn()
    If light = 1 Then
        keybd_event vk_numlock, 0, 1, 0
        keybd_event vk_numlock, 0, 2, 0
        light = 2
    ElseIf light = 2 Then
        keybd_event vk_capslock, 0, 1, 0
        keybd_event vk_capslock, 0, 2, 0
        light = 3
    ElseIf light = 3 Then
        keybd_event vk_scrolllock, 0, 1, 0
        keybd_event vk_scrolllock, 0, 2, 0
        light = 1
    End If
End Sub
