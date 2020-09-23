Attribute VB_Name = "Module1"
Option Explicit
'Intelligent Solutions For CommandButtons
'By; Jovica Mizdrak
'Date: Feb,27,2004
'
'PLEASE VOTE!!!
'
'As you can see in "BS_CAP_LEFT" its value
'means the object is in the middle. Alignment is changed
'every 100th number as u can see, like
'&H100& , &H200& , ... and cuts at , &H900& WHY?
'Because is only allows 3 digits meaning not more than 1000
'if you type &H1000& it would mean &H100&
'
'anything under 100 means some object
'and every 100th number means objects alignment
'
'HAVE FUN !!!

'VOTE PLEASE!!!

'---------- Settings ----------
Public Enum BUTTON_STYLE
    BS_CAP_CENTER = &H300&
    BS_CAP_LEFT = &H100&
    BS_CAP_RIGHT = &H200&
    BS_CAP_TOP = &H400&
    BS_CAP_TOPLEFT = &H500&
    BS_CAP_TOPRIGHT = &H600&
    BS_CAP_BOTTOM = &H800&
    BS_CAP_BOTTOMLEFT = &H900&
    BS_AS_CHECKBOX_LEFT = &H2&
    BS_AS_CHECKBOX_LEFT_UNCHECKABLE = &H4&
    BS_AS_FRAME = &H6&
    BS_AS_OPTION_BUTTON = &H8&
    BS_AS_CHECKBOX_RIGHT = &H22&
    BS_AS_CHECKBOX_RIGHT_UNCHECKABLE = &H24&
End Enum
'------------------------------------------
Private Const GWL_STYLE& = (-16) ' Alignment setting

'Declarations
Private Declare Function GetWindowLong& Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)

Private Declare Function SetWindowLong& Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long)

           
'----- Command Button Settings --------
Public Sub AlignCmdBtnTxt(Button As Object, _
   Style As BUTTON_STYLE)

Dim lHwnd As Long

On Error Resume Next
lHwnd = Button.hwnd
Dim lWnd As Long
Dim lRet As Long
If lHwnd = 0 Then Exit Sub

lWnd = GetWindowLong(lHwnd, GWL_STYLE)
lRet = SetWindowLong(Button.hwnd, GWL_STYLE, Style Or lWnd)

Button.Refresh
End Sub
