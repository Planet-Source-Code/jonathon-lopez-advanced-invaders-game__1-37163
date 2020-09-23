Attribute VB_Name = "mod_key"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Const kLeft = &H25
Public Const kRight = &H27
Public Const kUp = &H26
Public Const kDown = &H28
Public Const kReturn = 13
Public Const kShift = &H10
Public Const kCtrl = &H11
