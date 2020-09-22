Attribute VB_Name = "mCCEx"
Option Explicit

Private Type uInitCommonControlsEx
   lSize As Long
   lICC  As Long
End Type

Private Const ICC_USEREX_CLASSES As Long = &H200

Private Declare Function InitCommonControlsEx Lib "comctl32" (iccex As uInitCommonControlsEx) As Boolean

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long



Private m_hMod As Long

Private Function InitCommonControlsVB() As Boolean
 
  Dim uICCEx As uInitCommonControlsEx
   
    On Error Resume Next
    
    With uICCEx
        .lSize = LenB(uICCEx)
        .lICC = ICC_USEREX_CLASSES
    End With
    Call InitCommonControlsEx(uICCEx)
    InitCommonControlsVB = (Err.Number = 0)
    
    On Error GoTo 0
End Function

Public Sub Main()
   m_hMod = LoadLibrary("shell32.dll") ' Prevents crashes on Windows XP
   Call InitCommonControlsVB
   Call fTest.Show
End Sub

Public Sub SafeEnd()
    Call FreeLibrary(m_hMod)
End Sub



