VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8475
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCustomColorVBHex 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "&H00000000&"
      Top             =   3765
      Width           =   1680
   End
   Begin Test.ucScrollbar ucScrollbar2 
      Height          =   225
      Index           =   0
      Left            =   360
      Top             =   2850
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   397
      Max             =   255
      Orientation     =   1
      Style           =   1
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   270
      Left            =   345
      TabIndex        =   2
      Top             =   1890
      Value           =   1  'Checked
      Width           =   1170
   End
   Begin VB.ComboBox cbStyle 
      Height          =   330
      Left            =   345
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1395
      Width           =   2175
   End
   Begin Test.ucScrollbar ucScrollbar1 
      Height          =   4650
      Left            =   3120
      Top             =   1200
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   8202
   End
   Begin VB.ListBox lstEvents 
      Appearance      =   0  'Flat
      Height          =   4650
      Left            =   3810
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4335
   End
   Begin Test.ucScrollbar ucScrollbar2 
      Height          =   225
      Index           =   1
      Left            =   360
      Top             =   3150
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   397
      Max             =   255
      Orientation     =   1
      Style           =   1
   End
   Begin Test.ucScrollbar ucScrollbar2 
      Height          =   225
      Index           =   2
      Left            =   360
      Top             =   3450
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   397
      Max             =   255
      Orientation     =   1
      Style           =   1
   End
   Begin VB.Label lblCustomColor 
      Caption         =   "Custom color (RGB):"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2565
      Width           =   1830
   End
   Begin VB.Label lblStyle 
      Caption         =   "Style:"
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Top             =   1155
      Width           =   1365
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private m_lEventCount As Long
Private m_clrMyColor  As Long



Private Sub Form_Load()

  Dim lClr As Long
  
    With cbStyle
        .AddItem "0 - [sClassic]"
        .AddItem "1 - [sFlat]"
        .AddItem "2 - [sThemed]"
        .AddItem "3 - [sCustomDraw]"
        .ListIndex = 0
    End With
    
    m_lEventCount = 0
    lClr = GetSysColor(vbButtonFace And &HFF)
    ucScrollbar2(0) = (lClr And &HFF&)
    ucScrollbar2(1) = (lClr And &HFF00&) \ &H100
    ucScrollbar2(2) = (lClr And &HFF0000) \ &H10000
End Sub

Private Sub Form_Paint()
    
    Me.Line (0, 0)-(Me.ScaleWidth, 50), vbWhite, BF
    
    Me.CurrentX = 10
    Me.CurrentY = 8
    Me.Font.Size = 12
    Me.Font.Bold = True
    Me.Print "ucScrollbar simple demo"
    
    Me.CurrentX = 10
    Me.CurrentY = 30
    Me.Font.Size = 9
    Me.Font.Bold = False
    Me.Print "Classic, flat, themed and custom draw styles"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mCCEx.SafeEnd
End Sub





Private Sub cbStyle_Click()
    ucScrollbar1.Style = cbStyle.ListIndex
    lblCustomColor.Visible = (ucScrollbar1.Style = [sCustomDraw])
    ucScrollbar2(0).Visible = lblCustomColor.Visible
    ucScrollbar2(1).Visible = lblCustomColor.Visible
    ucScrollbar2(2).Visible = lblCustomColor.Visible
    txtCustomColorVBHex.Visible = lblCustomColor.Visible
End Sub

Private Sub chkEnabled_Click()
    ucScrollbar1.Enabled = CBool(chkEnabled)
End Sub

 
 
 
 
Private Sub ucScrollbar1_Change()
    Call pvAddEventString("_Change [Value: " & ucScrollbar1 & "]")
End Sub

Private Sub ucScrollbar1_Scroll()
    Call pvAddEventString("_Scroll [Value: " & ucScrollbar1 & "]")
End Sub

Private Sub ucScrollbar1_OnPaint( _
            ByVal lhDC As Long, _
            ByVal x1 As Long, ByVal y1 As Long, _
            ByVal x2 As Long, ByVal y2 As Long, _
            ByVal ePart As sbOnPaintPartCts, _
            ByVal eState As sbOnPaintPartStateCts _
            )

    '-- Testing a simple Custom-Draw method
    
    Select Case ePart
        
        Case [ppTLButton], [ppBRButton]
        
            If (ucScrollbar1.Enabled) Then
                If (eState = [ppsHot]) Then
                    Call DrawArrowEx(lhDC, x1, y1, x2, y2, afDFCS_SCROLLDOWN * -(ePart = [ppBRButton]), ShiftColor(m_clrMyColor, -25), ShiftColor(m_clrMyColor, -100))
                ElseIf (eState = [ppsPressed]) Then
                    Call DrawArrowEx(lhDC, x1, y1, x2, y2, afDFCS_SCROLLDOWN * -(ePart = [ppBRButton]), ShiftColor(m_clrMyColor, -50), ShiftColor(m_clrMyColor, -100))
                  Else
                    Call DrawArrowEx(lhDC, x1, y1, x2, y2, afDFCS_SCROLLDOWN * -(ePart = [ppBRButton]), m_clrMyColor, ShiftColor(m_clrMyColor, -50))
                End If
              Else
                Call DrawArrowEx(lhDC, x1, y1, x2, y2, afDFCS_SCROLLDOWN * -(ePart = [ppBRButton]), TranslateColor(vbButtonFace), TranslateColor(vb3DShadow))
            End If
            
        Case [ppTLTrack], [ppBRTrack]
        
            If (ucScrollbar1.Enabled) Then
                If (eState = [ppsPressed]) Then
                    Call FillRectEx(lhDC, x1, y1, x2, y2, ShiftColor(m_clrMyColor, -100))
                  Else
                    Call FillRectEx(lhDC, x1, y1, x2, y2, ShiftColor(m_clrMyColor, 25))
                End If
              Else
                Call FillRectEx(lhDC, x1, y1, x2, y2, TranslateColor(vb3DShadow))
            End If
            
        Case [ppNullTrack]
        
            If (ucScrollbar1.Enabled) Then
                Call FillRectEx(lhDC, x1, y1, x2, y2, ShiftColor(m_clrMyColor, -100))
              Else
                Call FillRectEx(lhDC, x1, y1, x2, y2, TranslateColor(vb3DShadow))
            End If
            
        Case [ppThumb]
        
            If (ucScrollbar1.Enabled) Then
                If (eState = [ppsHot]) Then
                    Call DrawRectEx(lhDC, x1, y1, x2, y2, ShiftColor(m_clrMyColor, -100))
                    Call FillRectEx(lhDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, ShiftColor(m_clrMyColor, -25))
                ElseIf (eState = [ppsPressed]) Then
                    Call DrawRectEx(lhDC, x1, y1, x2, y2, ShiftColor(m_clrMyColor, -150))
                    Call FillRectEx(lhDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, ShiftColor(m_clrMyColor, -50))
                  Else
                    Call DrawRectEx(lhDC, x1, y1, x2, y2, ShiftColor(m_clrMyColor, -50))
                    Call FillRectEx(lhDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, m_clrMyColor)
                End If
              Else
                Call DrawRectEx(lhDC, x1, y1, x2, y2, TranslateColor(vb3DShadow))
                Call FillRectEx(lhDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, TranslateColor(vbButtonFace))
            End If
    End Select
End Sub

Private Sub pvAddEventString(ByVal sString As String)
    
    With lstEvents
        m_lEventCount = m_lEventCount + 1
        
        If (.ListCount = 22) Then
            Call .RemoveItem(0)
        End If
        Call .AddItem(Format$(m_lEventCount, "00000 ") & sString)
        .ListIndex = .ListCount - 1
    End With
End Sub

'//

Private Sub ucScrollbar2_Change(Index As Integer)
    
    m_clrMyColor = RGB(ucScrollbar2(0), ucScrollbar2(1), ucScrollbar2(2))
    txtCustomColorVBHex = "&H00" & Format$(Hex(ucScrollbar2(2)), "00") & Format$(Hex(ucScrollbar2(1)), "00") & Format$(Hex(ucScrollbar2(0)), "00") & "&"
    Call ucScrollbar1.Refresh
End Sub

Private Sub ucScrollbar2_Scroll(Index As Integer)
    Call ucScrollbar2_Change(Index)
End Sub
