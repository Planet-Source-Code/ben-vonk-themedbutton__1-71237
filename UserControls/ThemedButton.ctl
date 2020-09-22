VERSION 5.00
Begin VB.UserControl ThemedButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   384
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   972
   DefaultCancel   =   -1  'True
   MouseIcon       =   "ThemedButton.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "ThemedButton.ctx":058A
End
Attribute VB_Name = "ThemedButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ThemedButton Control
'
'Author Ben Vonk
'15-10-2008 First version, included: Paul Caton's self Subclass v1.1.0008 and thanks to LaVolpe for his DrawTransparentPicture routine
'08-11-2011 Second version, Fixed some bugs and add properties so the user can customize the button
'09-11-2011 Third version, Add single corner roundings for User themed button and fixed some bugs
'30-11-2011 Fourth version, Fixed some bugs, add OptionButton and CheckBox properties (the OptionButton can also be used in a multi selection array)
'09-12-2011 Fifth version, Fixed some bugs for non themed windows buttons, add Defaulted property for the custom CommandButton picture

Option Explicit
Option Compare Text

' Public Events
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

' Private Constants
Private Const BPS_DEFAULTED       As Long = 5
Private Const BPS_HOT             As Long = 2
Private Const BPS_NORMAL          As Long = 1
Private Const BPS_PRESSED         As Long = 3
Private Const ALL_MESSAGES        As Long = -1
Private Const GWL_WNDPROC         As Long = -4
Private Const PATCH_05            As Long = 93
Private Const PATCH_09            As Long = 137
Private Const WM_LBUTTONDBLCLK    As Long = &H203
Private Const WM_LBUTTONDOWN      As Long = &H201
Private Const WM_LBUTTONUP        As Long = &H202
Private Const WM_MOUSELEAVE       As Long = &H2A3
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_THEMECHANGED     As Long = &H31A
Private Const WM_TIMER            As Long = &H113

' Public Enumerations
Public Enum ButtonCorners
   AllCorners
   TopCorners
   TopLeftCorner
   TopRightCorner
   LeftCorners
   RightCorners
   BottomCorners
   BottomLeftCorner
   BottomRightCorner
End Enum

Public Enum ButtonThemeTypes
   Windows
   User
End Enum

Public Enum ButtonTypeConstants
   CommandButton
   OptionButton
   CheckBox
End Enum

Public Enum FocusStyles
   Button
   Text
End Enum

Public Enum PictureAlignConstants
   TopCenter
   LeftAlign
   Center
   RightAlign
   BottomCenter
End Enum

Public Enum PictureSizeConstants
   ps16x16
   ps24x24
   ps32x32
   ps48x48
   ps64x64
End Enum

' Private Enumerations
Private Enum MsgWhen
   MSG_AFTER = 1
   MSG_BEFORE = 2
   MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

Private Enum TrackMouseEventFlags
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

Public Enum ValueConstants
   Unchecked
   Checked
   Grayed
End Enum

' Private Types
Private Type MouseStateType
   Button                         As Integer
   Shift                          As Integer
   X                              As Single
   Y                              As Single
End Type

Private Type OSVersionInfo
   OSVSize                        As Long
   dwVerMajor                     As Long
   dwVerMinor                     As Long
   dwBuildNumber                  As Long
   PlatformID                     As Long
   szCSDVersion                   As String * 128
End Type

Private Type PointAPI
   X                              As Long
   Y                              As Long
End Type

Private Type Rect
   Left                           As Long
   Top                            As Long
   Right                          As Long
   Bottom                         As Long
End Type

Private Type ButtonPropertiesType
   ButtonRect                     As Rect
   FocusRect                      As Rect
   CaptionRect                    As Rect
   PictureSize                    As Long
End Type

Private Type SubclassDataType
   hWnd                           As Long
   nAddrSclass                    As Long
   nAddrOrig                      As Long
   nMsgCountA                     As Long
   nMsgCountB                     As Long
   aMsgTabelA()                   As Long
   aMsgTabelB()                   As Long
End Type

Private Type TrackMouseEventStruct
   cbSize                         As Long
   dwFlags                        As TrackMouseEventFlags
   hwndTrack                      As Long
   dwHoverTime                    As Long
End Type

' Private Variables
Private m_CaptionAlign            As AlignmentConstants
Private m_CaptionShadow           As Boolean
Private InControl                 As Boolean
Private IsFocused                 As Boolean
Private IsHit                     As Boolean
Private IsThemed                  As Boolean
Private IsThemedWindows           As Boolean
Private m_OptionButtonMultiSelect As Boolean
Private m_ShowFocusRect           As Boolean
Private m_UseParentBackColor      As Boolean
Private MouseDown                 As Boolean
Private SpaceKeyPressed           As Boolean
Private TrackUser32               As Boolean
Private m_ButtonCorner            As ButtonCorners
Private ButtonProperties          As ButtonPropertiesType
Private m_ButtonThemeType         As ButtonThemeTypes
Private m_ButtonType              As ButtonTypeConstants
Private m_FocusStyle              As FocusStyles
Private AccessKeyPointer          As Integer
Private ButtonState               As Integer
Private m_ButtonRounding          As Integer
Private m_BackColor               As Long
Private m_ForeColor               As Long
Private m_OverColor               As Long
Private SubclassMemory            As Long
Private TimerID                   As Long
Private MouseState                As MouseStateType
Private m_PictureAlign            As PictureAlignConstants
Private m_PictureSize             As PictureSizeConstants
Private m_ButtonPicture(12)       As StdPicture
Private m_Picture                 As StdPicture
Private m_Caption                 As String
Private SubclassData()            As SubclassDataType
Private m_Value                   As ValueConstants

' Private API's
Private Declare Function TrackMouseEventComCtl Lib "ComCtl32" Alias "_TrackMouseEvent" (lpEventTrack As TrackMouseEventStruct) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Integer
Private Declare Function GetBkColor Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetMapMode Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTextColor Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function PatBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "GDI32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function timeGetTime Lib "WinMM" () As Long
Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "User32" (ByVal hDC As Long, ByRef lpRect As Rect) As Long
Private Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "User32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function OffsetRect Lib "User32" (ByRef lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TrackMouseEvent Lib "User32" (lpEventTrack As TrackMouseEventStruct) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme" (ByVal lngTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UxTheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function GetCurrentThemeName Lib "UxTheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "UxTheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub MouseEvents Lib "User32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Dim blnMouseLeave          As Boolean
Dim blnMouseMove           As Boolean
Dim lngButtonState         As Long

   lngButtonState = ButtonState
   
   Select Case uMsg
      Case WM_LBUTTONDBLCLK
         Call MouseEvents(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
         
      Case WM_LBUTTONDOWN
         MouseDown = True
         IsFocused = True
         ButtonState = BPS_PRESSED
         SetCapture UserControl.hWnd
         
      Case WM_LBUTTONUP
         MouseDown = False
         ReleaseCapture
         
         If IsHit Then
            ButtonState = BPS_HOT
            
         ElseIf m_ButtonType = CommandButton Then
            ButtonState = BPS_DEFAULTED
            
         Else
            ButtonState = BPS_NORMAL
         End If
         
         If InControl Then Call UserControl_Click
         
      Case WM_MOUSELEAVE
         If SpaceKeyPressed Then Exit Sub
         
         If IsHit Or IsFocused Then
            If m_ButtonType = CommandButton Then
               If ButtonState = BPS_DEFAULTED Then lngButtonState = -1
               
               ButtonState = BPS_DEFAULTED
               
            Else
               If ButtonState = BPS_NORMAL Then lngButtonState = -1
               
               ButtonState = BPS_NORMAL
            End If
            
            Call SetBackColor(m_BackColor)
            
         Else
            If ButtonState = BPS_NORMAL Then lngButtonState = -1
            
            ButtonState = BPS_NORMAL
            
            Call SetBackColor(m_BackColor)
         End If
         
         InControl = False
         blnMouseLeave = True
         
      Case WM_MOUSEMOVE
         If SpaceKeyPressed Then Exit Sub
         
         Call TrackMouseLeave(lhWnd)
         
         If InControl Then
            Call SetBackColor(m_OverColor)
            
            If lngButtonState = ButtonState Then lngButtonState = -1
            
         Else
            ButtonState = BPS_HOT
            
            Call SetBackColor(m_BackColor)
         End If
         
         blnMouseMove = True
         
      Case WM_THEMECHANGED
         ' Wait a while so all controls can change the theme
         lngButtonState = timeGetTime
         
         Do
            DoEvents
         Loop Until (timeGetTime - lngButtonState) > 60
         
         IsThemed = CheckIsThemed
         lngButtonState = -1
         
         If InControl Then
            Call SetBackColor(m_OverColor)
            
         Else
            Call SetBackColor(m_BackColor)
         End If
         
         Call Refresh
         
      Case WM_TIMER
         Call ResetOptionButtons(ByTimer:=True)
   End Select
   
   If ButtonState <> lngButtonState Then Call Refresh
   
   With MouseState
      If blnMouseLeave Then
         RaiseEvent MouseLeave(.Button, .Shift, .X, .Y)
         
      ElseIf blnMouseMove Then
         RaiseEvent MouseMove(.Button, .Shift, .X, .Y)
      End If
   End With

End Sub

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

   Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   Debug.Assert Subclass_AddrFunc

End Function

Private Function Subclass_Index(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean) As Long

   For Subclass_Index = UBound(SubclassData) To 0 Step -1
      If SubclassData(Subclass_Index).hWnd = lhWnd Then
         If Not bAdd Then Exit Function
         
      ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
         If bAdd Then Exit Function
      End If
   Next 'Subclass_Index
   
   If Not bAdd Then Debug.Assert False

End Function

Private Function Subclass_InIDE() As Boolean

   Debug.Assert Subclass_SetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Initialize(ByVal lhWnd As Long) As Long

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

   If bytBuffer(1) Then
      lngIndex = Subclass_Index(lhWnd, True)
      
      If lngIndex = -1 Then
         lngIndex = UBound(SubclassData()) + 1
         
         ReDim Preserve SubclassData(lngIndex) As SubclassDataType
      End If
      
      Subclass_Initialize = lngIndex
      
   Else
      strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
      
      For lngCount = 1 To CODE_LEN
         bytBuffer(lngCount) = Val("&H" & Left(strHex, 2))
         strHex = Mid(strHex, 3)
      Next 'lngCount
      
      If Subclass_InIDE Then
         bytBuffer(16) = &H90
         bytBuffer(17) = &H90
         lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
         
         If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
      End If
      
      lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
      lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
      
      ReDim SubclassData(0) As SubclassDataType
   End If
   
   With SubclassData(lngIndex)
      .hWnd = lhWnd
      .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
      .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)
      
      Call CopyMemory(ByVal .nAddrSclass, bytBuffer(1), CODE_LEN)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
   End With

End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean

   Subclass_SetTrue = True
   bValue = True

End Function

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lhWnd))
      If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

   ReDim lngOffset(1) As Long
   
   If uMsg = ALL_MESSAGES Then
      nMsgCount = ALL_MESSAGES
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = 0 Then
            aMsgTabel(lngEntry) = uMsg
            
            GoTo ExitSub
            
         ElseIf aMsgTabel(lngEntry) = uMsg Then
            GoTo ExitSub
         End If
      Next 'lngEntry
      
      nMsgCount = nMsgCount + 1
      
      ReDim Preserve aMsgTabel(1 To nMsgCount) As Long
      
      aMsgTabel(nMsgCount) = uMsg
   End If
   
   If When = MSG_BEFORE Then
      lngOffset(0) = PATCH_04
      lngOffset(1) = PATCH_05
      
   Else
      lngOffset(0) = PATCH_08
      lngOffset(1) = PATCH_09
   End If
   
   If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
   
   Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
   
ExitSub:
   Erase lngOffset

End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

Private Sub Subclass_Stop(ByVal lhWnd As Long)

   With SubclassData(Subclass_Index(lhWnd))
      SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig
      
      Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
      
      GlobalFree .nAddrSclass
      .hWnd = 0
      .nMsgCountA = 0
      .nMsgCountB = 0
      Erase .aMsgTabelA, .aMsgTabelB
   End With

End Sub

Private Sub Subclass_Terminate()

Dim lngCount As Long

   For lngCount = UBound(SubclassData) To 0 Step -1
      If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)
   Next 'lngCount

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. (Only if Windows is not themed or ButtonType is set as OptionButton or CheckBox!)"

   BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   If m_BackColor <> NewBackColor Then
      m_UseParentBackColor = False
      PropertyChanged "UseParentBackColor"
   End If
   
   m_BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call SetBackColor(m_BackColor)
   Call Refresh

End Property

Public Property Get ButtonCorner() As ButtonCorners
Attribute ButtonCorner.VB_Description = "Returns/sets the corners that will be rounded. (Only if ButtonThemeType is set as User!)"

   ButtonCorner = m_ButtonCorner

End Property

Public Property Let ButtonCorner(ByVal NewButtonCorner As ButtonCorners)

   m_ButtonCorner = NewButtonCorner
   PropertyChanged "ButtonCorner"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonDefaulted() As StdPicture
Attribute ButtonDefaulted.VB_Description = "Returns/sets a graphic to be displayed when the control is defaulted. (Only if ButtonThemeType is set as User and ButtonType is set as CommandButton!)"

   Set ButtonDefaulted = m_ButtonPicture(12)

End Property

Public Property Let ButtonDefaulted(ByRef NewButtonDefaulted As StdPicture)

   Set ButtonDefaulted = NewButtonDefaulted

End Property

Public Property Set ButtonDefaulted(ByRef NewButtonDefaulted As StdPicture)

   Set m_ButtonPicture(12) = NewButtonDefaulted
   PropertyChanged "ButtonDefaulted"
   CheckButtonThemeType
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonDisabled() As StdPicture
Attribute ButtonDisabled.VB_Description = "Returns/sets a graphic to be displayed when the control is disabled. (Only if ButtonThemeType is set as User!)"

   Set ButtonDisabled = m_ButtonPicture(3)

End Property

Public Property Let ButtonDisabled(ByRef NewButtonDisabled As StdPicture)

   Set ButtonDisabled = NewButtonDisabled

End Property

Public Property Set ButtonDisabled(ByRef NewButtonDisabled As StdPicture)

   Set m_ButtonPicture(3) = NewButtonDisabled
   PropertyChanged "ButtonDisabled"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonDisabledGrayed() As StdPicture
Attribute ButtonDisabledGrayed.VB_Description = "Returns/sets a graphic to be displayed when the control is disabled. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonDisabledGrayed = m_ButtonPicture(11)

End Property

Public Property Let ButtonDisabledGrayed(ByRef NewButtonDisabledGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonDisabledGrayed = NewButtonDisabledGrayed

End Property

Public Property Set ButtonDisabledGrayed(ByRef NewButtonDisabledGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(11) = NewButtonDisabledGrayed
   PropertyChanged "ButtonDisabledGrayed"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonDisabledValued() As StdPicture
Attribute ButtonDisabledValued.VB_Description = "Returns/sets a graphic to be displayed when the control is disabled. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonDisabledValued = m_ButtonPicture(7)

End Property

Public Property Let ButtonDisabledValued(ByRef NewButtonDisabledValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonDisabledValued = NewButtonDisabledValued

End Property

Public Property Set ButtonDisabledValued(ByRef NewButtonDisabledValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(7) = NewButtonDisabledValued
   PropertyChanged "ButtonDisabledValued"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonNormal() As StdPicture
Attribute ButtonNormal.VB_Description = "Returns/sets a graphic to be displayed in an button normal state of the control. (Only if ButtonThemeType is set as User!)"

   Set ButtonNormal = m_ButtonPicture(0)

End Property

Public Property Let ButtonNormal(ByRef NewButtonNormal As StdPicture)

   Set ButtonNormal = NewButtonNormal

End Property

Public Property Set ButtonNormal(ByRef NewButtonNormal As StdPicture)

   Set m_ButtonPicture(0) = NewButtonNormal
   PropertyChanged "ButtonNormal"
   CheckButtonThemeType
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonNormalGrayed() As StdPicture
Attribute ButtonNormalGrayed.VB_Description = "Returns/sets a graphic to be displayed in an button normal state of the control. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonNormalGrayed = m_ButtonPicture(8)

End Property

Public Property Let ButtonNormalGrayed(ByRef NewButtonNormalGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonNormalGrayed = NewButtonNormalGrayed

End Property

Public Property Set ButtonNormalGrayed(ByRef NewButtonNormalGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(8) = NewButtonNormalGrayed
   PropertyChanged "ButtonNormalGrayed"
   CheckButtonThemeType
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonNormalValued() As StdPicture
Attribute ButtonNormalValued.VB_Description = "Returns/sets a graphic to be displayed in an button normal state of the control. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonNormalValued = m_ButtonPicture(4)

End Property

Public Property Let ButtonNormalValued(ByRef NewButtonNormalValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonNormalValued = NewButtonNormalValued

End Property

Public Property Set ButtonNormalValued(ByRef NewButtonNormalValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(4) = NewButtonNormalValued
   PropertyChanged "ButtonNormalValued"
   CheckButtonThemeType
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonOver() As StdPicture
Attribute ButtonOver.VB_Description = "Returns/sets a graphic to be displayed in an button over state of the control. (Only if ButtonThemeType is set as User!)"

   Set ButtonOver = m_ButtonPicture(1)

End Property

Public Property Let ButtonOver(ByRef NewButtonOver As StdPicture)

   Set ButtonOver = NewButtonOver

End Property

Public Property Set ButtonOver(ByRef NewButtonOver As StdPicture)

   Set m_ButtonPicture(1) = NewButtonOver
   PropertyChanged "ButtonOver"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonOverGrayed() As StdPicture
Attribute ButtonOverGrayed.VB_Description = "Returns/sets a graphic to be displayed in an button over state of the control. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonOverGrayed = m_ButtonPicture(9)

End Property

Public Property Let ButtonOverGrayed(ByRef NewButtonOverGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonOverGrayed = NewButtonOverGrayed

End Property

Public Property Set ButtonOverGrayed(ByRef NewButtonOverGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(9) = NewButtonOverGrayed
   PropertyChanged "ButtonOverGrayed"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonOverValued() As StdPicture
Attribute ButtonOverValued.VB_Description = "Returns/sets a graphic to be displayed in an button over state of the control. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonOverValued = m_ButtonPicture(5)

End Property

Public Property Let ButtonOverValued(ByRef NewButtonOverValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonOverValued = NewButtonOverValued

End Property

Public Property Set ButtonOverValued(ByRef NewButtonOverValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(5) = NewButtonOverValued
   PropertyChanged "ButtonOverValued"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonPressed() As StdPicture
Attribute ButtonPressed.VB_Description = "Returns/sets a graphic to be displayed in an button pressed state of the control. (Only if ButtonThemeType is set as User!)"

   Set ButtonPressed = m_ButtonPicture(2)

End Property

Public Property Let ButtonPressed(ByRef NewButtonPressed As StdPicture)

   Set ButtonPressed = NewButtonPressed

End Property

Public Property Set ButtonPressed(ByRef NewButtonPressed As StdPicture)

   Set m_ButtonPicture(2) = NewButtonPressed
   PropertyChanged "ButtonPressed"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonPressedGrayed() As StdPicture
Attribute ButtonPressedGrayed.VB_Description = "Returns/sets a graphic to be displayed in an button pressed state of the control. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonPressedGrayed = m_ButtonPicture(10)

End Property

Public Property Let ButtonPressedGrayed(ByRef NewButtonPressedGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonPressedGrayed = NewButtonPressedGrayed

End Property

Public Property Set ButtonPressedGrayed(ByRef NewButtonPressedGrayed As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(10) = NewButtonPressedGrayed
   PropertyChanged "ButtonPressedGrayed"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonPressedValued() As StdPicture
Attribute ButtonPressedValued.VB_Description = "Returns/sets a graphic to be displayed in an button pressed state of the control. (Only if ButtonThemeType is set as User and ButtonType is set as OptionButton or CheckBox!)"

   Set ButtonPressedValued = m_ButtonPicture(6)

End Property

Public Property Let ButtonPressedValued(ByRef NewButtonPressedValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set ButtonPressedValued = NewButtonPressedValued

End Property

Public Property Set ButtonPressedValued(ByRef NewButtonPressedValued As StdPicture)

   If m_ButtonType = CommandButton Then Exit Property
   
   Set m_ButtonPicture(6) = NewButtonPressedValued
   PropertyChanged "ButtonPressedValued"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonRounding() As Integer
Attribute ButtonRounding.VB_Description = "Returns/sets the curve value to rounding the control corners."

   ButtonRounding = m_ButtonRounding

End Property

Public Property Let ButtonRounding(ByVal NewButtonRounding As Integer)

   If m_ButtonRounding < 0 Then m_ButtonRounding = 0
   
   m_ButtonRounding = NewButtonRounding
   PropertyChanged "ButtonRounding"
   
   Call RoundControl
   Call Refresh

End Property

Public Property Get ButtonThemeType() As ButtonThemeTypes
Attribute ButtonThemeType.VB_Description = "Returns/sets a theme type of the ThemedButton control."

   ButtonThemeType = m_ButtonThemeType

End Property

Public Property Let ButtonThemeType(ByVal NewButtonThemeType As ButtonThemeTypes)

   m_ButtonThemeType = NewButtonThemeType
   PropertyChanged "ButtonThemeType"
   CheckButtonThemeType
   
   Call Refresh

End Property

Public Property Get ButtonType() As ButtonTypeConstants
Attribute ButtonType.VB_Description = "Returns/sets a button type of the ThemedButton control."

   ButtonType = m_ButtonType

End Property

Public Property Let ButtonType(ByVal NewButtonType As ButtonTypeConstants)

   m_ButtonType = NewButtonType
   PropertyChanged "ButtonType"
   
   If m_ButtonType <> CommandButton Then
      If m_ButtonPicture(0) Is Nothing Then PictureSize = ps16x16
      If m_CaptionAlign = vbCenter Then CaptionAlign = vbLeftJustify
      If m_ButtonType = OptionButton Then If m_Value = Grayed Then Value = Unchecked
      
      OverColor = m_BackColor
      FocusStyle = Text
      
      Call UserControl_Resize
   End If
   
   Call Refresh

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an ThemedButton control."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"

   Caption = m_Caption

End Property

Public Property Let Caption(ByVal NewCaption As String)

Dim intCount As Integer

   intCount = InStrRev(NewCaption, "&")
   AccessKeyPointer = 0
   
   Do While intCount
      If Mid(NewCaption, intCount, 2) = "&&" Then
         intCount = InStrRev(intCount - 1, NewCaption, "&")
         
      Else
         AccessKeyPointer = intCount + 1
         intCount = 0
      End If
   Loop
   
   If AccessKeyPointer Then AccessKeys = UCase(Mid(NewCaption, AccessKeyPointer, 1))
   
   m_Caption = NewCaption
   PropertyChanged "Caption"
   
   If (m_PictureAlign = Center) And Len(m_Caption) Then
      m_PictureAlign = TopCenter
      PropertyChanged "PictureAlign"
   End If
   
   Call GetPictureSize
   Call Refresh

End Property

Public Property Get CaptionAlign() As AlignmentConstants
Attribute CaptionAlign.VB_Description = "Returns/sets a alignment value for the caption of the ThemedButton control."

   CaptionAlign = m_CaptionAlign

End Property

Public Property Let CaptionAlign(ByVal NewCaptionAlign As AlignmentConstants)

   If NewCaptionAlign < vbLeftJustify Then NewCaptionAlign = vbLeftJustify
   If NewCaptionAlign > vbCenter Then NewCaptionAlign = vbCenter
   If (m_ButtonType <> CommandButton) And (NewCaptionAlign = vbCenter) Then NewCaptionAlign = vbLeftJustify
   
   m_CaptionAlign = NewCaptionAlign
   PropertyChanged "CaptionAlign"
   
   Call Refresh

End Property

Public Property Get CaptionShadow() As Boolean
Attribute CaptionShadow.VB_Description = "Determines whether the caption will being displayed with a shadow."

   CaptionShadow = m_CaptionShadow

End Property

Public Property Let CaptionShadow(ByVal NewCaptionShadow As Boolean)

   m_CaptionShadow = NewCaptionShadow
   PropertyChanged "CaptionShadow"
   
   Call Refresh

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   UserControl.Enabled = NewEnabled
   PropertyChanged "Enabled"
   
   If UserControl.Enabled Then ButtonState = BPS_NORMAL
   
   Call Refresh

End Property

Public Property Get FocusStyle() As FocusStyles
Attribute FocusStyle.VB_Description = "Returns/sets a focus style of the ThemedButton control."

   FocusStyle = m_FocusStyle

End Property

Public Property Let FocusStyle(ByVal NewFocusStyle As FocusStyles)

   If m_ButtonType <> CommandButton Then NewFocusStyle = Text
   
   m_FocusStyle = NewFocusStyle
   PropertyChanged "FocusStyle"
   
   Call Refresh

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = UserControl.Font

End Property

Public Property Let Font(ByVal NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByVal NewFont As StdFont)

   Set UserControl.Font = NewFont
   PropertyChanged "Font"
   
   Call Refresh

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."

   ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)

   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   
   Call Refresh

End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."

   Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Let MouseIcon(ByRef NewMouseIcon As StdPicture)

   Set MouseIcon = NewMouseIcon

End Property

Public Property Set MouseIcon(ByRef NewMouseIcon As StdPicture)

   On Local Error GoTo ErrorProperty
   Set UserControl.MouseIcon = NewMouseIcon
   
   If Not NewMouseIcon Is Nothing Then
      MousePointer = vbCustom
      PropertyChanged "MouseIcon"
   End If
   
   GoTo ExitProperty
   
ErrorProperty:
   If Not Ambient.UserMode Then MsgBox "Error: #" & Err.Number & vbCrLf & Err.Description & vbCrLf & "Select .ico or .cur files only.", vbCritical + vbOKOnly, Extender.Name
   
ExitProperty:
   On Local Error GoTo 0

End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."

   MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal NewMousePointer As MousePointerConstants)

   If NewMousePointer < vbDefault Then NewMousePointer = vbDefault
   If (NewMousePointer > vbSizeAll) And (NewMousePointer <> vbCustom) Then NewMousePointer = vbSizeAll
   
   UserControl.MousePointer = NewMousePointer
   PropertyChanged "MousePointer"

End Property

Public Property Get OptionButtonMultiSelect() As Boolean
Attribute OptionButtonMultiSelect.VB_Description = "Returns/sets a value to use the OptionButton for a multi selection array."

   OptionButtonMultiSelect = m_OptionButtonMultiSelect

End Property

Public Property Let OptionButtonMultiSelect(ByVal NewOptionButtonMultiSelect As Boolean)

   m_OptionButtonMultiSelect = NewOptionButtonMultiSelect
   PropertyChanged "OptionButtonMultiSelect"

End Property

Public Property Get OverColor() As OLE_COLOR
Attribute OverColor.VB_Description = "Returns/sets the color used when the mouse is in an object. (Only if Windows is not themed!)"

   OverColor = m_OverColor

End Property

Public Property Let OverColor(ByVal NewOverColor As OLE_COLOR)

   If ButtonType <> CommandButton Then NewOverColor = m_BackColor
   
   m_OverColor = NewOverColor
   PropertyChanged "OverColor"
   
   Call Refresh

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in an ThemedButton control. (Only if ButtonType is set as CommandButton!)"

   Set Picture = m_Picture

End Property

Public Property Let Picture(ByRef NewPicture As StdPicture)

   Set Picture = NewPicture

End Property

Public Property Set Picture(ByRef NewPicture As StdPicture)

   Set m_Picture = NewPicture
   PropertyChanged "Picture"
   
   Call Refresh

End Property

Public Property Get PictureAlign() As PictureAlignConstants
Attribute PictureAlign.VB_Description = "Returns/sets a alignment value for the picture in relation to the caption and/or ThemedButton control. (Only if ButtonType is set as CommandButton!)"

   PictureAlign = m_PictureAlign

End Property

Public Property Let PictureAlign(ByVal NewPictureAlign As PictureAlignConstants)

   If NewPictureAlign < TopCenter Then NewPictureAlign = TopCenter
   If NewPictureAlign > BottomCenter Then NewPictureAlign = BottomCenter
   If (NewPictureAlign = Center) And Len(m_Caption) Then NewPictureAlign = TopCenter
   
   m_PictureAlign = NewPictureAlign
   PropertyChanged "PictureAlign"
   
   Call GetPictureSize
   Call Refresh

End Property

Public Property Get PictureSize() As PictureSizeConstants
Attribute PictureSize.VB_Description = "Returns/sets a size value for the picture. (Only if ButtonType is set as CommandButton!)"

   PictureSize = m_PictureSize

End Property

Public Property Let PictureSize(ByVal NewPictureSize As PictureSizeConstants)

   If m_ButtonType <> CommandButton Then Exit Property
   If NewPictureSize < ps16x16 Then NewPictureSize = ps16x16
   If NewPictureSize > ps64x64 Then NewPictureSize = ps64x64
   
   m_PictureSize = NewPictureSize
   PropertyChanged "PictureSize"
   
   Call GetPictureSize
   Call Refresh

End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Determines whether a focus rectangle will being displayed."

   ShowFocusRect = m_ShowFocusRect

End Property

Public Property Let ShowFocusRect(ByVal NewShowFocusRect As Boolean)

   m_ShowFocusRect = NewShowFocusRect
   PropertyChanged "ShowFocusRect"
   
   Call Refresh

End Property

Public Property Get UseParentBackColor() As Boolean
Attribute UseParentBackColor.VB_Description = "Determines whether the parent background color can be used as background color. (Only if Windows is not themed or ButtonType is set as OptionButton or CheckBox!)"

   UseParentBackColor = m_UseParentBackColor

End Property

Public Property Let UseParentBackColor(ByVal NewUseParentBackColor As Boolean)

   m_UseParentBackColor = NewUseParentBackColor
   PropertyChanged "UseParentBackColor"
   
   If m_UseParentBackColor Then
      m_BackColor = Parent.BackColor
      
      Call SetBackColor(m_BackColor)
   End If
   
   Call Refresh

End Property

Public Property Get Value() As ValueConstants
Attribute Value.VB_Description = "Returns/sets the value of an object."

   Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As ValueConstants)

   If (m_ButtonType = OptionButton) And (NewValue = Grayed) Then NewValue = Unchecked
   
   m_Value = NewValue
   PropertyChanged "Value"
   
   Call Refresh

End Property

Public Function hWnd() As Long

   hWnd = UserControl.hWnd

End Function

Public Sub Refresh()

Const BPS_CHECKED  As Long = 4
Const BPS_DISABLED As Long = 4
Const BPS_MIXED    As Long = 8
Const BDR_RAISED   As Long = &H5
Const BDR_SUNKEN   As Long = &HA
Const BF_RIGHT     As Long = &H4
Const BF_TOP       As Long = &H2
Const BF_LEFT      As Long = &H1
Const BF_BOTTOM    As Long = &H8
Const BF_RECT      As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Dim lngBorder      As Long
Dim lngButtonState As Long
Dim lngButtonType  As Long
Dim lngColorItem   As Long
Dim lngColorWindow As Long
Dim lngLeft        As Long
Dim lngTheme       As Long
Dim lngTop         As Long
Dim rctButtonRect  As Rect

   lngButtonType = m_ButtonType + 1
   
   With UserControl
      If Not .Enabled Then ButtonState = BPS_DISABLED
      
      With ButtonProperties
         Cls
         
         With .ButtonRect
            rctButtonRect.Top = .Top
            rctButtonRect.Bottom = .Bottom
            
            If m_ButtonType = CommandButton Then
               rctButtonRect.Left = .Left
               rctButtonRect.Right = .Right
               
            Else
               rctButtonRect.Left = 1 + ((ScaleWidth - 17) And (m_CaptionAlign = vbRightJustify))
               rctButtonRect.Right = 17 + ((ScaleWidth - 18) And (m_CaptionAlign = vbRightJustify))
            End If
         End With
         
         If ButtonThemeType = User Then
            lngButtonState = ButtonState - 1
            
            If lngButtonState < 0 Then lngButtonState = 0
            
            If m_ButtonType = CommandButton Then
               ' 12 is the Index for Defaulted state
               If lngButtonState + 1 = BPS_DEFAULTED Then lngButtonState = 12
               
            Else
               lngButtonState = lngButtonState + (BPS_CHECKED And (m_Value = Checked)) + (BPS_MIXED And (m_Value = Grayed))
            End If
            
            If Not m_ButtonPicture(lngButtonState) Is Nothing Then
               If m_ButtonType = CommandButton Then
                  PaintPicture m_ButtonPicture(lngButtonState), 0, 0, ScaleWidth, ScaleHeight, , , , , vbSrcCopy
                  
               Else
                  lngTop = (ScaleHeight - 18) / 2
                  lngLeft = ((ScaleWidth - 18) And (m_CaptionAlign = vbRightJustify))
                  PaintPicture m_ButtonPicture(lngButtonState), lngLeft, lngTop, 18, 18, , , , , vbSrcCopy
               End If
            End If
            
         ElseIf IsThemed Then
            lngButtonState = ButtonState
            
            If (m_ButtonType <> CommandButton) Then
               If (ButtonState = BPS_NORMAL) And InControl Then ButtonState = BPS_HOT
               If m_Picture Is Nothing Then .PictureSize = 16
               
               lngButtonState = ButtonState + (BPS_CHECKED And (m_Value = Checked)) + (BPS_MIXED And (m_Value = Grayed))
            End If
            
            lngTheme = OpenThemeData(hWnd, StrPtr("Button"))
            DrawThemeBackground lngTheme, hDC, lngButtonType, lngButtonState, rctButtonRect, rctButtonRect
            CloseThemeData lngTheme
            
         ' Not Themed Windows
         Else
            If ButtonType = CommandButton Then
               If ButtonState = BPS_PRESSED Then
                  lngBorder = BDR_SUNKEN
                  
               ' BPS_NORMAL, BPS_HOT or BPS_DEFAULTED
               Else
                  lngBorder = BDR_RAISED
               End If
               
               If ButtonState = BPS_PRESSED Then
                  Line (1, 1)-(ScaleWidth - 3, ScaleHeight - 3), vbWhite, B
                  Line (1, 1)-(ScaleWidth - 2, ScaleHeight - 2), m_BackColor, B
                  Line (1, 1)-(ScaleWidth - 2, ScaleHeight - 2), vb3DShadow, B
                  
               Else
                  DrawEdge hDC, .ButtonRect, lngBorder, BF_RECT
               End If
               
               If IsFocused Or Ambient.DisplayAsDefault Then Line (0, 0)-(ScaleWidth - 2, ScaleHeight - 2), vbBlack, B
               
            Else
               .PictureSize = 16
               lngTop = 1
               lngLeft = 1 + ((ScaleWidth - 19) And (m_CaptionAlign = vbRightJustify))
               
               With rctButtonRect
                  .Top = lngTop
                  .Left = lngLeft
                  .Right = .Left + 16
                  .Bottom = .Top + 16
               End With
               
               If ButtonType = CheckBox Then
                  lngColorItem = vbBlack
                  
                  If ButtonState = BPS_PRESSED Then
                     lngColorWindow = vb3DLight
                     
                  ElseIf m_Value = Grayed Then
                     lngColorWindow = &HF2F1F1
                     lngColorItem = vbInactiveTitleBar
                     
                  Else
                     lngColorWindow = vbWindowBackground
                  End If
                  
                  DrawEdge hDC, rctButtonRect, BDR_SUNKEN, BF_RECT
                  Line (lngLeft + 2, lngTop + 2)-(lngLeft + 13, lngTop + 13), lngColorWindow, BF
                  
                  If m_Value <> Unchecked Then
                     DrawWidth = 2
                     Line (lngLeft + 4, lngTop + 7)-(lngLeft + 7, lngTop + 10), lngColorItem
                     Line (lngLeft + 6, lngTop + 10)-(lngLeft + 11, lngTop + 5), lngColorItem
                     DrawWidth = 1
                  End If
                  
               Else
                  ' OptionButton
                  lngLeft = lngLeft + 8
                  lngTop = lngTop + 8
                  Circle (lngLeft, lngTop), 7, vbInactiveTitleBar, 1, 4
                  Circle (lngLeft, lngTop), 7, vbWhite, 4, 1
                  Circle (lngLeft, lngTop), 6, vbBlack, 1, 4
                  Circle (lngLeft, lngTop), 6, &HF2F1F1, 4, 1
                  
                  If ButtonState = BPS_PRESSED Then
                     lngColorWindow = vb3DLight
                     
                  Else
                     lngColorWindow = vbWindowBackground
                  End If
                  
                  FillStyle = vbFSSolid
                  FillColor = lngColorWindow
                  Circle (lngLeft, lngTop), 5, lngColorWindow
                  FillColor = vbBlack
                  
                  If m_Value <> Unchecked Then Circle (lngLeft, lngTop), 4, lngColorWindow
                  
                  FillStyle = vbFSTransparent
               End If
            End If
         End If
         
         If Len(m_Caption) Then Call DrawCaption(.PictureSize + 5)
         If (Not m_Picture Is Nothing) And (m_ButtonType = CommandButton) Then Call DrawPicture(.PictureSize)
         
         If IsFocused And m_ShowFocusRect Then
            SetTextColor hDC, vbBlack
            DrawFocusRect hDC, .FocusRect
         End If
      End With
      
      .Refresh
   End With

End Sub

Private Function CheckButtonThemeType() As Boolean

Dim intIndex As Integer

   If Not m_ButtonPicture(0) Is Nothing Then
      For intIndex = 1 To 12
         If m_ButtonPicture(intIndex) Is Nothing Then Set m_ButtonPicture(intIndex) = m_ButtonPicture(0)
      Next 'intIndex
      
      If m_ButtonThemeType = User Then
         If m_ButtonType = CommandButton Then
            Height = ScaleY(m_ButtonPicture(0).Height, vbHimetric, vbTwips)
            Width = ScaleX(m_ButtonPicture(0).Width, vbHimetric, vbTwips)
            CheckButtonThemeType = True
         End If
      End If
   End If

End Function

Private Function CheckIsThemed() As Boolean

Const VER_PLATFORM_WIN32_NT As Long = 2

Dim lngLibrary              As Long
Dim osvInfo                 As OSVersionInfo
Dim strTheme                As String
Dim strName                 As String

   With osvInfo
      .OSVSize = Len(osvInfo)
      GetVersionEx osvInfo
      
      If .PlatformID = VER_PLATFORM_WIN32_NT Then
         If ((.dwVerMajor > 4) And .dwVerMinor) Or (.dwVerMajor > 5) Then
            IsThemedWindows = True
            lngLibrary = LoadLibrary("UXTheme")
            
            If lngLibrary Then
               strTheme = String(255, 0)
               GetCurrentThemeName StrPtr(strTheme), Len(strTheme), 0, 0, 0, 0
               strTheme = StripNull(strTheme)
               
               If Len(strTheme) Then
                  strName = String(255, 0)
                  GetThemeDocumentationProperty StrPtr(strTheme), StrPtr("ThemeName"), StrPtr(strName), Len(strName)
                  CheckIsThemed = (StripNull(strName) <> "")
               End If
               
               FreeLibrary lngLibrary
            End If
         End If
      End If
   End With

End Function

Private Function IsFunctionSupported(ByVal sFunction As String, ByVal sModule As String) As Boolean

Dim lngModule As Long

   lngModule = GetModuleHandle(sModule)
   
   If lngModule = 0 Then lngModule = LoadLibrary(sModule)
   
   If lngModule Then
      IsFunctionSupported = GetProcAddress(lngModule, sFunction)
      FreeLibrary lngModule
   End If

End Function

Private Function StripNull(ByVal Text As String) As String

   StripNull = Split(Text, vbNullChar, 2)(0)

End Function

Private Sub CheckIsDefault()

   If m_ButtonType = CommandButton Then
      If Not Ambient.UserMode Then
         If Ambient.DisplayAsDefault Then
            ButtonState = BPS_DEFAULTED
            
         Else
            ButtonState = BPS_NORMAL
         End If
         
         Call Refresh
      End If
      
   Else
      ButtonState = BPS_NORMAL
      
      Call Refresh
   End If

End Sub

Private Sub DrawCaption(ByVal PictureSpace As Long)

Const DT_CALCRECT  As Long = &H400
Const DT_CENTER    As Long = &H1
Const DT_LEFT      As Long = &H0
Const DT_RIGHT     As Long = &H2
Const DT_WORDBREAK As Long = &H10
Const vbShadow     As Long = &HDCDCDC

Dim blnDown        As Boolean
Dim lngAlignment   As Long
Dim lngBottom      As Long
Dim lngRight       As Long
Dim rctCaption(1)  As Rect
Dim strCaption     As String

   With rctCaption(0)
      If Not m_Picture Is Nothing Then
         If (m_PictureAlign = TopCenter) Or (m_PictureAlign = BottomCenter) Then
            lngBottom = PictureSpace
            
         Else
            lngRight = PictureSpace
         End If
         
      ElseIf m_ButtonType <> CommandButton Then
         lngRight = PictureSpace
      End If
      
      .Right = ScaleWidth - 12 - lngRight
      .Bottom = ScaleHeight - 12 - lngBottom
      strCaption = m_Caption
      OffsetRect rctCaption(0), 6, 6
      
      If IsThemedWindows Then
         DrawTextW hDC, StrPtr(strCaption), Len(strCaption), rctCaption(0), DT_CALCRECT Or DT_WORDBREAK
         
      Else
         DrawText hDC, strCaption, Len(strCaption), rctCaption(0), DT_CALCRECT Or DT_WORDBREAK
      End If
      
      lngRight = -lngRight   ' Picture = LeftAlign
      lngBottom = -lngBottom ' Picture = TopCenter
      OffsetRect rctCaption(0), (ScaleWidth - .Right - 6 - lngRight) / 2, (ScaleHeight - .Bottom - 6 - lngBottom) / 2
      blnDown = ((ButtonState = BPS_PRESSED) And (m_ButtonType = CommandButton))
      .Top = .Top - blnDown
      .Bottom = .Bottom - blnDown
      .Left = 6 - blnDown
      .Right = ScaleWidth - 6 - blnDown
      
      If m_PictureAlign = BottomCenter Then
         .Top = .Top + lngBottom
         
      ElseIf m_PictureAlign = LeftAlign Then
         .Left = .Left - lngRight
         
      ElseIf m_PictureAlign = RightAlign Then
         .Right = .Right + lngRight
      End If
      
      If m_CaptionAlign = vbLeftJustify Then
         lngAlignment = DT_LEFT
         
      ElseIf m_CaptionAlign = vbRightJustify Then
         lngAlignment = DT_RIGHT
         
      ' vbCenter
      ElseIf m_ButtonType = CommandButton Then
         lngAlignment = DT_CENTER
      End If
      
      If m_ButtonType <> CommandButton Then
         lngAlignment = DT_LEFT
         ButtonProperties.FocusRect = rctCaption(0)
         PictureSpace = PictureSpace
         .Left = PictureSpace - ((PictureSpace - 6) And (m_CaptionAlign = vbRightJustify))
         .Right = ScaleWidth - PictureSpace + ((PictureSpace - 6) And (m_CaptionAlign = vbLeftJustify))
      End If
      
      rctCaption(1) = rctCaption(0)
      OffsetRect rctCaption(1), 0, 0
      
      If IsThemedWindows Then
         DrawTextW hDC, StrPtr(strCaption), Len(strCaption), rctCaption(1), DT_CALCRECT Or DT_WORDBREAK
         
      Else
         DrawText hDC, strCaption, Len(strCaption), rctCaption(1), DT_CALCRECT Or DT_WORDBREAK
      End If
      
      ButtonProperties.CaptionRect = rctCaption(0)
      
      If m_ButtonType = CommandButton Then
         If m_FocusStyle = Text Then
            ButtonProperties.FocusRect = rctCaption(1)
            
            With ButtonProperties.FocusRect
               If m_CaptionAlign = vbLeftJustify Then
                  .Left = .Left - 2
                  
               ElseIf m_CaptionAlign = vbRightJustify Then
                  .Left = .Left + (ScaleWidth - .Right) - 8
                  .Right = .Left + .Right - 4
                  
               ' vbCenter
               Else
                  .Left = .Left + (ScaleWidth - .Right) / 2 - 4
                  .Right = .Right + .Left - 4 - (1 And (ButtonState = BPS_PRESSED))
               End If
            End With
         End If
         
      Else
         ButtonProperties.FocusRect = rctCaption(1)
         
         With ButtonProperties.FocusRect
            .Left = rctCaption(1).Left - 2
            .Right = rctCaption(1).Right + 1
            .Bottom = rctCaption(1).Bottom + 1
            
            If .Top <= 0 Then .Top = 1
            If .Bottom >= ScaleHeight Then .Bottom = ScaleHeight - 1
         End With
      End If
      
      If UserControl.Enabled Then
         If m_CaptionShadow Then
            SetTextColor hDC, GetPixel(hDC, .Left, .Top) And vbShadow
            OffsetRect rctCaption(0), 1, 1
            
            If IsThemedWindows Then
               DrawTextW hDC, StrPtr(strCaption), Len(strCaption), rctCaption(0), DT_WORDBREAK Or lngAlignment
               
            Else
               DrawText hDC, strCaption, Len(strCaption), rctCaption(0), DT_WORDBREAK Or lngAlignment
            End If
         End If
         
         OffsetRect rctCaption(0), -1, -1
         SetTextColor hDC, m_ForeColor
         
      Else
         SetTextColor hDC, vbGrayText
      End If
      
      OffsetRect rctCaption(0), 0, 0
      
      If IsThemedWindows Then
         DrawTextW hDC, StrPtr(strCaption), Len(strCaption), rctCaption(0), DT_WORDBREAK Or lngAlignment
         
      Else
         DrawText hDC, strCaption, Len(strCaption), rctCaption(0), DT_WORDBREAK Or lngAlignment
      End If
   End With
   
   Erase rctCaption

End Sub

Private Sub DrawPicture(ByVal Size As Long)

Const DI_NORMAL     As Long = &H3
Const vbGray        As Long = &H808080
Const vbSrcReplace  As Long = &H220326
Const vbSrcGrayed   As Long = &HBEBABE
Const WHITENESS     As Long = &HFF0062

Dim blnDown         As Boolean
Dim lngBitmap       As Long
Dim lngBrush        As Long
Dim lngColor        As Long
Dim lngColorDC      As Long
Dim lngLeft         As Long
Dim lngMaskDC       As Long
Dim lngMemoryDC     As Long
Dim lngOldBackColor As Long
Dim lngOldBitmap    As Long
Dim lngOldBrush     As Long
Dim lngOldColor     As Long
Dim lngOldMemory    As Long
Dim lngOldObject    As Long
Dim lngSourceDC     As Long
Dim lngSourceWidth  As Long
Dim lngSourceHeight As Long
Dim lngTop          As Long
Dim rctPicture      As Rect

   With ButtonProperties.CaptionRect
      If m_Caption = "" Then
         .Top = Size + 14
         .Left = Size + 10
         .Right = ScaleWidth - Size - 9
         .Bottom = .Top + Size + 10
      End If
      
      lngLeft = (ScaleWidth - Size) \ 2
      blnDown = (ButtonState = BPS_PRESSED)
      
      If m_PictureAlign = TopCenter Then
         lngTop = (.Top - Size) \ 2
         lngLeft = lngLeft - blnDown
         
      ElseIf m_PictureAlign = BottomCenter Then
         lngTop = ScaleHeight - (.Bottom - .Top + Size) \ 2 - 2
         lngLeft = lngLeft - blnDown
         
      ' Center, LeftAlign or RightAlign
      Else
         lngTop = (ScaleHeight - Size) \ 2
         
         If m_PictureAlign = LeftAlign Then
            lngLeft = (.Left - Size) \ 2 + 2
            
         ElseIf m_PictureAlign = RightAlign Then
            lngLeft = .Right + (ScaleWidth - .Right - Size) \ 2 - 2
         End If
      End If
   End With
   
   lngTop = lngTop - blnDown
   
   If m_Picture.Type = vbPicTypeIcon Then
      DrawIconEx hDC, lngLeft, lngTop, m_Picture.Handle, Size, Size, 0, 0, DI_NORMAL
      
   Else
      lngSourceDC = CreateCompatibleDC(hDC)
      
      With m_Picture
         SelectObject lngSourceDC, .Handle
         lngSourceWidth = ScaleX(.Width, vbHimetric, vbPixels)
         lngSourceHeight = ScaleY(.Height, vbHimetric, vbPixels)
      End With
      
      lngColor = GetPixel(lngSourceDC, 0, 0)
      
      If lngColor < 0 Then lngColor = GetSysColor(lngColor And &HFF&)
      
      lngMaskDC = CreateCompatibleDC(hDC)
      lngMemoryDC = CreateCompatibleDC(hDC)
      lngColorDC = CreateCompatibleDC(hDC)
      lngOldColor = SelectObject(lngColorDC, CreateCompatibleBitmap(hDC, lngSourceWidth, lngSourceHeight))
      lngOldMemory = SelectObject(lngMemoryDC, CreateCompatibleBitmap(hDC, Size, Size))
      lngOldObject = SelectObject(lngMaskDC, CreateBitmap(lngSourceWidth, lngSourceHeight, 1, 1, ByVal 0&))
      SetMapMode lngMemoryDC, GetMapMode(hDC)
      SelectPalette lngMemoryDC, 0, True
      RealizePalette lngMemoryDC
      BitBlt lngMemoryDC, 0, 0, Size, Size, hDC, lngLeft, lngTop, vbSrcCopy
      SelectPalette lngColorDC, 0, True
      RealizePalette lngColorDC
      SetBkColor lngColorDC, GetBkColor(lngSourceDC)
      SetTextColor lngColorDC, GetTextColor(lngSourceDC)
      BitBlt lngColorDC, 0, 0, lngSourceWidth, lngSourceHeight, lngSourceDC, 0, 0, vbSrcCopy
      SetBkColor lngColorDC, lngColor
      SetTextColor lngColorDC, vbWhite
      BitBlt lngMaskDC, 0, 0, lngSourceWidth, lngSourceHeight, lngColorDC, 0, 0, vbSrcCopy
      SetBkColor lngColorDC, vbWhite
      SetTextColor lngColorDC, vbBlack
      BitBlt lngColorDC, 0, 0, lngSourceWidth, lngSourceHeight, lngMaskDC, 0, 0, vbSrcReplace
      StretchBlt lngMemoryDC, 0, 0, Size, Size, lngMaskDC, 0, 0, lngSourceWidth, lngSourceHeight, vbSrcAnd
      StretchBlt lngMemoryDC, 0, 0, Size, Size, lngColorDC, 0, 0, lngSourceWidth, lngSourceHeight, vbSrcPaint
      BitBlt hDC, lngLeft, lngTop, Size, Size, lngMemoryDC, 0, 0, vbSrcCopy
      DeleteObject SelectObject(lngColorDC, lngOldColor)
      DeleteObject SelectObject(lngMaskDC, lngOldObject)
      DeleteObject SelectObject(lngMemoryDC, lngOldMemory)
      DeleteDC lngMemoryDC
      DeleteDC lngMaskDC
      DeleteDC lngColorDC
      DeleteDC lngSourceDC
   End If
   
   If Not UserControl.Enabled Then
      lngMemoryDC = CreateCompatibleDC(hDC)
      lngBitmap = CreateCompatibleBitmap(hDC, Size, Size)
      lngOldBitmap = SelectObject(lngMemoryDC, lngBitmap)
      PatBlt lngMemoryDC, 0, 0, Size, Size, WHITENESS
      
      With rctPicture
         .Top = lngTop
         .Left = lngLeft
         .Right = Size
         .Bottom = Size
         OffsetRect rctPicture, -.Left, -.Top
      End With
      
      lngOldBackColor = SetBkColor(hDC, vbWhite)
      lngBrush = CreateSolidBrush(vbGray)
      lngOldBrush = SelectObject(hDC, lngBrush)
      BitBlt hDC, lngLeft, lngTop, Size, Size, lngMemoryDC, 0, 0, vbSrcGrayed
      SetBkColor hDC, lngOldBackColor
      SelectObject hDC, lngOldBrush
      SelectObject lngMemoryDC, lngOldBitmap
      DeleteObject lngBrush
      DeleteObject lngBitmap
      DeleteDC lngMemoryDC
   End If

End Sub

Private Sub GetPictureSize()

   ButtonProperties.PictureSize = 16 + 8 * (m_PictureSize + ((m_PictureSize - 2) And (m_PictureSize > 2)))

End Sub

Private Sub ResetOptionButtons(Optional ByVal ByTimer As Boolean)

Dim ctlControl As Control
Dim intIndex   As Integer
Dim intPointer As Integer
Dim strName    As String

   If ByTimer Then
      DoEvents
      
      For Each ctlControl In Parent.Controls
         If TypeOf ctlControl Is OptionButton Then
            If ctlControl.Value Then
               Value = Unchecked
               TimerID = KillTimer(hWnd, TimerID)
               Exit For
            End If
         End If
      Next 'ctlControl
      
   Else
      intIndex = -1
      strName = Ambient.DisplayName
      intPointer = InStr(strName, "(")
      
      If intPointer And (Right(strName, 1) = ")") Then
         intIndex = Val(Mid(strName, intPointer + 1))
         strName = Left(strName, intPointer - 1)
      End If
      
      For Each ctlControl In Parent.Controls
         If TypeOf ctlControl Is OptionButton Then ctlControl.Value = False
         
         If TypeOf ctlControl Is ThemedButton Then
            If ctlControl.ButtonType = OptionButton Then
               If InStr(ctlControl.Name, strName) Then
                  If intIndex > -1 Then
                     If ctlControl.Index <> intIndex Then
                        ctlControl.Value = Unchecked
                        TimerID = KillTimer(hWnd, TimerID)
                     End If
                  End If
                  
               ElseIf Not ctlControl.OptionButtonMultiSelect Then
                  ctlControl.Value = Unchecked
                  TimerID = KillTimer(hWnd, TimerID)
               End If
            End If
         End If
      Next 'ctlControl
      
      TimerID = SetTimer(hWnd, TimerID, 50, SubclassData(Subclass_Index(hWnd)).nAddrSclass)
   End If

End Sub

Private Sub RoundControl()

Const RGN_OR     As Long = 2

Dim intCurve     As Integer
Dim intX1        As Integer
Dim intX2        As Integer
Dim intY1        As Integer
Dim intY2        As Integer
Dim lngRegion(1) As Long

   If m_ButtonThemeType = User Then
      intCurve = m_ButtonRounding
      intX2 = 1
      intY2 = 1
      
      If (m_ButtonCorner = TopLeftCorner) Or (m_ButtonCorner = BottomLeftCorner) Then intX2 = intX2 + intCurve
      If (m_ButtonCorner = TopRightCorner) Or (m_ButtonCorner = BottomRightCorner) Then intX1 = -intCurve
      
   Else
      intX1 = 1
      intY1 = 1
      intCurve = (3 And IsThemed)
   End If
   
   lngRegion(0) = CreateRoundRectRgn(intX1, intY1, ScaleWidth + intX2, ScaleHeight + intY2, intCurve, intCurve)
   
   If m_ButtonThemeType = User Then
      If (m_ButtonCorner > AllCorners) And (m_ButtonCorner <= BottomRightCorner) Then
         Select Case m_ButtonCorner
            Case TopCorners, TopLeftCorner, TopRightCorner
               lngRegion(1) = CreateRectRgn(0, intCurve, ScaleWidth, ScaleHeight)
               
            Case LeftCorners
               lngRegion(1) = CreateRectRgn(intCurve, 0, ScaleWidth, ScaleHeight)
               
            Case RightCorners
               lngRegion(1) = CreateRectRgn(0, 0, ScaleWidth - intCurve, ScaleHeight)
               
            Case BottomCorners, BottomLeftCorner, BottomRightCorner
               lngRegion(1) = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight - intCurve)
         End Select
         
         CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
         DeleteObject lngRegion(1)
      End If
   End If
   
   SetWindowRgn hWnd, lngRegion(0), True
   DeleteObject lngRegion(0)
   Erase lngRegion

End Sub

Private Sub SetBackColor(ByVal Color As Long)

   If m_ButtonPicture(0) Is Nothing Then UserControl.BackColor = Color

End Sub

Private Sub SetOptionButtonCheckBoxValue()

   If m_ButtonType = CheckBox Then
      If m_Value = Checked Then
         Value = Unchecked
         
      Else
         Value = Checked
      End If
      
   ElseIf m_Value <> Checked Then
      Value = Checked
      
      If Not m_OptionButtonMultiSelect Then Call ResetOptionButtons
      
   ElseIf m_OptionButtonMultiSelect Then
      Value = Unchecked
   End If

End Sub

Private Sub TrackMouseLeave(ByVal lhWnd As Long)

Const TME_LEAVE   As Long = &H2&

Dim tmeMouseTrack As TrackMouseEventStruct

   With tmeMouseTrack
      .cbSize = Len(tmeMouseTrack)
      .dwFlags = TME_LEAVE
      .hwndTrack = lhWnd
   End With
   
   If TrackUser32 Then
      TrackMouseEvent tmeMouseTrack
      
   Else
      TrackMouseEventComCtl tmeMouseTrack
   End If

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

   If (KeyAscii = vbKeyReturn) Or UCase(Chr(KeyAscii)) = AccessKeys Then
      MouseState.Button = vbLeftButton
      
      Call UserControl_GotFocus
      Call UserControl_Click
   End If
   
   MouseState.Button = vbDefault

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

   If PropertyName = "DisplayAsDefault" Then Call CheckIsDefault
   
   If (PropertyName = "BackColor") And m_UseParentBackColor Then
      m_BackColor = Ambient.BackColor
      
      Call SetBackColor(m_BackColor)
      Call Refresh
   End If

End Sub

Private Sub UserControl_Click()

   Call Refresh
   
   DoEvents
   IsHit = True
   
   If MouseState.Button = vbLeftButton Then
      If m_ButtonType <> CommandButton Then Call SetOptionButtonCheckBoxValue
      
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub UserControl_GotFocus()

   IsFocused = True
   SpaceKeyPressed = False
   
   If Not MouseDown Then
      If Not InControl Then
         If m_ButtonType = CommandButton Then
            ButtonState = BPS_DEFAULTED
            
         Else
            ButtonState = BPS_NORMAL
         End If
      End If
      
      If ((ButtonType = OptionButton) And Not OptionButtonMultiSelect) Then
         Call SetOptionButtonCheckBoxValue
         
         RaiseEvent Click
      End If
   End If
   
   Call Refresh

End Sub

Private Sub UserControl_Initialize()

   IsThemed = CheckIsThemed

End Sub

Private Sub UserControl_InitProperties()

   m_CaptionAlign = vbCenter
   m_BackColor = vbButtonFace
   m_Caption = Ambient.DisplayName
   m_ForeColor = vbButtonText
   Font = Ambient.Font
   m_OverColor = &HE0E0E0
   m_PictureSize = ps32x32
   m_ShowFocusRect = True
   m_Value = Unchecked
   ButtonState = BPS_NORMAL
   
   Call GetPictureSize

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

Const WM_KEYDOWN As Long = &H100
Const VK_DOWN    As Long = &H28
Const VK_LEFT    As Long = &H25
Const VK_RIGHT   As Long = &H27
Const VK_UP      As Long = &H26

Dim lngKey       As Long
Dim lngParam     As Long

   Select Case KeyCode
      Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown
         Select Case KeyCode
            Case vbKeyLeft
               lngKey = VK_LEFT
               lngParam = &H4B0001
               
            Case vbKeyRight
               lngKey = VK_RIGHT
               lngParam = &H4D0001
               
            Case vbKeyUp
               lngKey = VK_UP
               lngParam = &H480001
               
            Case vbKeyDown
               lngKey = VK_DOWN
               lngParam = &H500001
         End Select
         
         KeyCode = 0
         PostMessage ContainerHwnd, WM_KEYDOWN, ByVal lngKey, ByVal lngParam
         
         If SpaceKeyPressed Then
            If (ButtonType = CheckBox) Or ((ButtonType = OptionButton) And OptionButtonMultiSelect) Then Call SetOptionButtonCheckBoxValue
            
            RaiseEvent Click
            SpaceKeyPressed = False
         End If
         
      Case vbKeySpace
         ButtonState = BPS_PRESSED
         SpaceKeyPressed = True
         
         Call Refresh
         
         RaiseEvent KeyDown(KeyCode, Shift)
   End Select

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)
   
   If (KeyCode = vbKeySpace) Then
      ButtonState = BPS_HOT
      
      If (ButtonType <> CommandButton) And SpaceKeyPressed Then Call SetOptionButtonCheckBoxValue
      
      Call Refresh
      
      If SpaceKeyPressed Then RaiseEvent Click
      
      SpaceKeyPressed = False
   End If

End Sub

Private Sub UserControl_LostFocus()

   If InControl Then
      ButtonState = BPS_HOT
      
   Else
      ButtonState = BPS_NORMAL
   End If
   
   SpaceKeyPressed = False
   IsFocused = False
   IsHit = False
   
   Call Refresh

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   With MouseState
      .Button = Button
      .Shift = Shift
      .X = X
      .Y = Y
   End With
   
   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim blnInRegion As Boolean
Dim ptaMouse    As PointAPI

   If SpaceKeyPressed Then Exit Sub
   
   GetCursorPos ptaMouse
   InControl = (WindowFromPoint(ptaMouse.X, ptaMouse.Y) = hWnd)
   
   If InControl And MouseDown Then
      ButtonState = BPS_PRESSED
      
   Else
      ButtonState = BPS_HOT
   End If
   
   Call Refresh

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)

   RaiseEvent OLECompleteDrag(Effect)

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

   RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)

End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)

   RaiseEvent OLESetData(Data, DataFormat)

End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

   RaiseEvent OLEStartDrag(Data, AllowedEffects)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_BackColor = .ReadProperty("BackColor", vbButtonFace)
      m_ButtonCorner = .ReadProperty("ButtonCorner", AllCorners)
      Set m_ButtonPicture(12) = .ReadProperty("ButtonDefaulted", Nothing)
      Set m_ButtonPicture(3) = .ReadProperty("ButtonDisabled", Nothing)
      Set m_ButtonPicture(11) = .ReadProperty("ButtonDisabledGrayed", Nothing)
      Set m_ButtonPicture(7) = .ReadProperty("ButtonDisabledValued", Nothing)
      Set m_ButtonPicture(0) = .ReadProperty("ButtonNormal", Nothing)
      Set m_ButtonPicture(8) = .ReadProperty("ButtonNormalGrayed", Nothing)
      Set m_ButtonPicture(4) = .ReadProperty("ButtonNormalValued", Nothing)
      Set m_ButtonPicture(1) = .ReadProperty("ButtonOver", Nothing)
      Set m_ButtonPicture(9) = .ReadProperty("ButtonOverGrayed", Nothing)
      Set m_ButtonPicture(5) = .ReadProperty("ButtonOverValued", Nothing)
      Set m_ButtonPicture(2) = .ReadProperty("ButtonPressed", Nothing)
      Set m_ButtonPicture(10) = .ReadProperty("ButtonPressedGrayed", Nothing)
      Set m_ButtonPicture(6) = .ReadProperty("ButtonPressedValued", Nothing)
      m_ButtonRounding = .ReadProperty("ButtonRounding", 0)
      m_ButtonThemeType = .ReadProperty("ButtonThemeType", Windows)
      m_ButtonType = .ReadProperty("ButtonType", CommandButton)
      m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
      m_CaptionAlign = .ReadProperty("CaptionAlign", vbCenter)
      m_CaptionShadow = .ReadProperty("CaptionShadow", False)
      UserControl.Enabled = .ReadProperty("Enabled", True)
      m_FocusStyle = .ReadProperty("FocusStyle", Button)
      Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
      m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
      MouseIcon = .ReadProperty("MouseIcon", Nothing)
      Me.MousePointer = .ReadProperty("MousePointer", vbDefault)
      m_OptionButtonMultiSelect = .ReadProperty("OptionButtonMultiSelect", False)
      m_OverColor = .ReadProperty("OverColor", &HE0E0E0)
      Set m_Picture = .ReadProperty("Picture", Nothing)
      m_PictureAlign = .ReadProperty("PictureAlign", TopCenter)
      m_PictureSize = .ReadProperty("PictureSize", ps32x32)
      m_ShowFocusRect = .ReadProperty("ShowFocusRect", True)
      m_UseParentBackColor = .ReadProperty("UseParentBackColor", False)
      m_Value = .ReadProperty("Value", Grayed)
      ButtonState = BPS_NORMAL
      Me.Caption = m_Caption
      CheckButtonThemeType
      
      If m_ButtonType <> CommandButton Then m_OverColor = m_BackColor
      
      Call SetBackColor(m_BackColor)
   End With
   
   Call GetPictureSize
   Call RoundControl
   Call Refresh
   
   If Ambient.UserMode Then
      IsThemed = CheckIsThemed
      TrackUser32 = IsFunctionSupported("TrackMouseEvent", "User32")
      
      If Not TrackUser32 Then IsFunctionSupported "_TrackMouseEvent", "ComCtl32"
      
      With UserControl
         Call Subclass_Initialize(.hWnd)
         Call Subclass_AddMsg(.hWnd, WM_LBUTTONDBLCLK, MSG_BEFORE)
         Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_BEFORE)
         Call Subclass_AddMsg(.hWnd, WM_LBUTTONUP, MSG_BEFORE)
         Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE)
         Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE)
         Call Subclass_AddMsg(.hWnd, WM_TIMER)
         
         If IsThemedWindows Then Call Subclass_AddMsg(.hWnd, WM_THEMECHANGED)
      End With
   End If

End Sub

Private Sub UserControl_Resize()

Static blnBusy As Boolean

   If blnBusy Then Exit Sub
   
   blnBusy = True
   
   If (m_ButtonType <> CommandButton) And (Height / Screen.TwipsPerPixelY < 18) Then Height = Screen.TwipsPerPixelY * 18
   
   If Not CheckButtonThemeType Then
      If Width / Screen.TwipsPerPixelX < 7 Then Width = 7 * Screen.TwipsPerPixelX
      If Height / Screen.TwipsPerPixelY < 16 Then Height = 16 * Screen.TwipsPerPixelY
   End If
   
   With ButtonProperties.ButtonRect
      .Right = ScaleWidth
      .Bottom = ScaleHeight
   End With
   
   With ButtonProperties.FocusRect
      .Top = 5
      .Left = 5
      .Right = ScaleWidth - 5
      .Bottom = ScaleHeight - 5
   End With
   
   Call RoundControl
   Call Refresh
   
   blnBusy = False

End Sub

Private Sub UserControl_Show()

   Call CheckIsDefault

End Sub

Private Sub UserControl_Terminate()

   On Local Error GoTo ExitSub
   
   Call Subclass_Terminate
   
ExitSub:
   On Local Error GoTo 0
   Set m_Picture = Nothing
   Erase SubclassData, m_ButtonPicture

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColor", m_BackColor, vbButtonFace
      .WriteProperty "ButtonCorner", m_ButtonCorner, AllCorners
      .WriteProperty "ButtonDefaulted", m_ButtonPicture(12), Nothing
      .WriteProperty "ButtonDisabled", m_ButtonPicture(3), Nothing
      .WriteProperty "ButtonDisabledGrayed", m_ButtonPicture(11), Nothing
      .WriteProperty "ButtonDisabledValued", m_ButtonPicture(7), Nothing
      .WriteProperty "ButtonNormal", m_ButtonPicture(0), Nothing
      .WriteProperty "ButtonNormalGrayed", m_ButtonPicture(8), Nothing
      .WriteProperty "ButtonNormalValued", m_ButtonPicture(4), Nothing
      .WriteProperty "ButtonOver", m_ButtonPicture(1), Nothing
      .WriteProperty "ButtonOverGrayed", m_ButtonPicture(9), Nothing
      .WriteProperty "ButtonOverValued", m_ButtonPicture(5), Nothing
      .WriteProperty "ButtonPressed", m_ButtonPicture(2), Nothing
      .WriteProperty "ButtonPressedGrayed", m_ButtonPicture(10), Nothing
      .WriteProperty "ButtonPressedValued", m_ButtonPicture(6), Nothing
      .WriteProperty "ButtonRounding", m_ButtonRounding, 0
      .WriteProperty "ButtonThemeType", m_ButtonThemeType, Windows
      .WriteProperty "ButtonType", m_ButtonType, CommandButton
      .WriteProperty "Caption", m_Caption, Ambient.DisplayName
      .WriteProperty "CaptionAlign", m_CaptionAlign, vbCenter
      .WriteProperty "CaptionShadow", m_CaptionShadow, False
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "FocusStyle", m_FocusStyle, Button
      .WriteProperty "Font", UserControl.Font
      .WriteProperty "ForeColor", m_ForeColor, vbButtonText
      .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
      .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
      .WriteProperty "OptionButtonMultiSelect", m_OptionButtonMultiSelect, False
      .WriteProperty "OverColor", m_OverColor, &HE0E0E0
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "PictureAlign", m_PictureAlign, TopCenter
      .WriteProperty "PictureSize", m_PictureSize, ps32x32
      .WriteProperty "ShowFocusRect", m_ShowFocusRect, True
      .WriteProperty "UseParentBackColor", m_UseParentBackColor, False
      .WriteProperty "Value", m_Value, Grayed
   End With

End Sub
