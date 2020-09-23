VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInfoTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: Skip
'---------------------------------------------------------
'
' Original implementation by 'AVE'(?), no readme attached.
' Heavily modified by Morten Hansen (eriol@e-dsign.com).
'
' In the initial modification, I added settings for:
'   BackColor
'   BorderColor
'   ForeColor
'   Padding
'   BorderStyle
'   TabSize
'   DropShadow
'   Autohiding
'   Custom Font Properties:
'     Name
'     Size
'     Bold
'     Italic
'     Underline
'     StrikeThrough
'
' ...and did some general code 'cleanup' to make it suit my own
' coding style better.
'
'---------------------------------------------------------
'
' Changes:
'         - 09.01.2003
'           Added 'ShowTitle','Title','TitleAlignment' and
'           'TitleColor' properties, and 'About' method.
'
'
' Planned:
'         - Rounded corners
'         - Fade in/out     (w2k/xp)
'         - Transparency    (w2k/xp)
'         - Error handling
'         - Icon
'         - Gradient background
'         - Header styles
'
' Bugs:
'         - No known bugs at the moment. (1.0.1)
'
'---------------------------------------------------------
Option Explicit

' [versioninfo (major.minor.revision)]
Private Const sVersion                  As String = "1.0.1"

' 1.0.2
Private Const AW_HOR_POSITIVE = &H1
Private Const AW_HOR_NEGATIVE = &H2
Private Const AW_VER_POSITIVE = &H4
Private Const AW_VER_NEGATIVE = &H8
Private Const AW_CENTER = &H10
Private Const AW_HIDE = &H10000
Private Const AW_ACTIVATE = &H20000
Private Const AW_SLIDE = &H40000
Private Const AW_BLEND = &H80000
Private Declare Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
'\1.0.2

' [constants]
Private Const DT_TOP                    As Long = &H0
Private Const DT_LEFT                   As Long = &H0
Private Const DT_CENTER                 As Long = &H1
Private Const DT_RIGHT                  As Long = &H2
Private Const DT_CALCRECT               As Long = &H400
Private Const DT_EXPANDTABS             As Long = &H40
Private Const LF_FACESIZE               As Long = 32
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_SHOWWINDOW            As Long = &H40
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SPI_GETNONCLIENTMETRICS   As Long = 41
Private Const BDR_RAISEDOUTER           As Long = &H1
Private Const BF_LEFT                   As Long = &H1
Private Const BF_TOP                    As Long = &H2
Private Const BF_RIGHT                  As Long = &H4
Private Const BF_BOTTOM                 As Long = &H8
Private Const BF_RECT                   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const HWND_TOPMOST              As Long = -1
Private Const TRANSPARENT               As Long = 1
Private Const OPAQUE                    As Long = 2
Private Const WS_POPUP                  As Long = &H80000000
Private Const WS_EX_TRANSPARENT         As Long = &H20&
Private Const WS_EX_TOPMOST             As Long = &H8&
Private Const WM_USER                   As Long = &H400
Private Const IDI_ASTERISK              As Long = 32516&


' [types]
Private Type POINTAPI
  X                 As Long
  Y                 As Long
End Type

Private Type RECT
  Left              As Long
  Top               As Long
  Right             As Long
  Bottom            As Long
End Type

Private Type LOGFONT
  lfHeight          As Long
  lfWidth           As Long
  lfEscapement      As Long
  lfOrientation     As Long
  lfWeight          As Long
  lfItalic          As Byte
  lfUnderline       As Byte
  lfStrikeOut       As Byte
  lfCharSet         As Byte
  lfOutPrecision    As Byte
  lfClipPrecision   As Byte
  lfQuality         As Byte
  lfPitchAndFamily  As Byte
  lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
  cbSize            As Long
  iBorderWidth      As Long
  iScrollWidth      As Long
  iScrollHeight     As Long
  iTextWidth        As Long
  iTextHeight       As Long
  lfTextFont        As LOGFONT
  iSMTextWidth      As Long
  iSMTextHeight     As Long
  lfSMTextFont      As LOGFONT
  iMenuWidth        As Long
  iMenuHeight       As Long
  lfMenuFont        As LOGFONT
  lfStatusFont      As LOGFONT
  lfMessageFont     As LOGFONT
End Type

Private Type COLORREF
  Red               As Byte
  Green             As Byte
  Blue              As Byte
  Mode              As Byte
End Type

Private Type SIZEL
  cX                As Long
  cY                As Long
End Type

' [enums]
Public Enum EInfoTipBorderStyle
  eitbs_Beveled = 0
  eitbs_Line
End Enum

Public Enum EInfoTipTextAlignment
  eitta_Left = 0
  eitta_Center
  eitta_Right
End Enum

' [API declares]
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEL) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal m_ParenthWnd As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBR As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, rc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' [private value holders]
Private m_BackColor         As OLE_COLOR
Private m_ForeColor         As OLE_COLOR
Private m_BorderStyle       As EInfoTipBorderStyle
Private m_BorderColor       As OLE_COLOR
Private hWndTip             As Long
Private m_ParenthWnd        As Long
Private mText               As String
Private m_Padding           As Long
Private m_UseSystemFont     As Boolean
Private m_FontName          As String
Private m_FontSize          As Long
Private m_FontBold          As Boolean
Private m_FontItalic        As Boolean
Private m_FontUnderline     As Boolean
Private m_FontStrikethrough As Boolean
Private m_TabSize           As Long
Private m_DropShadow        As Boolean
Private m_TimeOut           As Long
Private m_UseTimeOut        As Boolean
Private m_TextAlignment     As EInfoTipTextAlignment
Private m_ShowTitle         As Boolean
Private m_Title             As String
Private m_TitleAlignment    As EInfoTipTextAlignment
Private m_TitleColor        As OLE_COLOR

' [public events]
Public Event Hide()

Public Property Get TitleColor() As OLE_COLOR
  TitleColor = m_TitleColor
End Property
Public Property Let TitleColor(ByVal Value As OLE_COLOR)
  m_TitleColor = Value
End Property


' [public functions]
Public Sub About()
  Debug.Assert m_ParenthWnd <> 0
  
  Dim lhIcon As Long
  lhIcon = LoadIcon(0&, IDI_ASTERISK)
  ShellAbout m_ParenthWnd, "CInfoTip", "CInfoTip v" & sVersion, lhIcon
  DestroyIcon lhIcon
End Sub

Public Sub Show(Optional sMessage As String = "", Optional sTitle As String = "")
  Debug.Assert m_ParenthWnd <> 0
  
  Dim PT As POINTAPI, RT As RECT, lDC As Long, hBR As Long
  Dim I As Long, sMsg As String, sTtl As String
  Dim X As Long, Y As Long, dX As Long, dY As Long
  Dim lPen As Long, lOldPen As Long, lFlags As Long
  
  sMsg = IIf(Len(sMessage) = 0, mText, sMessage)
  sMsg = Replace(sMsg, vbTab, Space(m_TabSize))
  sTtl = IIf(Len(sTitle) = 0, m_Title, sTitle)
  sTtl = Replace(sTtl, vbTab, Space(0))
  
  ' [create the window and give it a font]
  If hWndTip = 0 Then hWndTip = CreateWindowEx(WS_EX_TOPMOST Or WS_EX_TRANSPARENT, "STATIC", vbNullString, WS_POPUP, 0, 0, 0, 0, m_ParenthWnd, 0, App.hInstance, 0)
  lDC = GetWindowDC(hWndTip)
  SetFont lDC
  
  ' [calculate the windows size]
  DrawText lDC, sMsg, Len(sMsg), RT, DT_TOP Or DT_LEFT Or DT_CALCRECT Or DT_EXPANDTABS
  With RT
    .Bottom = (.Bottom + m_Padding) + IIf(m_ShowTitle, GetStringHeight(lDC, sTtl) + m_Padding, 0)
    .Right = .Right + m_Padding * 2
  End With
  dX = RT.Right - RT.Left
  dY = RT.Bottom - RT.Top
    
  ' [get current mouse position]
  GetCursorPos PT
  X = PT.X
  Y = PT.Y + 22
    
  ' [check window position versus screen size]
  If X + dX > Screen.Width / Screen.TwipsPerPixelX Then X = Screen.Width / Screen.TwipsPerPixelX - dX - m_Padding
  If Y + dY > Screen.Height / Screen.TwipsPerPixelY Then Y = PT.Y - m_Padding - dY
  
  ' [show the window]
  SetWindowPos hWndTip, HWND_TOPMOST, X, Y, IIf(m_DropShadow, dX + 4, dX), IIf(m_DropShadow, dY + 4, dY), SWP_NOACTIVATE Or SWP_SHOWWINDOW
  'AnimateWindow hWndTip, 100, AW_SLIDE Or AW_HOR_POSITIVE Or AW_VER_POSITIVE
  DoEvents
  
  ' [fill the window with our backcolor]
  hBR = CreateSolidBrush(WinColor(m_BackColor))
  FillRect lDC, RT, hBR
  DeleteObject hBR

  ' [draw the border in the selected style]
  Select Case m_BorderStyle
    Case 0
      DrawEdge lDC, RT, BDR_RAISEDOUTER, BF_RECT
    Case 1
      lPen = CreatePen(0, 1, WinColor(m_BorderColor))
      lOldPen = SelectObject(lDC, lPen)

      MoveToEx lDC, RT.Left, RT.Top, PT
      LineTo lDC, RT.Right - 1, RT.Top
      LineTo lDC, RT.Right - 1, RT.Bottom - 1
      LineTo lDC, RT.Left, RT.Bottom - 1
      LineTo lDC, RT.Left, RT.Top
      
      SelectObject lDC, lOldPen
      DeleteObject lPen: DeleteObject lOldPen
  End Select

  ' [adjust text rectangle for padding]
  InflateRect RT, -m_Padding, -m_Padding / 2
  
  ' [set title text color]
  SetBkMode lDC, TRANSPARENT
  SetTextColor lDC, WinColor(m_TitleColor)
  
  ' [paint the message text]
  If m_ShowTitle Then
    Select Case m_TitleAlignment
      Case 0: lFlags = DT_LEFT
      Case 1: lFlags = DT_CENTER
      Case 2: lFlags = DT_RIGHT
    End Select
    SetFont lDC, True
    DrawText lDC, sTtl, Len(sTtl), RT, DT_TOP Or lFlags Or DT_EXPANDTABS
  End If
  
  ' [set message text color]
  SetBkMode lDC, TRANSPARENT
  SetTextColor lDC, WinColor(m_ForeColor)
  
  ' [paint the message text]
  SetFont lDC
  If m_ShowTitle Then
    lPen = CreatePen(0, 1, WinColor(m_BorderColor))
    lOldPen = SelectObject(lDC, lPen)
    RT.Top = RT.Top + GetStringHeight(lDC, sTtl) + (m_Padding / 2)
    MoveToEx lDC, RT.Left, RT.Top, PT
    LineTo lDC, RT.Right, RT.Top
    RT.Top = RT.Top + (m_Padding / 2)
    SelectObject lDC, lOldPen
    DeleteObject lPen: DeleteObject lOldPen
  End If
  Select Case m_TextAlignment
    Case 0: lFlags = DT_LEFT
    Case 1: lFlags = DT_CENTER
    Case 2: lFlags = DT_RIGHT
  End Select
  DrawText lDC, sMsg, Len(sMsg), RT, DT_TOP Or lFlags Or DT_EXPANDTABS
  
  ' [draw window dropshadow]
  If m_DropShadow Then
    GetWindowRect hWndTip, RT
    DrawDropShadow hWndTip, lDC, RT.Left, RT.Top, RT.Right - RT.Left, RT.Bottom - RT.Top
  End If
  
  ' [set the timer if autohide is selected]
  If m_UseTimeOut Then
    Set MInfoTip.CTipReference = Me
    SetTimer hWndTip, 0, m_TimeOut, AddressOf tipTimerProc
  End If
End Sub

Public Sub Hide()
  If hWndTip <> 0 Then DestroyWindow hWndTip
  hWndTip = 0
  RaiseEvent Hide
End Sub

' [public properties]
Public Property Get Version() As String
  Version = sVersion
End Property

Public Property Get TitleAlignment() As EInfoTipTextAlignment
  TitleAlignment = m_TitleAlignment
End Property
Public Property Let TitleAlignment(ByVal Value As EInfoTipTextAlignment)
  m_TitleAlignment = Value
End Property

Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Let Title(ByVal Value As String)
  m_Title = Value
End Property

Public Property Get ShowTitle() As Boolean
  ShowTitle = m_ShowTitle
End Property
Public Property Let ShowTitle(ByVal Value As Boolean)
  m_ShowTitle = Value
End Property

Public Property Get TextAlignment() As EInfoTipTextAlignment
  TextAlignment = m_TextAlignment
End Property
Public Property Let TextAlignment(ByVal Value As EInfoTipTextAlignment)
  m_TextAlignment = Value
End Property

Public Property Get UseTimeOut() As Boolean
  UseTimeOut = m_UseTimeOut
End Property
Public Property Let UseTimeOut(ByVal Value As Boolean)
  m_UseTimeOut = Value
End Property

Public Property Get TimeOut() As Long
  TimeOut = m_TimeOut
End Property
Public Property Let TimeOut(ByVal Value As Long)
  m_TimeOut = Value
End Property

Public Property Get hWnd() As Long
  hWnd = hWndTip
End Property

Public Property Get DropShadow() As Boolean
  DropShadow = m_DropShadow
End Property
Public Property Let DropShadow(ByVal Value As Boolean)
  m_DropShadow = Value
End Property

Public Property Get TabSize() As Long
  TabSize = m_TabSize
End Property
Public Property Let TabSize(ByVal Value As Long)
  m_TabSize = Value
End Property

Public Property Get FontStrikethrough() As Boolean
  FontStrikethrough = m_FontStrikethrough
End Property
Public Property Let FontStrikethrough(ByVal Value As Boolean)
  m_FontStrikethrough = Value
End Property

Public Property Get FontUnderline() As Boolean
  FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(ByVal Value As Boolean)
  m_FontUnderline = Value
End Property

Public Property Get FontItalic() As Boolean
  FontItalic = m_FontItalic
End Property
Public Property Let FontItalic(ByVal Value As Boolean)
  m_FontItalic = Value
End Property

Public Property Get FontBold() As Boolean
  FontBold = m_FontBold
End Property
Public Property Let FontBold(ByVal Value As Boolean)
  m_FontBold = Value
End Property

Public Property Get FontSize() As Long
  FontSize = m_FontSize
End Property
Public Property Let FontSize(ByVal Value As Long)
  m_FontSize = Value
End Property

Public Property Get FontName() As String
  FontName = m_FontName
End Property
Public Property Let FontName(ByVal Value As String)
  m_FontName = Value
End Property

Public Property Get UseSystemFont() As Boolean
  UseSystemFont = m_UseSystemFont
End Property
Public Property Let UseSystemFont(ByVal Value As Boolean)
  m_UseSystemFont = Value
End Property

Public Property Get Padding() As Long
  Padding = m_Padding
End Property
Public Property Let Padding(ByVal Value As Long)
  m_Padding = Value
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal Value As OLE_COLOR)
  m_BorderColor = Value
End Property

Public Property Get BorderStyle() As EInfoTipBorderStyle
  BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal Value As EInfoTipBorderStyle)
  m_BorderStyle = Value
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
  m_ForeColor = Value
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal Value As OLE_COLOR)
  m_BackColor = Value
End Property

Property Let Text(txt As String)
  mText = txt
End Property
Property Get Text() As String
  Text = mText
End Property

Property Let ParenthWnd(ByVal Value As Long)
  m_ParenthWnd = Value
End Property

' [class instance handling]
Private Sub Class_Initialize()
  m_BackColor = vbInfoBackground
  m_ForeColor = vbInfoText
  m_BorderStyle = 0
  m_BorderColor = vb3DDKShadow
  m_Padding = 5
  m_UseSystemFont = True
  m_FontName = "Arial"
  m_FontSize = 8
  m_TabSize = 8
  m_DropShadow = True
  m_TimeOut = 2500
  m_UseTimeOut = True
  m_TextAlignment = 0
  m_ShowTitle = False
  m_TitleAlignment = 0
  m_TitleColor = vbInfoText
End Sub

Private Sub Class_Terminate()
  Call Me.Hide
End Sub

' [private helper functions]
Private Function WinColor(lVBColor As Long) As Long
  Dim SysClr As COLORREF
  CopyMemory SysClr, lVBColor, Len(SysClr)
  If SysClr.Mode = &H80 Then WinColor = GetSysColor(SysClr.Red) Else WinColor = lVBColor
End Function

Private Sub DrawDropShadow(ByVal hWnd As Long, ByVal hDC As Long, ByVal xOrg As Long, ByVal yOrg As Long, ByVal winW As Long, ByVal winH As Long)
  Dim hDcDsk As Long, XO2 As Long, YO2 As Long
  Dim X As Long, Y As Long, C As Long
  XO2 = xOrg + winW: YO2 = yOrg + winH
  hDcDsk = GetWindowDC(GetDesktopWindow())

  For X = 1 To 4
    For Y = 0 To 7 '3
      C = GetPixel(hDcDsk, XO2 - X, yOrg + Y)
      SetPixelV hDC, winW - X, Y, IIf((Y >= 4), pMask(3 * X * (Y - 3), C), C)
    Next
    For Y = 8 To winH - 5
      C = GetPixel(hDcDsk, XO2 - X, yOrg + Y)
      SetPixelV hDC, winW - X, Y, pMask(15 * X, C)
    Next
    For Y = winH - 4 To winH - 1
      C = GetPixel(hDcDsk, XO2 - X, yOrg + Y)
      SetPixelV hDC, winW - X, Y, pMask(3 * X * -(Y - winH), C)
    Next
  Next

  For Y = 1 To 4
    For X = 0 To 7 '3
      C = GetPixel(hDcDsk, xOrg + X, YO2 - Y)
      SetPixelV hDC, X, winH - Y, IIf((X > 3), pMask(3 * (X - 3) * Y, C), C)
    Next
    For X = 8 To winW - 5
      C = GetPixel(hDcDsk, xOrg + X, YO2 - Y)
      SetPixelV hDC, X, winH - Y, pMask(15 * Y, C)
    Next
  Next

  ReleaseDC GetDesktopWindow, hDcDsk
End Sub

Private Function pMask(ByVal lScale As Long, ByVal lColor As Long) As Long
  Dim R As Long, G As Long, B As Long

  R = ClrComp(lColor, vbRed, lScale)
  G = ClrComp(lColor, vbGreen, lScale)
  B = ClrComp(lColor, vbBlue, lScale)
  
  pMask = RGB(R, G, B)
End Function

Private Function ClrComp(Color As Long, Component As Long, ByVal lScale As Long) As Long 'Byte
  Dim tCREF As COLORREF
  CopyMemory tCREF, Color, LenB(tCREF)
  Select Case Component
    Case vbRed:   ClrComp = tCREF.Red - Int(tCREF.Red * lScale / 255)
    Case vbGreen: ClrComp = tCREF.Green - Int(tCREF.Green * lScale / 255)
    Case vbBlue:  ClrComp = tCREF.Blue - Int(tCREF.Blue * lScale / 255)
  End Select
End Function

Private Function SetFont(ByVal lhDC As Long, Optional ByVal ForceBold As Boolean = False)
  Dim FontName As String, fntTip As New StdFont, fnt As IFont, I As Long, NCM As NONCLIENTMETRICS
' [getting system tooltip font information]
  NCM.cbSize = Len(NCM)
  SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(NCM), NCM, 0
  
  ' [creating the font]
  With NCM.lfStatusFont
    For I = 1 To LF_FACESIZE
      If .lfFaceName(I) = 0 Then
        Exit For
      End If
      FontName = FontName & Chr$(.lfFaceName(I))
    Next
    fntTip.Charset = .lfCharSet
    fntTip.Name = IIf(m_UseSystemFont, FontName, m_FontName)
    fntTip.Size = IIf(m_UseSystemFont, -.lfHeight, m_FontSize) * 72 / GetDeviceCaps(GetWindowDC(0), LOGPIXELSY)
    fntTip.Italic = IIf(m_UseSystemFont, .lfItalic, m_FontItalic)
    fntTip.Strikethrough = IIf(m_UseSystemFont, .lfStrikeOut, m_FontStrikethrough)
    fntTip.Underline = IIf(m_UseSystemFont, .lfUnderline, m_FontUnderline)
    fntTip.Weight = IIf(m_UseSystemFont, .lfWeight, IIf(m_FontBold, 900, 300))
    If ForceBold Then fntTip.Weight = 900
  End With
  Set fnt = fntTip
  SelectObject lhDC, fnt.hFont
End Function

Private Function GetStringWidth(ByVal lhDC As Long, ByVal sString As String) As Long
  Dim txS As SIZEL
  GetTextExtentPoint32 lhDC, sString, Len(sString), txS
  GetStringWidth = txS.cX
End Function
Private Function GetStringHeight(ByVal lhDC As Long, ByVal sString As String) As Long
  Dim txS As SIZEL
  GetTextExtentPoint32 lhDC, sString, Len(sString), txS
  GetStringHeight = txS.cY
End Function

Public Function GetAddress(ByVal lngAddr As Long) As Long
  ''Used with AddressOf to return the address in memory of a procedure.
  GetAddress = lngAddr&
End Function
