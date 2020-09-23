VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CInfoTip Demo"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Custom font and colors."
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   810
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "About"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2385
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   885
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rightclick here!"
      Height          =   195
      Left            =   915
      TabIndex        =   0
      Top             =   315
      Width           =   1110
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents CTip As CInfoTip
Attribute CTip.VB_VarHelpID = -1
Private sText           As String
Private sTitle          As String

Private Sub CTip_Hide()
  Set CTip = Nothing
End Sub

Private Sub Form_Load()
  Set CTip = New CInfoTip
  Set Me.Icon = Nothing
  Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 3 - Me.Height / 2
  sText = "Hi ppl!" & vbCrLf & "As you can see in this demonstration, CInfoTip is very customizable." & vbCrLf & _
          "Smart use with a timer (API or VB) makes for an excellent replacement" & vbCrLf & _
          "for your regular tooltips." & vbCrLf & _
          "I developed this class for use as tooltips in my ActiveX control projects." & vbCrLf & vbCrLf & "Enjoy."
  sTitle = "CInfoTip demonstration (" & CTip.Version & ")"
  Set CTip = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not (CTip Is Nothing) Then CTip.Hide
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not (CTip Is Nothing) Then CTip.Hide
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Call ShowTip
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not (CTip Is Nothing) Then CTip.Hide
End Sub

Private Sub ShowTip()
  Set CTip = New CInfoTip
  CTip.ParenthWnd = Me.hWnd
  CTip.ShowTitle = True
  CTip.UseTimeOut = False
  CTip.DropShadow = False
  CTip.Title = sTitle
  CTip.Text = sText
  If Check1.Value = vbChecked Then
    CTip.BorderStyle = eitbs_Line
    CTip.UseSystemFont = False
    CTip.FontName = "arial"
    CTip.FontSize = 12
    CTip.BackColor = vbWhite
    CTip.BorderColor = vbGrayText
    CTip.Padding = 8
    CTip.TitleAlignment = eitta_Right
    'Randomize Now
    CTip.TitleColor = &HAF9A7A 'Rnd(&HBBBBBB - &H0) * &HBBBBBB
    CTip.DropShadow = True
  End If
  CTip.Show
End Sub

Private Sub Label2_Click()
  Set CTip = New CInfoTip
  CTip.ParenthWnd = Me.hWnd
  CTip.About
  Set CTip = Nothing
End Sub
