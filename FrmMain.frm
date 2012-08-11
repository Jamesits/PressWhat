VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PressWhat"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2580
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3600
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   4635
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   1320
      Left            =   60
      TabIndex        =   4
      Top             =   300
      Width           =   3780
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press a key,and I'll tell what key you are pressing."
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
Text1.SetFocus
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "Another Presswhat is running!", , "Presswhat"
End
End If
Label2.Caption = "Usage:" & vbCrLf & "1.Press a key and you can see what you are pressing." & vbCrLf & "2.For technical reasons,some keys may not be displayed." _
 & vbCrLf & "e.g. Print Screen/System Request(SysRq) and Pause/Break." & vbCrLf & "3.Click on the text area to clear message."
Label1.Top = 0
Label1.Left = 0
Label1.Refresh
Label2.Top = Label1.Top + Label1.Height
Label2.Left = Label1.Left
Label2.Width = Label1.Width
Label2.Refresh
Text1.Left = Label1.Left
Text1.Top = Label2.Top + Label2.Height + 15
Text1.Width = Label1.Width
Me.Width = Text1.Width + 90
Me.Height = Text1.Top + Text1.Height + 420
Me.Show
End Sub

Private Sub Text1_GotFocus()
Text2.SetFocus
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If (65 <= KeyCode And KeyCode <= 90) Then
Text1.Text = "Pressed Letter " + UCase(Left(Chr(KeyCode), 1))
ElseIf (48 <= KeyCode And KeyCode <= 57) Then
Text1.Text = "Pressed Number " + UCase(Left(Chr(KeyCode), 1))
ElseIf (vbKeyNumpad0 <= KeyCode And KeyCode <= vbKeyNumpad9) Then
Text1.Text = "Pressed Number On Numpad " + CStr(KeyCode - vbKeyNumpad0)
ElseIf (vbKeyF1 <= KeyCode And KeyCode <= vbKeyF16) Then
Text1.Text = "Pressed Function Key F" + CStr(KeyCode - vbKeyF1 + 1)
Else
Dim a As String
a = ""
Select Case KeyCode
Case vbKeyLButton: a = "Mouse Left"
Case vbKeyRButton: a = "Mouse Right"
Case vbKeyCancel: a = "Cancel"
Case vbKeyMButton: a = "Mouse Middle"
Case vbKeyBack: a = "Backspace"
Case vbKeyTab: a = "Tab"
Case vbKeyClear: a = "Clear"
Case vbKeyReturn: a = "Enter"
Case vbKeyShift: a = "Shift"
Case vbKeyControl: a = "Ctrl"
Case vbKeyMenu: a = "Alt"
Case vbKeyPause: a = "Pause/Break"
Case vbKeyCapital: a = "CAPS LOCK"
Case vbKeyEscape: a = "ESC"
Case vbKeySpace: a = "Spacebar"
Case vbKeyPageUp: a = "Page Up"
Case vbKeyPageDown: a = "Page Down"
Case vbKeyEnd: a = "End"
Case vbKeyHome: a = "Home"
Case vbKeyLeft: a = "Left Arrow"
Case vbKeyUp: a = "Up Arrow"
Case vbKeyRight: a = "Right Arrow"
Case vbKeyDown: a = "Down Arrow"
Case vbKeySelect: a = "Select"
Case vbKeyPrint: a = "Print Screen/System Request(SysRq)"
Case vbKeyExecute: a = "Execute"
Case vbKeySnapshot: a = "Snapshot"
Case vbKeyInsert: a = "Insert"
Case vbKeyDelete:  a = "Delete"
Case vbKeyHelp:  a = "Help"
Case vbKeyNumlock:  a = "NUM LOCK"
Case vbKeyMultiply: a = "Multiply(*)"
Case vbKeyAdd: a = "Add(+)"
Case vbKeySeparator: a = "Enter/Separator"
Case vbKeySubtract: a = "Subtract(-)"
Case vbKeyDecimal: a = "Decimal Point(.)"
Case vbKeyDivide: a = "Divide(/)"
Case 91: a = "Windows Logo/Option(Left Side)"
Case 92: a = "Windows Logo/Option(Right Side)"
Case 220: a = "Backslash(\)"
Case 192: a = "Interpunct(`)"
Case 145: a = "Scroll Lock"
Case 93: a = "Menu"
Case 186: a = "Semicolon(;)"
Case 222: a = "Single quotation marks(')"
Case 219: a = "([)"
Case 221: a = "(])"
Case 187: a = "Dash(-)"
Case 189: a = "Equal(=)"
Case 188: a = "Comma(,)"
Case 190: a = "Period(.)"
Case 191: a = "Slash(/)"
Case Else: Text1.Text = "Key name not found in database,Keycode=" & CStr(KeyCode)
End Select
If a <> "" Then Text1.Text = "Pressed " & a
End If
AppActivate App.Title
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_GotFocus()
Text1.Text = "Pressed Tab"
Text2.SetFocus
End Sub
