VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email entity converter"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "email.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCC 
      Caption         =   "Copy to clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Picture         =   "email.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "email.frx":5C24
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "name@email.com"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Email address to convert"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCC_Click()
Clipboard.Clear
Clipboard.SetText txtOutput.Text

End Sub

Private Sub cmdConvert_Click()
txtOutput.Text = ""

If Len(txtEmail.Text) > 255 Then
MsgBox "Email address is too big.", vbCritical, "Error"
Exit Sub
End If

If Len(txtEmail.Text) = 0 Then
MsgBox "Enter an email address.", vbCritical, "Error"
Exit Sub
End If

Dim x%
Dim CurrentChar$

For x = 0 To Len(txtEmail.Text) - 1
txtEmail.SelStart = x
txtEmail.SelLength = 1
CurrentChar = Asc(txtEmail.SelText)
If Len(CurrentChar) = 2 Then
txtOutput.Text = txtOutput.Text & "&#0" & Asc(txtEmail.SelText) & ";"
Else
txtOutput.Text = txtOutput.Text & "&#" & Asc(txtEmail.SelText) & ";"
End If

Next x


End Sub

