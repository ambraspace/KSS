VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Štampanje"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Štampaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtKopija 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   585
      Width           =   495
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Broj kopija:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bOKPressed As Boolean

Public Property Get OKPressed() As Boolean
  OKPressed = m_bOKPressed
End Property

Public Property Let OKPressed(bValue As Boolean)
  m_bOKPressed = bValue
End Property

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  If Val(Me.txtKopija) < 1 Then Exit Sub
  Me.OKPressed = True
  Me.Hide
End Sub

Private Sub Form_Load()
  Dim c As Printer, i As Integer
  For Each c In Printers
    Me.cboPrinter.AddItem c.DeviceName
    If c.DeviceName = Printer.DeviceName Then i = Me.cboPrinter.ListCount - 1
  Next
  Me.cboPrinter.ListIndex = i
End Sub

Public Sub Display()
  Me.OKPressed = False
  Me.Show 1, frmMain
End Sub

Private Sub txtKopija_GotFocus()
  Me.txtKopija.SelStart = 0
  Me.txtKopija.SelLength = Len(Me.txtKopija)
End Sub
