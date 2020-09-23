VERSION 5.00
Begin VB.Form frmDialogExample 
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Set Path"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open Text"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open All"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmDialogExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdl As New clsDialog
Dim xFlags As DialogFlags

Private Sub Command1_Click()
  Dim sz As String 'will go to last folder if InitDir is blank
  sz = cdl.ShowOpen(Me.hWnd, "Sample Code, Open anything", , , xFlags)
  If Len(sz) = 0 Then Exit Sub
  MsgBox sz
End Sub

Private Sub Command2_Click()
  Dim sz As String
  'DefExt make sure it will be saved with whatever you typed at the end
  'Just make sure there is no filter
  '0 Means no flags, when using multiple flags, use or instead of +. or is faster
  sz = cdl.ShowSave(Me.hWnd, "Where to save", , , "pdf", 0)
  If Len(sz) = 0 Then Exit Sub
  MsgBox "Saving at " & sz
End Sub

Private Sub Command3_Click()
  Dim sz As String
  sz = cdl.ShowOpen(Me.hWnd, "Sample Code, Open Text", "c:\", "Text Files|*.doc;*.txt")
  If Len(sz) = 0 Then Exit Sub
  MsgBox sz
End Sub

Private Sub Command4_Click()
  MsgBox "Currently " & cdl.FileName
  cdl.FileName = InputBox("New Path ?")
  MsgBox "it is now " & cdl.FileName
End Sub
