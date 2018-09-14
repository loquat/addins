VERSION 5.00
Begin VB.Form frmToolTip 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1485
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   1485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   570
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Friend Sub ShowMe(hWnd As Long, r As RECT, newTxt As String)
    If lastRect.Top <> r.Top Or lastRect.Right <> r.Right Or lastRect.Left <> r.Left Then
        Label1.Caption = newTxt
        Call MoveMe(r)
        lastRect = r
    End If
End Sub

Private Sub MoveMe(r As RECT)
    Debug.Print "MoveMe " & r.Top, r.Right
    
    Me.Visible = True
    
    Me.Top = (r.Top) * Screen.TwipsPerPixelY
    Me.Top = Me.Top + 15
    Me.Left = (r.Right) * Screen.TwipsPerPixelX
    
    Me.Width = Label1.Width + 2 * Screen.TwipsPerPixelX
    Me.Height = Label1.Height + 2 * Screen.TwipsPerPixelY
    
    Call SetTopMostWindow(Me.hWnd, True)
    
End Sub



