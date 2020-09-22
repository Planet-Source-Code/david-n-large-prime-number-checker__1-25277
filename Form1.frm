VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prime Number Checker"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAbout 
      Caption         =   "?"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtTime 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox lRe 
      Height          =   2595
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   1920
      MaxLength       =   28
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton CmdSlow 
      Caption         =   "Start Checking Number"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Results:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Processing Time (secs):"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Number to Check:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'9999999999999999999999999999 = largest number it can do....
Public Num1, Num2 As Variant
Public CurSec As Integer
Public PCancel As Boolean 'the P is so i can stop the processing in Form_Unload

Private Sub CmdAbout_Click()
    'my little about box
    MsgBox "Made 100% By: David N", vbInformation + vbOKOnly, "About Prime Number Checker"
End Sub

Private Sub CmdCancel_Click()
    PCancel = True 'cancel processing
End Sub

Private Sub CmdSlow_Click()
    On Error GoTo Er
    If Len(txtNumber.Text) = 0 Then Exit Sub
    Dim Num, Num3 As Variant
    Num = CDec(txtNumber.Text) 'CDec(Expression) allows you to deal with LARGE numbers
    CmdCancel.Enabled = True 'Enable cancel button
    CmdSlow.Caption = "Thinking..." 'Set current button caption
    CmdSlow.Enabled = False 'disable current button (prevent future clicks to same button
    txtNumber.Enabled = False 'disable the number textbox (just cause!)
    CmdAbout.Enabled = False 'the about box stop processing so we can't have that button enabled!
    lRe.Clear 'clear results listbox
    CurSec = 0 'number of elapsed seconds = 0
    PCancel = False 'reset cancel state
    Dim Found As Boolean
    PBar.Max = Sqr(Num) 'set progressbar max value
    Timer1.Enabled = True 'start the timer
    For Num1 = CDec(2) To CDec(Sqr(Num))
        Num2 = CDec(Num / Num1) 'Divide
        Num3 = Int(Num2) 'Take off remainder
        If Num3 * Num1 = Num Then 'if remainderless_number*for_next_number = number_to_check
            'if the first number is less than the second number add it to the list
            If Num1 <= Num3 Then lRe.AddItem Num1 & " * " & Num3 & " = " & Num1 * Num3
            Found = True 'we found at least one diviser
        End If
        PBar.Value = Num1 'give user something to look @ while waiting!
        DoEvents 'allow the computer to think while doing this processing...
        If PCancel = True Then Exit For 'Cancel = true when user clicks cancel button
    Next
    Timer1.Enabled = False 'stop timer
    CmdCancel.Enabled = False 'disable cancel button
    CmdSlow.Enabled = True 'enable current button
    CmdAbout.Enabled = True 'enable about button
    txtNumber.Enabled = True 'enable number textbox
    If PCancel = True Then
        CmdSlow.Caption = "Canceled"
    ElseIf Found = True Then
        CmdSlow.Caption = "Not Prime"
    ElseIf Found = False Then
        CmdSlow.Caption = "Prime"
    End If
    Exit Sub
Er:
    If Err.Number = 13 Then 'Type mismatch from having text in the number text box
        MsgBox "The number/data you typed in can't be checked."
    Else 'who knows what kind of error happened!
        MsgBox Err.Number & ":" & Err.Description, vbInformation + vbOKOnly, "Error:"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PCancel = True
End Sub

Private Sub Timer1_Timer()
    CurSec = CurSec + 1
    txtTime.Enabled = True
    txtTime.Text = CurSec & " - " & Num1 & " * " & Num2
    txtTime.Enabled = False
End Sub

Private Sub txtNumber_Change()
    CmdSlow.Caption = "Start Checking Number" 'change buttons caption back to normal
End Sub
