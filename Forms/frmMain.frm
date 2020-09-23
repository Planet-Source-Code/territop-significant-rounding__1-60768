VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Round Significant Example"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbSigValues 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "Ok"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "Reset"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtInitValue 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Enter a Number"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   4080
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      Index           =   0
      X1              =   120
      X2              =   4080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblRoundSig 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblVBARound 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblDescription 
      Caption         =   "Sig. Digits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label lblDescription 
      Caption         =   "RoundSig:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label lblDescription 
      Caption         =   "VBA Round:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblDescription 
      Caption         =   "Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+  File Description:
'       frmMain - Test harness for the Custom Routines Module
'
'   Product Name:
'       frmMain.frm
'
'   Compatability:
'       Windows: 95, 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'
'   Legal Copyright & Trademarks:
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and may be used
'       or distributed if the above Copyright and Trademark statments are
'       retained.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       27May05 - Initial routine and test harness completed
'
'   Force Declarations
Option Explicit

Private Sub cmdSeries_Click(Index As Integer)
    With Me
        Select Case Index
            Case 0
                '   Is it a valid number?
                If IsNumeric(.txtInitValue.Text) Then
                    .lblVBARound.Caption = Round(CDbl(.txtInitValue.Text), .cbSigValues.List(.cbSigValues.ListIndex))
                    .lblRoundSig.Caption = RoundSig(CDbl(.txtInitValue.Text), .cbSigValues.List(.cbSigValues.ListIndex))
                Else
                '   Opps, its not, so tell the user....
                    MsgBox "Please enter a valid number.", vbExclamation, "Round Significant Example"
                    '   Set the focus, as the Msgbox causes LostFocus
                    .txtInitValue.SetFocus
                    '   Now reset everything
                    .txtInitValue.Text = "Enter a Number"
                    Call SelectText(.txtInitValue)
                End If
            Case 1
                If MsgBox("Close the Application?", vbQuestion + vbYesNo, "Round Significant Example") = vbYes Then
                    Unload Me
                End If
            Case 2
                '   Now reset everything
                .txtInitValue.SetFocus
                .txtInitValue.Text = "Enter a Number"
                Call SelectText(.txtInitValue)
                .lblVBARound.Caption = ""
                .lblRoundSig.Caption = ""
        End Select
    End With
End Sub

Private Sub Form_Load()
    Dim i       As Long
    With Me
        With .cbSigValues
            '   Fill the combobox with significant digits
            For i = 1 To 20
                .AddItem i
            Next i
            .ListIndex = 0
        End With
        '   Select the text to change...
        Call SelectText(.txtInitValue)
    End With
End Sub

Private Sub txtInitValue_Click()
    With Me
        '   Select the text automatically
        Call SelectText(.txtInitValue)
    End With
End Sub

Private Sub SelectText(TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
