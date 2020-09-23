Attribute VB_Name = "modMath"
'+  File Description:
'       modMath.bas - Custom Routines for Scientific Computing...
'
'   Product Name:
'       modMath.bas
'
'   Compatability:
'       Windows: 95, 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Inspired by the following On-Line Articles
'       URL: http://vertex42.com/ExcelTips/significant-figures.html
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
'       28May05 - Initial routine and test harness completed
'
'   Force Declarations
Option Explicit

Public Function RoundSig(dVal As Double, SigFigures As Integer) As Double
    Dim Exponent As Double

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

     '  Compute the significant digits of the passed value
    If IsNumeric(dVal) And IsNumeric(SigFigures) Then
        If SigFigures < 1 Then
            '   Return the value unchanged, since the Num of Sig Digits are
            '   negative values....
            RoundSig = dVal
        Else
            If dVal > 0 Then
                '   Compute the exponent as a Log10
                Exponent = Int(Log(Abs(dVal)) / Log(10#))
                '   Now force the VBA round to give the correct rounding
                RoundSig = Round(dVal, SigFigures - (1 + Exponent))
            Else
                '   We can't round Zero, so just return it...
                RoundSig = 0
            End If
        End If
    Else
        '   Return the value unchanged, as it is text...
        RoundSig = dVal
    End If

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    If Err.Number = 5 Then
        RoundSig = Round(dVal, 0)
    Else
        Err.Raise Err.Number, "modMath.RoundSig", Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Resume Func_ErrHandlerExit:
End Function
