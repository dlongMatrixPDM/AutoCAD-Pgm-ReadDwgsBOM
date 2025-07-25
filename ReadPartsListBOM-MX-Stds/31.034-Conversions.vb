'Option Strict Off
'Option Explicit On
'Option Compare Text

'Imports System
'Imports System.Reflection

'Module Conversions31_062
'    Public ValidChars As String

'  sInput -> The feet and inches string to convert
'    Public Function FToD(ByVal sInput As String) As Double

'        Dim sTemp, foot, inch, Numerator, Denominator As String
'        Dim i As Long
'        Dim j As Long
'        Dim tmpValue As Double
'        Dim StepSize As Long

'        ValidChars = "0123456789 .'-/" & Chr(34) & Chr(34)

'        On Error GoTo Err_FToD
'        If sInput = "" Then
'            FToD = 0.0#
'            Exit Function
'        End If

'        Denominator = vbNullString
'        Numerator = vbNullString
'        inch = vbNullString
'        foot = vbNullString
'        sTemp = sInput
'        sTemp = Trim(sTemp)

'        For i = 1 To Len(sTemp)
'            If InStr(ValidChars, Mid(sTemp, i, 1)) = 0 Then
'                FToD = 0
'                GoTo Exit_FToD
'            End If
'        Next i

'        If InStr(sTemp, ".") > 0 Then
'            If InStr(sTemp, " ") Then
'                FToD = 0
'                GoTo Exit_FToD
'            ElseIf InStr(sTemp, "-") Then
'                FToD = 0
'                GoTo Exit_FToD
'            ElseIf InStr(sTemp, "/") Then
'                FToD = 0
'                GoTo Exit_FToD
'            End If

'            If InStr(sTemp, "'") > 0 Then
'                FToD = CDbl(Left(sTemp, InStr(sTemp, "'") - 1)) * 12.0#
'            Else
'                FToD = CDbl(sTemp)
'            End If
'            Exit Function
'        End If

'        If InStr(sTemp, "/") > 0 Then
'            Denominator = LStrip(Mid(sTemp, InStr(sTemp, "/") + 1, Len(sTemp)))
'            Numerator = RStrip(Left(sTemp, InStr(sTemp, "/") - 1))
'            StepSize = Len(Numerator) - 1

'            If StepSize <= 0 Then StepSize = 1
'            For i = InStr(sTemp, "/") - 1 To 0 Step StepSize \ -1
'                If Mid(sTemp, i - Len(Numerator) + 1, Len(Numerator)) = Numerator Then
'                    j = i - Len(Numerator)
'                    Exit For
'                End If
'            Next i

'            If j > 0 Then
'                sTemp = Left(sTemp, j)
'            Else
'                sTemp = ""
'            End If
'        End If

'        If InStr(sTemp, "'") > 0 Then
'            foot = Left(sTemp, InStr(sTemp, "'") - 1)
'            inch = Mid(sTemp, InStr(sTemp, "'") + 1, Len(sTemp))
'        Else
'            inch = sTemp
'        End If

'        'Remove leading and trialing non-numeric characters from the foot and inch values
'        If inch <> "" Then inch = RStrip(inch)
'        If foot <> "" Then foot = Strip(foot)

'        'generate the decimal value by looking at each component. if it is not empty and the is value to the temporaty value
'        If Numerator <> "" And Denominator <> "" Then
'            tmpValue = CDbl(Numerator) / CDbl(Denominator)
'        End If
'        If inch <> "" Then tmpValue = tmpValue + CDbl(inch)
'        If foot <> "" Then tmpValue = tmpValue + (CDbl(foot) * 12.0#)

'        FToD = tmpValue                         'Return the calculated value

'Exit_FToD:
'        Exit Function

'Err_FToD:
'        FToD = 0
'        Resume Exit_FToD
'    End Function

'Private Function Strip(ByVal sStr As String) As String
'    Dim i As Long
'    Dim sTemp As String
'    sTemp = vbNullString

'    If sStr <> "" Then
'        sStr = Trim(sStr)                   'Remove all leading and trailing spaces

'        For i = 1 To Len(sStr)
'            If IsNumeric(Mid(sStr, i, 1)) Then
'                sTemp = sTemp & Mid(sStr, i, 1)
'            End If
'        Next i
'    End If
'    Strip = sTemp
'End Function

'Private Function RStrip(ByVal sStr As String) As String
'    Dim i As Long
'    Dim StartPos As Long
'    Dim EndPos As Long
'    Dim sTemp As String

'    sTemp = vbNullString
'    If sStr <> "" Then
'        StartPos = -1
'        EndPos = -1
'        sTemp = Trim(sStr)
'        For i = Len(sStr) To 0 Step -1
'            If i > 0 Then
'                If IsNumeric(Mid(sStr, i, 1)) Then
'                    If StartPos = -1 Then StartPos = i
'                    If StartPos > -1 Then EndPos = i
'                Else
'                    If StartPos > -1 Then
'                        Exit For
'                    End If
'                End If
'            End If
'        Next i
'        sTemp = Mid(sStr, EndPos, StartPos - EndPos + 1)
'    End If

'    RStrip = sTemp

'End Function

'Private Function LStrip(ByVal sStr As String) As String
'    Dim i As Long
'    Dim StartPos As Long
'    Dim EndPos As Long
'    Dim sTemp As String

'    sTemp = vbNullString
'    If sStr <> "" Then
'        StartPos = -1
'        EndPos = -1
'        sTemp = Trim(sStr)
'        For i = 1 To Len(sStr)
'            If IsNumeric(Mid(sStr, i, 1)) Then
'                If StartPos = -1 Then StartPos = i
'                If StartPos > -1 Then EndPos = i
'            Else
'                If StartPos > -1 Then
'                    Exit For
'                End If
'            End If
'        Next i
'        sTemp = Mid(sStr, StartPos, EndPos - StartPos + 1)
'    End If

'    LStrip = sTemp
'End Function
'End Module
