Option Strict Off
Option Explicit On
Option Compare Text

Imports System
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Module Formatting31_078
    Dim WorkShtName, PriPrg, ErrNo, ErrMsg, ErrSource, ErrDll, ErrLastLineX, PrgName As String
    Dim ErrException As System.Exception
    Public ExcelApp As Object
    Public GetShipMk, GetNextShipMk, GetShipMkYesNo As String

    Function HighlightLine(ByRef LineNo As Short, ByRef Highlight As String, ByRef BOMSheet As String) As Object
        Dim xlColorIndexNone As Object
        Dim Range As Object
        Dim RevColor, DescColor As String
        Dim WorkBooks As Workbooks
        PrgName = "HighlightLine"

        On Error GoTo Err_HighlightLine

        WorkBooks = ExcelApp.Workbooks
        NewBulkBOM = WorkBooks.Application.Worksheets(BOMSheet)
        NewBulkBOM.Activate()
        With NewBulkBOM
            With .Range("A" & LineNo & ":L" & LineNo).Interior
                Select Case Highlight
                    Case "Y"                        'color yellow for revised
                        .ColorIndex = 6
                    Case "G"                        'color green for new
                        .ColorIndex = 4
                    Case "R"                        'color red for deleted
                        .ColorIndex = 3
                    Case "N"                        'no color for unchanged cells
                        .ColorIndex = -4142
                    Case "X"
                        GoTo BOMNumOnly
                End Select
            End With
        End With

        GoTo EndFunction
BOMNumOnly:
        With NewBulkBOM
            With .Range("C" & LineNo & ":C" & LineNo).Interior
                .ColorIndex = 6
            End With
        End With

EndFunction:

        With NewBulkBOM
            RevColor = .Range("C" & LineNo).Interior.ColorIndex
            DescColor = .Range("G" & LineNo).Interior.ColorIndex

            If LineNo = 5 Then
                .Range("R" & (LineNo - 1)).Value = "Color Code No"
                .Range("R" & (LineNo - 1)).Font.Bold = True
                .Range("R" & (LineNo - 1)).Font.Size = 12
                .Range("R" & (LineNo - 1)).Orientation = 90
            End If

            If RevColor = DescColor Then
FixColor:
                Select Case RevColor
                    Case 4
                        .Range("R" & LineNo).Value = 1           'Green                        
                    Case 6
                        .Range("R" & LineNo).Value = 2           'Yellow
                    Case 3
                        .Range("R" & LineNo).Value = 3           'Red
                    Case 7
                        .Range("R" & LineNo).Value = 4           'Purple
                    Case 8
                        .Range("R" & LineNo).Value = 5           'Cyan
                    Case 45
                        .Range("R" & LineNo).Value = 6           'Color Orange
                    Case -4142
                        .Range("R" & LineNo).Value = 8           'No Color, No Change
                    Case Else
                        MsgBox("Error has been found in Function ColorToNumber on Program BulkBOM.")
                End Select
            Else
                If DescColor = -4142 And RevColor = 6 Then
                    .Range("R" & LineNo).Value = 7
                Else
                    If DescColor = -4142 Then
                        GoTo FixColor
                    Else
                        If DescColor = 7 And RevColor = 6 Then
                            .Range("R" & LineNo).Value = 7
                        Else
                            If DescColor = 8 And RevColor = 6 Then
                                .Range("R" & LineNo).Value = 7
                            Else
                                If DescColor = 45 And RevColor = 6 Then
                                    .Range("R" & LineNo).Value = 7
                                Else
                                    .Range("R" & LineNo).Value = RevColor
                                    .Range("S" & LineNo).Value = DescColor
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With

Err_HighlightLine:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = "20" And ErrMsg = "Resume without error." Then
                Exit Function
            End If

            If ErrNo = "91" And ErrMsg = "" Then
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            ReadDwgs.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function FormatLine(ByRef LineNo As Object, ByRef FileToOpen As String, Optional ByRef MultiLineMatl As Boolean = False) As Object
        Dim xlInsideHorizontal As Object, xlLeft As Object, xlInsideVertical As Object, xlEdgeRight As Object
        Dim xlEdgeBottom As Object, xlEdgeTop As Object, xlAutomatic As Object, xlThin As Object, xlEdgeLeft As Object
        Dim xlContinuous As Object, xlDiagonalUp As Object, xlDiagonalDown As Object, xlNone As Object, xlCenter As Object
        Dim Range As Object, xlDown As Object, Rows As Object, ExcelApp As Object
        Dim CallPos, ExceptionPos As Integer
        Dim WorkBooks As Workbooks
        Dim BOMWrkSht As Worksheet
        Dim WorkSht As Worksheet
        PrgName = "FormatLine"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_FormatLine

        WorkBooks = ExcelApp.Workbooks

        If FileToOpen = "Purchase BOM" Then
            WorkShtName = "Other BOM"
        Else
            WorkShtName = FileToOpen
        End If

        BOMWrkSht = WorkBooks.Application.Worksheets(WorkShtName)

        With BOMWrkSht
            .Rows(LineNo + 2 & ":" & LineNo + 2).Insert()
            .Rows(LineNo & ":" & LineNo).RowHeight = 18

            If LineNo = 5 Then
                .Range("W" & (LineNo - 1)).Value = "PROCUREMENT"
                .Range("W" & (LineNo - 1)).Font.Bold = True
                .Range("W" & (LineNo - 1)).Font.Size = 12

                With .Range("W" & (LineNo - 1) & ":W" & (LineNo - 1))
                    With .Borders.Item(XlBordersIndex.xlEdgeTop)
                        .LineStyle = XlLineStyle.xlContinuous
                        .Weight = XlBorderWeight.xlThin
                        .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    End With

                    With .Borders.Item(XlBordersIndex.xlEdgeRight)
                        .LineStyle = XlLineStyle.xlContinuous
                        .Weight = XlBorderWeight.xlThin
                        .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                    End With
                End With
            End If

            With .Range("A" & LineNo & ":W" & LineNo)                           'With .Range("A" & LineNo & ":L" & LineNo)      '-------DJL-10-11-2023
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Name = "Arial"
                .Font.FontStyle = "Regular"
                .Font.Size = 9
                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
            End With

            With .Range("F" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
            End With

            With .Range("C" & LineNo)
                .NumberFormat = "@"
            End With

            With .Range("D" & LineNo)
                .NumberFormat = "@"
            End With

            With .Range("M" & LineNo)
                .NumberFormat = "@"
            End With

            With .Range("N" & LineNo)
                .NumberFormat = "General"
            End With

            With .Range("O" & LineNo)
                .NumberFormat = "General"
            End With

            With .Range("P" & LineNo)
                .NumberFormat = "General"
            End With

            With .Range("I" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Name = "Arial"
                .Font.Size = 7
            End With
        End With

Err_FormatLine:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            ReadDwgs.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function FormatLine3(ByRef LineNo As Object, ByRef FileToOpen As String, Optional ByRef MultiLineMatl As Boolean = False) As Object
        Dim xlInsideHorizontal As Object, xlLeft As Object, xlInsideVertical As Object, xlEdgeRight As Object
        Dim xlEdgeBottom As Object, xlEdgeTop As Object, xlAutomatic As Object, xlThin As Object, xlEdgeLeft As Object
        Dim xlContinuous As Object, xlDiagonalUp As Object, xlDiagonalDown As Object, xlNone As Object, xlCenter As Object
        Dim Range As Object, xlDown As Object, Rows As Object, ExcelApp As Object
        Dim CallPos, ExceptionPos As Integer
        Dim WorkBooks As Workbooks
        Dim BOMWrkSht As Worksheet
        Dim WorkSht As Worksheet
        PrgName = "FormatLine3"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Function
            End If
        End If

        On Error GoTo Err_FormatLine3

        WorkBooks = ExcelApp.Workbooks

        If FileToOpen = "Purchase BOM" Then
            WorkShtName = "Other BOM"
        Else
            WorkShtName = FileToOpen
        End If

        BOMWrkSht = WorkBooks.Application.Worksheets(WorkShtName)

        With BOMWrkSht
            .Rows(LineNo & ":" & LineNo).Insert()
            .Rows(LineNo & ":" & LineNo).RowHeight = 18

            With .Range("A" & LineNo & ":W" & LineNo)                           'With .Range("A" & LineNo & ":L" & LineNo)      '-------DJL-10-11-2023
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .Font.Name = "Arial"
                .Font.FontStyle = "Regular"
                .Font.Size = 9
                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
            End With

            With .Range("F" & LineNo)
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter

                With .Borders.Item(XlBordersIndex.xlDiagonalDown)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlDiagonalUp)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeLeft)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeTop)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeBottom)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlEdgeRight)
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = XlBorderWeight.xlThin
                    .ColorIndex = XlColorIndex.xlColorIndexAutomatic
                End With
                With .Borders.Item(XlBordersIndex.xlInsideVertical)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
                With .Borders.Item(XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = XlLineStyle.xlLineStyleNone
                End With
            End With

            If MultiLineMatl = True Then
                With .Range("I" & LineNo)
                    .HorizontalAlignment = XlHAlign.xlHAlignCenter
                    .VerticalAlignment = XlVAlign.xlVAlignCenter
                    .Font.Size = 7
                End With
            End If
        End With

Err_FormatLine3:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            ReadDwgs.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Function UpdateShpMarksArray(ByRef BOMList As Object) As Object
        '-------DJL-10-11-2023
        '-------Do it all in an array
        'This program updates the ShipMark when blank from PrevShipMark
        '-------DJL-06-09-2025              '----------------------------------For some reason they are out of order again after I saw them in order.
        '-------Push to Manway then read manway.BOMListSort
        Dim GetDwgNo, OldGetDwgNo, GetShipMk, GetShopMk, FndGetDwgNo, FndGetShipMk, FndGetShopMk, FndGetPartDesc As String
        Dim FindShipMarks, PrevGetDwgNo, PrevGetShipMk, PrevGetShopMk, GetPartDesc, PrevGetPartDesc, PrevCcolumn, GetShopPartDesc As String
        Dim Style, Msg, Title, GetNextShopMk, GetNextShipMk As String
        Dim Response As Object
        Dim Workbooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim ExceptionPos, CallPos, LineNo, CntBOMItems As Integer
        Dim ShpMkList(3, 1)

        PrgName = "UpdateShpMarks"
        On Error GoTo Err_UpdateShpMarksArray

        '-------DJL-------12-20-2023-------Not sure ShpMkList is needed anymore?

        For i = 1 To (UBound(BOMList, 2) - 1)
            GetPartDesc = BOMList(7, i)
            GetDwgNo = BOMList(1, i)
            GetShipMk = BOMList(3, i)
            GetShipMk = LTrim(GetShipMk)
            GetShipMk = RTrim(GetShipMk)
            GetShopMk = BOMList(5, i)

            If GetShipMk <> "" Then
                If GetShipMk = " " Then
                    GoTo EmptySpace
                End If

                If GetDwgNo = "" And BOMList(7, i) <> "" Then
                    GetDwgNo = "XX"
                    BOMList(1, i) = GetDwgNo
                End If

                If GetPartDesc = Nothing And GetShopPartDesc <> Nothing Then
                    GetPartDesc = GetShopPartDesc
                End If

                ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo              'If assembly is called out on another drawing, Need way to Update Ship Mark.
                ShpMkList(2, UBound(ShpMkList, 2)) = GetShipMk
                ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                FindShipMarks = "No"
            Else
EmptySpace:

                If GetDwgNo = "" And BOMList(7, i) <> "" Then
                    GetDwgNo = "XX"
                    BOMList(1, i) = GetDwgNo
                End If

                If GetShipMk = " " Then
                    GetShipMk = ""
                End If

                If GetShipMk = Nothing Then
                    If FndGetShipMk = "FCO" And GetShipMk = Nothing Then
                        BOMList(3, i) = FndGetShipMk

                        If GetPartDesc = "" And GetShopPartDesc = "" Then
                            GetPartDesc = GetShopPartDesc
                        End If

                        ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                        ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo              'If assembly is called out on another drawing, Need way to Update Ship Mark.
                        ShpMkList(2, UBound(ShpMkList, 2)) = FndGetShipMk
                        ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                        ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                        GoTo EndFixFCO
                    End If

                    If i > 1 Then
                        PrevGetShopMk = BOMList(5, (i - 1))
                        PrevGetPartDesc = BOMList(7, (i - 1))
                    End If

                    If i = 0 And GetPartDesc = Nothing Then
                        GoTo NextDwg
                    End If

                    If GetDwgNo = Nothing And GetShipMk = Nothing Then
                        GoTo NextDwg
                    End If

                    If GetShopMk = PrevGetShopMk And GetDwgNo <> PrevGetDwgNo Then      'New Drawing but Piecemarks are the same and so is the Descriptions
                        If GetPartDesc = PrevGetPartDesc And GetShopMk <> "" Then
                            If i > 1 Then
                                PrevGetDwgNo = BOMList(1, (i - 1))
                            End If

                            Msg = "Is this the same part " & PrevGetShopMk & " on drawing " & PrevGetDwgNo & " Description = " & GetPartDesc & " as item " & GetShopMk & " on drawing " & GetDwgNo & " Description = " & GetPartDesc & "?"
                            Style = MsgBoxStyle.YesNo
                            Title = "Bulk BOM"
                            Response = MsgBox(Msg, Style, Title)

                            If Response = 6 Then
                                GetShipMk = PrevGetShipMk
                                BOMList(3, i) = GetShipMk
                            Else
                                GoTo NextDwg                'This is not a match go to next item in list.
                            End If

                            If GetPartDesc = "" And GetShopPartDesc = "" Then
                                GetPartDesc = GetShopPartDesc
                            End If

                            ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                            ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo              'If assembly is called out on another drawing, Need way to Update Ship Mark.
                            ShpMkList(2, UBound(ShpMkList, 2)) = GetShipMk
                            ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                            ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                            GoTo NextDwg
                        End If
                    Else
                        If GetDwgNo = PrevGetDwgNo And GetShipMk = Nothing Then              'Same Drawing Number but Ship marks are nothing, copy previous Ship marks.
                            If PrevGetShipMk <> "" And BOMList(3, (i - 1)) <> Nothing Then    'if Previous GetShipMk not nothing.
                                If PrevGetShipMk <> GetShopMk Then
                                    If Regex.IsMatch(Strings.Right(Mid(GetShopMk, 1, 1), 1), "[0-99]") Then
                                        GetShipMk = PrevGetShipMk
                                        BOMList(3, i) = GetShipMk
                                    Else
                                        GoTo SolutionNotFound
                                    End If
                                End If

                                If GetPartDesc = "" And IsNothing(GetShopPartDesc) = False Then
                                    GetPartDesc = GetShopPartDesc
                                End If

                                ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo          'If assembly is called out on another drawing, Need way to Update Ship Mark.
                                ShpMkList(2, UBound(ShpMkList, 2)) = GetShipMk
                                ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                                GoTo NextDwg
                            End If
                        End If
SolutionNotFound:
                    End If

                    If Mid(GetDwgNo, 1, 2) = Mid(OldGetDwgNo, 1, 2) And GetDwgNo <> OldGetDwgNo Then
                        FindShipMarks = "Yes"
                        If GetShipMk = Nothing Or GetShipMk = "" Then
FindNextShopMak:
                            If Len(GetShopMk) > 2 Then
                                For j = 0 To UBound(ShpMkList, 2)
                                    FndGetPartDesc = ShpMkList(0, j)
                                    FndGetDwgNo = ShpMkList(1, j)
                                    FndGetShipMk = ShpMkList(2, j)
                                    FndGetShopMk = ShpMkList(3, j)

                                    If Mid(FndGetDwgNo, 1, 2) = Mid(GetDwgNo, 1, 2) Then
                                        Select Case FndGetShipMk
                                            Case "FCO"
FixFCO:
                                                BOMList(3, i) = FndGetShipMk

                                                If GetPartDesc = "" And GetShopPartDesc = "" Then
                                                    GetPartDesc = GetShopPartDesc
                                                End If

                                                ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                                ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                                ShpMkList(2, UBound(ShpMkList, 2)) = FndGetShipMk
                                                ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                                ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                                                GoTo EndFixFCO
                                            Case Else
                                                If FndGetShopMk = GetShopMk And FndGetDwgNo = GetDwgNo Then
                                                    BOMList(3, i) = FndGetShipMk

                                                    If GetPartDesc = "" And GetShopPartDesc = "" Then
                                                        GetPartDesc = GetShopPartDesc
                                                    End If

                                                    ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                                    ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                                    ShpMkList(2, UBound(ShpMkList, 2)) = FndGetShipMk
                                                    ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                                    ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                                                    GoTo EndFixFCO
                                                Else
                                                    If FndGetShopMk = GetShopMk And FndGetDwgNo <> GetDwgNo Then
                                                        Msg = "Is this the same part " & FndGetShopMk & " on drawing " & FndGetDwgNo & " Description = " & FndGetPartDesc & " as item " & GetShopMk & " on drawing " & GetDwgNo & " Description = " & GetPartDesc & "?"
                                                        Style = MsgBoxStyle.YesNo
                                                        Title = "Bulk BOM"
                                                        Response = MsgBox(Msg, Style, Title)

                                                        If Response = 6 Then                    'if user clicks yes
                                                            GetShipMk = FndGetShipMk
                                                            BOMList(3, i) = GetShipMk
                                                        Else
                                                            GoTo NextItemInList                 'This is not a match go to next item in list.
                                                        End If

                                                        If GetPartDesc = "" And GetShopPartDesc = "" Then
                                                            GetPartDesc = GetShopPartDesc
                                                        End If

                                                        ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                                        ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                                        ShpMkList(2, UBound(ShpMkList, 2)) = FndGetShipMk
                                                        ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                                        ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                                                        GoTo EndFixFCO

                                                    Else
NextItemInList:
                                                        If FndGetShipMk = GetShopMk Then
                                                            BOMList(3, i) = FndGetShipMk
                                                            GetShipMk = FndGetShipMk

                                                            If GetPartDesc = "" And GetShopPartDesc = "" Then
                                                                GetPartDesc = GetShopPartDesc
                                                            End If

                                                            ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                                            ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                                            ShpMkList(2, UBound(ShpMkList, 2)) = FndGetShipMk
                                                            ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                                            ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                                                            GoTo EndFixFCO
                                                        End If
                                                    End If
                                                End If
                                        End Select
                                    End If
                                Next j
                            End If
EndFixFCO:
                        End If
                    Else
                        If FindShipMarks = "Yes" Then
                            If Regex.IsMatch(Strings.Right(Mid(GetShopMk, 1, 1), 1), "[0-9]") Then
                                BOMList(3, i) = FndGetShipMk

                                If GetPartDesc = "" And GetShopPartDesc = "" Then
                                    GetPartDesc = GetShopPartDesc
                                End If

                                ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                ShpMkList(2, UBound(ShpMkList, 2)) = FndGetShipMk
                                ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                            Else
                                If Regex.IsMatch(Strings.Right(Mid(GetShopMk, 1, 1), 1), "[A-Z]") Then
                                    GoTo FindNextShopMak
                                Else
                                    If i > 1 Then
                                        PrevGetDwgNo = BOMList(1, (i - 1))
                                        PrevGetShopMk = BOMList(5, (i - 1))
                                        PrevGetPartDesc = BOMList(7, (i - 1))
                                    End If

                                    If GetShopMk = Nothing Then
                                        If GetDwgNo = PrevGetDwgNo And FndGetShipMk <> Nothing Then
                                            BOMList(3, i) = FndGetShipMk
                                        End If

                                        If GetPartDesc = "" And GetShopPartDesc = "" Then
                                            GetPartDesc = GetShopPartDesc
                                        End If

                                        ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                        ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                        ShpMkList(2, UBound(ShpMkList, 2)) = GetShipMk
                                        ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                        ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                                    End If
                                End If
                            End If
                        Else
                            If PrevGetDwgNo = GetDwgNo And InStr(GetDwgNo, "10D") > 0 Then
                                BOMList(3, i) = PrevGetShipMk
                            Else
                                If InStr(GetDwgNo, "10D") = 0 And FndGetShipMk <> Nothing Then
                                    BOMList(3, i) = PrevGetShipMk
                                Else
                                    If PrevGetDwgNo = GetDwgNo And InStr(GetDwgNo, "10E") > 0 Then
                                        BOMList(3, i) = PrevGetShipMk
                                    Else
                                        If InStr(GetDwgNo, "10E") = 0 And FndGetShipMk <> Nothing Then
                                            BOMList(3, i) = PrevGetShipMk
                                        End If
                                    End If
                                End If
                            End If

                            If GetPartDesc = "" And GetShopPartDesc = "" Then
                                GetPartDesc = GetShopPartDesc
                            End If

                            If FindShipMarks = "Yes" Then                           '-------DJL-07-02-2025      'Only collect when Ship Mark is found.
                                ShpMkList(0, UBound(ShpMkList, 2)) = GetPartDesc
                                ShpMkList(1, UBound(ShpMkList, 2)) = GetDwgNo
                                ShpMkList(2, UBound(ShpMkList, 2)) = GetShipMk

                                If InStr(GetShipMk, "SR") = 1 Then
                                    BOMList(7, i) = "SHELL PLATE " & GetShipMk & BOMList(7, i)
                                End If

                                ShpMkList(3, UBound(ShpMkList, 2)) = GetShopMk
                                ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                            End If
                        End If
                        End If

                    If PrevGetDwgNo = GetDwgNo And GetShipMk = "" Then
                        If PrevGetShipMk <> "" And BOMList(3, i) = "" Then
                            If BOMList(3, (i - 1)) <> "" Then

                                If Regex.IsMatch(Strings.Right(Mid(GetShopMk, 1, 1), 1), "[A-Z]") Then
                                    BOMList(3, i) = GetShopMk
                                    PrevGetShipMk = GetShopMk
                                    BOMList(8, i) = "Yes"
                                Else
                                    BOMList(3, i) = PrevGetShipMk
                                End If
                            End If
                        End If
                    Else
                        If PrevGetDwgNo <> GetDwgNo And GetShipMk = "" Then                         'Job Found new drawing with no Ship marks for Assembly drawings.
                            GetNextShipMk = BOMList(3, (i + 1))

                            If PrevGetShipMk <> "" And GetNextShipMk <> "" Then
                                Msg = ("An error was found on your drawings were " & GetNextShipMk & " is not on the Assembly Items or Plates! Description = " & GetPartDesc & ". Is this the correct Shop Mark?")
                                Style = MsgBoxStyle.YesNo
                                Title = "Bulk BOM"
                                Response = MsgBox(Msg, Style, Title)

                                If Response = 6 Then                                    'if user clicks yes
                                    GetShipMk = GetNextShipMk
                                    BOMList(3, i) = GetNextShipMk
                                    GoTo FoundGetShipMk
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If Mid(OldGetDwgNo, 1, 2) <> Mid(GetDwgNo, 1, 2) Then
                FindShipMarks = "No"
                OldGetDwgNo = GetDwgNo
            Else
                OldGetDwgNo = GetDwgNo
            End If

            If PrevGetDwgNo <> GetDwgNo And InStr(GetDwgNo, "10D") > 0 Then
                If PrevGetShipMk <> GetShipMk And GetShipMk <> Nothing Then
                    GoTo FoundGetShipMk
                End If

                If PrevGetShipMk <> GetShipMk And IsNothing(PrevGetShipMk) = False Then
                    FndGetShipMk = BOMList(3, i)
                    OldGetDwgNo = BOMList(1, i)

                    If GetDwgNo = OldGetDwgNo And FndGetShipMk <> Nothing Then
                        PrevGetShipMk = FndGetShipMk
                    End If
                Else
                    FndGetShipMk = Nothing
                End If
            Else
                If PrevGetDwgNo <> GetDwgNo And InStr(GetDwgNo, "10E") > 0 Then
                    If PrevGetShipMk <> GetShipMk And GetShipMk <> Nothing Then
                        GoTo FoundGetShipMk
                    End If

                    If PrevGetShipMk <> GetShipMk And IsNothing(PrevGetShipMk) = False Then
                        FndGetShipMk = BOMList(3, i)
                        OldGetDwgNo = BOMList(1, i)

                        If GetDwgNo = OldGetDwgNo And FndGetShipMk <> Nothing Then
                            PrevGetShipMk = FndGetShipMk
                        End If
                    Else
                        FndGetShipMk = Nothing
                    End If
                Else
FoundGetShipMk:
                    If GetShipMk <> "" Then
                        PrevGetShipMk = GetShipMk
                    End If
                End If
            End If

NextDwg:

            PrevGetDwgNo = GetDwgNo
        Next i

        GenInfo3233.BOMList = BOMList

Err_UpdateShpMarksArray:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            PrgName = "UpdateShpMarksArray"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            ReadDwgs.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        Resume
                    End If
                    If CallPos > 0 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function
End Module