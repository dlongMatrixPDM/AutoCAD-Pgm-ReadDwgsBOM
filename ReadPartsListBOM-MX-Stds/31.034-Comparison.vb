'Option Strict Off
'Option Explicit On
'Option Compare Text

'Imports System
'Imports System.Reflection
'Imports System.Runtime.InteropServices
'Imports System.Text.RegularExpressions
'Imports Microsoft.Office.Interop.Excel

'Module Comparison31_142
'Public Structure InputType3
'    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As VariantType) As Integer
'    Public FirstTimeThru As String
'    Public FuncGetDataNew As String
'    Public NewBulkBOM As Object
'    Public MainBOMFile As Microsoft.Office.Interop.Excel.Workbook
'    Public NewPlateBOM As Object
'    Public NewStickBOM As Object
'    Public NewPurchaseBOM As Object
'    Public OldBOMFile As Microsoft.Office.Interop.Excel.Workbook
'    Public OldBulkBOM As Object
'    Public OldBulkBOMFile As String
'    Public OldPlateBOM As Object
'    Public OldStickBOM As Object
'    Public OldPurchaseBOM As Object
'    Public NewBOM As Object
'    Public OldBOM As Object
'    Public OldStdItems As Object
'    Public BOMType As String
'    Public BOMSheet As String
'    Public RowNo As String
'    Public RowNo2 As String
'    Public OldStdDwg As String
'    Public NewStdDwg As String
'    Public ExceptionPos As Integer
'    Public CallPos As Integer
'    Public CntExcept As Integer
'    Public ErrNo As String
'    Public ErrMsg As String
'    Public ErrSource As String
'    Public ErrDll As String
'    Public PriPrg As String
'    Public PrgName As String
'    Public ErrException As System.Exception
'    Public ErrLastLineX As Integer
'    Public Count As Integer
'    Public PassFilename As String
'    Public ReadyToContinue As Boolean
'    Public CBclicked As Boolean
'    Public errorExist As Boolean
'    Public AcadApp As Object
'    Public AcadDoc As Object
'    Public RevNo As String
'    Public RevNo2 As String
'    Public Continue_Renamed As Boolean
'    Public SortListing As Boolean
'    Public MatInch As Double
'    Public FoundDir As String
'    Public SearchException As String
'    Public ExceptPos As Integer
'    Public ThisDrawing As AutoCAD.AcadDocument
'    Public LytHid As Boolean

'    '        Public Shared Function ReadBOM(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
'    '            Dim iA As Object                                                'Used to read contents of "BulkBOMFab3D-Intent"
'    '            'Dim FoundLast As Boolean                                        'Create Array with all information on the BOM
'    '            Dim LineNo As Short, jA As Short, LineDel As Short
'    '            Dim Test As String

'    '            LineDel = 0
'    '            LineNo = SheetToUse.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

'    'LastLineFound:
'    '            LineNo = LineNo - LineDel
'    '            ReDim BomArray(14, LineNo - 4)
'    '            ReadDwgs.ProgressBar1.Maximum = LineNo

'    '            'Look at rewriting the AutoCAD BOM to use the existing array, and delete below'-------DJL-10-11-2023
'    '            For iA = 5 To LineNo
'    '                For jA = 1 To 13
'    '                    If SheetToUse.Range("A" & iA).Interior.ColorIndex <> 3 Then
'    '                        Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '                        BomArray(jA, iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '                    Else
'    '                        GoTo NextiA
'    '                    End If
'    '                Next jA

'    '                BomArray(14, iA - 4) = iA
'    '                ReadDwgs.ProgressBar1.Value = iA
'    'NextiA:
'    '            Next iA

'    '        End Function

'    '        Public Shared Function ReadBOMOld(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
'    '            Dim iA As Object                                                'Used to read contents of "BulkBOMFab3D-Intent"
'    '            Dim FoundLast As Boolean                                        'Create Array with all information on the BOM
'    '            Dim LineNo, jA, LineDel As Short
'    '            Dim Test, Test2 As String

'    '            Test = SheetToUse.Name
'    '            FoundLast = False
'    '            LineNo = 4
'    '            LineDel = 0

'    '            Do Until FoundLast = True
'    '                LineNo = LineNo + 1
'    '                Test = SheetToUse.Range("G" & LineNo).Value
'    '                Test2 = SheetToUse.Range("G" & LineNo).Interior.ColorIndex

'    '                If SheetToUse.Range("G" & LineNo).Value = "" Then
'    '                    LineNo = LineNo - 1
'    '                    FoundLast = True
'    '                End If
'    '                If SheetToUse.Range("G" & LineNo).Interior.ColorIndex = 3 Then
'    '                    LineDel = LineDel + 1
'    '                End If
'    '            Loop

'    '            LineNo = LineNo - LineDel
'    '            ReDim BomArray(13, LineNo - 4)
'    '            ReadDwgs.ProgressBar1.Maximum = LineNo

'    '            For iA = 5 To LineNo
'    '                For jA = 2 To 13
'    '                    If SheetToUse.Range("G" & iA).Interior.ColorIndex <> 3 Then
'    '                        Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '                        BomArray((jA - 1), iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '                    Else
'    '                        GoTo NextiA
'    '                    End If
'    '                Next jA
'    '                ReadDwgs.ProgressBar1.Value = iA
'    'NextiA:
'    '            Next iA

'    '        End Function

'    'Public Shared Function ReadBulkBOM(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
'    '    Dim iA As Object                                                'Used to read contents of "BulkBOMFab3D-Intent"
'    '    'Dim FoundLast As Boolean                                        'Create Array with all information on the BOM
'    '    Dim LineNo
'    '    'Dim RangeArea
'    '    Dim jA As Short
'    '    Dim Test As String          ', Test2 As String

'    '    LineNo = SheetToUse.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

'    '    ReDim BomArray(11, LineNo - 4)

'    '    'Look at rewriting the AutoCAD BOM to use the existing array, and delete below'-------DJL-10-11-2023
'    '    For iA = 5 To LineNo
'    '        For jA = 1 To 11
'    '            Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '            BomArray(jA, iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '        Next jA
'    '    Next iA

'    'End Function

'    'Public Shared Function ReadFindSTDs(ByRef BomArray As Object, ByRef SheetToUse As Object) As Object
'    '    Dim iA As Object                        'Used to read contents of STDs BOM.
'    '    Dim LineNo As Short, jA As Short
'    '    Dim Test As String

'    '    LineNo = SheetToUse.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

'    '    'Look at rewriting the AutoCAD BOM to use the existing array, and delete below'-------DJL-10-11-2023
'    '    ReDim BomArray(14, LineNo - 4)

'    '    For iA = 5 To LineNo
'    '        For jA = 1 To 13
'    '            Test = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '            BomArray(jA, iA - 4) = SheetToUse.Range(Chr(jA + 64) & iA).Value
'    '        Next jA
'    '    Next iA

'    'End Function
'End Structure
'End Module