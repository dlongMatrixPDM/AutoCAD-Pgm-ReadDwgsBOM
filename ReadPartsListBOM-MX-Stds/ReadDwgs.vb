Option Strict Off
Option Explicit On
Option Compare Text

Imports System.IO
Imports System.IO.Stream
Imports System.Data.SqlClient                   'Used for SQL Connection

Imports System.Reflection
Imports System.Text.RegularExpressions
Imports AutoCAD
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.SqlServer

Public Structure GenInfo3233
    Public Shared FileName As String
    Public Shared FullJobNo As String
    Public Shared Customer As String            '-------DJL-07-14-2025      'Added for Shipping Info.

    Public Shared BOMList(20, 1)            'BOM Items on drawings.+
    Public Shared STDsList(20, 1)           'Standards found On drawings
    Public Shared StdsFnd2(2, 1)
    Public Shared StdsBOMList(20, 1)        'All Stds for MX-Standards
    'Public Shared CollectSTDsList(20, 1)

    Public Shared FileDir As String
    Public Shared JobDir As String
    Public Shared BomListFileName As String         '-------DJL-12-19-2024
    Public Shared StartAdept As Boolean

    Public Shared ExDwgs1 As Object
    Public Shared PrgReqSpreadSht As String            'Program Requesting SpreadSheet Info.
    Public Shared SpreadSht As String                   'Spreadsheet user Selected.
    Public Shared SpreadshtLoc As String                'Spreadsheet location.

    Public Shared UserName As String                   '-------DJL-06-10-2025
    Public Shared OldDesc As String
    Public Shared GetDesc As String
    Public Shared GetMatLen As String
    Public Shared GetMatQty As String
    Public Shared GetShipMk As String
    Public Shared GetPieceMk As String
    Public Shared MatType As String
    Public Shared RevNo As String

    Public Shared AddMat As String
    Public Shared AddMatSize As String
    Public Shared AddMatQty As String
    Public Shared SubMFGData(3, 1) As Object
End Structure

Public Class ReadDwgs                      'BulkBOMFab3D
    Inherits System.Windows.Forms.Form
    Public Shared AcadApp As AutoCAD.AcadApplication
    'Public Shared ExcelApp As Excel.Application
    Dim ErrMsg, ErrNo, ErrSource, ErrDll, PriPrg, PrgName As String
    Dim ErrException As System.Exception
    Dim ErrLastLineX As Integer
    Dim FileToOpen As String
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As VariantType) As Integer

    Public FirstTimeThru As String
    Public FuncGetDataNew As String
    Public MainBOMFile
    Public NewBOM As Object
    Public OldBOM As Object
    Public FindSTD As Object
    Public OldStdItems As Object
    Public BOMType As String
    Public RowNo As String
    Public RowNo2 As String
    Public OldStdDwg As String
    Public NewStdDwg As String
    Public Count As Integer
    Public BOMList(21, 1) As Object                   '-------DJL-06-10-2025   'Moved below Just collect per drawing.          'Added 21 for indexing-------DJL-06-06-2025
    Public BOMListNew(21, 1) As Object                          '-------DJL-06-09-2025
    Public BOMListSort As New ArrayList                         '(1, 1) As ArrayList
    'Public BOMListColl(20, 1) As Object
    'Public BOMListArray(20, 1) As Array
    'Public BOMListTemp(20, 1) As Object                         'Tried Array-------DJL
    Public NewBOMList(20, 1) As String
    'Public FindStdList(20, 1) As String
    'Public FoundStdList(20, 1) As String
    Public ShpMkList() As String
    Public STDsList(20, 1)              'As String
    Public StdsBOMList(20, 1)
    'Public CollectSTDsList(13, 1)

    Public AcadDoc As AutoCAD.AcadDocument
    Public RevNo As String
    Public SearchException As String
    Public ExceptPos As Integer
    Dim CountNewItems As Integer
    Dim JobNoIndex As Integer
    Dim Test, StartDir, SecondChk, WorkShtName, PartFound, GetPartDesc, GetMat, GetMat2, GetMat3, NewDir, NTest5, Ntest6, Inv1 As String
    Dim FType, FirstJobNo, FileSaveAS, FileNam, DwgTitle1, DwgTitle2, DwgTitle3, BadDwgFound As String

    Public GetLen As String
    Public Workbooks As Microsoft.Office.Interop.Excel.Workbooks
    Public WorkSht As Worksheet

    Const SearchInch As String = Chr(34)                'Find " Inch
    Const SearchFoot As String = Chr(39)                'Find ' Feet
    Const SearchSpace As String = " "
    Public LenPrgLineNo As Integer
    Public StartAdept As Boolean
    Public BOMWrkSht As Worksheet
    Public ShippingWrkSht As Worksheet
    Public db_String As String
    Public OldFileNam As String
    Public Sapi As Object = CreateObject("SAPI.spvoice")

    Private Sub AddAllButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnAddAll.Click
        Dim Dupe As Boolean
        Dim Count, i, j, CountDwgs, CntDwgs As Integer
        Dim VarSelArray As Object
        Dim DwgListArray As Object
        Dim Test2 As String

        PrgName = "AddAllButton_Click"

        On Error GoTo Err_AddAllButton_Click

        CntDwgs = 0
        CountDwgs = DwgList.Items.Count
        ReDim DwgListArray(CountDwgs)

        For i = 0 To (CountDwgs - 1)
            SelectList.Items.Add(DwgList.Items.Item(i))
            DwgListArray(CntDwgs) = DwgList.Items.Item(i)
            CntDwgs = (CntDwgs + 1)
        Next i

        For i = 0 To (CountDwgs - 1)
            DwgList.Items.Remove(DwgListArray(i))
        Next

        DwgList.Sorted = True
        SelectList.Sorted = True
        BtnRemove.Enabled = True
        BtnClear.Enabled = True
        BtnStart2.Enabled = True

        If SelectList.Items.Count > 0 Then
            SelectList.BackColor = System.Drawing.Color.GreenYellow
            DwgList.BackColor = System.Drawing.Color.LightBlue
        End If
Err_AddAllButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

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

    End Sub

    Private Sub AddButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnAdd.Click
        Dim Dupe As Boolean
        Dim Count, i, j, CountDwgs, CntDwgs As Integer
        Dim DwgListArray As Object
        Dim Test, Test2 As String

        PrgName = "AddButton_Click"

        On Error GoTo Err_AddButton_Click

        CntDwgs = 0
        CountDwgs = DwgList.SelectedItems.Count
        ReDim DwgListArray(CountDwgs)

        For i = 0 To (CountDwgs - 1)                                'Do not need to look at everydwg just look at selected Dwgs.
            SelectList.Items.Add(DwgList.SelectedItems.Item(i))
            DwgListArray(CntDwgs) = DwgList.SelectedItems.Item(i)   'So create Array and delete after complete.
            CntDwgs = (CntDwgs + 1)
        Next i

        For i = 0 To (CountDwgs - 1)
            DwgList.Items.Remove(DwgListArray(i))
        Next

        DwgList.Sorted = True
        SelectList.Sorted = True

        If SelectList.Items.Count = 0 Then
            BtnRemove.Enabled = False
            BtnClear.Enabled = False
        Else
            BtnRemove.Enabled = True
            BtnClear.Enabled = True
        End If

        If Me.SelectList.Items.Count <> 0 Then
            Me.BtnStart2.Enabled = True
        End If

        If SelectList.Items.Count > 0 Then
            SelectList.BackColor = System.Drawing.Color.GreenYellow
            DwgList.BackColor = System.Drawing.Color.LightBlue
        End If

Err_AddButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

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

    End Sub

    Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelButton_Renamed.Click
        Me.Close()
    End Sub

    Private Sub ClearButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnClear.Click
        SelectList.Items.Clear()
        BtnRemove.Enabled = False
        BtnClear.Enabled = False
        BtnStart2.Enabled = False
    End Sub

    Private Sub RemoveButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles BtnRemove.Click
        Dim Dupe As Boolean
        Dim Count, i, j, DwgPos, LenDwg1, LenDwg2, LenPath, CountDwgs, CntDwgs As Integer
        Dim VarSelArray As Object
        Dim GetDwg, FilePath, FindFile As String
        Dim DwgListArray As Object

        PrgName = "RemoveButton_Click"

        On Error GoTo Err_RemoveButton_Click

        Count = Me.SelectList.Items.Count - 1
        CountDwgs = SelectList.SelectedItems.Count
        CntDwgs = 0
        ReDim DwgListArray(CountDwgs)

        For i = 0 To (CountDwgs - 1)    'Do not need to look at every drawing, and Array is not required.
            DwgPos = 0
            GetDwg = SelectList.SelectedItems.Item(i).ToString
            Me.DwgList.Items.Add(GetDwg)
            CntDwgs = (CntDwgs + 1)                         '-------DJL-08-08-2025      'Added
            DwgListArray(CntDwgs) = SelectList.SelectedItems.Item(i)  'So create Array and delete after complete.
        Next i

        If SelectList.Items.Count = 0 Then
            BtnRemove.Enabled = False
            BtnClear.Enabled = False
            BtnStart2.Enabled = False
        End If

        For i = 0 To (CntDwgs)                            '-------DJL-08-08-2025      'For i = 0 To (CountDwgs - 1)
            SelectList.Items.Remove(DwgListArray(i))
        Next

        DwgList.Sorted = True
        SelectList.Sorted = True

Err_RemoveButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

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

    End Sub

    Private Sub UserForm_Initialize()
        Dim RevList(20) As Short
        Dim i As Short
        PrgName = "UserForm_Initialize"

        On Error GoTo Err_UserForm_Initialize

        Me.LblProgress.Text = "Progress........"

        For i = 0 To 20
            RevList(i) = i
        Next i

        Me.ComboBxRev.Items.Add(RevList)

Err_UserForm_Initialize:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

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

    End Sub

    Private Async Sub BOM_Menu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Short
        PrgName = "BOM_Menu_Load"

        'On Error GoTo Err_BOM_Menu_Load

        Me.LblProgress.Text = "Progress........"

        For i = 0 To 20
            Me.ComboBxRev.Items.Add(i)
        Next i

        Me.PathBox_Click(sender, e)

        '-------Email Bill Sieg, Since you  “I don’t even look at them, I just run them and send them.” Then the process Is being removed.

Err_BOM_Menu_Load:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                'Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                'Resume
            Else
                ExceptionPos = InStr(1, ErrMsg, "Exception")
                CallPos = InStr(1, ErrMsg, "Call was rejected by callee")
                CntExcept = (CntExcept + 1)

                If CntExcept < 20 Then
                    If ExceptionPos > 0 Then
                        'Resume
                    End If
                    If CallPos > 0 Then
                        'Resume
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub BtnStart2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnStart2.Click
        Dim VarSelArray As Object, BlockData(1) As Object, DwgItem As Object
        Dim TempAttributes As Object, Temparray As Object, InsertionPT As Object, gvntSDIvar As Object
        Dim DwgIndex As Single              '-------DJL-06-27-2025
        Dim Title, Msg, Style, Response As Object
        Dim BlockSel As AutoCAD.AcadSelectionSet
        Dim AcadOpen As Boolean
        Dim GroupCode(1) As Short
        Dim AcadPref As AutoCAD.AcadPreferencesSystemClass
        Dim Dimscale, CompareX2 As Double, CompareX1 As Double     '-------DJL-07-03-2025      'Not Required.
        Dim a, b, c, d, f, i, s, t, v, x, y, z, ShtPos, QtyMarkIssue, NotePos, DecPos, CountVal, CntTitleBlks, CntDwgs, LinesRemoved, CntCollected, CntBOMList, RowNoPlus4 As Integer
        Dim TestTags, Test1, OldFileNam, MX2, MissingTxt1, MissingTxt2, LookForStd, InsPt0, InsPt1, LineNo, LineNo2, LineNo3 As String
        Dim SpacePos, CntBlks, CntCollect, CntSortList4, CntAttFound, DwgNoPos As Integer
        Dim CurrentDwgRev, CurrentDwgNo, FilePath, CurrentDWG, PrgName, WorkBookName, AdeptStd, AdeptPrg, BOMListSortChg As String
        Dim GetShipStds, GetInv1, GetInv2, Get2DShipQty, FullJobNo, FoundIndex, FirstDwg, FindIndex, ExistSTD, DrawingIndex, CurrentStdNo As String
        Dim GetPartNo, Get2DShipMk, GetQty, GetShipDesc, GetDesc, GetLen, GetMat, GetMat2, GetMat3, GetNotes, GetWt, GetProc, DwgItem2, Dwg As String
        Dim BOMItemNam, Customer, ChkSort As String
        Dim StdsWrkSht As Worksheet, StdItemsWrkSht As Worksheet

        PrgName = "StartButton_Click-Part1"

        On Error GoTo Err_StartButton_Click 'On Error Resume Next

        ExceptPos = 0
        BadDwgFound = "No"
        LinesRemoved = 0
        QtyMarkIssue = 0
        SecondChk = "First"
        Me.BtnStart2.Enabled = False

        If Directory.Exists("K:\AWA\" & System.Environment.UserName & "\AdeptWork\") = False Then                            'DJL 9-17-2024
            ClosePrg("Excel", SecondChk, StartAdept)
            'ClosePrg("Adept", SecondChk, StartAdept)                               '-------Hold for now.
        End If

        PrgName = "StartButton_Click-Part2"

        On Error Resume Next

        ExcelApp = GetObject(, "Excel.Application")

        If Err.Number Then
            Information.Err.Clear()
            ExcelApp = CreateObject("Excel.Application")
            If Err.Number Then
                MsgBox(Err.Description)
                Exit Sub
            End If
        End If

        On Error GoTo Err_StartButton_Click

        BOMType = "Tank"

        If Me.ComboBxRev.Text = vbNullString Then
            Sapi.Speak("Please select a revision number for Bulk BOM")
            MsgBox("Please select a revision number for Bulk BOM")
            Exit Sub
        End If

        Me.LblProgress.Text = "Gathering Information from AutoCAD Drawings........Please Wait"
        Me.Refresh()

        PrgName = "StartButton_Click-Part3"

        On Error Resume Next
        If Err.Number Then
            Err.Clear()
        End If

        AcadApp = GetObject(, "AutoCAD.Application")
        AcadOpen = True

        If Err.Number Then
            Information.Err.Clear()
            AcadApp = CreateObject("AutoCAD.Application")
            AcadOpen = False
            If Err.Number Then
                Information.Err.Clear()
                AcadApp = CreateObject("AutoCAD.Application")
                If Err.Number Then
                    MsgBox(Err.Description)
                    MsgBox("Instance of 'AutoCAD.Application' could not be created.")

                    AcadApp.Visible = False
                    MsgBox("Now running " & AcadApp.Name & " version " & AcadApp.Version)

                    If System.Environment.UserName = "dlong" Then
                        Stop
                        Resume
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If

        AcadApp.Visible = False

        On Error GoTo Err_StartButton_Click

        AcadPref = AcadApp.Preferences.System

        If AcadPref.SingleDocumentMode = True Then
            gvntSDIvar = AcadPref.SingleDocumentMode
        End If

        If gvntSDIvar = True Then
            AcadPref.SingleDocumentMode = False
        End If

        PrgName = "StartButton_Click-Part4"

        AcadApp.WindowState = AutoCAD.AcWindowState.acMin
        FirstDwg = "NotFound"
        FirstDwg = Nothing
        Me.Refresh()
        Count = Me.SelectList.Items.Count
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = Count
        ProgressBar1.Visible = True
        CountVal = 0

        If Count <> 0 Then
            ReDim VarSelArray(Count)
            For a = 1 To Count
                VarSelArray(a) = Me.SelectList.Items(a - 1)
            Next a

            CntDwgs = UBound(VarSelArray)

            For z = 1 To CntDwgs
                'FindNextDwg:
                ProgressBar1.Value = z
                DwgItem = VarSelArray(z)

                If IsNothing(DwgItem) = True Then
                    GoTo NextDwg
                End If

                PrgName = "StartButton_Click-Part5"
                AcadApp.Documents.Open(DwgItem.FullName)
                Dim BOMList(21, 1) As Object                            '-------DJL-06-10-2025-REFINE IT EACH TIME IT STARTS COLLECTING

                AcadApp.WindowState = AutoCAD.AcWindowState.acMin
                AcadApp.Visible = False
                CurrentDwgNo = Nothing
                Me.TxtBoxDwgsToProcess.Visible = True
                Me.LblDwgsToProcess.Visible = True
                Me.TxtBoxDwgsToProcess.Text = CntDwgs - z
                Me.TxtBoxBOMItemsToProcess.Visible = True
                Me.LblBOMItemsToProess.Visible = True
                Me.Refresh()

                If BadDwgFound = "Yes" Then
                    BadDwgFound = "No"
                    GoTo NextDwg
                End If

                On Error GoTo Err_StartButton_Click

                Me.Refresh()
                AcadDoc = AcadApp.ActiveDocument

                PrgName = "StartButton_Click-Part6"

                '-------------------------------------------There is no reason to look at every titleblock ?
                BlockSel = AcadDoc.SelectionSets.Add("Titleblock")
                GroupCode(0) = 0
                BlockData(0) = "INSERT"
                GroupCode(1) = 2                                                                                                    'Who added this ?   '2211-4070-BRDER_11x17                      
                BlockData(1) = "AMW_TITLE,OSF_TITLE,OSF_TITLE_D,MX_TITLE,LNG_TITLE_D,MX_TITLE_SP,MX_TITLE-11x17,MTRX_PDM-BRDER_11x17,2211-4070-BRDER_11x17,Title Blocks Matrix,MATRIX TITLEBLOCK,MPDM_STD_TITLE*,MPDM_MEC_TITLE*,MPDM TULSA*"          'DJL-------06-27-2025        'Added ,MPDM TULSA* 
                BlockSel.Select(AutoCAD.AcSelect.acSelectionSetAll, , , GroupCode, BlockData)
                CntBlks = BlockSel.Count

                If CntBlks = Nothing Or CntBlks = 0 Then
                    MsgBox("Drawing " & AcadDoc.Name & " has no titleblock and will be skipped.")
                    AcadDoc.Close()
                    GoTo NextDwg
                End If

                CntTitleBlks = 0
                Temparray = BlockSel.Item(CntTitleBlks).GetAttributes

                '                For f = 1 To CntBlks
                'FindNextBlk:
                PrgName = "StartButton_Click-Part7"

FindAttributes:
                    For t = 0 To UBound(Temparray)
                        Test = Temparray(t).TagString                          '-------Requisition Number not on BOM.
                        Test1 = Temparray(t).TextString

                        Select Case Temparray(t).TagString
                            Case "DN"
                                CurrentDwgNo = Temparray(t).TextString

                            If InStr(CurrentDwgNo, "(SH") > 0 Then          '-------New problem Job 2212-1001-HVEC has put the (sht 1 of 2) on DN tag for dwg no.
                                ShtPos = InStr(CurrentDwgNo, "(SH")
                                CurrentDwgNo = Mid(CurrentDwgNo, 1, (ShtPos - 1))
                            Else
                                If InStr(CurrentDwgNo, "(") > 0 Then
                                    ShtPos = InStr(CurrentDwgNo, "(")
                                    CurrentDwgNo = Mid(CurrentDwgNo, 1, (ShtPos - 1))
                                End If
                            End If

                            DwgNoPos = InStr(CurrentDwgNo, "-DW-")             '-------DJL-07-07-2025

                            If DwgNoPos > 0 Then                                '5606-1107-210A-DW-     'Shiela Ganote
                                CurrentDwgNo = Mid(CurrentDwgNo, (DwgNoPos + 4), Len(CurrentDwgNo))             '-------DJL-07-07-2025
                            End If

                            CurrentDwgNo = LTrim(CurrentDwgNo)
                                CurrentDwgNo = RTrim(CurrentDwgNo)
                                CntCollect = (CntCollect + 1)
                            Case "TITLE"
                                CurrentDwgNo = Temparray(t).TextString                          '-------Drawing Converted for Inventor to AutoCAD
                                CntCollect = (CntCollect + 1)
                            Case "DT"
                                DwgTitle1 = Temparray(t).TextString
                            Case "DT1"
                                If IsNothing(DwgTitle1) = True Then
                                    DwgTitle1 = Temparray(t).TextString
                                End If
                        'Case "DT2"                          '-------DJL-06-27-2025         'Not requied for BOM, or Shipping List.
                        '    DwgTitle2 = Temparray(t).TextString
                        'Case "DT3"                          '-------DJL-06-27-2025
                        '    If IsNothing(DwgTitle3) = True Then
                        '        DwgTitle3 = Temparray(t).TextString
                        '    End If
                        Case "C"      '-------DJL-07-25-2025      'Added Customer so that Shipping list could be produced also.
                            If Customer = "" Then
                                Customer = Temparray(t).TextString
                                GenInfo3233.Customer = Customer
                                CntCollect = (CntCollect + 1)
                            End If
                        Case "CT"      '-------DJL-07-14-2025      'Added Customer so that Shipping list could be produced also.
                            If Customer = "" Then
                                Customer = Temparray(t).TextString
                                GenInfo3233.Customer = Customer
                                CntCollect = (CntCollect + 1)
                            End If
                        Case "C1"      '-------DJL-07-14-2025      'Added Customer so that Shipping list could be produced also.
                            If Customer = "" Then
                                Customer = Temparray(t).TextString
                                GenInfo3233.Customer = Customer
                                CntCollect = (CntCollect + 1)
                            End If
                        Case "JN"
                            FullJobNo = Temparray(t).TextString
                                JobNoIndex = t
                                CntCollect = (CntCollect + 1)
                            Case "PROJECT"
                                FullJobNo = Temparray(t).TextString
                                JobNoIndex = t
                                CntCollect = (CntCollect + 1)
                            Case "PN"
                                FullJobNo = Temparray(t).TextString
                                JobNoIndex = t
                                CntCollect = (CntCollect + 1)
                            Case "RN"
                                CurrentDwgRev = Temparray(t).TextString
                                CntCollect = (CntCollect + 1)
                            Case "REV"
                                CurrentDwgRev = Temparray(t).TextString
                                CntCollect = (CntCollect + 1)
                            Case "REVISION_NUMBER"
                                CurrentDwgRev = Temparray(t).TextString
                                CntCollect = (CntCollect + 1)
                                'Case Else
                                '    Stop
                        End Select

                    If CntCollect = 4 Then      '-------DJL-07-14-2025      'Added Customer so that Shipping list could be produced also.      'If CntCollect = 3 Then 
                        GoTo FoundAllItems
                    End If

                    PrgName = "StartButton_Click-Part8"
                    Next t

                    '    If UBound(Temparray, 2) = 0 And CntBlks > 1 Then            '-------DJL-07-03-2025      'Found issue were drawing has two references to a titleblock one of them has nothing in the attributes.
                    '        CntTitleBlks = (CntTitleBlks + 1)                       'Index starts at 0 to 1 for two items.
                    '        Temparray = BlockSel.Item(CntTitleBlks).GetAttributes

                    '        GoTo FindNextBlk
                    '    End If
                    'Next f                          '-------DJL-07-03-2025

FoundAllItems:
                    If CurrentDwgNo = Nothing And CntBlks > 0 Then
                        CntTitleBlks = (CntTitleBlks + 1)

                        If CntTitleBlks <= (CntBlks - 1) Then
                            Temparray = BlockSel.Item(CntTitleBlks).GetAttributes
                        End If

                        GoTo FindAttributes
                    End If

                    PrgName = "StartButton_Click-Part9"
                    CntCollect = Nothing
                    GenInfo3233.FullJobNo = FullJobNo

                    If FirstJobNo <> FullJobNo Then
                        If IsNothing(FirstJobNo) = True Then
                            FirstJobNo = FullJobNo
                            GoTo JobNoChecked
                        End If

                        Msg = "Drawing " & CurrentDwgNo & " has a different Job Number example " & FirstJobNo & " and " & FullJobNo & " do not match, Do you want to use Job No ? " & FirstJobNo
                        Style = MsgBoxStyle.YesNo
                        Title = "Job Number Issue"
                        Response = MsgBox(Msg, Style, Title)

                        If Response = 6 Then
                            FullJobNo = FirstJobNo
                            Temparray(JobNoIndex).TextString = FirstJobNo
                        Else
                            FirstJobNo = FullJobNo
                            Temparray(JobNoIndex).TextString = FullJobNo
                        End If
                    End If

JobNoChecked:
                    PrgName = "StartButton_Click-Part10"
                    BlockSel = AcadDoc.SelectionSets.Add("BillOfMaterial")
                    GroupCode(0) = 0
                    BlockData(0) = "INSERT"
                    GroupCode(1) = 2
                    BlockData(1) = "STANDARD_BILL_OF_MATERIAL,B_BILL_OF_MATERIAL,SP_BILL_OF_MATERIAL,BILL OF MATERIAL,STANDARD_BILL_OF_MATERIAL_Erectino Double Digit,STANDARD_BILL_OF_MATERIAL Assembly,STANDARD_BILL_OF_MATERIAL Erection Single Digit"
                    BlockSel.Select(AutoCAD.AcSelect.acSelectionSetAll, , , GroupCode, BlockData)           'Some drawings are giving BlockSel.Count = 0 RW-------DJL this is true if no BOM Blocks exist on drawing.     

                    'Did not remove the problem
                    'BOMListSort.Add("0000")           '-------DJL 06-27-2025      'Modified record zero to = 0 was having problems resorting started @ 1

                    If BlockSel.Count <> 0 Then
                        CntCollected = 0

                        For Each BomItem In BlockSel
                            PrgName = "StartButton_Click-Part11"
                            CntCollected = CntCollected + 1
                            Me.TxtBoxBOMItemsToProcess.Text = BlockSel.Count - CntCollected
                            Me.Refresh()
                            TempAttributes = BomItem.GetAttributes

                            For s = 0 To UBound(TempAttributes)     'Not Required Found Job Information Previuosly 'Need info for Standards Dwg Number.
                                Test1 = TempAttributes(s).TagString

                            Select Case TempAttributes(s).TagString
                                Case "SLM"                          '-------MK"
                                    Get2DShipMk = TempAttributes(s).TextString

                                    '-------Temp test to find Qty error.
                                    'If Get2DShipMk = "E38-11" Or Get2DShipMk = "E38-03" Then          '-------DJL-07-28-2025
                                    '    Stop
                                    'ElseIf Get2DShipMk = "E38-04" Or Get2DShipMk = "E38-05" Then          '-------DJL-07-28-2025
                                    '    Stop
                                    'End If

                                        'If Get2DShipMk <> "" Then          '-------DJL-07-03-2025      'Moved below at Array collection if Inv1 and Inv2 are Not Nothing.
                                        '    GetShipStds = "Yes"
                                        'Else
                                        '    GetShipStds = "No"
                                        'End If
                                Case "SLQ"                          '-------MK"
                                    Get2DShipQty = TempAttributes(s).TextString
                                Case "SM"                           '-------SP"
                                    GetPartNo = TempAttributes(s).TextString            'Found drawing with two different Attributes labeled both m 1 = A105 2 = 5
                                Case "Q"
                                    GetQty = TempAttributes(s).TextString

                                    If Get2DShipMk = "" And Get2DShipQty = "" Then      '-------07-14-2025  'Added
                                        If GetQty <> "" And GetMat <> "" Then
                                            GetPartNo = GetMat
                                        End If
                                    End If
                                Case "SD"
                                    GetShipDesc = TempAttributes(s).TextString
                                Case "D"
                                    GetDesc = TempAttributes(s).TextString
                                Case "D2"
                                    GetShipDesc = TempAttributes(s).TextString
                                Case "IU"
                                    GetInv1 = TempAttributes(s).TextString
                                Case "IL"
                                    GetInv2 = TempAttributes(s).TextString
                                Case "SP"
                                    GetPartNo = TempAttributes(s).TextString
                                Case "MK"
                                    Get2DShipMk = TempAttributes(s).TextString
                                        'GetShipStds = "No"                         '-------DJL-07-03-2025          'Moved below at array collection.
                                Case "L"
                                    GetLen = TempAttributes(s).TextString
                                Case "M"      'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                                    If GetMat <> "" Then            '-------DJL-07-14-2025      'May need to look at material types in order to find Piece marks.
                                        If GetPartNo <> GetMat Then
                                            GetPartNo = GetMat
                                        End If

                                        GetMat = TempAttributes(s).TextString
                                    Else
                                        GetMat = TempAttributes(s).TextString
                                    End If
                                Case "M2"      'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                                    GetMat2 = TempAttributes(s).TextString
                                Case "M3"      'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                                    GetMat3 = TempAttributes(s).TextString
                                Case "N"
                                    GetNotes = TempAttributes(s).TextString
                                Case "W"
                                    GetWt = TempAttributes(s).TextString
                                Case "P"
                                    GetProc = TempAttributes(s).TextString
                            End Select
                        Next s

BOMInfoCollected:
                        If GetMat <> "" And MX2 = "Found" Then
                            If GetPartNo = "" And Get2DShipMk = "" Then
                                GetPartNo = GetMat            'For some reason EEP Users have BOM insert with two m = A105 M2= Part number

                                If GetMat2 <> "" Then
                                    GetMat = ""
                                End If
                            End If
                        End If

                        'Dim BOMList(21, 1) As Object        'Moved above at open each drawing                    '-------DJL-06-10-2025-REFINE IT EACH TIME IT STARTS COLLECTING
                        InsertionPT = BomItem.InsertionPoint
                            Dimscale = BomItem.XScaleFactor
                        'CompareX1 = 10.5 * Dimscale                         '-------DJL-07-03-2025      'Not Required.
                        'CompareX1 = InsertionPT(0) - CompareX1
                        'CompareX1 = CompareX1 / Dimscale

                        'CompareX2 = 6 * Dimscale                         '-------DJL-07-03-2025      'Not Required.
                        'CompareX2 = InsertionPT(0) - CompareX2
                        'CompareX2 = CompareX2 / Dimscale

                        'If CompareX1 <1 Or CompareX2 <1 Then                         '-------DJL-07-03-2025      ' Not Required.
                        '    If CompareX1 > 0 Or CompareX2 > 0 Then
                                        '        BOMList(16, UBound(BOMList, 2)) = CStr(1)
                                        '    Else
                                        '        BOMList(16, UBound(BOMList, 2)) = CStr(2)
                                        '    End If
                                        'Else
                                        '    BOMList(16, UBound(BOMList, 2)) = CStr(2)
                                        'End If

                                        BOMList(1, UBound(BOMList, 2)) = CurrentDwgNo                               '-------Dwg number
                        BOMList(2, UBound(BOMList, 2)) = CurrentDwgRev                              '-------Rev Number
                        BOMList(3, UBound(BOMList, 2)) = Get2DShipMk

                        If GetQty = Nothing Or GetQty = "" Then                         '-------DJL-07-28-2025      'If GetQty = Nothing Then
                            BOMList(4, UBound(BOMList, 2)) = Get2DShipQty
                        Else
                            If GetQty = " " Then                         '-------DJL-07-28-2025
                                BOMList(4, UBound(BOMList, 2)) = Get2DShipQty
                            Else
                                BOMList(4, UBound(BOMList, 2)) = GetQty
                            End If
                        End If

                        If MX2 = "Found" Then
                            If GetPartNo = "A105" Then
                                Test1 = GetMat
                                GetMat = GetPartNo
                                GetPartNo = Test1
                            Else
                                If GetPartNo = "A106" Then
                                    Test1 = GetMat
                                    GetMat = GetPartNo
                                    GetPartNo = Test1
                                Else
                                    If GetPartNo = "A36" Then
                                        Test1 = GetMat
                                        GetMat = GetPartNo
                                        GetPartNo = Test1
                                    Else
                                        If GetPartNo <> "" And GetMat <> "" Then
                                            Test1 = GetMat
                                            GetMat = GetPartNo
                                            GetPartNo = Test1
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        'If GetPartNo = "" Then
                        '    BOMList(5, UBound(BOMList, 2)) = "-"            '-------DJL-07-24-2025      'Do not do this anymore was just done as a quick fix.
                        'Else

                        BOMList(5, UBound(BOMList, 2)) = GetPartNo
                        'End If

                        BOMList(6, UBound(BOMList, 2)) = GetShipDesc
                        BOMList(7, UBound(BOMList, 2)) = GetDesc

                        If GetInv1 <> "" And GetInv2 <> "" Then
                            BOMList(8, UBound(BOMList, 2)) = "Yes"            '-------DJL-07-03-2025      'BOMList(8, UBound(BOMList, 2)) = GetShipStds
                        End If

                        BOMList(9, UBound(BOMList, 2)) = GetInv1
                        BOMList(10, UBound(BOMList, 2)) = GetInv2

                        'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                        GetMat = LTrim(GetMat)          'Below has a problem due to someone putting spaces in the values
                        GetMat2 = LTrim(GetMat2)
                        GetMat3 = LTrim(GetMat3)

                        'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                        If IsNothing(GetMat) = False And GetMat <> "" Then
UpdateMat2:
                            If GetMat = " " Or GetMat = "  " Then        'Found material m = " " had a space   "Question if someone is using space to blank out data using replace?
                                If (GetMat2) <> "" And GetMat3 <> "" Then
                                    GoTo UpdateMat4
                                End If
                            End If

                            BOMList(11, UBound(BOMList, 2)) = GetMat
                        Else
UpdateMat4:
                            If (GetMat2) <> "" And GetMat3 <> "" Then
                                BOMList(11, UBound(BOMList, 2)) = (GetMat2 & " " & GetMat3)             'Trevor requested to have the dash between Material 1 and 2 to be removed.  'BOMList(11, UBound(BOMList, 2)) = (GetMat2 & "-" & GetMat3)    'BOMList(11, UBound(BOMList, 2)) = (GetMat2 & "~" & GetMat3)
                            Else
                                GoTo UpdateMat2
                            End If
                        End If

                        BOMList(13, UBound(BOMList, 2)) = GetLen
                        BOMList(14, UBound(BOMList, 2)) = GetWt
                        'BOMList(15, UBound(BOMList, 2)) = CurrentDwgNo          '-------DJL-07-03-2025  'Not required same as BOMList(1, i)
                        BOMList(17, UBound(BOMList, 2)) = InsertionPT(1).ToString
                        BOMList(18, UBound(BOMList, 2)) = GetProc

                        DecPos = Nothing
                        InsPt0 = InsertionPT(0).ToString
                        DecPos = InStr(InsPt0, ".")
                        If DecPos = 2 Then
                            InsPt0 = "0" & InsPt0
                        End If

                        DecPos = Nothing
                        InsPt1 = InsertionPT(1).ToString
                        DecPos = InStr(InsPt1, ".")
                        If DecPos = 2 Then
                            InsPt1 = "0" & InsPt1
                        End If

                        BOMList(19, UBound(BOMList, 2)) = (InsPt0)      'Required for Shipping List.

                        'BOMList(0, UBound(BOMList, 2)) = InsertionPT(0).ToString     '-------DJL-07-03-2025      'Not required.
                        '-------DJL-07-03-2025      'Not required.
                        'DrawingIndex = InsertionPT(1).ToString     '-------DJL-06-27-2025       'DrawingIndex = CurrentDwgNo & "-" & InsertionPT(1).ToString & "-" & InsertionPT(0).ToString
                        'DwgIndex = DrawingIndex
                        'BOMList(21, UBound(BOMList, 2)) = DwgIndex              'DrawingIndex
                        ReDim Preserve BOMList(21, UBound(BOMList, 2) + 1)

                            '-------DJL-06-27-2025      'Still having issues on sort, Because we are looking at items per drawing we can remove some of the complexity.
                            BOMListSort.Add(InsertionPT(1).ToString)     '-------DJL-06-27-2025       'BOMListSort.Add(DrawingIndex)  '-------DJL 06-27-2025      'Modified record zero to = 0 above
NextBOMItem:
                            GetPartNo = Nothing
                            Get2DShipMk = Nothing
                            Get2DShipQty = Nothing
                            GetPartNo = Nothing
                            GetQty = Nothing
                            GetShipDesc = Nothing
                            GetDesc = Nothing
                            GetLen = Nothing
                            GetNotes = Nothing
                            GetMat = Nothing    'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                            GetMat2 = Nothing
                            GetMat3 = Nothing
                            MX2 = Nothing
                            GetShipStds = Nothing
                            GetNotes = Nothing
                            GetWt = Nothing
                        Next BomItem
                    End If
                    'End If

                    PrgName = "StartButton_Click-Part12"

                    ProblemAt = "CloseDwg"
                    CountVal = (CountVal + 1)
                    ProgressBar1.Value = CountVal
                    AcadDoc.Close(SaveChanges:=False)        'AcadDoc.Close()        Changed AcadDoc from Object to AcadDocument type    RW 8/17/2023
                    ProblemAt = ""
NextDwg:
                    GetPartNo = Nothing
                    Get2DShipMk = Nothing
                    Get2DShipQty = Nothing
                    GetPartNo = Nothing
                    GetQty = Nothing
                    GetShipDesc = Nothing
                    GetDesc = Nothing
                    GetLen = Nothing
                    GetNotes = Nothing
                    GetMat = Nothing        'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                    GetMat2 = Nothing
                    GetMat3 = Nothing
                    GetNotes = Nothing
                    GetWt = Nothing

                BOMListSort.Sort()          '-------DJL-07-07-2025      'Must be done Sorts from 0 to 99.000000 'Below form 99.0000000 to 0
                BOMListSort.Reverse()   '-------DJL-06-27-2025      'For some reason this does not resort from position 0 but from 1 to 2
                    Dim Val1, Val2 As String        '-------DJL-07-03-2025      'Dim Val1, Val2 As Single
                    Dim Val1Dbl, Val2Dbl As Double
                    Dim ValChg As String
ChkSortAgain:                                       '-------DJL-07-28-2025  'Added
                For v = 0 To (BOMListSort.Count - 1)
                        Val1 = BOMListSort(v)
                        Val1Dbl = Val1

                        If v = (BOMListSort.Count - 1) Then
                            GoTo SortNoDone
                        End If

                        Val2 = BOMListSort(v + 1)
                        Val2Dbl = Val2
                        ValChg = 0

                        If Val1Dbl < Val2Dbl Then             'If Val1 < Val2 Then            'If BOMListSort(v) < BOMListSort(v + 1) Then    'If BOMListSort(0) < BOMListSort(1) Then
                            ValChg = BOMListSort(v)                 'BOMListSortChg = BOMListSort(v)    'BOMListSortChg = BOMListSort(0)
                            BOMListSort(v) = BOMListSort(v + 1)     'BOMListSort(0) = BOMListSort(1)
                        BOMListSort(v + 1) = ValChg             'BOMListSort(v + 1) = BOMListSortChg    'BOMListSort(1) = BOMListSortChg
                        ChkSort = "Found"                   '-------DJL-07-28-2025  'Added
                    End If
                    Next v

SortNoDone:
                If ChkSort = "Found" Then                   '-------DJL-07-28-2025  'Added
                    ChkSort = ""
                    GoTo ChkSortAgain
                End If
                y = 1

                    For x = 0 To (BOMListSort.Count)                           'For x = 0 To (BOMListSort.Count - 1)                                  '-------DJL-06-09-2025
FindNextItem:
                    If BOMListSort.Count <> 0 Then
                        FindIndex = BOMListSort(0)
                    Else
                        'If z < CntDwgs Then                         '-------DJL-07-03-2025      'If all of the drawings have not been found.
                        '    GoTo Nextz
                        'Else
                        GoTo SortFinished
                        'End If
                    End If

                        x = y
                        CntBOMList = UBound(BOMList, 2)

                        If y > CntBOMList And BOMListSort.Count > 0 Then        '--------DJL-07-03-2025     'Need to reset y when not all of the parts where found.
                            x = 1
                        End If

                        For y = x To UBound(BOMList, 2)                     'For y = 1 To UBound(BOMList)     '-------DJL-06-09-2025
                            FoundIndex = BOMList(17, (CntBOMList - y))      '-----Look at sort  '-------DJL-07-02-2025      'FoundIndex = BOMList(21, (CntBOMList - y))

                        If FindIndex = FoundIndex Then                                                      '-------DJL-06-09-2025
                            BOMListNew(1, UBound(BOMListNew, 2)) = BOMList(1, (CntBOMList - y))     'CurrentDwgNo '-------DJL-06-27-2025      'BOMListNew(1, UBound(BOMListNew, 2)) = BOMList(1, y)
                            BOMListNew(2, UBound(BOMListNew, 2)) = BOMList(2, (CntBOMList - y))     'CurrentDwgRev
                            BOMListNew(3, UBound(BOMListNew, 2)) = BOMList(3, (CntBOMList - y))     'Get2DShipMk
                            BOMListNew(4, UBound(BOMListNew, 2)) = BOMList(4, (CntBOMList - y))     'GetQty
                            BOMListNew(5, UBound(BOMListNew, 2)) = BOMList(5, (CntBOMList - y))     'GetPartNo
                            BOMListNew(6, UBound(BOMListNew, 2)) = BOMList(6, (CntBOMList - y))     'GetShipDesc
                            BOMListNew(7, UBound(BOMListNew, 2)) = BOMList(7, (CntBOMList - y))     'GetDesc
                            BOMListNew(8, UBound(BOMListNew, 2)) = BOMList(8, (CntBOMList - y))     'Yes or No Stds Found   'Not Required.
                            BOMListNew(9, UBound(BOMListNew, 2)) = BOMList(9, (CntBOMList - y))     'GetInv1 or Standard Part No

                            If BOMList(9, (CntBOMList - y)) <> "" And IsNothing(BOMList(9, (CntBOMList - y))) = False Then
                                STDsList(1, UBound(STDsList, 2)) = BOMList(1, (CntBOMList - y))     'CurrentDwgNo       '-------DJL-07-07-2025
                                STDsList(2, UBound(STDsList, 2)) = BOMList(2, (CntBOMList - y))     'CurrentDwgRev
                                STDsList(3, UBound(STDsList, 2)) = BOMList(3, (CntBOMList - y))     'Get2DShipMk
                                STDsList(4, UBound(STDsList, 2)) = BOMList(4, (CntBOMList - y))     'GetQty
                                STDsList(5, UBound(STDsList, 2)) = BOMList(5, (CntBOMList - y))     'GetPartNo
                                STDsList(6, UBound(STDsList, 2)) = BOMList(6, (CntBOMList - y))     'GetShipDesc
                                STDsList(7, UBound(STDsList, 2)) = BOMList(7, (CntBOMList - y))     'GetDesc
                                'STDsList(8, UBound(STDsList, 2)) = BOMList(8, (CntBOMList - y))     'Yes or No Stds Found   'Not Required.
                                STDsList(9, UBound(STDsList, 2)) = BOMList(9, (CntBOMList - y))     'GetInv1 or Standard Part No
                                STDsList(10, UBound(STDsList, 2)) = BOMList(10, (CntBOMList - y))   'GetInv2 or Standard Drawings
                                STDsList(11, UBound(STDsList, 2)) = BOMList(11, (CntBOMList - y))   'GetMat
                                'STDsList(12, UBound(STDsList, 2)) = BOMList(12, (CntBOMList - y))   'Does Not Exist
                                'STDsList(13, UBound(STDsList, 2)) = BOMList(13, (CntBOMList - y))   'GetLen        'Not required for AutoCAD.
                                STDsList(14, UBound(STDsList, 2)) = BOMList(14, (CntBOMList - y))   'GetWt
                                STDsList(17, UBound(STDsList, 2)) = BOMList(17, (CntBOMList - y))   'InsertionPT(1)
                                STDsList(18, UBound(STDsList, 2)) = BOMList(18, (CntBOMList - y))   'GetProc
                                STDsList(19, UBound(STDsList, 2)) = BOMList(19, (CntBOMList - y))   'InsertionPT(0) required Shipping List.     '-------DJL-07-07-2025
                                RowNoPlus4 = (UBound(BOMListNew, 2) + 4)
                                STDsList(20, UBound(STDsList, 2)) = RowNoPlus4                          '-------DJL-07-07-2025      'Record number or row number
                                STDsList(20, UBound(STDsList, 2)) = UBound(BOMListNew, 2)               '-------DJL-07-07-2025      'Record number or row number

                                'STDsList(1, UBound(STDsList, 2)) = BOMList(9, (CntBOMList - y))        '-------DJL-06-09-2025 collect standards
                                'STDsList(2, UBound(STDsList, 2)) = BOMList(10, (CntBOMList - y))
                                ReDim Preserve STDsList(20, UBound(STDsList, 2) + 1)                    '-------DJL-07-07-2025
                            End If

                            BOMListNew(10, UBound(BOMListNew, 2)) = BOMList(10, (CntBOMList - y))       'GetInv2 or Standard Drawings
                            BOMListNew(11, UBound(BOMListNew, 2)) = BOMList(11, (CntBOMList - y))       'Material
                            'BOMListNew(12, UBound(BOMListNew, 2)) = BOMList(12, (CntBOMList - y))       'Does Not Exist
                            'BOMListNew(13, UBound(BOMListNew, 2)) = BOMList(13, (CntBOMList - y))       'GetLen        'Not required for AutoCAD.
                            BOMListNew(14, UBound(BOMListNew, 2)) = BOMList(14, (CntBOMList - y))       'GetWt
                            'BOMListNew(15, UBound(BOMListNew, 2)) = BOMList(15, (CntBOMList - y))       '-------DJL-07-03-2025      'Not required.
                            'BOMListNew(16, UBound(BOMListNew, 2)) = BOMList(16, (CntBOMList - y))       '-------DJL-07-03-2025      'Not required.
                            BOMListNew(17, UBound(BOMListNew, 2)) = BOMList(17, (CntBOMList - y))          'InsertionPT(1)      '-------DJL-07-03-2025
                            BOMListNew(18, UBound(BOMListNew, 2)) = BOMList(18, (CntBOMList - y))       'GetProc
                            BOMListNew(19, UBound(BOMListNew, 2)) = BOMList(19, (CntBOMList - y))       ''InsertionPT(0) required Shipping List.        '-------DJL-07-03-2025      'Required for Shipping List sort.       
                            BOMListNew(20, UBound(BOMListNew, 2)) = UBound(BOMListNew, 2)           'Record Number or Row Number.     '-------DJL-07-03-2025        'BOMListNew(20, UBound(BOMListNew, 2)) = UBound(BOMListNew, 2)
                            'BOMListNew(21, UBound(BOMListNew, 2)) = BOMList(21, (CntBOMList - y))       '-------DJL-07-03-2025      'Not Required.
                            ReDim Preserve BOMListNew(21, UBound(BOMListNew, 2) + 1)

                            BOMListSort.RemoveAt(0)
                            y = (y + 1)

                            If BOMListSort.Count > 0 Then
                                GoTo FindNextItem
                            End If

                            GoTo Nextx
                        End If
                    Next y
Nextx:
                        If BOMListSort.Count > 0 Then
                            GoTo FindNextItem
                        End If
                    Next x
Nextz:
                'If z < CntDwgs Then
                '    z = (z + 1)
                '    GoTo FindNextDwg
                'End If
SortFinished:
            Next z                           'Next DwgItem

                PrgName = "StartButton_Click-Part13"

            'SortFinished:                          '-------DJL-07-03-2025      'moved above
            WriteToExcel(BOMListNew)     '-------DJL-06-09-2025          'Above removes most of this.

            '-------DJL-07-07-2025     'Below is fixed above as the Array is collected.
            'UpdateShpMarksArray(BOMListNew)     '-------DJL-06-10-2025     'UpdateShpMarksArray(BOMList)    '-------DJL-06-09-2025
            PrgName = "StartButton_Click-Part14"
        End If

        WorkShtName = "Stds BOM"         'Move to Shipping List    '-------DJL-06-06-2025
        Workbooks = ExcelApp.Workbooks
        StdsWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        StdsWrkSht.Activate()

        With StdsWrkSht
            .Range("B3").Value = FullJobNo
            .Range("I3").Value = Today
            .Range("F3").Value = Me.ComboBxRev.Text
        End With

        FileToOpen = "Stds BOM"
        PrgName = "StartButton_Click-Part15"

        For i = 1 To (UBound(STDsList, 2) - 1)
            RowNo = i + 4
            ProgressBar1.Maximum = (UBound(STDsList, 2) - 1)
            StdsWrkSht.Activate()

            If RowNo = "5" Then
                FormatLine(RowNo, FileToOpen)

                With StdsWrkSht
                    .Rows(RowNo & ":" & RowNo).Select()
                    .Rows(RowNo & ":" & RowNo).Insert()             '-------DJL-07-07-2025
                End With
                LineNo = RowNo
            Else
                With StdsWrkSht
                    .Rows(RowNo & ":" & RowNo).Select()
                    .Rows(RowNo & ":" & RowNo).Insert()             '-------DJL-07-07-2025
                    LineNo = RowNo
                End With
            End If

            '---------------------------------------Add Data on Sheet STDs BOM
            With StdsWrkSht
                .Range("A" & RowNo).Value = STDsList(1, i)      'Dwg No
                .Range("B" & RowNo).Value = STDsList(2, i)      'Rev No
                .Range("C" & RowNo).Value = STDsList(3, i)      'Ship Mk
                .Range("E" & RowNo).Value = STDsList(4, i)      'Qty
                .Range("D" & RowNo).Value = STDsList(5, i)      'Part No
                .Range("F" & RowNo).Value = STDsList(6, i)      'Desc                      '-------DJL-07-07-2025      '.Range("I" & RowNo).Value = STDsList(6, i)
                '.Range("F" & RowNo).Value = STDsList(7, i)      'Desc
                '.Range("H" & RowNo).Value = STDsList(8, i)     'Not required just "Yes" for when Standards is found.
                .Range("G" & RowNo).Value = STDsList(9, i)          'Std Part Number
                .Range("H" & RowNo).Value = STDsList(10, i)         'Standard number MX0104A
                .Range("U" & RowNo).Value = STDsList(10, i)         'In Order to update all standards found on BOM      'DJL-12-29-2023

                .Range("I" & RowNo).Value = STDsList(11, i)         'Material
                .Range("J" & RowNo).Value = STDsList(14, i)         'Weight
                .Range("M" & RowNo).Value = STDsList(1, i)          'Dwg No     '-------DJL-07-07-2025        STDsList(15, i)
                '.Range("N" & RowNo).Value = STDsList(16, i)        'Blank      '-------DJL-07-07-2025
                .Range("O" & RowNo).Value = STDsList(19, i)         'InsertionPt(0)    '-------DJL-07-07-2025     '.Range("O" & RowNo).Value = STDsList(0, i)
                .Range("P" & RowNo).Value = STDsList(17, i)         'InsertionPt(1)     '-------DJL-07-07-2025
                '.Range("X" & RowNo).Value = STDsList(18, i)        'Blank      '-------DJL-07-07-2025
                '.Range("Y" & RowNo).Value = STDsList(19, i)        'Nothing    '-------DJL-07-07-2025
                .Range("AA" & RowNo).Value = STDsList(20, i)         'Record no or Row No        '-------DJL-07-07-2025     '.Range("Z" & RowNo).Value = STDsList(20, i) 
            End With
            ProgressBar1.Value = i
        Next i

        Dim StdsBOMList(20, 1)
        LineNo3 = StdsWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        With StdsWrkSht
            With .Range("A4:Z" & LineNo3)
                .Sort(Key1:= .Range("H5"), Order1:=XlSortOrder.xlAscending, Key2:= .Range("G5"), Order2:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=XlSortOrientation.xlSortColumns)
            End With
        End With

        PrgName = "StartButton_Click-Part16"

        Dim StdsFnd(1, 1)

        With StdsWrkSht
            ExistSTD = .Range("H" & 5).Value
            StdsFnd(0, UBound(StdsFnd, 2)) = ExistSTD
            ReDim Preserve StdsFnd(1, UBound(StdsFnd, 2) + 1)
        End With

        With StdsWrkSht
            For c = (1 + 5) To LineNo3
                If .Range("H" & c).Value = ExistSTD Then
                    .Range("H" & c).Value = ""
                Else
                    ExistSTD = .Range("H" & c).Value
                    StdsFnd(0, UBound(StdsFnd, 2)) = ExistSTD
                    ReDim Preserve StdsFnd(1, UBound(StdsFnd, 2) + 1)
                End If
            Next c
        End With

        PrgName = "StartButton_Click-Part17"
        '----------Resort to record order number descending

        With StdsWrkSht
            With .Range("A4:Z" & LineNo3)
                .Sort(Key1:= .Range("Z5"), Order1:=XlSortOrder.xlDescending, Header:=XlYesNoGuess.xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=XlSortOrientation.xlSortColumns)
            End With
        End With

        Dim StdsFnd2(2, 1)

        With StdsWrkSht
            For d = (1 + 4) To LineNo3                  'For d = (1 + 5) To LineNo3
                StdsFnd2(0, UBound(StdsFnd2, 2)) = .Range("G" & d).Value
                StdsFnd2(1, UBound(StdsFnd2, 2)) = .Range("U" & d).Value                    'StdsFnd2(1, UBound(StdsFnd2, 2)) = .Range("H" & d).Value
                StdsFnd2(2, UBound(StdsFnd2, 2)) = .Range("Z" & d).Value
                ReDim Preserve StdsFnd2(2, UBound(StdsFnd2, 2) + 1)
            Next d
        End With

        GenInfo3233.StdsFnd2 = StdsFnd2
        Me.LblProgress.Text = "Making sure all Standards are found in user Directory........Please Wait"
        Me.Refresh()
        CountVal = 0
        ProgressBar1.Maximum = LineNo3
        Dim l, CntDwgsNotFound As Integer
        Dim TestForYes, DwgsFound, DwgsFound2 As String

        PrgName = "StartButton_Click-Part18"

        For l = 1 To (UBound(StdsFnd, 2) - 1)                            'For l = 1 To LineNo3
            'With StdsWrkSht
            LookForStd = StdsFnd(0, l)            '.Range("H" & (l + 4)).Value            'LookForStd = .Range("N" & (l + 4)).Value

            If LookForStd = "" Or LookForStd = Nothing Or LookForStd = " " Then             '-------DJL-08-08-2025      'If LookForStd = "" Or LookForStd = Nothing Then
                GoTo Nextl
            End If

            AdeptPrg = "\\Adept01\Tulsa\STD\PROGRAM\"
            AdeptStd = "\\Adept01\Tulsa\STD\TANKSTANDARDS\"
            Dwg = ".dwg"

            If Dir(PathBox.Text & FullJobNo & "_" & LookForStd & Dwg) <> "" Then
                DwgItem = PathBox.Text & FullJobNo & "_" & LookForStd & Dwg
            Else
                If Dir(PathBox.Text & FullJobNo & "._" & LookForStd & Dwg) <> "" Then
                    DwgItem = PathBox.Text & FullJobNo & "._" & LookForStd & Dwg
                Else
                    If Dir(PathBox.Text & FullJobNo & "-" & LookForStd & Dwg) <> "" Then
                        DwgItem = PathBox.Text & FullJobNo & "-" & LookForStd & Dwg
                    Else
                        If Dir(PathBox.Text & FullJobNo & ".-" & LookForStd & Dwg) <> "" Then
                            DwgItem = PathBox.Text & FullJobNo & ".-" & LookForStd & Dwg
                        Else
                            Me.DwgsNotFoundList.Items.Add(FullJobNo & "._" & LookForStd & Dwg)
                        End If
                    End If
                End If
            End If
Nextl:
        Next l

        PrgName = "StartButton_Click-Part19"
        CntDwgsNotFound = Me.DwgsNotFoundList.Items.Count

        If CntDwgsNotFound > 0 Then                                 '-----------------------Notify user of Standards missing.
            DwgsFound2 = Nothing

            For l = 0 To (CntDwgsNotFound - 1)
                If CntDwgsNotFound > 0 Then
                    DwgsFound = Me.DwgsNotFoundList.Items(l)
                End If

                If l = 0 Then
                    DwgsFound2 = DwgsFound
                Else
                    DwgsFound2 = DwgsFound2 & ", " & Chr(10) & DwgsFound
                End If
            Next l

            MissingTxt1 = "The following Standards are Missing" & Chr(10) & DwgsFound2 & Chr(10) & "Do you want to Continue?    Type   Yes or No" & Chr(10)
            MissingTxt2 = Chr(10) & "If you copy the standards to this directory," & Chr(10) & PathBox.Text & Chr(10) & " Then type Yes in the Box below the program will contine without having to be restarted."

            Sapi.Speak("Some Standard Drawings are missing, If you copy the files to your working Directory, then type Yes in the box below the program will continue without having to be restarted.")
            TestForYes = InputBox(MissingTxt1 & MissingTxt2)            ' & MissingTxt3)

            Select Case TestForYes
                Case "Yes"
                    GoTo Continue_Dwgs
                Case "YES"
                    TestForYes = "Yes"
                    GoTo Continue_Dwgs
                Case "Y"
                    TestForYes = "Yes"
                    GoTo Continue_Dwgs
                Case "y"
                    TestForYes = "Yes"
                    GoTo Continue_Dwgs
                Case "No"
                    Sapi.Speak("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    MsgBox("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    GoTo Cancel
                Case "NO"
                    Sapi.Speak("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    MsgBox("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    GoTo Cancel
                Case "N"
                    Sapi.Speak("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    MsgBox("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    GoTo Cancel
                Case "n"
                    Sapi.Speak("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    MsgBox("After you have put your Standard Drawings in your working directory, you will need to rerun the program.")
                    GoTo Cancel
                Case Else
                    Sapi.Speak("You must enter Yes or No, you will need to rerun the program.")
                    MsgBox("You must enter Yes or No, you will need to rerun the program.")
                    GoTo Cancel
            End Select
        End If

Continue_Dwgs:  '------------------------------Get Standard drawing information for BOM Items.
        PrgName = "StartButton_Click-Part20"
        Me.LblProgress.Text = "Opening Matrix Standard Drawings........Please Wait"
        Me.Refresh()

        CountVal = 0
        ProgressBar1.Maximum = LineNo3
        CntDwgs = (UBound(StdsFnd, 2) - 1)

        For j = 1 To (UBound(StdsFnd, 2) - 1)
            LookForStd = StdsFnd(0, j)

            If LookForStd = "" Or LookForStd = " " Then         '-------DJL-08-08-2025      'If LookForStd = "" Then
                GoTo NextDwg2
            End If

            AdeptPrg = "\\Adept01\Tulsa\STD\PROGRAM\"
            AdeptStd = "\\Adept01\Tulsa\STD\TANKSTANDARDS\"
            Dwg = ".dwg"

            PrgName = "StartButton_Click-Part21"

            If Dir(PathBox.Text & FullJobNo & "_" & LookForStd & Dwg) <> "" Then
                DwgItem = PathBox.Text & FullJobNo & "_" & LookForStd & Dwg
            Else
                If Dir(PathBox.Text & FullJobNo & "._" & LookForStd & Dwg) <> "" Then
                    DwgItem = PathBox.Text & FullJobNo & "._" & LookForStd & Dwg
                Else
                    If Dir(PathBox.Text & FullJobNo & "-" & LookForStd & Dwg) <> "" Then
                        DwgItem = PathBox.Text & FullJobNo & "-" & LookForStd & Dwg
                    Else
                        If Dir(PathBox.Text & FullJobNo & ".-" & LookForStd & Dwg) <> "" Then
                            DwgItem = PathBox.Text & FullJobNo & ".-" & LookForStd & Dwg
                        Else
                            GoTo NextDwg2
                        End If
                    End If
                End If
            End If

            'End With

            AcadApp.Documents.Open(DwgItem)
            System.Threading.Thread.Sleep(50)
            AcadApp.Visible = False
            Me.Refresh()
            AcadDoc = AcadApp.ActiveDocument

            BlockSel = AcadDoc.SelectionSets.Add("Titleblock")
            GroupCode(0) = 0
            BlockData(0) = "INSERT"
            GroupCode(1) = 2
            BlockData(1) = "AMW_TITLE,OSF_TITLE,OSF_TITLE_D,MX_TITLE,LNG_TITLE_D,MX_TITLE_SP,Title Blocks Matrix"
            BlockSel.Select(AutoCAD.AcSelect.acSelectionSetAll, , , GroupCode, BlockData)

            Temparray = BlockSel.Item(0).GetAttributes
            CntAttFound = 0

            For i = 0 To UBound(Temparray)
                Select Case Temparray(i).TagString
                    Case "DN"
                        CurrentDwgNo = Temparray(i).TextString
                        CntAttFound = (CntAttFound + 1)
                    Case "RN"
                        CurrentDwgRev = Temparray(i).TextString
                        CntAttFound = (CntAttFound + 1)
                    Case "SN"
                        CurrentStdNo = Temparray(i).TextString
                        CntAttFound = (CntAttFound + 1)
                End Select

                If CntAttFound > 2 Then
                    GoTo FoundAtts                      'Why look for all when only three are needed.-------DJL-12-29-2023
                End If
            Next i

FoundAtts:
            If IsNothing(CurrentDwgNo) = True Or CurrentDwgNo = "" Then
                Sapi.Speak("You have a Drawing with no drawing number, Please check your drawings for missing drawing Numbers.")
                MsgBox("You have a Drawing with no drawing number, Please check your drawings for missing drawing Numbers.")

                Sapi.Speak("This standard " & CurrentStdNo & " needs to have a drawing number fixed before this information will be put on BOM List.")
                MsgBox("This standard " & CurrentStdNo & " needs to have the drawing number fixed before this information will be put on the BOM List.")

                Sapi.Speak("Bulk Bom List will show Standard Not Found for" & CurrentStdNo & ".")
                MsgBox("Bulk Bom List will show Standard Not Found for" & CurrentStdNo & ".")
                GoTo NextDwg2
            End If

            PrgName = "StartButton_Click-Part22"
            BlockSel = AcadDoc.SelectionSets.Add("BillOfMaterial")
            GroupCode(0) = 0
            BlockData(0) = "INSERT"
            GroupCode(1) = 2
            BlockData(1) = "STANDARD_BILL_OF_MATERIAL,B_BILL_OF_MATERIAL,SP_BILL_OF_MATERIAL"
            BlockSel.Select(AutoCAD.AcSelect.acSelectionSetAll, , , GroupCode, BlockData)

            Me.TxtBoxDwgsToProcess.Text = CntDwgs - j
            CntCollected = 0

            If BlockSel.Count <> 0 Then
                For Each BomItem In BlockSel
                    CntCollected = (CntCollected + 1)
                    Me.TxtBoxBOMItemsToProcess.Text = BlockSel.Count - CntCollected
                    Me.Refresh()

                    BOMItemNam = BomItem.Name
                    TempAttributes = BomItem.GetAttributes

                    '-------Get Items from Standards sheet were Tag D2 equal to Description
                    'NTest1 = TempAttributes(0).TextString                 'Mark1
                    'NTest2 = TempAttributes(1).TextString                 'Qty1
                    'NTest3 = TempAttributes(2).TextString                 'Mark2
                    'Ntest4 = TempAttributes(3).TextString                 'Qty2
                    NTest5 = TempAttributes(4).TextString                 'Description1
                    Ntest6 = TempAttributes(5).TextString                 'Description2 -Centered
                    'NTest7 = TempAttributes(6).TextString                 'Inv-1
                    'NTest8 = TempAttributes(7).TextString                 'Inv -2
                    'NTest9 = TempAttributes(8).TextString                 'Material1
                    'NTest10 = TempAttributes(9).TextString                 'Material2A
                    'NTest11 = TempAttributes(10).TextString                'Material2B
                    'NTest12 = TempAttributes(11).TextString                'Weight
                    ''Test = TempAttributes(12).TextString               'Only 11 Items exist for BOM Items.
                    ''Test = TempAttributes(13).TextString
                    ''Test = TempAttributes(14).TextString

                    If NTest5 = Nothing Or NTest5 = "" Then                 'New problem remove blank lines....
                        If Ntest6 = Nothing Or Ntest6 = "" Then
                            GoTo NextBOMItem2
                        End If
                    End If

                    PrgName = "StartButton_Click-Part23"

                    For r = 0 To UBound(TempAttributes)                         'Not Required Found Job Information Previuosly 'Need info for Standards Dwg Number.
                        TestTags = TempAttributes(r).TagString

                        Select Case TempAttributes(r).TagString
                            Case "SLM"                                          '-------MK"
                                Get2DShipMk = TempAttributes(r).TextString
                            Case "SLQ"                                          '-------MK"
                                Get2DShipQty = TempAttributes(r).TextString
                            Case "SM"                                           '-------SP"
                                GetPartNo = TempAttributes(r).TextString
                            Case "Q"
                                GetQty = TempAttributes(r).TextString
                            Case "SD"
                                GetShipDesc = TempAttributes(r).TextString
                            Case "D"
                                GetDesc = TempAttributes(r).TextString
                            Case "D2"
                                GetShipDesc = TempAttributes(r).TextString
                            Case "IU"
                                GetInv1 = TempAttributes(r).TextString
                            Case "IL"
                                GetInv2 = TempAttributes(r).TextString
                            'Case "L"                                            '-------DJL-07-07-2025      'Not Required for AutoCAD.
                            '    GetLen = TempAttributes(r).TextString
                            Case "M"        'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                                GetMat = TempAttributes(r).TextString

                                If InStr(GetMat, "NOTE") > 0 Then
                                    NotePos = InStr(GetMat, "NOTE")
                                    GetNotes = Mid(GetMat, NotePos, Len(GetMat))
                                    GetMat = Mid(GetMat, 1, (NotePos - 1))
                                End If
                            Case "M2"       'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                                GetMat2 = TempAttributes(r).TextString
                            Case "M3"       'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                                GetMat3 = TempAttributes(r).TextString
                            Case "N"
                                GetNotes = TempAttributes(r).TextString
                            Case "W"
                                GetWt = TempAttributes(r).TextString
                        End Select
                    Next r

                    PrgName = "StartButton_Click-Part24"

                    StdsBOMList(1, UBound(StdsBOMList, 2)) = CurrentDwgNo
                    StdsBOMList(2, UBound(StdsBOMList, 2)) = CurrentDwgRev
                    StdsBOMList(3, UBound(StdsBOMList, 2)) = Get2DShipMk
                    StdsBOMList(4, UBound(StdsBOMList, 2)) = GetQty
                    StdsBOMList(5, UBound(StdsBOMList, 2)) = GetPartNo
                    StdsBOMList(6, UBound(StdsBOMList, 2)) = GetShipDesc            'Assembly Description
                    StdsBOMList(7, UBound(StdsBOMList, 2)) = GetDesc                'Part Description
                    StdsBOMList(9, UBound(StdsBOMList, 2)) = GetInv1
                    StdsBOMList(10, UBound(StdsBOMList, 2)) = GetInv2

                    'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                    If IsNothing(GetMat) = False And GetMat <> "" Then
UpdateMat3:
                        StdsBOMList(11, UBound(StdsBOMList, 2)) = GetMat
                    Else
                        If GetMat2 <> "" And GetMat3 <> "" Then
                            StdsBOMList(11, UBound(StdsBOMList, 2)) = (GetMat2 & " " & GetMat3)         'Requested by Trevor Ruffin  'StdsBOMList(11, UBound(StdsBOMList, 2)) = (GetMat2 & "-" & GetMat3)   'StdsBOMList(11, UBound(StdsBOMList, 2)) = (GetMat2 & "~" & GetMat3)
                        Else
                            If GetMat2 <> "" And GetMat3 = "" Then
                                StdsBOMList(11, UBound(StdsBOMList, 2)) = GetMat2
                            Else
                                If GetMat2 = "" And GetMat3 <> "" Then
                                    StdsBOMList(11, UBound(StdsBOMList, 2)) = GetMat3
                                Else
                                    GoTo UpdateMat3
                                End If
                            End If
                        End If
                    End If

                    'StdsBOMList(13, UBound(StdsBOMList, 2)) = GetLen            '-------DJL-07-07-2025      'Not Required for AutoCAD.
                    StdsBOMList(14, UBound(StdsBOMList, 2)) = GetWt

BOMInfoCollected2:
                    PrgName = "StartButton_Click-Part25"
                    InsertionPT = BomItem.InsertionPoint
                    Dimscale = BomItem.XScaleFactor
                    CompareX1 = 10.5 * Dimscale                         '-------DJL-07-03-2025      'Is required for sort below     'Not Required.
                    CompareX1 = InsertionPT(0) - CompareX1
                    CompareX1 = CompareX1 / Dimscale

                    CompareX2 = 6 * Dimscale                           '-------DJL-07-23-2025      'Not Required.
                    CompareX2 = InsertionPT(0) - CompareX2
                    CompareX2 = CompareX2 / Dimscale

                    If CompareX1 < 1 Or CompareX2 < 1 Then
                        If CompareX1 > 0 Or CompareX2 > 0 Then
                            StdsBOMList(16, UBound(StdsBOMList, 2)) = CStr(1)
                        Else
                            StdsBOMList(16, UBound(StdsBOMList, 2)) = CStr(2)
                        End If
                    Else
                        StdsBOMList(16, UBound(StdsBOMList, 2)) = CStr(2)
                    End If

                    StdsBOMList(17, UBound(StdsBOMList, 2)) = InsertionPT(1)
                    StdsBOMList(18, UBound(StdsBOMList, 2)) = InsertionPT(0)
                    'StdsBOMList(15, UBound(StdsBOMList, 2)) = CurrentDwgNo             '-------DJL-07-03-2025      'same as StdsBOMList(1
                    ReDim Preserve StdsBOMList(20, UBound(StdsBOMList, 2) + 1)
NextBOMItem2:
                Next BomItem
            End If

            CountVal = (CountVal + 1)
            ProgressBar1.Value = j
            ProblemAt = "CloseDwg"
            AcadDoc.Close()
            ProblemAt = ""
NextDwg2:
        Next j                                    'Next DwgItem

        If Not IsNothing(gvntSDIvar) Then
            AcadPref.SingleDocumentMode = gvntSDIvar    'reset to sdi mode
        End If

        PrgName = "StartButton_Click-Part26"
        Me.LblProgress.Text = "Placing MX Standards List on BOM........Please Wait"
        Me.Refresh()

        '---------------------------------------------Create new sheet for Standards found.
        '-------DJL-07-08-2025      'Need to sort Array instead of writing to Spread sheet then sort, and write back to array.
        '-------Leave for now.
        '-----------------------------------------------------------------------------------------------------------------

        WorkShtName = "STD Items"
        StdItemsWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        StdItemsWrkSht.Activate()

        With StdItemsWrkSht
            .Range("B3").Value = FullJobNo
            .Range("I3").Value = Today
            .Range("F3").Value = Me.ComboBxRev.Text
        End With

        For i = 1 To (UBound(StdsBOMList, 2) - 1)
            PrgName = "StartButton_Click-Part27"
            RowNo = i + 4
            ProgressBar1.Maximum = (UBound(StdsBOMList, 2) - 1)
            FileToOpen = "STD Items"
            StdItemsWrkSht.Activate()

            If RowNo = "5" Then
                FormatLine(RowNo, FileToOpen)

                With StdItemsWrkSht
                    .Rows(RowNo & ":" & RowNo).Select()
                    .Rows(RowNo & ":" & RowNo).Insert()
                    LineNo = RowNo
                End With

                LineNo = RowNo
            Else
                With StdItemsWrkSht
                    .Rows(RowNo & ":" & RowNo).Select()
                    .Rows(RowNo & ":" & RowNo).Insert()
                    LineNo = RowNo
                End With
            End If

            With StdItemsWrkSht
                .Range("A" & RowNo).Value = StdsBOMList(1, i)
                .Range("B" & RowNo).Value = StdsBOMList(2, i)
                .Range("C" & RowNo).Value = StdsBOMList(3, i)
                .Range("D" & RowNo).Value = StdsBOMList(5, i)

                .Range("E" & RowNo).Value = StdsBOMList(4, i)

                '-------8 does not exist at this time, or is not used.

                If StdsBOMList(6, i) = vbNullString Or StdsBOMList(6, i) = " " Then
                    If InStr(1, StdsBOMList(7, i), "%%D") <> 0 Then
                        GetDesc = StdsBOMList(7, i)
                        GetDesc = GetDesc.Replace("%%D", " DEG.")
                        StdsBOMList(7, i) = GetDesc
                    End If
                    .Range("F" & RowNo).Value = StdsBOMList(7, i)
                Else
                    If InStr(1, StdsBOMList(6, i), "%%D") <> 0 Then
                        GetDesc = StdsBOMList(6, i)
                        GetDesc = GetDesc.Replace("%%D", " DEG.")
                        StdsBOMList(6, i) = GetDesc
                    End If
                    .Range("F" & RowNo).Value = StdsBOMList(6, i)
                End If

                '-------8 does not exist at this time, or is not used.

                .Range("G" & RowNo).Value = StdsBOMList(9, i)
                .Range("H" & RowNo).Value = StdsBOMList(10, i)
                .Range("I" & RowNo).Value = StdsBOMList(11, i)          '11 will always be material even when two types of material.
                .Range("J" & RowNo).Value = StdsBOMList(14, i)          'Weight
                .Range("M" & RowNo).NumberFormat = "@"
                .Range("M" & RowNo).Value = StdsBOMList(1, i)          'Standard No        '-------DJL-07-07-2025      ' .Range("M" & RowNo).Value = StdsBOMList(15, i)
                .Range("N" & RowNo).NumberFormat = "General"
                .Range("N" & RowNo).Value = StdsBOMList(16, i)          'one or two
                .Range("O" & RowNo).NumberFormat = "General"
                .Range("O" & RowNo).Value = StdsBOMList(17, i)          'Y point for insertion
                .Range("P" & RowNo).NumberFormat = "General"
                .Range("P" & RowNo).Value = StdsBOMList(18, i)          'X point for insertion
            End With
            ProgressBar1.Value = i
        Next i

        PrgName = "StartButton_Click-Part28"

        With StdItemsWrkSht
            With .Range("A4:P" & RowNo)
                .Sort(Key1:= .Range("M4"), Order1:=XlSortOrder.xlAscending, Key2:= .Range("N4"), Order2:=XlSortOrder.xlAscending, Key3:= .Range("O4"), Order3:=XlSortOrder.xlDescending, Header:=XlYesNoGuess.xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=XlSortOrientation.xlSortColumns)
            End With
        End With

        LineNo2 = StdItemsWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        RowNo = 5
        ReDim StdsBOMList(20, 1)

        '-----------------------------------------------------------------------------------------------------------------
        '-------DJL-07-08-2025      'Need to sort Array instead of writing to Spread sheet then sort, and write back to array.
        '-----------------------------------------------------------------------------------------------------------------

        For i = 1 To (LineNo2 - 4)
            With StdItemsWrkSht
                StdsBOMList(1, i) = .Range("A" & RowNo).Value           'MX-Std
                StdsBOMList(2, i) = .Range("B" & RowNo).Value           'Rev
                StdsBOMList(3, i) = .Range("C" & RowNo).Value           'Ship Mark
                StdsBOMList(5, i) = .Range("D" & RowNo).Value           'Piece Mark
                StdsBOMList(4, i) = .Range("E" & RowNo).Value           'Qty

                '-------6 does not exist at this time, or is not used.

                StdsBOMList(7, i) = .Range("F" & RowNo).Value           'Description

                '-------8 does not exist at this time, or is not used.

                StdsBOMList(9, i) = .Range("G" & RowNo).Value           'Std Part No = 2RR
                StdsBOMList(10, i) = .Range("H" & RowNo).Value          'MX0104A
                StdsBOMList(11, i) = .Range("I" & RowNo).Value         'material.
                StdsBOMList(14, i) = .Range("J" & RowNo).Value          'Weight
                StdsBOMList(15, i) = .Range("M" & RowNo).Value          'MX Std read MX0104A
                StdsBOMList(16, i) = .Range("N" & RowNo).Value          'one or two column one or two for BOM columns on drawings
                StdsBOMList(17, i) = .Range("O" & RowNo).Value         'Y point for insertion
                StdsBOMList(18, i) = .Range("P" & RowNo).Value          'X point for insertion
                ReDim Preserve StdsBOMList(20, UBound(StdsBOMList, 2) + 1)
                RowNo = (RowNo + 1)
            End With
            ProgressBar1.Value = i
        Next i

        PrgName = "StartButton_Click-Part29"
        GenInfo3233.StdsBOMList = StdsBOMList
        WorkShtName = "BulK BOM"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name
        BOMWrkSht.Activate()

        LineNo2 = BOMWrkSht.Range("B4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        PrgName = "StartButton_Click-Part30"

        WriteToExcelAfterSort(STDsList)           'Write new data to Spreadsheet          '-------DJL-07-08-2025      'WriteToExcelAfterSort(BOMListNew)
        PrgName = "StartButton_Click-Part31"
        AcadApp.WindowState = AutoCAD.AcWindowState.acMin
        LineNo2 = BOMWrkSht.Range("B4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        With BOMWrkSht
            .Range("G3").Value = Me.ComboBxRev.Text
            .Range("C3").Value = GenInfo3233.FullJobNo
            .Range("J3").Value = Today

            RevNo = Me.ComboBxRev.Text          '.Range("G3").Value
            z = 5
            .Range("X" & z & ":AB" & (LineNo2)).Delete()
            .Range("N1" & ":V" & (LineNo2)).Delete()
            .Range("A1" & ":A" & (LineNo2)).Delete()

            .Columns("A:A").EntireColumn.AutoFit
            .Columns("B:B").EntireColumn.AutoFit
            .Columns("C:C").EntireColumn.AutoFit
            '.Columns("D:D").EntireColumn.AutoFit
            .Columns("D:D").ColumnWidth = 12.0
            .Columns("E:E").EntireColumn.AutoFit
            .Columns("F:F").ColumnWidth = 49.0
            .Columns("F:F").EntireColumn.AutoFit
            .Columns("G:G").EntireColumn.AutoFit
            .Columns("H:H").EntireColumn.AutoFit
            .Columns("I:I").ColumnWidth = 13.0
            .Columns("I:I").EntireColumn.AutoFit
            .Columns("J:J").EntireColumn.AutoFit
            .Columns("M:M").ColumnWidth = 24.43
            .Columns("M:M").EntireColumn.AutoFit

            .Rows("5:" & LineNo2).RowHeight = 21.75
        End With

        PrgName = "StartButton_Click-Part32"
        'OldFileNam = Me.PathBox.Text          'Not Required.     '-------DJL-08-07-2025
        'FileToOpen = "K:\CAD\VBA\XLTSheets\BOM-New-1-15-2024.xltm"          'Not Required.     '-------DJL-08-07-2025      '-------DJL-11-10-2024          'FileToOpen = "K:\CAD\VBA\XLTSheets\BOM-New-1-15-2023.xltm"
        CopyBOMFile(OldFileNam, RevNo, ExcelApp)                          '-------DJL-08-08-2025         'Moved request by IT.
        ProgramFinished()
Cancel:

Err_StartButton_Click:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            DwgItem2 = CurrentDwgNo
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException
            Test = DwgItem2

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If ErrNo = 9 And InStr(ErrMsg, "Index was outside the bounds of the array") > 0 Then
                Resume Next
            End If

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = -2147417848 And InStr(ErrMsg, "The object invoked has disconnected") > 0 Then
                AcadApp = GetObject(, "AutoCAD.Application")
                System.Threading.Thread.Sleep(25)

                If ProblemAt = "CloseDwg" Then
                    Resume Next
                Else
                    Resume
                End If
            End If

            If ErrNo = -2145320885 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next                     'Layout.dvb was not found to be loaded
            End If

            If ErrNo = -2145320924 And InStr(ErrMsg, "is not found.") < 0 Then
                'DwgItem = VarSelArray(z)
                Resume
            End If

            CntDwgsNotFound = InStr(ErrMsg, "not a valid drawing")

            If ErrNo = -2145320825 And CntDwgsNotFound > 0 Then
                Sapi.Speak("AutoCAD found a bad Drawing, " & DwgItem & ", Going to next drawing.")
                MsgBox("AutoCAD found a bad Drawing, " & DwgItem & ", Going to next drawing.")
                BadDwgFound = "Yes"
                CntDwgsNotFound = 0
                Resume Next
            End If

            If ErrNo = -2145320851 And ErrMsg = "The named selection set exists" Then
                BlockSel.Delete()
                Resume
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next                     'Layout.dvb was not found to be loaded
            End If

            If ErrNo = 462 And Mid(ErrMsg, 1, 29) = "The RPC server is unavailable" Then
                Information.Err.Clear()
                AcadApp = CreateObject("AutoCAD.Application")
                AcadOpen = False
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            'PrgLineNo = st.GetFrame(3).GetFileLineNumber().ToString
            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

            If ErrNo = -2145320900 And ErrMsg = "Failed to get the Document object" Then
                If FirstDwg = "NotFound" Then
                    AcadApp.Application.Documents.Add()
                    Resume
                End If
            End If

            If GenInfo3233.UserName = "dlong" Then
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If

                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If

                If ErrNo = -2147418113 And ErrMsg = "Internal application error." Then
                    Information.Err.Clear()
                    AcadApp = CreateObject("AutoCAD.Application")
                    AcadOpen = False
                    Resume
                End If
            End If
        End If

    End Sub

    Function WriteToExcelAfterSort(STDsList)                         '-------DJL-07-08-2025      'Function WriteToExcelAfterSort(BOMList)
        '-------Move to new function-------WritetoExcel-------DJL-10-11-2023            
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           1/2/2024
        '-------Description:    Write data to Excel Spreadsheet
        '-------
        '-------Updates:        Description:
        '-------1-2-2024       Read Array and write to Excel what was collected from AutoCAD.     
        '-------07-08-2025      No longer need to wite to Spread Sheet then sort and Rewrite to Spread Sheet.                
        '-------                
        '------------------------------------------------------------------------------------------------
        Dim i, j, k As Integer
        Dim DwgItem2, CurrentDwgNo, FirstDwg, GetDwgNo, GetRowNo, GetX, GetY, FoundDwgNo, FoundX, FoundY, FoundItem, TotalCnt As String
        Dim GetStdPartNo, GetMXStd, FoundStdPartNo, FoundMXStd, GetRecNo, FoundDesc, FoundMat, DelItem, StdPart, StdDwg, FoundStdQty As String
        Dim GetPrevPartNo, GetPartNo, GetDesc, GetStdDesc, LookStdPartNo, LookStdDesc As String
        Dim CntDwgsNotFound, StrLineNo, PrevCnt, CntStdFound, LastRowCnt, StdDwgRow, ItemsAdded As Integer
        Dim AcadOpen As Boolean
        Dim FoundStds(20, 1)
        Dim StdsFnd2(2, 1)

        On Error GoTo Err_WriteToExcelAfterSort

        ProgressBar1.Value = 0
        Me.LblProgress.Text = "Outputting Information To Bulk BOM........Please Wait"
        Me.Refresh()

        FileSaveAS = PathBox.Text & "\" & GenInfo3233.FullJobNo & "BOM.xls"
        Workbooks = ExcelApp.Workbooks

        WorkShtName = "BulK BOM"                            '-------DJL-07-08-2025      'WorkShtName = "BOMList"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name

        '---------------------------------------------------------------------------------------------------------------------------------------------------
        '-------DJL-07-08-2025      'Below is not required anymore program is sorted at the array and wrote to spread sheet before now.
        '---------------------------------------------------------------------------------------------------------------------------------------------------
        'With BOMWrkSht
        '    .Range("C3").Value = GenInfo3233.FullJobNo
        '    .Range("J3").Value = Today
        '    .Range("G3").Value = Me.ComboBxRev.Text
        'End With

        FileToOpen = "BulK BOM"                         '-------DJL-07-08-2025      'FileToOpen = "Bulk BOM"
        ExcelApp.Visible = True
        'Test = (UBound(BOMList, 2) - 1)
        'TotalCnt = (UBound(BOMList, 2) - 1)
        ExcelApp.WindowState = XlWindowState.xlMinimized            'Minimize Excel so user can see dialogs from program
        FirstTimeThru = "Yes"
        StdsFnd2 = GenInfo3233.StdsFnd2
        StdsBOMList = GenInfo3233.StdsBOMList            'CollectSTDsList = GenInfo3233.CollectSTDsList

        For i = 0 To (UBound(STDsList, 2) - 1)           '-------DJL-07-08-2025      'Instead of looking at every line again in BOMList Look at StdsList?     'For i = 1 To (UBound(BOMList, 2) - 1)
            If i = 0 Then                               '-------DJL-07-08-2025      'If i = 1 Then
                RowNo = i + 5                           '-------DJL-07-10-2025      'RowNo = i + 4
                'Else                                   '-------DJL-07-10-2025      'Not Required.
                '    RowNo = RowNo + 1
            End If

            TotalCnt = (UBound(STDsList, 2) - 1)

            ProgressBar1.Maximum = (UBound(STDsList, 2) - 1)     '-------DJL-07-08-2025      'ProgressBar1.Maximum = (UBound(BOMList, 2) - 1)

            If RowNo = "5" And FirstTimeThru = "Yes" Then
                BOMWrkSht.Activate()
                '---------------------------------------------------------------------------------------------------------------------------------------------------
                '-------DJL-07-08-2025      'Below is not required anymore program is sorted at the array and wrote to spread sheet before now.
                '---------------------------------------------------------------------------------------------------------------------------------------------------
                'FormatLine(RowNo, FileToOpen)
                FirstTimeThru = "No"

                'With BOMWrkSht         '-----------------------Do Not need to format everyline, Speeds up Process.
                '    .Rows(RowNo & ":" & RowNo).Select()
                '    .Rows(RowNo & ":" & RowNo).Insert()
                'End With
            End If

            DelItem = STDsList(1, (TotalCnt - i))                         '-------DJL-07-08-2025       'DelItem = BOMList(1, i)
            StdPart = STDsList(9, (TotalCnt - i))                         '-------DJL-07-08-2025       'StdPart = BOMList(9, i)
            StdDwg = STDsList(10, (TotalCnt - i))                         '-------DJL-07-08-2025       'StdDwg = BOMList(10, i)
            StdDwgRow = STDsList(20, (TotalCnt - i))                         '-------DJL-07-08-2025       'StdDwg = BOMList(10, i)

            If InStr(DelItem, "Delete") = 0 Then

                With BOMWrkSht
                    '.Rows(RowNo & ":" & RowNo).Select()
                    '.Rows(RowNo & ":" & RowNo).Insert()

                    ''.Range("A" & RowNo).Value = STDsList(13, i)          'GetLen            '-------DJL-07-08-2025      13 = Nothing
                    '.Range("B" & RowNo).Value = STDsList(1, i)           'CurrentDwgNo 
                    '.Range("C" & RowNo).Value = STDsList(2, i)           'CurrentDwgRev 
                    '.Range("D" & RowNo).Value = STDsList(3, i)          'Get2DShipMk
                    '.Range("E" & RowNo).Value = STDsList(5, i)          'GetPartNo
                    '.Range("F" & RowNo).Value = STDsList(4, i)          'GetQty

                    'If STDsList(7, i) = "" Then                                              '-------DJL-07-08-0225      'Sometimes need to look at BOMList(6, i)
                    '    .Range("G" & RowNo).Value = STDsList(6, i)          'GetDesc
                    'Else
                    '    .Range("G" & RowNo).Value = STDsList(7, i)          'GetDesc
                    'End If

                    '.Range("H" & RowNo).Value = STDsList(9, i)          'GetInv1
                    '.Range("I" & RowNo).Value = STDsList(10, i)          'GetInv2
                    '.Range("J" & RowNo).Value = STDsList(11, i)          'GetMat
                    '.Range("K" & RowNo).Value = STDsList(14, i)          'GetWt
                    ''.Range("N" & RowNo).Value = STDsList(15, i)          'Blank              'CurrentDwgNo
                    '.Range("O" & RowNo).Value = STDsList(16, i)          'CStr(1) or CStr(2)
                    ''.Range("P" & RowNo).Value = STDsList(0, i)          'Blank
                    '.Range("Q" & RowNo).Value = STDsList(17, i)          'InsertionPT(1).ToString
                    '.Range("R" & RowNo).Value = STDsList(8, i)          'Yes when it is a standard.
                    ''.Range("W" & RowNo).Value = STDsList(13, i)        'Length            '-------DJL-07-08-2025      13 = Nothing
                    '.Range("W" & RowNo).Value = STDsList(18, i)          'GetProc
                    '.Range("Z" & RowNo).Value = STDsList(19, i)          '(InsPt0)
                    '.Range("AA" & RowNo).Value = STDsList(0, i)
                    '.Range("AB" & RowNo).Value = STDsList(19, i)         '(InsPt0)

                    If STDsList(9, (TotalCnt - i)) <> "" Or STDsList(10, (TotalCnt - i)) <> "" Then         '-------DJL-07-08-2025        'If BOMList(9, i) <> "" Or BOMList(10, i) <> "" Then
                        StdDwgRow = STDsList(20, (TotalCnt - i))                                      '-------DJL-07-08-2025        'RowNo = STDsList(20, (TotalCnt - i))       'RowNo = BOMList(20, i)
                        LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-08-2025
                        LastRowCnt = (LastRowCnt + 1)

                        '-------DJL-08-08-2025  'Next three lines not required.
                        '.Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025
                        '.Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()       '-------DJL-07-10-2025      'StdDwgRow      '.Rows(RowNo & ":" & RowNo).Insert() 
                        'RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)    '-------DJL-07-10-2025

                        GetStdPartNo = STDsList(9, (TotalCnt - i))                        '-------DJL-07-08-2025      'GetStdPartNo = BOMList(9, i)
                        GetMXStd = STDsList(10, (TotalCnt - i))                        '-------DJL-07-08-2025      'GetMXStd = BOMList(10, i) 

                        For l = 1 To (UBound(StdsFnd2, 2) - 1)
                            For m = 5 To (UBound(StdsBOMList, 2) - 1)
                                FoundStdQty = StdsBOMList(4, m)
                                FoundStdPartNo = StdsBOMList(9, m)
                                FoundMXStd = StdsBOMList(10, m)

                                If GetStdPartNo = FoundStdPartNo And GetMXStd = FoundMXStd Then
                                    PartFound = "Yes"

                                    '-------New problem Trevor is wanting the Bolt, Nut, and Gasket information on the standards.
                                    If STDsList(7, (TotalCnt - i)) = "" Then                  '-------DJL-07-08-0225      'If BOMList(7, i) = "" Then       'Sometimes need to look at BOMList(6, i)        'GetDesc = BOMList(7, i)
                                        GetDesc = STDsList(6, (TotalCnt - i))          'GetDesc       '-------DJL-07-08-0225      'GetDesc = BOMList(6, i)
                                    Else
                                        GetDesc = STDsList(7, (TotalCnt - i))          'GetDesc       '-------DJL-07-08-0225      'GetDesc = BOMList(7, i)
                                    End If

                                    GetPartNo = STDsList(5, (TotalCnt - i))          'GetPartNo       '-------DJL-07-08-0225      'GetPartNo = BOMList(5, i)

                                    GetPrevPartNo = StdsBOMList(5, (m - 1))         'GetPartNo
                                    GetStdDesc = StdsBOMList(7, m)                  'GetDesc
                                    FoundStdQty = StdsBOMList(4, m)                 'GetQty

                                    If InStr(GetDesc, "BOLT") > 0 Then
                                        'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                        LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                        LastRowCnt = (LastRowCnt + 1)
                                        .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()
                                        .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-10-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                        RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)        '-------DJL-07-11-2025

                                        .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                        .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                        .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                        .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                        .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                        .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                        .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                        .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                        .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                        .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight    
                                        ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                        GoTo Nextm
                                        'End If
                                    Else
                                        If InStr(GetDesc, "NUTS") > 0 Then
                                            'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                            LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                            LastRowCnt = (LastRowCnt + 1)
                                            .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()
                                            .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-10-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                            RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)        '-------DJL-07-10-2025

                                            .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                            .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                            .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                            .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                            .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                            .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                            .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                            .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                            .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                            .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight    
                                            ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                            GoTo Nextm
                                            'End If
                                        Else
                                            If InStr(GetDesc, "GASKET") > 0 Then
                                                'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                                LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                                LastRowCnt = (LastRowCnt + 1)
                                                .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()
                                                .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-10-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                                RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)        '-------DJL-07-10-2025

                                                .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                                .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                                .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                                .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                                .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                                .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                                .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                                .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                                .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                                .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight    
                                                ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                                GoTo Nextm
                                                'End If
                                            End If
                                        End If
                                    End If
                                    '--------------------------------------------------------------------------------------------

                                Else
                                    If PartFound = "Yes" Then
                                        'If RowNo = 5 Then                          'No do not do this
                                        'RowNo = (StdDwgRow + 4)        '-------DJL-07-10-2025       '? is this causing problems?
                                        'End If

                                        If FoundStdPartNo = "" And FoundMXStd = "" Then         'There are times when parts only have one reference, and there are times when many parts are part of an assembly or Standard.
                                                If FoundStdQty <> "" Then
                                                'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                                LookStdPartNo = .Range("B" & RowNo).Value
                                                    LookStdDesc = .Range("G" & RowNo).Value

                                                    If LookStdPartNo <> "" And LookStdDesc <> "" Then        '-------DJL-07-10-2025
                                                        LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                                        LastRowCnt = (LastRowCnt + 1)
                                                        .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()

                                                        .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                                    'ItemsAdded = (ItemsAdded + 1)         'Only after Item is added below.   '-------DJL-07-10-2025
                                                    'RowNo = (RowNo + 1)
                                                    RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)                         '-------DJL-07-10-2025
                                                End If

                                                    If RowNo < (StdDwgRow + 4 + 1 + ItemsAdded) Then        '-------DJL-07-10-2025
                                                        RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)
                                                    End If

                                                    .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                                    .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                                    .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                                    .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                                    .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                                    .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                                    .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                                    .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                                    .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                                    .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight
                                                    ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                                End If
                                            Else
                                                GetDesc = StdsBOMList(7, m)
                                                GetPartNo = StdsBOMList(5, m)
                                                GetPrevPartNo = StdsBOMList(5, (m - 1))

                                                If InStr(GetDesc, "BOLT") > 0 And GetPartNo > GetPrevPartNo Then
                                                    If FoundStdQty <> "" Then
                                                    'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                                    LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                                    LastRowCnt = (LastRowCnt + 1)
                                                    .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()
                                                    .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-10-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                                    RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)        '-------DJL-07-10-2025

                                                    .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                                        .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                                        .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                                        .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                                        .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                                        .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                                        .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                                        .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                                        .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                                        .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight    
                                                        ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                                        GoTo Nextm
                                                    End If
                                                Else
                                                    If InStr(GetDesc, "NUTS") > 0 And GetPartNo > GetPrevPartNo Then
                                                        If FoundStdQty <> "" Then
                                                        'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                                        LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                                        LastRowCnt = (LastRowCnt + 1)
                                                        .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()
                                                        .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-10-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                                        RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)        '-------DJL-07-10-2025

                                                        .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                                            .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                                            .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                                            .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                                            .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                                            .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                                            .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                                            .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                                            .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                                            .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight    
                                                            ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                                            GoTo Nextm
                                                        End If
                                                    Else
                                                        If InStr(GetDesc, "GASKET") > 0 And GetPartNo > GetPrevPartNo Then
                                                            If FoundStdQty <> "" Then
                                                            'RowNo = (RowNo + 1)                         '-------DJL-07-10-2025
                                                            LastRowCnt = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row         '-------DJL-07-10-2025
                                                            LastRowCnt = (LastRowCnt + 1)
                                                            .Rows(LastRowCnt & ":" & LastRowCnt).Select()       '-------DJL-07-08-2025      '.Rows(RowNo & ":" & RowNo).Select()
                                                            .Rows((StdDwgRow + 4 + 1 + ItemsAdded) & ":" & (StdDwgRow + 4 + 1 + ItemsAdded)).Insert()     '-------DJL-07-10-2025      '.Rows(RowNo & ":" & RowNo).Insert()
                                                            RowNo = (StdDwgRow + 4 + 1 + ItemsAdded)        'RowNo = (StdDwgRow + 4 + 1)        '-------DJL-07-10-2025

                                                            .Range("B" & RowNo).Value = StdsBOMList(1, m) 'ColumnA    'Dwg No.
                                                                .Range("C" & RowNo).Value = StdsBOMList(2, m) 'ColumnB    'Rev No.
                                                                .Range("D" & RowNo).Value = StdsBOMList(3, m) 'ColumnC    'Ship No.
                                                                .Range("F" & RowNo).Value = StdsBOMList(4, m) 'ColumnD    'Qty 
                                                                .Range("E" & RowNo).Value = StdsBOMList(5, m) 'ColumnE    'Part No.
                                                                .Range("G" & RowNo).Value = StdsBOMList(7, m) 'ColumnH   'Desc 
                                                                .Range("H" & RowNo).Value = StdsBOMList(9, m) 'ColumnL    'Std No. 
                                                                .Range("I" & RowNo).Value = StdsBOMList(10, m) 'ColumnK    'Inv No.
                                                                .Range("J" & RowNo).Value = StdsBOMList(11, m) 'ColumnM    'Material                                   
                                                                .Range("K" & RowNo).Value = StdsBOMList(14, m) 'ColumnN    'Weight    
                                                                ItemsAdded = (ItemsAdded + 1)                           '-------DJL-07-10-2025
                                                                GoTo Nextm
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                If GetStdPartNo <> FoundStdPartNo Then           'If GetStdPartNo <> FoundStdPartNo And GetMXStd <> FoundMXStd Then
                                                    GoTo LastItemFound                          'It is possible to have additional parts on same Standard.
                                                End If
                                            End If
                                        End If
                                    End If

Nextm:
                                If m = (UBound(StdsBOMList, 2) - 1) Then
                                    LookStdPartNo = .Range("B" & (StdDwgRow + 4 + 1)).Value        '-------DJL-07-10-2025
                                    LookStdDesc = .Range("G" & (StdDwgRow + 4 + 1)).Value

                                    If LookStdPartNo = "" And LookStdDesc = "" Then        '-------DJL-07-10-2025
                                        .Rows((StdDwgRow + 4 + 1) & ":" & (StdDwgRow + 4 + 1)).Delete()
                                    End If

                                    GoTo LastItemFound
                                End If
                            Next m

NextL:
                        Next l
                    End If
                End With
            Else
                If StdPart <> "" Or StdDwg <> "" Then
                    If StdPart = " " Or StdDwg = " " Then
                        GoTo FoundSpace
                    Else
                        Stop
                    End If
                End If

FoundSpace:
                RowNo = (RowNo - 1)
            End If

LastItemFound:
            PartFound = "No"
            ProgressBar1.Value = i
            ItemsAdded = 0
        Next i

        ExcelApp.WindowState = XlWindowState.xlNormal

        If IsNothing(RowNo) Then
            MsgBox("There were no BOM blocks found in any of the drawings. Please try again.")
            MainBOMFile.Close(SaveChanges:=False)
            ExcelApp.Quit()
            LblProgress.Text = ""
            Exit Function
        End If

        k = 1
        ProgressBar1.Maximum = UBound(STDsList, 2)           '-------DJL-07-08-2025      'ProgressBar1.Maximum = UBound(BOMList, 2)
Err_WriteToExcelAfterSort:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            DwgItem2 = CurrentDwgNo
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException
            Test = DwgItem2

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If ErrNo = 9 And InStr(ErrMsg, "Index was outside the bounds of the array") > 0 Then
                Resume Next
            End If

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = -2147417848 And InStr(ErrMsg, "The object invoked has disconnected") > 0 Then
                AcadApp = GetObject(, "AutoCAD.Application")
                System.Threading.Thread.Sleep(25)

                If ProblemAt = "CloseDwg" Then
                    Resume Next
                Else
                    Resume
                End If
            End If

            If ErrNo = -2145320885 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next                     'Layout.dvb was not found to be loaded
            End If

            If ErrNo = -2145320924 And InStr(ErrMsg, "is not found.") < 0 Then
                'DwgItem = VarSelArray(z)
                Resume
            End If

            CntDwgsNotFound = InStr(ErrMsg, "not a valid drawing")

            If ErrNo = -2145320825 And CntDwgsNotFound > 0 Then
                Sapi.Speak("AutoCAD found a bad Drawing, " & DwgItem & ", Going to next drawing.")
                MsgBox("AutoCAD found a bad Drawing, " & DwgItem & ", Going to next drawing.")
                BadDwgFound = "Yes"
                CntDwgsNotFound = 0
                Resume Next
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next                     'Layout.dvb was not found to be loaded
            End If

            If ErrNo = 462 And Mid(ErrMsg, 1, 29) = "The RPC server is unavailable" Then
                Information.Err.Clear()
                AcadApp = CreateObject("AutoCAD.Application")
                AcadOpen = False
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

            If ErrNo = -2145320900 And ErrMsg = "Failed to get the Document object" Then
                If FirstDwg = "NotFound" Then
                    AcadApp.Application.Documents.Add()
                    Resume
                End If
            End If

            If GenInfo3233.UserName = "dlong" Then
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If

                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If

                If ErrNo = -2147418113 And ErrMsg = "Internal application error." Then
                    Information.Err.Clear()
                    AcadApp = CreateObject("AutoCAD.Application")
                    AcadOpen = False
                    Resume
                End If
            End If
        End If

    End Function

    Public Function FindStdsBOM() As Object
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           Sometime before 4/2/2015    ?2013?
        '-------Description:    Find material used on Standard drawings, and add to Bulk BOM
        '-------
        '-------Updates:        Description:
        '-------    4/2/2015    Fab has requested to remove some Standard Drawings from being broke down        
        '-------                to all parts.
        '-------                
        '------------------------------------------------------------------------------------------------
        Dim i, j, jA, GetNewjA, SeeNotePos, SeeDwgPos, SeeDwg2Pos, OldQtyInt, NewCount, NewChkPos, NewBOMPos, FoundSpace, FoundPart As Integer   ', ChkPos, ExceptPos, CntExcept, RevPos, ExtPos As Integer
        Dim TotaljA, Totalj, Testi, StartCnt, StartjA, Startj, Note2Pos, NewSeeDwgPos, NewSeeNotePos, NewQtyInt, NewNote2Pos, HoldCnt, FoundInch As Integer
        Dim FoundFoot, CntTest, CountVal, CntOldStdItems, CntNewBulkBOM, ChkPos As Integer
        Dim NewDwg, NewPcMk, NewQty, NewDesc, NewDesc2, NewDesc3, NewInv, NewMatl, NewProd, GetRecNo, LookForStd, LineNo, LineNo2, LineNo4 As String
        Dim NTest, NTest2, NewStd, NewShpMk, NewRev, NewReq, FoundItem, DescFixed, CompDesc As String ', NTest1, NTest3, Ntest4, NTest5, Ntest6, NTest7, NTest8, NTest9, NTest10 As String
        Dim OTest, OTest1, OTest2, OldRecNo, OldWht, OldQty, OldInv, OldDesc, OldDesc2, OldDesc3, NewWht, FirstPart As String
        Dim SecondPart, SearchNote, SearchNote2, SearchSeeNote, SearchDwg, SearchDwg2, pattern, OPds, NPds As String
        Dim BOMSTDsSht As Worksheet
        Dim StdItemsWrkSht As Worksheet
        Dim BOMMnu As ReadDwgs
        BOMMnu = Me
        PrgName = "FindStdsBOM"
        HoldCnt = 0

        On Error GoTo Err_FindStdsBOM

        If GenInfo3233.RevNo = Nothing Then
            GenInfo3233.RevNo = Me.ComboBxRev.Text
        End If

        BOMMnu.LblProgress.Text = "Copy Standards to Bulk BOM........Please Wait"
        BOMMnu.Refresh()
        Workbooks = ExcelApp.Workbooks
        WorkShtName = "STDs BOM"
        BOMSTDsSht = Workbooks.Application.Worksheets(WorkShtName)
        BOMSTDsSht.Activate()
        WorkShtName = "STD Items"
        StdItemsWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        LineNo2 = BOMSTDsSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        LineNo4 = StdItemsWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        Count = 0
        CountVal = 0
        i = 1
        SearchSeeNote = "(SEE NOTE"
        SearchNote = " ("
        SearchDwg = "(SEE DWG"
        SearchDwg2 = "(DWG"
        SearchNote2 = "NOTE"

RptGetData: Testi = (i)
        If Testi > (LineNo2 + CountVal) Then
            BOMMnu.ProgressBar1.Maximum = Testi
        Else
            BOMMnu.ProgressBar1.Maximum = (LineNo2 + CountVal)
        End If

        For i = Testi To (LineNo2 + CountVal)
            If i > 1 Then
                GoTo GetDataNew
            End If

            RowNo = i + 4
            NewBulkBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("Bulk BOM")
            FindStdsBOM = MainBOMFile.Application.ActiveWorkbook.Sheets("STDs BOM")
            OldStdItems = MainBOMFile.Application.ActiveWorkbook.Sheets("STD Items")
            NewBOM = Nothing
            OldBOM = Nothing
            FindSTD = Nothing

            NewBOM = GenInfo3233.BOMList           'Comparison31_142.InputType3.ReadBulkBOM(NewBOM, NewBulkBOM)        'Already Done
            FindSTD = GenInfo3233.STDsList          'Comparison31_142.InputType3.ReadFindSTDs(FindSTD, FindStdsBOM)     'Already Done
            OldBOM = GenInfo3233.StdsBOMList        'Comparison31_142.InputType3.ReadBOM(OldBOM, OldStdItems)     'Already Done

            OldDesc = ""
            OldInv = ""
            OldStdDwg = ""
            OTest = ""
            OTest2 = ""
            NTest = ""
            NTest2 = ""
            Dim CntFindStds As String

            If HoldCnt = 0 Then
                HoldCnt = UBound(NewBOM, 2)
            End If

            CntNewBulkBOM = UBound(NewBOM, 2)
            CntOldStdItems = UBound(OldBOM, 2)
            CntFindStds = UBound(FindSTD, 2)
GetDataNew:
            If GetNewjA < 1 Then
                GetNewjA = 1
            Else
                GetNewjA = (GetNewjA + 1)
            End If

            If GetNewjA > BOMMnu.ProgressBar1.Maximum Then
                BOMMnu.ProgressBar1.Maximum = (GetNewjA + CountVal + 4)
                BOMMnu.ProgressBar1.Value = GetNewjA
            Else
                BOMMnu.ProgressBar1.Value = GetNewjA
            End If

            DescFixed = "No"
            TotaljA = UBound(FindSTD, 2)

            If GetNewjA > TotaljA Then
                GoTo FoundAllParts
            End If

            For jA = GetNewjA To UBound(FindSTD, 2)
                NewStdDwg = FindSTD(10, jA)          'NewStdDwg = FindSTD(8, jA)

                If Mid(NewStdDwg, 1, 2) = "MX" Or Mid(NewStdDwg, 1, 2) = "CH" Then
                    'NTest1 = FindSTD(0, jA)                    '?X Position
                    'NTest1 = FindSTD(1, jA)               'A   'Dwg
                    'NTest2 = FindSTD(2, jA)               'B   'Rev
                    'NTest2 = FindSTD(3, jA)               'D   'Piece Mark
                    NewQty = FindSTD(4, jA)                'E   'Qty               'NewQty = FindSTD(5, jA) 
                    'NTest3 = FindSTD(5, jA)               'C   'Ship Mark
                    'NTest3 = FindSTD(6, jA)               '
                    NewDesc = FindSTD(7, jA)               'F   'Description       'NewDesc = FindSTD(6, jA)
                    NewDesc = RTrim(NewDesc)
                    NewDesc2 = NewDesc

                    NewInv = FindSTD(9, jA)                'G   'INV-1               'NewInv = FindSTD(7, jA)  
                    NewStdDwg = FindSTD(10, jA)            'H   'Std Dwg No.        'NewStdDwg = FindSTD(8, jA)
                    GetMat = FindSTD(11, jA)               'I   'Material       'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                    NPds = FindSTD(12, jA)                 'J   'Weight or pounds   'NPds = FindSTD(10, jA)

                    'NTest12 = FindSTD(12, jA)             'L
                    NewBOMPos = FindSTD(13, jA)             'M
                    GetRecNo = FindSTD(20, jA)             'T
                    Dim FindSc, LookForSC, ItemNo As String

                    If NewInv = Nothing Then
                        FindSc = Nothing
                    Else
                        FindSc = NewInv & NewStdDwg
                    End If

                    For x = 0 To UBound(GenInfo3233.SubMFGData, 2)          'Do not need to look up Part numbers anymore.
                        LookForSC = GenInfo3233.SubMFGData(1, x)

                        If FindSc = LookForSC Then
                            ItemNo = GenInfo3233.SubMFGData(0, x)

                            If IsNothing(BOMWrkSht) = True Then
                                WorkShtName = "Bulk BOM"
                                BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                            End If

                            Count = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                            Totalj = UBound(NewBOM, 2)

                            If HoldCnt <> Totalj Then
                                Count = (Count - HoldCnt)
                                Count = (Count - 4)
                            Else
                                Count = (Count - UBound(NewBOM, 2))
                                Count = (Count - 4)
                            End If

                            With BOMWrkSht
                                .Range("M" & (NewBOMPos + Count)).Value = "SUB-MFG Part, Per FAB Do Not break down to BOM parts."
                                OldQtyInt = OldQty
                                NewQtyInt = NewQty
                                .Range("L" & (NewBOMPos + Count)).Value = NewQtyInt

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & (NewBOMPos + Count)).Value = "1"
                                    .Range("A" & (NewBOMPos + Count) & ":L" & (NewBOMPos + Count)).Interior.ColorIndex = 4
                                End If

                                If GenInfo3233.GetMatLen = 0 Or GenInfo3233.GetMatLen = Nothing Then
                                    .Range("L" & (NewBOMPos + Count)).Value = NewQty
                                Else
                                    .Range("L" & (NewBOMPos + Count)).Value = GenInfo3233.GetMatLen
                                End If
                            End With

                            GoTo GetDataNew
                        End If
                    Next x

                    GoTo GetStdInfo
                End If

            Next jA

GetStdInfo:
            Totalj = UBound(OldBOM, 2)

            '--------------------------------------------------------------------
            For j = GetRecNo To UBound(OldBOM, 2)          'For j = 0 To UBound(OldBOM, 2)         '-------DJL-12-27-2023
                LookForStd = OldBOM(1, j)
                OldRecNo = OldBOM(20, j)                    'T
                'May want to look at modifing program to look for Standard match before checking every line in Spreadsheet see example below:   (Hold for now.)<---Has Been Done
                '(Program needs to find second line before getting parts on Standard.
                '24"x12" FLAT BOTTOM SUMP (TYPE-A)
                '24"x10" DISHED BOTTOM SUMP (TYPE-B)

                If LookForStd = NewStdDwg And OldRecNo = GetRecNo Then                          'If Mid(LookForStd, 1) = NewStdDwg Then     '-------DJL-12-27-2023
                    'OTest1 = OldBOM(1, j)               'A
                    'OTest2 = OldBOM(2, j)               'B
                    'OTest3 = OldBOM(3, j)               'C
                    'OTest5 = OldBOM(4, j)               'D
                    OldQty = OldBOM(5, j)               'E          "Looking at list of Standards found here.
                    OldDesc = OldBOM(6, j)              'F  'Description
                    OldDesc2 = OldDesc
                    OldInv = OldBOM(7, j)               'G  'INV-1
                    OldStdDwg = OldBOM(8, j)            'H  'Std Dwg No.
                    'OTest9 = OldBOM(9, j)               'I
                    OPds = OldBOM(10, j)                'J     'Weight or Pounds
                    'OTest11 = OldBOM(11, j)             'K
                    'OTest12 = OldBOM(12, j)             'L

                    If IsNothing(OldInv) = True Then
                        GoTo GetData4
                    Else
                        If OldInv <> NewInv Then
                            GoTo GetData4
                        End If
                    End If

                    SeeNotePos = 0
                    ChkPos = 0
                    SeeDwgPos = 0
                    Note2Pos = 0
                    SeeNotePos = InStr(1, OldDesc, SearchSeeNote)
                    ChkPos = InStr(1, OldDesc, SearchNote)
                    SeeDwgPos = InStr(1, OldDesc, SearchDwg)                            '-------(SEE DWG
                    SeeDwg2Pos = InStr(1, OldDesc, SearchDwg2)                          '-------(DWG
                    Note2Pos = InStr(1, OldDesc, SearchNote2)

                    Select Case 0
                        Case Is < SeeNotePos
                            OldDesc = Mid(OldDesc, 1, (SeeNotePos - 2))         'Question should this be minus 1
                        Case Is < ChkPos
                            OldDesc2 = Mid(OldDesc, 1, (ChkPos - 1))
                        Case Is < SeeDwgPos
                            GoTo GetData4
                        Case Is < SeeDwg2Pos
                            GoTo GetData4
                        Case Is < Note2Pos
                            OldDesc2 = Mid(OldDesc, 1, (Note2Pos - 1))
                    End Select

                    NewDesc = NewDesc
                    OldDesc2 = OldDesc2
                    NewInv = NewInv
                    Dim NewSeeDwg2Pos As Integer
                    NewSeeNotePos = 0
                    NewChkPos = 0
                    NewSeeDwgPos = 0
                    NewSeeDwg2Pos = 0
                    NewNote2Pos = 0
                    NewSeeNotePos = InStr(1, NewDesc, SearchSeeNote)
                    NewChkPos = InStr(1, NewDesc, SearchNote)
                    NewSeeDwgPos = InStr(1, NewDesc, SearchDwg)
                    NewSeeDwg2Pos = InStr(1, NewDesc, SearchDwg2)
                    NewNote2Pos = InStr(1, NewDesc, SearchNote2)

                    Select Case 0
                        Case Is < NewSeeNotePos
                            NewDesc = Mid(NewDesc, 1, (NewSeeNotePos - 2))
                            NewDesc = RTrim(NewDesc)
                        Case Is < NewChkPos
                            NewDesc2 = Mid(NewDesc, 1, (NewChkPos - 1))
                            NewDesc2 = RTrim(NewDesc2)
                        Case Is < NewSeeDwgPos
                            GoTo GetData4
                        Case Is < NewSeeDwg2Pos
                            GoTo GetData4
                        Case Is < NewNote2Pos
                            NewDesc2 = Mid(NewDesc, 1, (NewNote2Pos - 2))
                            NewDesc2 = RTrim(NewDesc2)
                    End Select

                    NewDesc = NewDesc
                    NewDesc2 = NewDesc2
                    OldDesc2 = OldDesc2
                    NewInv = NewInv

                    If OldInv = NewInv Then
                        NewDesc3 = NewDesc2
                        OldDesc3 = OldDesc2

                        If NewDesc2 <> OldDesc2 Then
                            FoundFoot = InStr(1, NewDesc3, SearchFoot)
                            While FoundFoot > 0
                                FirstPart = Mid(NewDesc3, 1, (FoundFoot - 1))
                                SecondPart = Mid(NewDesc3, (FoundFoot + 1), (Len(NewDesc3) - FoundFoot))
                                NewDesc3 = FirstPart & SecondPart
                                FoundFoot = InStr(1, NewDesc3, SearchFoot)
                            End While

                            FoundInch = InStr(1, NewDesc3, SearchInch)
                            While FoundInch > 0
                                FirstPart = Mid(NewDesc3, 1, (FoundInch - 1))
                                SecondPart = Mid(NewDesc3, (FoundInch + 1), (Len(NewDesc3) - FoundInch))
                                NewDesc3 = FirstPart & SecondPart
                                FoundInch = InStr(1, NewDesc3, SearchInch)
                            End While

                            FoundSpace = InStr(1, NewDesc3, SearchSpace)
                            While FoundSpace > 0
                                FirstPart = Mid(NewDesc3, 1, (FoundSpace - 1))
                                SecondPart = Mid(NewDesc3, (FoundSpace + 1), (Len(NewDesc3) - FoundSpace))
                                NewDesc3 = FirstPart & SecondPart
                                FoundSpace = InStr(1, NewDesc3, SearchSpace)
                            End While

                            FoundFoot = InStr(1, OldDesc3, SearchFoot)
                            While FoundFoot > 0
                                FirstPart = Mid(OldDesc3, 1, (FoundFoot - 1))
                                SecondPart = Mid(OldDesc3, (FoundFoot + 1), (Len(OldDesc3) - FoundFoot))
                                OldDesc3 = FirstPart & SecondPart
                                FoundFoot = InStr(1, OldDesc3, SearchFoot)
                            End While

                            FoundInch = InStr(1, OldDesc3, SearchInch)
                            While FoundInch > 0
                                FirstPart = Mid(OldDesc3, 1, (FoundInch - 1))
                                SecondPart = Mid(OldDesc3, (FoundInch + 1), (Len(OldDesc3) - FoundInch))
                                OldDesc3 = FirstPart & SecondPart
                                FoundInch = InStr(1, OldDesc3, SearchInch)
                            End While

                            FoundSpace = InStr(1, OldDesc3, SearchSpace)
                            While FoundSpace > 0
                                FirstPart = Mid(OldDesc3, 1, (FoundSpace - 1))
                                SecondPart = Mid(OldDesc3, (FoundSpace + 1), (Len(OldDesc3) - FoundSpace))
                                OldDesc3 = FirstPart & SecondPart
                                FoundSpace = InStr(1, OldDesc3, SearchSpace)
                            End While
                        End If
                    End If

                    NewDesc3 = NewDesc3
                    OldDesc3 = OldDesc3

                    Select Case NewDesc
                        Case OldDesc
                            If OldInv = NewInv Then
                                Startj = j
                                StartCnt = 0
                                GoTo GetData5
                            Else
                                GoTo GetData4
                            End If
                        Case OldDesc2
                            If OldInv = NewInv Then
                                Startj = j
                                GoTo GetData5
                            Else
                                GoTo GetData4
                            End If
                        Case Else
                            Select Case NewDesc2
                                Case OldDesc
                                    If OldInv = NewInv Then
                                        Startj = j
                                        GoTo GetData5
                                    Else
                                        GoTo GetData4
                                    End If
                                Case OldDesc2
                                    If OldInv = NewInv Then
                                        Startj = j
                                        GoTo GetData5
                                    Else
                                        GoTo GetData4
                                    End If
                                Case Else
                                    If NewDesc3 = OldDesc3 Then
                                        If OldInv = NewInv Then
                                            Startj = j
                                            GoTo GetData5
                                        Else
                                            GoTo GetData4
                                        End If
                                    Else
                                        j = (UBound(OldBOM, 2) + 1)
                                        GoTo GetData6
                                    End If
                            End Select
                    End Select
                Else
                    GoTo GetData4
                End If

                pattern = NTest
                If pattern <> "" Then
                    Dim matches As MatchCollection = Regex.Matches(OTest, pattern)
                    If Regex.IsMatch(OTest, pattern) Then
                        For Each match As Match In matches
                            FoundPart = match.Value
                        Next
                    End If
                End If
GetData5:
                FirstTimeThru = "Yes"
                OldQty = NewQty
                CountNewItems = (CountNewItems + 1)

                For l = (Startj + 1) To UBound(OldBOM, 2)       '-------Found Part Now get Items for Standard.
                    NewDwg = OldBOM(1, l)               'A      'Dwg
                    NewRev = OldBOM(2, l)               'B      'Rev
                    NewShpMk = OldBOM(3, l)             'C      'Ship Mark
                    NewPcMk = OldBOM(4, l)              'D      'New Piece Mark
                    NewPcMk = LTrim(NewPcMk)            '----------New Problem found NewPcMk's with blank spaces...........
                    NewPcMk = RTrim(NewPcMk)

                    NewQty = OldBOM(5, l)               'E      'Qty
                    NewDesc = OldBOM(6, l)              'F      'Description
                    NewInv = OldBOM(7, l)               'G      'Inv-1
                    NewStd = OldBOM(8, l)               'H      'Standard Number Example:MX1001A
                    NewMatl = OldBOM(9, l)              'I      'Material
                    NewWht = OldBOM(10, l)              'J      'Weight
                    NewReq = OldBOM(11, l)              'K      'Required Type
                    'NewProd = OldBOM(12, l)             'L         'Index Out of Bounds OldBOM only looks at first 11 columns.

                    If IsNothing(BOMWrkSht) = True Then
                        WorkShtName = "Bulk BOM"
                        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                    End If

                    If NewPcMk = Nothing Then       '--Question is there a mistake here"If IsNothing(NewPcMk) = True Then".
                        If CompDesc <> "" Then
                            With BOMWrkSht
                                If FirstTimeThru = "Yes" Then
                                    If CompDesc = "" Then
                                        .Range("M" & (jA + 4)).Value = "Standard was not found."
                                        With .Range("A" & (jA + 4) & ":M" & (jA + 4))
                                            With .Interior
                                                .ColorIndex = 7
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With
                                    Else
                                        .Range("M" & (jA + 4)).Value = "Standard Reference only, No additional parts. " & CompDesc
                                        With .Range("A" & (jA + 4) & ":M" & (jA + 4))
                                            With .Interior
                                                .ColorIndex = 45
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With
                                    End If

                                    If GenInfo3233.RevNo = 0 Then
                                        .Range("Q" & (jA + 4)).Value = "1"
                                        .Range("A" & (jA + 4) & ":L" & (jA + 4)).Interior.ColorIndex = 4
                                    End If
                                Else
                                    .Range("M" & (jA + 4)).Value = CompDesc
                                    With .Range("A" & (jA + 4) & ":M" & (jA + 4))
                                        With .Interior
                                            .ColorIndex = 45
                                            .Pattern = Constants.xlSolid
                                        End With
                                    End With

                                    If GenInfo3233.RevNo = 0 Then
                                        .Range("Q" & (jA + 4)).Value = "1"
                                        .Range("A" & (jA + 4) & ":L" & (jA + 4)).Interior.ColorIndex = 4
                                    End If
                                End If
                            End With
                        Else
                            If IsNothing(BOMWrkSht) = True Then
                                WorkShtName = "Bulk BOM"
                                BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                            End If

                            If FirstTimeThru = "Yes" Then
                                Count = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                                Totalj = UBound(NewBOM, 2)

                                If HoldCnt <> Totalj Then
                                    Count = (Count - HoldCnt)
                                    Count = (Count - 4)
                                Else
                                    Count = (Count - UBound(NewBOM, 2))
                                    Count = (Count - 4)
                                End If

                                With BOMWrkSht
                                    .Range("M" & (NewBOMPos + Count)).Value = "Standard Reference only, No additional parts."
                                    OldQtyInt = OldQty
                                    NewQtyInt = NewQty

                                    If InStr(OldDesc, "SHT") = 0 Then
                                        If OldQtyInt > 0 Then
                                            If NewQtyInt > 0 Then
                                                .Range("L" & (NewBOMPos + Count)).Value = (OldQtyInt * NewQtyInt)
                                            Else
                                                .Range("L" & (NewBOMPos + Count)).Value = OldQtyInt
                                            End If
                                        Else
                                            .Range("L" & (NewBOMPos + Count)).Value = NewQtyInt
                                        End If
                                    Else
                                        .Range("L" & (NewBOMPos + Count)).Value = OldWht
                                    End If

                                    With .Range("A" & (NewBOMPos + Count) & ":M" & (NewBOMPos + Count))
                                        With .Interior
                                            .ColorIndex = 8
                                            .Pattern = Constants.xlSolid
                                        End With
                                    End With

                                    If GenInfo3233.RevNo = 0 Then
                                        .Range("Q" & (NewBOMPos + Count)).Value = "1"
                                        .Range("A" & (NewBOMPos + Count) & ":L" & (NewBOMPos + Count)).Interior.ColorIndex = 4
                                    End If
                                End With

                                GoTo GetDataNew
                            End If
                        End If
                        CompDesc = ""
                        FirstTimeThru = "No"
                        GetNewjA = jA
                        GoTo GetDataNew
                    End If

                    WorkShtName = "Bulk BOM"
                    BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                    Count = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                    Totalj = UBound(NewBOM, 2)

                    If HoldCnt <> Totalj Then
                        Count = (Count - HoldCnt)
                        Count = (Count - 4)
                    Else
                        Count = (Count - UBound(NewBOM, 2))
                        Count = (Count - 4)
                    End If

                    With BOMWrkSht
                        FirstTimeThru = "No"

                        If FoundItem = "Yes" Then
                            LineNo = jA
                        Else
                            LineNo = jA
                            FoundItem = "Yes"
                        End If

                        If Count <> NewCount Then
                            If DescFixed = "No" Then
                                LineNo = (LineNo + Count + 3)
                            Else
                                LineNo = (LineNo + Count + 2)
                            End If
                        Else
                            If Startj > LineNo Then
                                LineNo = (Startj + Count + 4)
                            Else
                                LineNo = (LineNo + Count + 3)
                            End If
                        End If

                        OldDesc = OldDesc

                        'Found problem here were program is inserting duplicates due to LineNo is wrong above for new part which is the same part on different drawing.
                        With BOMWrkSht                          ' Standard Found.
                            If Count > 0 Then                   'Make sure to format all lines inserted if not you will have no lines.
                                LineNo = (NewBOMPos + Count + 1)
                                CntTest = (Totalj - HoldCnt)

                                If CntTest > 0 Then
                                    LineNo = LineNo + CntTest
                                End If

                                FileToOpen = "Bulk BOM"
                                FormatLine3(LineNo, FileToOpen)  'This inserts a line and formats the line replaced above insert
                            Else
                                LineNo = (NewBOMPos + 1)
                                CntTest = (Totalj - HoldCnt)                'Then this is used everytime after.

                                If CntTest > 0 Then
                                    LineNo = LineNo + CntTest
                                End If

                                FileToOpen = "Bulk BOM"
                                FormatLine3(LineNo, FileToOpen)
                            End If
                        End With

                        .Range("A" & LineNo).Value = NewDwg
                        .Range("B" & LineNo).Value = NewRev
                        .Range("C" & LineNo).Value = NewShpMk
                        .Range("D" & LineNo).Value = NewPcMk

                        If InStr(OldQty, "?") > 0 Then
                            OldQty = 0
                        End If

                        If OldQty = 0 Then
                            .Range("E" & LineNo).Value = NewQty
                        Else
                            .Range("E" & LineNo).Value = (NewQty * OldQty)
                        End If

                        .Range("F" & LineNo).Value = NewDesc
                        .Range("G" & LineNo).Value = NewInv
                        .Range("H" & LineNo).Value = NewStd
                        .Range("I" & LineNo).Value = NewMatl

                        Select Case NewWht
                            Case "-"
                                .Range("J" & LineNo).Value = NewWht
                            Case "??"
                                .Range("J" & LineNo).Value = "0.000"
                            Case Else
                                If OldQty = 0 Then
                                    .Range("J" & LineNo).Value = (NewWht * NewQty)
                                Else
                                    .Range("J" & LineNo).Value = (NewWht * (NewQty * OldQty))
                                End If
                        End Select

                        .Range("K" & LineNo).Value = NewReq
                        '.Range("L" & LineNo).Value = NewProd

                        If CompDesc <> "" Then
                            .Range("M" & RowNo).Value = CompDesc
                            With .Range("A" & RowNo & ":M" & RowNo)
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & RowNo).Value = "1"
                                .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                            End If

                            With .Range("A" & (RowNo + 1) & ":M" & (RowNo + 1))
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & (RowNo + 1)).Value = "1"
                                .Range("A" & (RowNo + 1) & ":L" & (RowNo + 1)).Interior.ColorIndex = 4
                            End If
                        Else
                            If DescFixed = "Yes" Then
                                With .Range("A" & LineNo & ":M" & LineNo)
                                    With .Interior
                                        .ColorIndex = 8
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & LineNo).Value = "1"
                                    .Range("A" & LineNo & ":L" & LineNo).Interior.ColorIndex = 4
                                End If
                            Else
                                .Range("M" & (LineNo - 1)).Value = "Found Standard Information"
                                With .Range("A" & (LineNo - 1) & ":M" & (LineNo - 1))
                                    With .Interior
                                        .ColorIndex = 8
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & (LineNo - 1)).Value = "1"
                                    .Range("A" & (LineNo - 1) & ":L" & (LineNo - 1)).Interior.ColorIndex = 4
                                End If

                                DescFixed = "Yes"
                                With .Range("A" & LineNo & ":M" & LineNo)
                                    With .Interior
                                        .ColorIndex = 8
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & LineNo).Value = "1"
                                    .Range("A" & LineNo & ":L" & LineNo).Interior.ColorIndex = 4
                                End If
                            End If
                        End If
                    End With
                    NewCount = (Count + 1)
                    Count = (Count + 1)

                Next l

                '-------------Check if all parts were found then go to next part, instead of looking for part again.
                If DescFixed = "Yes" Then
                    GoTo GetDataNew
                End If
GetData4:
            Next j
GetData6:
            With BOMWrkSht
                If j > UBound(OldBOM, 2) Then
                    Totalj = UBound(OldBOM, 2)          'Description does not match
                    CompDesc = ""

                    For j = 0 To UBound(OldBOM, 2)
                        LookForStd = OldBOM(1, j)

                        If Mid(LookForStd, 1) = NewStdDwg Then
                            'OTest1 = OldBOM(1, j)               'A
                            'OTest2 = OldBOM(2, j)               'B
                            'OTest3 = OldBOM(3, j)               'C
                            'Otest4 = OldBOM(4, j)               'D
                            'OTest5 = OldBOM(5, j)               'E

                            OldDesc = OldBOM(6, j)              'F  'Description
                            OldInv = OldBOM(7, j)               'G  'INV-1
                            OldStdDwg = OldBOM(8, j)            'H  'Std Dwg No.

                            'OTest9 = OldBOM(9, j)               'I
                            'OTest10 = OldBOM(10, j)             'J
                            'OTest11 = OldBOM(11, j)             'K
                            'OTest12 = OldBOM(12, j)             'L

                            If OldInv = Nothing Then
                                'Do nothing except go to next line
                            Else
                                If OldInv = NewInv Then
                                    CompDesc = NewDesc & " and " & OldDesc & " --Description Did Not Match double check Item"
                                    Startj = j
                                    StartjA = (jA + 4)
                                    OldQty = NewQty
                                    GetStdInfo(CompDesc, Startj, StartjA, OldQty, FuncGetDataNew, NewBOMPos, HoldCnt)   'GoTo GetData5  'GoTo GetData
                                    DescFixed = "Yes"
                                    CompDesc = ""
                                    FirstTimeThru = "No"
                                    GetNewjA = jA
                                    GoTo GetDataNew
                                End If
                            End If
                        End If

                    Next j

                    If IsNothing(BOMWrkSht) = True Then
                        WorkShtName = "Bulk BOM"
                        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
                    End If

                    With BOMWrkSht                  '----------Insert Items into Bulk BOM.
                        DescFixed = DescFixed
                        If DescFixed = "No" Then
                            OldDesc = OldDesc
                            NewDesc = NewDesc
                            Count = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                            Totalj = UBound(NewBOM, 2)

                            If HoldCnt <> Totalj Then
                                Count = (Count - HoldCnt)
                                Count = (Count - 4)
                            Else
                                Count = (Count - UBound(NewBOM, 2))
                                Count = (Count - 4)
                            End If

                            '----------------------First time to find part missing
                            .Range("M" & (NewBOMPos + Count)).Value = "Standard was not found."
                            If OldQty = "" Then
                                OldQty = 1
                            End If

                            OldQtyInt = OldQty

                            If InStr(NewQty, "?") > 0 Then
                                NewQtyInt = 0
                            Else
                                NewQtyInt = NewQty
                            End If

                            If OldQtyInt > 0 Then
                                .Range("L" & (NewBOMPos + Count)).Value = (OldQtyInt * NewQtyInt)
                            Else
                                .Range("L" & (NewBOMPos + Count)).Value = NewQtyInt
                            End If

                            With .Range("A" & (NewBOMPos + Count) & ":M" & (NewBOMPos + Count))
                                With .Interior
                                    .ColorIndex = 7
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & (NewBOMPos + Count)).Value = "1"
                                .Range("A" & (NewBOMPos + Count) & ":L" & (NewBOMPos + Count)).Interior.ColorIndex = 4
                            End If
                        End If
                    End With
                    If jA > 1 Then
                        GoTo GetDataNew
                    Else
                        GoTo GetData2
                    End If
                End If

                If NewPcMk = Nothing Then
                    NewPcMk = "GetData"
                End If
GetData:
                FirstTimeThru = "Yes"

                For j = (RowNo2 + 1) To LineNo4
                    With StdItemsWrkSht
                        NewDwg = .Range("A" & j).Value         'Dwg
                        NewRev = .Range("B" & j).Value         'Rev
                        NewShpMk = .Range("C" & j).Value       'Ship Mark
                        NewPcMk = .Range("D" & j).Value        'New Piece Mark
                        NewPcMk = LTrim(NewPcMk)            '----------New Problem found NewPcMk's with blank spaces...........
                        NewPcMk = RTrim(NewPcMk)

                        NewQty = .Range("E" & j).Value         'Qty
                        NewDesc = .Range("F" & j).Value        'Description
                        NewInv = .Range("G" & j).Value         'Inv-1
                        NewStd = .Range("H" & j).Value         'Standard Number Example:MX1001A
                        NewMatl = .Range("I" & j).Value        'Material
                        NewWht = .Range("J" & j).Value         'Weight
                        NewReq = .Range("K" & j).Value         'Required Type
                        NewProd = .Range("L" & j).Value        'Production Code
                        If NewPcMk = Nothing Then
                            If CompDesc <> "" Then
                                With BOMWrkSht
                                    If FirstTimeThru = "Yes" Then
                                        If CompDesc = "" Then
                                            .Range("M" & RowNo).Value = "Standard was not found."
                                            With .Range("A" & RowNo & ":M" & RowNo)
                                                With .Interior
                                                    .ColorIndex = 7
                                                    .Pattern = Constants.xlSolid
                                                End With
                                            End With

                                            If GenInfo3233.RevNo = 0 Then
                                                .Range("Q" & RowNo).Value = "1"
                                                .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                                            End If
                                        Else
                                            .Range("M" & RowNo).Value = "Standard Reference only, No additional parts. " & CompDesc
                                            With .Range("A" & RowNo & ":M" & RowNo)
                                                With .Interior
                                                    .ColorIndex = 45
                                                    .Pattern = Constants.xlSolid
                                                End With
                                            End With

                                            If GenInfo3233.RevNo = 0 Then
                                                .Range("Q" & RowNo).Value = "1"
                                                .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                                            End If

                                        End If
                                    Else
                                        .Range("M" & RowNo).Value = CompDesc
                                        With .Range("A" & RowNo & ":M" & RowNo)
                                            With .Interior
                                                .ColorIndex = 45
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With

                                        If GenInfo3233.RevNo = 0 Then
                                            .Range("Q" & RowNo).Value = "1"
                                            .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                                        End If
                                    End If
                                End With
                            Else
                                If FirstTimeThru = "Yes" Then
                                    With BOMWrkSht
                                        .Range("M" & RowNo).Value = "Standard Reference only, No additional parts."
                                        With .Range("A" & RowNo & ":M" & RowNo)
                                            With .Interior
                                                .ColorIndex = 8
                                                .Pattern = Constants.xlSolid
                                            End With
                                        End With

                                        If GenInfo3233.RevNo = 0 Then
                                            .Range("Q" & RowNo).Value = "1"
                                            .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                                        End If

                                    End With
                                End If
                            End If
                            CompDesc = ""
                            FirstTimeThru = "No"
                            GoTo GetData2
                        End If

                    End With

                    With BOMWrkSht                  '----------Insert Items into Bulk BOM.
                        FirstTimeThru = "No"
                        If FoundItem = "Yes" Then
                            LineNo = ((RowNo - 1) + Count)
                        Else
                            LineNo = (RowNo - 1)
                            FoundItem = "Yes"
                        End If

                        FileToOpen = "Bulk BOM"
                        FormatLine(LineNo, FileToOpen)
                        .Range("A" & ((RowNo + 1) + Count)).Value = NewDwg
                        .Range("B" & ((RowNo + 1) + Count)).Value = NewRev
                        .Range("C" & ((RowNo + 1) + Count)).Value = NewShpMk
                        .Range("D" & ((RowNo + 1) + Count)).Value = NewPcMk

                        If OldQty = 0 Then
                            .Range("E" & ((RowNo + 1) + Count)).Value = NewQty
                        Else
                            .Range("E" & ((RowNo + 1) + Count)).Value = (NewQty * OldQty)
                        End If

                        .Range("F" & ((RowNo + 1) + Count)).Value = NewDesc
                        .Range("G" & ((RowNo + 1) + Count)).Value = NewInv
                        .Range("H" & ((RowNo + 1) + Count)).Value = NewStd
                        .Range("I" & ((RowNo + 1) + Count)).Value = NewMatl
                        If NewWht = "-" Then
                            .Range("J" & ((RowNo + 1) + Count)).Value = NewWht
                        Else
                            .Range("J" & ((RowNo + 1) + Count)).Value = (NewWht * OldQty)
                        End If
                        .Range("K" & ((RowNo + 1) + Count)).Value = NewReq
                        .Range("L" & ((RowNo + 1) + Count)).Value = NewProd

                        If CompDesc <> "" Then
                            .Range("M" & RowNo).Value = CompDesc
                            With .Range("A" & RowNo & ":M" & RowNo)
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & RowNo).Value = "1"
                                .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                            End If

                            With .Range("A" & (RowNo + 1) & ":M" & (RowNo + 1))
                                With .Interior
                                    .ColorIndex = 45
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & (RowNo + 1)).Value = "1"
                                .Range("A" & (RowNo + 1) & ":L" & (RowNo + 1)).Interior.ColorIndex = 4
                            End If
                        Else
                            .Range("M" & RowNo).Value = "Found Standard Information"
                            With .Range("A" & RowNo & ":M" & RowNo)
                                With .Interior
                                    .ColorIndex = 8
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & RowNo).Value = "1"
                                .Range("A" & RowNo & ":L" & RowNo).Interior.ColorIndex = 4
                            End If

                            With .Range("A" & (RowNo + 1) & ":M" & (RowNo + 1))
                                With .Interior
                                    .ColorIndex = 8
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & (RowNo + 1)).Value = "1"
                                .Range("A" & (RowNo + 1) & ":L" & (RowNo + 1)).Interior.ColorIndex = 4
                            End If
                        End If
                    End With
                    Count = (Count + 1)

                Next j
GetData2:
                CountVal = (CountVal + Count)
                Count = 0
                FoundItem = "No"
                CompDesc = ""
                BOMMnu.ProgressBar1.Value = i
            End With
        Next i

        If (LineNo2 + CountVal) > i Then
            GoTo RptGetData
        End If

FoundAllParts:

Err_FindStdsBOM:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Public Function GetStdInfo(ByVal CompDesc As String, ByVal Startj As Integer, ByVal StartjA As Integer, ByVal OldQty As Integer, ByVal FuncGetDataNew As String, ByVal NewBOMPos As Integer, ByVal HoldCnt As Integer) As Object
        Dim ExceptPos, jA, l, GetNewjA, TotalOnNewBOM, Totalj, CntTest, OldQtyInt, NewQtyInt As Integer
        Dim WorkShtName, FoundLast, LineNo, SearchException, FoundItem, SearchSeeNote As String
        Dim SearchNote2, SearchNote, SearchDwg, pattern As String
        Dim FileToOpen, NewDwg, NewRev, NewShpMk, NewPcMk, NewQty, NewDesc, NewInv, NewStd As String
        Dim NewProd, DescFixed, NewMatl, NewWht, NewReq As String
        Dim Workbooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim BOMWrkSht As Worksheet
        Dim ExcelApp As Object
        Dim BOMMnu As ReadDwgs
        BOMMnu = Me

        jA = StartjA
        PrgName = "GetStdInfo"
        DescFixed = "No"

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

        On Error GoTo Err_GetStdInfo

        Workbooks = ExcelApp.Workbooks
        WorkShtName = "Bulk BOM"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        BOMWrkSht.Activate()
        Count = BOMWrkSht.Range("A4000").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        Totalj = Count
        TotalOnNewBOM = UBound(NewBOM, 2)
        Count = (Count - UBound(NewBOM, 2))
        Count = (Count - 4)

        Dim OldCount As Integer
        OldCount = UBound(OldBOM, 2)
        FirstTimeThru = "Yes"
        Startj = 4

        For l = (Startj + 1) To UBound(OldBOM, 2)                           'For l = (Startj + 1) To UBound(OldBOM, 2)
            '                                   With StdItemsWrkSht
            NewDwg = OldBOM(1, l)               'A  'NewDwg = .Range("A" & j).Value         'Dwg
            NewRev = OldBOM(2, l)               'B  'NewRev = .Range("B" & j).Value         'Rev
            NewShpMk = OldBOM(3, l)             'C  'NewShpMk = .Range("C" & j).Value       'Ship Mark
            NewPcMk = OldBOM(4, l)              'D  'NewPcMk = .Range("D" & j).Value        'New Piece Mark
            NewPcMk = LTrim(NewPcMk)            '----------New Problem found NewPcMk's with blank spaces...........
            NewPcMk = RTrim(NewPcMk)

            NewQty = OldBOM(5, l)               'E  'NewQty = .Range("E" & j).Value         'Qty
            NewDesc = OldBOM(6, l)              'F  'NewDesc = .Range("F" & j).Value        'Description
            NewInv = OldBOM(7, l)               'G  'NewInv = .Range("G" & j).Value         'Inv-1
            NewStd = OldBOM(8, l)               'H  'NewStd = .Range("H" & j).Value         'Standard Number Example:MX1001A
            NewMatl = OldBOM(9, l)              'I  'NewMatl = .Range("I" & j).Value        'Material
            NewWht = OldBOM(10, l)              'J  'NewWht = .Range("J" & j).Value         'Weight
            NewReq = OldBOM(11, l)              'K  'NewReq = .Range("K" & j).Value         'Required Type
            'NewProd = OldBOM(12, l)            'L  'Index Out of Bounds OldBOM only looks at first 11 columns.     'Production Code

            If NewPcMk = Nothing Then
                If CompDesc <> "" Then
                    With BOMWrkSht
                        If FirstTimeThru = "Yes" Then
                            If CompDesc = "" Then
                                .Range("M" & (NewBOMPos + Count)).Value = "Standard was not found."
                                With .Range("A" & (NewBOMPos + Count) & ":M" & (NewBOMPos + Count))
                                    With .Interior
                                        .ColorIndex = 7
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & (NewBOMPos + Count)).Value = "1"
                                    .Range("A" & (NewBOMPos + Count) & ":L" & (NewBOMPos + Count)).Interior.ColorIndex = 4
                                End If
                            Else
                                .Range("M" & (NewBOMPos + Count)).Value = "Standard Reference only, No additional parts. " & CompDesc
                                OldQtyInt = OldQty
                                NewQtyInt = NewQty

                                If OldQtyInt > 0 Then
                                    If NewQtyInt = 0 Then
                                        .Range("L" & (NewBOMPos + Count)).Value = OldQtyInt
                                    Else
                                        .Range("L" & (NewBOMPos + Count)).Value = (OldQtyInt * NewQtyInt)
                                    End If
                                Else
                                    .Range("L" & (NewBOMPos + Count)).Value = NewQtyInt
                                End If

                                With .Range("A" & (NewBOMPos + Count) & ":M" & (NewBOMPos + Count))
                                    With .Interior
                                        .ColorIndex = 45
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & (NewBOMPos + Count)).Value = "1"
                                    .Range("A" & (NewBOMPos + Count) & ":L" & (NewBOMPos + Count)).Interior.ColorIndex = 4
                                End If
                            End If
                        Else
                            If FoundItem <> "Yes" Then
                                .Range("M" & NewBOMPos).Value = CompDesc
                                With .Range("A" & NewBOMPos & ":M" & NewBOMPos)
                                    With .Interior
                                        .ColorIndex = 45
                                        .Pattern = Constants.xlSolid
                                    End With
                                End With

                                If GenInfo3233.RevNo = 0 Then
                                    .Range("Q" & NewBOMPos).Value = "1"
                                    .Range("A" & NewBOMPos & ":L" & NewBOMPos).Interior.ColorIndex = 4
                                End If
                            End If
                        End If
                    End With
                Else
                    If FirstTimeThru = "Yes" Then
                        With BOMWrkSht
                            .Range("M" & NewBOMPos).Value = "Standard Reference only, No additional parts."
                            With .Range("A" & NewBOMPos & ":M" & NewBOMPos)
                                With .Interior
                                    .ColorIndex = 8
                                    .Pattern = Constants.xlSolid
                                End With
                            End With

                            If GenInfo3233.RevNo = 0 Then
                                .Range("Q" & NewBOMPos).Value = "1"
                                .Range("A" & NewBOMPos & ":L" & NewBOMPos).Interior.ColorIndex = 4
                            End If
                        End With
                    End If
                End If
                CompDesc = ""
                FirstTimeThru = "No"
                GetNewjA = jA
                FuncGetDataNew = "GetDataNew"
                GoTo Err_GetStdInfo
            End If

            RowNo = StartjA

            With BOMWrkSht
                FirstTimeThru = "No"
                If FoundItem = "Yes" Then
                    LineNo = (NewBOMPos - 1)
                Else
                    LineNo = (NewBOMPos - 1)
                    FoundItem = "Yes"
                End If

                FileToOpen = "Bulk BOM"
                LineNo = (LineNo + Count)
                FormatLine(LineNo, FileToOpen)
                .Range("A" & ((NewBOMPos + 1) + Count)).Value = NewDwg
                .Range("B" & ((NewBOMPos + 1) + Count)).Value = NewRev
                .Range("C" & ((NewBOMPos + 1) + Count)).Value = NewShpMk
                .Range("D" & ((NewBOMPos + 1) + Count)).Value = NewPcMk

                If OldQty = 0 Then
                    .Range("E" & ((NewBOMPos + 1) + Count)).Value = NewQty
                Else
                    .Range("E" & ((NewBOMPos + 1) + Count)).Value = (NewQty * OldQty)
                End If

                .Range("F" & ((NewBOMPos + 1) + Count)).Value = NewDesc     '
                .Range("G" & ((NewBOMPos + 1) + Count)).Value = NewInv      '
                .Range("H" & ((NewBOMPos + 1) + Count)).Value = NewStd      '
                .Range("I" & ((NewBOMPos + 1) + Count)).Value = NewMatl     '
                If NewWht = "-" Then
                    .Range("J" & ((NewBOMPos + 1) + Count)).Value = NewWht
                Else
                    .Range("J" & ((NewBOMPos + 1) + Count)).Value = (NewWht * OldQty)
                End If
                .Range("K" & ((NewBOMPos + 1) + Count)).Value = NewReq

                If CompDesc <> "" Then
                    If DescFixed = "No" Then
                        .Range("M" & (NewBOMPos + Count)).Value = CompDesc
                        OldQtyInt = OldQty
                        NewQtyInt = NewQty

                        If OldQtyInt > 0 Then
                            If NewQtyInt > 0 Then
                                .Range("L" & (NewBOMPos + Count)).Value = (OldQtyInt * NewQtyInt)
                            Else
                                .Range("L" & (NewBOMPos + Count)).Value = OldQtyInt
                            End If
                        Else
                            .Range("L" & (NewBOMPos + Count)).Value = NewQtyInt
                        End If

                        DescFixed = "Yes"
                    End If

                    With .Range("A" & (NewBOMPos + Count) & ":M" & (NewBOMPos + Count))
                        With .Interior
                            .ColorIndex = 45
                            .Pattern = Constants.xlSolid
                        End With
                    End With

                    If GenInfo3233.RevNo = 0 Then
                        .Range("Q" & (NewBOMPos + Count)).Value = "1"
                        .Range("A" & (NewBOMPos + Count) & ":L" & (NewBOMPos + Count)).Interior.ColorIndex = 4
                    End If

                    With .Range("A" & ((NewBOMPos + 1) + Count) & ":M" & ((NewBOMPos + 1) + Count))
                        With .Interior
                            .ColorIndex = 45
                            .Pattern = Constants.xlSolid
                        End With
                    End With

                    If GenInfo3233.RevNo = 0 Then
                        .Range("Q" & ((NewBOMPos + 1) + Count)).Value = "1"
                        .Range("A" & ((NewBOMPos + 1) + Count) & ":L" & ((NewBOMPos + 1) + Count)).Interior.ColorIndex = 4
                    End If
                Else
                    .Range("M" & NewBOMPos).Value = "Found Standard Information"

                    With .Range("A" & NewBOMPos & ":M" & NewBOMPos)
                        With .Interior
                            .ColorIndex = 8
                            .Pattern = Constants.xlSolid
                        End With
                    End With

                    If GenInfo3233.RevNo = 0 Then
                        .Range("Q" & NewBOMPos).Value = "1"
                        .Range("A" & NewBOMPos & ":L" & NewBOMPos).Interior.ColorIndex = 4
                    End If

                    With .Range("A" & (NewBOMPos + 1) & ":M" & (NewBOMPos + 1))
                        With .Interior
                            .ColorIndex = 8
                            .Pattern = Constants.xlSolid
                        End With
                    End With

                    If GenInfo3233.RevNo = 0 Then
                        .Range("Q" & (NewBOMPos + 1)).Value = "1"
                        .Range("A" & (NewBOMPos + 1) & ":L" & (NewBOMPos + 1)).Interior.ColorIndex = 4
                    End If
                End If
            End With

            FirstTimeThru = "No"
            Count = (Count + 1)
        Next l

Err_GetStdInfo:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If GenInfo3233.UserName = "dlong" Then
                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If
            End If
        End If

    End Function

    Private Sub PathBox_Click(sender As Object, e As EventArgs) Handles PathBox.Click
        Dim FileNam2 As String
        Dim lReturn As Int64
        Dim UnderPos, MXPos, DwPos As Integer

        StartDir = System.Environment.SpecialFolder.Recent

        If StartDir = "0" Then
            StartDir = Me.PathBox.Text                          'StartDir = "C:\AdeptWork\"
        End If

        OpenFileDialog1.InitialDirectory = StartDir
        OpenFileDialog1.Filter = "AutoCAD Drawings(*.dwg)|*.dwg;"
        OpenFileDialog1.Title = "Select file to Open"
        OpenFileDialog1.FileName = "Select Drawing"           '-------DJL-12-18-2024
        lReturn = OpenFileDialog1.ShowDialog()

        FileNam = OpenFileDialog1.FileName                          'Has to be 64 bit Inorder for this to work.
        FileNam2 = OpenFileDialog1.SafeFileName
        NewDir = FileNam.Replace(FileNam2, "")

        GenInfo3233.JobDir = NewDir
        PathBox.Text = NewDir
        Me.PathBox.BackColor = System.Drawing.Color.White
        Me.BtnGetMWInfo.BackColor = System.Drawing.Color.White
        Me.DwgList.BackColor = System.Drawing.Color.GreenYellow

        sender = "Excel"
        SecondChk = "First"

        If Directory.Exists("K:\AWA\" & System.Environment.UserName & "\AdeptWork\") = False Then                            'DJL 9-17-2024                    
            ClosePrg("Excel", "First", StartAdept)
            'ClosePrg("senddmp", "First", StartAdept)
            'ClosePrg("senddmp", "Second", StartAdept)
        End If

        Dim Dir1 As DirectoryInfo = New DirectoryInfo(GenInfo3233.JobDir)

        'New Problem file names so long and many variables 208-25-001-512-DW-10D06-01._INNER TANK BOTTOM SKETCH PLATE NESTING.dwg
        'Look for "-DW-" and "._" then check next two Characters are MX, CA, or CH for standards.

        If DwgList.Items.Count = 0 Then
            'For Each DwgItem1 In Dir1.GetFiles("*.idw")     '-------DJL-08-07-2025     'Not required.
            '    If InStr(DwgItem1.Name, "MX") = 0 And InStr(DwgItem1.Name, "CH") = 0 Then       '-------DJL-08-06-2025      'If InStr(DwgItem1.Name, "MX") = 0 And InStr(DwgItem1.Name, "CH") = 0 Then
            '        DwgList.Items.Add(DwgItem1)
            '    End If
            'Next DwgItem1
            For Each DwgItem1 In Dir1.GetFiles("*.dwg")                     'Found error below where Bottom layouts was left off due to naming had "CH"
                If InStr(DwgItem1.Name, "-DW-") > 0 Then     '-------DJL-08-07-2025      'Pittsburgh numbering solution
                    DwPos = InStr(DwgItem1.Name, "-DW-")
                    UnderPos = InStr(DwgItem1.Name, "_")
                    FileNam = Mid(DwgItem1.Name, (UnderPos + 1), Len(DwgItem1.Name))
                Else
                    If InStr(DwgItem1.Name, "_MX") > 0 Then     '-------DJL-08-07-2025      'Tulsa numbering solution
                        MXPos = InStr(DwgItem1.Name, "_MX")
                        FileNam = Mid(DwgItem1.Name, (MXPos + 1), Len(DwgItem1.Name))
                    Else
                        If InStr(DwgItem1.Name, "-MX") > 0 Then     '-------DJL-08-07-2025      'Tulsa numbering solution
                            MXPos = InStr(DwgItem1.Name, "-MX")
                            FileNam = Mid(DwgItem1.Name, (MXPos + 1), Len(DwgItem1.Name))
                        End If
                    End If
                End If

                If InStr(1, FileNam, "MX") <> 1 And InStr(1, FileNam, "CH") <> 1 Then     '-------DJL-08-07-2025      'Remove standards until primary drawings are read.       'If InStr(DwgItem1.Name, "MX") = 0 And InStr(DwgItem1.Name, "CH") = 0 Then 
                    If InStr(1, FileNam, "CA") <> 1 Then     '-------DJL-08-07-2025
                        DwgList.Items.Add(DwgItem1)
                    End If
                End If
            Next DwgItem1
        End If

        If DwgList.Items.Count > 0 Then
            Me.BtnAddAll.Enabled = True
            Me.BtnAdd.Enabled = True
            Me.BtnRemove.Enabled = True
            Me.BtnClear.Enabled = True
        Else
            Me.BtnAddAll.Enabled = False
            Me.BtnAdd.Enabled = False
            Me.BtnRemove.Enabled = False
            Me.BtnClear.Enabled = False
        End If

        DwgList.Sorted = True
        Me.Show()
        Me.BringToFront()
        Me.Refresh()
    End Sub

    Public Function ClosePrg(ByVal sender As System.Object, ByVal SecondChk As String, ByVal StartAdept As Boolean)
        Dim myProcesses() As Process
        Dim instance As Process
        Dim Title, Msg, Style, Response As Object

        'ChkForAdditional:  'For some reason if Application does not close, this thoughs prg into endless loop.

        myProcesses = Process.GetProcessesByName(sender)                '-------Get Process if open example Excel.
        StartAdept = "False"

        If IsNothing(myProcesses) <> True Then
            For Each instance In myProcesses
                Select Case sender & SecondChk
                    Case "Adept" & "First"                            '-----------------------First Check
                        GoTo NextInstance
                    Case "acad" & "First"
                        MsgBox("Do you have any AutoCAD files open? If so save and close them then pick ok.")
                    Case "Excel" & "First"
                        '-------Looking for solution to remove message boxes in the back ground. & skip when users want to work on files in excel at the same time.
                        SecondChk = "Second"

                        Msg = "Do you have any Excel spreadsheets open if so please save and close them, Have you saved your work?"
                        Style = MsgBoxStyle.YesNo
                        Title = "Found Excel is open and need to make sure user saves files."
                        Response = MsgBox(Msg, Style, Title)

                        If Response = 6 Then
                            GoTo KillOpenFiles
                        Else
                            GoTo NextInstance
                        End If
                    Case "Word" & "First"
                        MsgBox("Word was closed due to Issues.")
                    Case "WinWord" & "First"
                        MsgBox("Word was closed due to Issues.")
                    Case "BulkBOM" & "First"
                        MsgBox("BulkBOM was closed due to Issues.")
                    Case "Adept" & "Second"                            '-----------------------Second Check
                        GenInfo3233.StartAdept = True
                    Case "acad" & "Second"
                    Case "Excel" & "Second"
                    Case "Word" & "First"
                        MsgBox("Word was closed due to Issues.")
                    Case "MatrixPrograms" & "First"
                        MsgBox("Matrix Programs was closed due to Issues.")
                    Case "Inventor" & "First"
                        SecondChk = "Second"
                        MsgBox("Inventor was closed due to Issues.")
                    Case "senddmp" & "First"
                        MsgBox("Inventor has had a Hard Crash, program will now try to recover.")
                        SecondChk = "Second"
                    Case "senddmp" & "Second"
                        MsgBox("Inventor has had a Hard Crash, program will now try to recover.")
                End Select
KillOpenFiles:
                instance.Kill()
                instance.CloseMainWindow()
                instance.Close()
NextInstance:
            Next instance
        End If

    End Function

    Function HandleErrSQL(ByVal PrgName As String, ByVal ErrNo As String, ByVal ErrMsg As String, ByVal ErrSource As String, ByVal PriPrg As String, ByVal ErrDll As String, ByVal DwgItem As String, ByVal PrgLineNo As String)
        Dim sqlConn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim sqlStr As String
        Dim ErrDate As Date
        Dim QuoteMkPos As Integer
        Dim UserName, ProgramNotes, ErrMsgPart1, ErrMsgPart2 As String

        sqlConn = New ADODB.Connection

        ErrDate = Now

        If IsNothing(GenInfo3233.UserName) = True Then
            GenInfo3233.UserName = System.Environment.UserName()
        End If

        UserName = GenInfo3233.UserName
        ProgramNotes = "VB Net 64bit for Inventor Programming Testing"

        QuoteMkPos = InStr(ErrMsg, Chr(39))

        While QuoteMkPos > 0                                        'Remove single quotes that cause database errors.
            ErrMsgPart1 = Mid(ErrMsg, 1, (QuoteMkPos - 1))
            ErrMsgPart2 = Mid(ErrMsg, (QuoteMkPos + 1), Len(ErrMsg))
            ErrMsg = (ErrMsgPart1 & ErrMsgPart2)

            QuoteMkPos = InStr(ErrMsg, Chr(39))
        End While

        '-----------------------------------------------------Save Errors to sql database
        sqlStr = "INSERT  INTO ErrCollection (PrimaryPrg, PrgName, ErrNo, ErrMsg, ErrDate, UserName, ProgramNotes, ErrSource, ErrDll, DwgName, [ProfessionalNotestoManagerson Err]) " &
        "VALUES ('" & PriPrg & "', '" & PrgName & "', '" & ErrNo & "', '" & ErrMsg & "', '" & ErrDate & "', '" & UserName & "', '" & ProgramNotes & "', '" & ErrSource & "', '" & ErrDll & "', '" & DwgItem & "', '" & PrgLineNo & "')"
        db_String = "Server=MTX16SQL09\Engineering;Database=HandleErrors;User=devDennis;Password=d3v3lop3r;Trusted_Connection=False"
        Dim ConJobLog As New SqlClient.SqlConnection
        ConJobLog = New SqlClient.SqlConnection
        ConJobLog.ConnectionString = db_String
        ConJobLog.Open()
        Dim command2 As New SqlCommand(sqlStr, ConJobLog)
        Dim Writer2
        Writer2 = command2.ExecuteReader
        ConJobLog.Close()

        ConJobLog = Nothing
        sqlConn = Nothing
        rs = Nothing
    End Function

    Function WriteToExcel(BOMList)
        '-------Move to new function-------WritetoExcel-------DJL-10-11-2023            
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           12/21/2023
        '-------Description:    After putting data on Excel Spreadsheet 'Sort Data is now done in the array per drawing     '-------DJL-06-27-2025      'sort the data
        '-------
        '-------Updates:        Description:
        '-------12-21-2023       Read Array and write to Excel what was collected from AutoCAD.     
        '-------                Remove Spreadsheet formation to speed up array sort.
        '-------                Must find away to sort complex Array's, instead of sorting excel
        '-------
        '-------                Produced sort for array's before this module and work 100%      '-------DJL-06-27-2025
        '-------07-02-2025      Produced new version that uses an array sort before spread sheet is produced.
        '-------                Below has been changed to just write to excel.
        '-------07-03-2025      Added new code to copy the Ship Mark to the parts that make up the assembly.
        '-------07-03-2025      Added new code to collect the shipping list information so that the Shipping List can be produced.
        '------------------------------------------------------------------------------------------------
        Dim i, j, k, ShipRowNo As Integer
        Dim DwgItem2, CurrentDwgNo, FirstDwg, GetDwgNo, GetRowNo, GetX, GetY, FoundDwgNo, FoundX, FoundY, FoundItem, TotalCnt, GetPrevShpMK As String
        Dim GetShpMK, GetPcMK As String
        Dim CntDwgsNotFound, StrLineNo, PrevCnt, CntStd As Integer
        Dim AcadOpen As Boolean
        'Dim ShpMkList(3, 1)

        On Error GoTo Err_WriteToExcel

        ProgressBar1.Value = 0
        Me.LblProgress.Text = "Writing Data Found on your Spread Sheets........Please Wait"          '-------DJL-06-27-2025     '"Sorting Data Found on your drawings........Please Wait"
        Me.Refresh()

        ProblemAt = "File not found for 2024"

        If Dir("K:\CAD\VBA\XLTSheets\ReadAutoCAD-Dwgs.xltm") <> vbNullString Then
            FileToOpen = "K:\CAD\VBA\XLTSheets\ReadAutoCAD-Dwgs.xltm"
            FileSaveAS = PathBox.Text & GenInfo3233.FullJobNo & "-BULKBOM-R" & Me.ComboBxRev.Text & ".xls"  '-------DJL-08-08-2025      'Moved     'FileSaveAS = PathBox.Text & "\" & GenInfo3233.FullJobNo & "ReadDwgsAutoCAD.xls"
            System.IO.File.Copy(FileToOpen, FileSaveAS, True)                           '-------DJL-08-08-2025      'Added
        End If

        If File.Exists(FileSaveAS) Then                                                 '-------DJL-08-08-2025
            Dim attributes As FileAttributes = File.GetAttributes(FileSaveAS)

            If (attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                File.SetAttributes(FileSaveAS, attributes And Not FileAttributes.ReadOnly)
            End If

            If (attributes And FileAttributes.Archive) = FileAttributes.Archive Then
                File.SetAttributes(FileSaveAS, attributes And Not FileAttributes.Archive)
            End If
        End If

        If Dir(FileSaveAS) <> "" Then                           '-------DJL-08-08-2025      'If Dir(FileToOpen) <> "" Then
            MainBOMFile = ExcelApp.Application.Workbooks.Open(FileSaveAS)               '-------DJL-08-08-2025      'MainBOMFile = ExcelApp.Application.Workbooks.Open(FileToOpen)
        End If

        'FileToOpen = "K:\CAD\VBA\XLTSheets\BOM-New-1-15-2024.xltm"          '-------DJL-08-07-2025      'Not Required
        'OldFileNam = Me.PathBox.Text            '-------DJL-08-07-2025
        'CopyBOMFile(OldFileNam, RevNo, ExcelApp)          '-------DJL-08-08-2025  'Not required anymore.    'Bill Sieg's machine is having problems open excel file and writing to it in the later part of progrm moving save spreadshet to here to see if this is the problem. 

        Workbooks = ExcelApp.Workbooks
        WorkShtName = "BulK BOM"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name

        '-------DJL-07-14-2025  'No need to write out BOM then collect from BOM and write to Shipping List.
        'WorkShtName = "Shipping List"
        'ShippingWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        'ShipRowNo = 44

        FileToOpen = "BulK BOM"
        ExcelApp.Visible = True
        TotalCnt = (UBound(BOMList, 2) - 1)

        'Minimize Excel so user can see dialogs from program
        ExcelApp.WindowState = XlWindowState.xlMinimized
        With BOMWrkSht                          'Want to remove this part of writing to excel to sort.-------DJL-2-5-2024

            '-------DJL-10-11-2023-------Need to move this to a later point.
            For i = 1 To (UBound(BOMList, 2) - 1)                           '-------Look at doing this later and using Array's for everything.
                RowNo = i + 4
                ProgressBar1.Maximum = (UBound(BOMList, 2) - 1)

                If RowNo = "5" Then           'Look at speeding up process by coping formatted lines   --- To Speed up Process
                    BOMWrkSht.Activate()
                    .Rows(RowNo & ":" & RowNo).Select()                         '-------DJL-07-03-2025
                    .Rows(RowNo & ":" & RowNo).Insert()                         '-------DJL-07-03-2025
                    .Range("AA3").Value = GenInfo3233.Customer                  '-------DJL-07-03-2025
                Else
                    .Rows((RowNo - 1) & ":" & (RowNo - 1)).Select()
                    .Rows(RowNo & ":" & RowNo).Insert()             '-------DJL-07-03-2025      '.Rows((RowNo - 1) & ":" & (RowNo - 1)).Insert()
                End If

                If InStr(BOMList(1, i), "Delete") > 0 Then
                    GoTo DeleteDup
                End If

                GetPcMK = BOMList(5, i)                            '-------DJL-07-03-2025      'GetShpMK = BOMList(5, i)
                GetShpMK = BOMList(3, i)                            '-------DJL-07-03-2025     'GetPcMK = BOMList(3, i)

                If GetShpMK = "" And GetPcMK <> "" Then                'If BOMList(5, i) = "" & BOMList(3, i) <> "" Then
                    .Range("A" & RowNo).Value = BOMList(1, i) & "_" & GetPcMK
                Else
                    If GetShpMK <> "" Then
                        .Range("A" & RowNo).Value = BOMList(1, i) & "_" & GetShpMK
                    Else
                        .Range("A" & RowNo).Value = BOMList(1, i)
                    End If
                End If

                .Range("B" & RowNo).Value = BOMList(1, i)
                .Range("C" & RowNo).Value = BOMList(2, i)

                If GetPcMK = "" And GetPcMK = "-" Then                            '-------DJL-07-03-2025    'If GetPcMK = "" Then
                    .Range("D" & RowNo).Value = GetPrevShpMK
                Else
                    If BOMList(3, i) <> "" Then                                     '-------DJL-07-03-2025
                        GetPrevShpMK = BOMList(3, i)
                        .Range("D" & RowNo).Value = BOMList(3, i)
                    Else
                        .Range("D" & RowNo).Value = GetPrevShpMK
                    End If

                    '-------DJL-07-14-2025  'No need to write out BOM then collect from BOM and write to Shipping List.
                    'If InStr(BOMList(3, i), "SR") = 1 And BOMList(7, i) <> "" Then                          '-------DJL-07-14-2025      'If InStr(GetShipMk, "SR") = 1 Then
                    '    BOMList(7, i) = "SHELL PLATE " & BOMList(3, i) & " " & BOMList(7, i)      '-------DJL-07-14-2025      'BOMList(7, i) = "SHELL PLATE " & GetShipMk & BOMList(7, i)
                    'End If

                    '-------DJL-07-14-2025  'No need to write out BOM then collect from BOM and write to Shipping List.
                    'If InStr(GetPrevShpMK, "SR") = 1 And BOMList(7, i) <> "" Then
                    '    ShippingWrkSht.Activate()

                    '    With ShippingWrkSht
                    '        .Range("H" & ShipRowNo).Value = "SHELL PLATE " & BOMList(3, i) & " " & BOMList(7, i)      '-------DJL-07-14-2025  
                    '    End With
                    'Else
                    '    If InStr(GetPrevShpMK, "SR") = 1 Then
                    '        ShippingWrkSht.Activate()

                    '        With ShippingWrkSht
                    '            .Range("H" & ShipRowNo).Value = "SHELL PLATE " & BOMList(3, i) & " " & BOMList(6, i)      '-------DJL-07-14-2025  
                    '        End With
                    '    End If
                    'End If

                    '-------DJL-07-14-2025  'No need to write out BOM then collect from BOM and write to Shipping List.
                    '-------DJL-07-14-2025      'Why not Just write it to Spreadsheet instead of collecting it again.
                    'If BOMList(3, i) <> "" Then
                    '    ShipRowNo = (ShipRowNo + 1)
                    '    ShippingWrkSht.Activate()
                    '    With ShippingWrkSht

                    '        If ShipRowNo = "45" Then           'Look at speeding up process by coping formatted lines   --- To Speed up Process
                    '            .Rows(ShipRowNo & ":" & ShipRowNo).Select()                         '-------DJL-07-14-2025
                    '            .Rows((ShipRowNo + 1) & ":" & (ShipRowNo + 1)).Insert()
                    '        Else
                    '            .Rows((ShipRowNo - 1) & ":" & (ShipRowNo - 1)).Select()
                    '            .Rows(ShipRowNo & ":" & ShipRowNo).Insert()
                    '        End If

                    '        .Range("C" & ShipRowNo).Value = (GenInfo3233.FullJobNo & "/" & GenInfo3233.Customer)       'Job Number/Customer
                    '        .Range("D" & ShipRowNo).Value = BOMList(1, i)       'CurrentDwgNo
                    '        .Range("E" & ShipRowNo).Value = BOMList(2, i)       'CurrentDwgRev
                    '        .Range("F" & ShipRowNo).Value = BOMList(3, i)       'Get2DShipMk
                    '        .Range("G" & ShipRowNo).Value = BOMList(4, i)       'Get2DShipQty
                    '        '.Range("E" & ShipRowNo).Value = BOMList(5, i)      'GetPartNo

                    '        If InStr(GetPrevShpMK, "SR") = 1 Then
                    '            .Range("H" & ShipRowNo).Value = "SHELL PLATE " & BOMList(3, i) & " " & BOMList(7, i)      '-------DJL-07-14-2025        'GetShipDesc
                    '        Else
                    '            .Range("H" & ShipRowNo).Value = BOMList(6, i)       'GetShipDesc
                    '        End If

                    '        '.Range("E" & ShipRowNo).Value = BOMList(7, i)      'GetDesc
                    '        '.Range("F" & ShipRowNo).Value = BOMList(8, i)      '"Yes"
                    '        .Range("K" & ShipRowNo).Value = BOMList(9, i)       'GetInv1
                    '        .Range("L" & ShipRowNo).Value = BOMList(10, i)      'GetInv2

                    '        If BOMList(11, i) = vbNullString Or Mid(BOMList(11, i), 1, 1) = " " Then
                    '            .Range("M" & ShipRowNo).Value = BOMList(12, i) & Chr(10) & BOMList(13, i)
                    '        Else
                    '            .Range("M" & ShipRowNo).Value = BOMList(11, i)
                    '        End If

                    '        '.Range("M" & ShipRowNo).Value = BOMList(11, i)      'GetMat     or (GetMat2 & " " & GetMat3)
                    '        '.Range("F" & ShipRowNo).Value = BOMList(12, i)      '
                    '        '.Range("G" & ShipRowNo).Value = BOMList(13, i)     'GetLen
                    '        .Range("N" & ShipRowNo).Value = BOMList(14, i)      'GetWt
                    '        '.Range("E" & ShipRowNo).Value = BOMList(15, i)     'CurrentDwgNo
                    '        '.Range("F" & ShipRowNo).Value = BOMList(16, i)     '
                    '        '.Range("G" & ShipRowNo).Value = BOMList(17, i)      'InsertionPT(1).ToString
                    '        '.Range("D" & ShipRowNo).Value = BOMList(18, i)      'GetProc
                    '        '.Range("E" & ShipRowNo).Value = BOMList(19, i)      '(InsPt0)
                    '        '.Range("F" & ShipRowNo).Value = BOMList(20, i)      'ShipRowNo
                    '    End With

                    '    'ShpMkList(0, UBound(ShpMkList, 2)) = BOMList(6, i)      'GetPartDesc        '-------DJL-07-03-2025      'ShpMkList(0, UBound(ShpMkList, 2)) = BOMList(7, i)
                    '    'ShpMkList(1, UBound(ShpMkList, 2)) = BOMList(1, i)      'GetDwgNo
                    '    'ShpMkList(2, UBound(ShpMkList, 2)) = BOMList(3, i)      'GetShipMk
                    '    'ShpMkList(3, UBound(ShpMkList, 2)) = BOMList(5, i)      'GetShopMk
                    '    'ReDim Preserve ShpMkList(3, UBound(ShpMkList, 2) + 1)
                    'End If
                End If

                BOMWrkSht.Activate()
                .Range("E" & RowNo).Value = BOMList(5, i)
                .Range("F" & RowNo).Value = BOMList(4, i)                           'Will always be BOMList(4, i)       '-------DJL-10-11-2023

                'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                'GetMat = BOMList(11, i)
                'GetMat2 = BOMList(12, i)
                'GetMat3 = BOMList(13, i)

                '-------Should below be looking at 6 and 7 instead of 7 and 8-------DJL-10-11-2023
                If BOMList(7, i) = vbNullString Or BOMList(7, i) = " " Then
                    If InStr(1, BOMList(6, i), "%%D") <> 0 And BOMList(6, i) <> Nothing Then
                        GetPartDesc = BOMList(6, i)
                        GetPartDesc = GetPartDesc.Replace("%%D", " DEG.")
                    Else
                        GetPartDesc = BOMList(6, i)
                    End If

                    'If IsNothing(GetLen) = False And GetLen <> "" Then         'Not required in AutoCAD
                    '    GetPartDesc = GetPartDesc & " x " & GetLen
                    'End If

                    'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                    '        GetPartDesc = GetPartDesc & " - " & GetMat2 & " " & GetMat3
                Else
                    If InStr(1, BOMList(7, i), "%%D") <> 0 Then
                        GetPartDesc = BOMList(7, i)
                        GetPartDesc = GetPartDesc.Replace("%%D", " DEG.")
                    Else
                        GetPartDesc = BOMList(7, i)
                    End If

                    'If IsNothing(GetLen) = False And GetLen <> "" Then                         '-------DJL-07-03-2025      'Length is in the description for AutoCAD.
                    '    GetPartDesc = GetPartDesc & " x " & GetLen
                    'End If

                    'Per request from Trevor Ruffin do not need to add Material to Description anymore.-------DJL 4-18-2024
                End If

                .Range("G" & RowNo).Value = GetPartDesc
                'BOMList(7, i) = GetPartDesc                         '-------DJL-07-03=2024      'Not Required

                Inv1 = BOMList(9, i)                                'Found problem were user are entering a space before the Inventory number not 
                Inv1 = LTrim(Inv1)                                  'allowing the program to find the matching Inventory number on the standards sheet.
                Inv1 = RTrim(Inv1)                                  'Modified the program here to fix issue.   DJL 03/21/2013

                .Range("H" & RowNo).Value = Inv1
                'BOMList(9, i) = Inv1                            '-------DJL-07-03-2025      'Not required

                .Range("I" & RowNo).Value = BOMList(10, i)

                If BOMList(11, i) = vbNullString Or Mid(BOMList(11, i), 1, 1) = " " Then
                    .Range("J" & RowNo).Value = BOMList(12, i) & Chr(10) & BOMList(13, i)
                Else
                    .Range("J" & RowNo).Value = BOMList(11, i)
                End If

                .Range("K" & RowNo).Value = BOMList(14, i)
                '.Range("N" & RowNo).NumberFormat = "@"                         '-------DJL-07-03-2025      'Not Required
                '.Range("N" & RowNo).Value = BOMList(15, i)
                '.Range("O" & RowNo).NumberFormat = "General"                         '-------DJL-07-03-2025      'Not Required
                '.Range("O" & RowNo).Value = BOMList(16, i)
                '.Range("P" & RowNo).NumberFormat = "General"                         '-------DJL-07-03-2025      'Not Required
                '.Range("P" & RowNo).Value = BOMList(0, i)
                '.Range("Q" & RowNo).NumberFormat = "General"                         '-------DJL-07-03-2025      'Not Required
                '.Range("Q" & RowNo).Value = BOMList(17, i)
                .Range("R" & RowNo).Value = BOMList(8, i)                      'See if Standard needs to be found on some references the parts are all listed out. Example Job 9318-0206_18B parts are simular but not the same is a reference only.
                '.Range("W" & RowNo).Value = BOMList(13, i)                          'Length        '-------DJL-07-03-2025      'Not Required for AutoCAD.
                .Range("W" & RowNo).Value = BOMList(18, i)                      'Procurement        '-------DJL-07-03-2025      '.Range("X" & RowNo).Value = BOMList(18, i)
                '.Range("Y" & RowNo).NumberFormat = "@"                         '-------DJL-07-03-2025      'Not Required
                '.Range("Y" & RowNo).Value = BOMList(19, i)
                '.Range("Z" & RowNo).Value = BOMList(0, i)                         '-------DJL-07-03-2025      'Not Required
                'BOMList(20, i) = RowNo                             '-------Replaced at start of Array so program knows what item it is collecting Standards for.
                .Range("AA" & RowNo).Value = BOMList(20, i)           'Below the sort is fixed and the recno must equal the sort order.

DontAddBlankLines:  '---------------Do not Add Blank BOM Lines:
DeleteDup:
                ProgressBar1.Value = i
            Next i

        End With
Err_WriteToExcel:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            DwgItem2 = CurrentDwgNo
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException
            Test = DwgItem2

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

            If ErrNo = 9 And InStr(ErrMsg, "Index was outside the bounds of the array") > 0 Then
                Resume Next
            End If

            If ErrNo = -2147418111 And InStr(ErrMsg, "Call was rejected by callee") Then
                System.Threading.Thread.Sleep(25)
                Resume
            End If

            If ErrNo = -2147417848 And InStr(ErrMsg, "The object invoked has disconnected") > 0 Then
                AcadApp = GetObject(, "AutoCAD.Application")
                System.Threading.Thread.Sleep(25)

                If ProblemAt = "CloseDwg" Then
                    Resume Next
                Else
                    Resume
                End If
            End If

            If ErrNo = -2145320885 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next                     'Layout.dvb was not found to be loaded
            End If

            If ErrNo = -2145320924 And InStr(ErrMsg, "is not found.") < 0 Then
                'DwgItem = VarSelArray(z)
                Resume
            End If

            CntDwgsNotFound = InStr(ErrMsg, "not a valid drawing")

            If ErrNo = -2145320825 And CntDwgsNotFound > 0 Then
                Sapi.Speak("AutoCAD found a bad Drawing, " & DwgItem & ", Going to next drawing.")
                MsgBox("AutoCAD found a bad Drawing, " & DwgItem & ", Going to next drawing.")
                BadDwgFound = "Yes"
                CntDwgsNotFound = 0
                Resume Next
            End If

            If ErrNo = 91 And ErrMsg = "Problem in unloading DVB file" Then
                Resume Next                     'Layout.dvb was not found to be loaded
            End If

            If ErrNo = 462 And Mid(ErrMsg, 1, 29) = "The RPC server is unavailable" Then
                Information.Err.Clear()
                AcadApp = CreateObject("AutoCAD.Application")
                AcadOpen = False
                Resume
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            LenPrgLineNo = (Len(PrgLineNo))
            PrgLineNo = Mid(PrgLineNo, 1, (LenPrgLineNo - 2))

            HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2, PrgLineNo)                         'DJL-10-11-2023-------HandleErrSQL(PrgName + " @ line " + st.GetFrame(3).GetFileLineNumber().ToString, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem2)

            If ErrNo = -2145320900 And ErrMsg = "Failed to get the Document object" Then
                If FirstDwg = "NotFound" Then
                    AcadApp.Application.Documents.Add()
                    Resume
                End If
            End If

            If GenInfo3233.UserName = "dlong" Then
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If

                MsgBox(ErrMsg)
                Stop
                Resume
            Else
                ExceptPos = 0
                SearchException = "Exception"
                ExceptPos = InStr(ErrMsg, 1)
                If ExceptPos > 0 Then
                    CntExcept = (CntExcept + 1)
                    If CntExcept < 6 Then
                        Resume
                    End If
                End If

                If ErrNo = -2147418113 And ErrMsg = "Internal application error." Then
                    Information.Err.Clear()
                    AcadApp = CreateObject("AutoCAD.Application")
                    AcadOpen = False
                    Resume
                End If
            End If
        End If

    End Function

End Class