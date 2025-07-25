Option Strict Off
Option Explicit On
Option Compare Text

Imports System
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Security.Permissions
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module File31_072
    Private Const BFFM_INITIALIZED As Short = 1
    Private Const BFFM_SETSELECTION As Integer = &H466
    Private Const BIF_DONTGOBELOWDOMAIN As Short = 2
    Private Const BIF_RETURNONLYFSDIRS As Short = 1
    Private Const MAX_PATH As Short = 260

    'Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByRef lpSecurityAttributes As ReadDwgs.SECURITY_ATTRIBUTES) As Integer           'Misc31_090.InputType2.SECURITY_ATTRIBUTES) As Integer
    Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Integer) As Integer
    'Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
    'Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Integer, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
    Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    'Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BrowseInfo) As Integer
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Integer, ByVal lpBuffer As String) As Integer
    Private Declare Function GetCurrentVBAProject Lib "vba332.dll" Alias "EbGetExecutingProj" (ByRef hProject As Integer) As Integer
    Private Declare Function GetAddr Lib "vba332.dll" Alias "TipGetLpfnOfFunctionId" (ByVal hProject As Integer, ByVal strFunctionId As String, ByRef lpfn As Integer) As Integer
    Private Declare Function GetFuncID Lib "vba332.dll" Alias "TipGetFunctionId" (ByVal hProject As Integer, ByVal strFunctionName As String, ByRef strFunctionId As String) As Integer
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Integer

    Dim WorkShtName, PriPrg, ErrNo, ErrMsg, ErrSource, ErrDll, ErrLastLineX, PrgName As String
    Dim ErrException As System.Exception
    'Public ExcelApp As Object
    'Public FirstTimeThru As String
    'Public FuncGetDataNew As String
    Public NewBulkBOM As Object
    'Public MainBOMFile As Microsoft.Office.Interop.Excel.Workbook
    'Public NewPlateBOM As Object
    'Public NewStickBOM As Object
    'Public NewPurchaseBOM As Object
    'Public OldBOMFile As Microsoft.Office.Interop.Excel.Workbook
    'Public OldBulkBOM As Object
    Public OldBulkBOMFile As String
    'Public OldPlateBOM As Object
    'Public OldStickBOM As Object
    'Public OldPurchaseBOM As Object
    'Public NewBOM As Object
    'Public OldBOM As Object
    'Public OldStdItems As Object
    'Public BOMType As String
    'Public BOMSheet As String
    'Public RowNo As String
    'Public RowNo2 As String
    'Public OldStdDwg As String
    'Public NewStdDwg As String
    Public ExceptionPos As Integer
    Public CallPos As Integer
    Public CntExcept As Integer
    'Public Count As Integer
    Public PassFilename As String
    Public ReadyToContinue As Boolean
    ''Public CBclicked As Boolean
    'Public errorExist As Boolean
    'Public AcadApp As Object
    'Public AcadDoc As Object
    'Public RevNo As String
    'Public RevNo2 As String
    'Public Continue_Renamed As Boolean
    'Public SortListing As Boolean
    'Public MatInch As Double
    'Public FoundDir As String
    'Public SearchException As String
    'Public ExceptPos As Integer
    ''Public ThisDrawing As AutoCAD.AcadDocument
    'Public LytHid As Boolean
    Public DwgItem As String
    'Public ErrFound As String
    'Public ErrorFoundNew As String
    'Public myarray As Object
    'Public myarray2 As Object

    Public GetFramesSrt
    Public PrgLineNo As String
    Public CntFrames As Integer
    Public ProblemAt As String                          '-------DJL-12-18-2024

    'Private Structure BrowseInfo
    '    Dim hwndOwner As Integer
    '    Dim pIDLRoot As Integer
    '    Dim pszDisplayName As String
    '    Dim lpszTitle As String
    '    Dim ulFlags As Integer
    '    Dim lpfnCallback As Integer
    '    Dim lParam As Integer
    '    Dim iImage As Integer
    'End Structure

    'Private Structure WIN32_FIND_DATA
    '    Dim dwFileAttributes As Integer
    '    Dim ftCreationTime As FILETIME
    '    Dim ftLastAccessTime As FILETIME
    '    Dim ftLastWriteTime As FILETIME
    '    Dim nFileSizeHigh As Integer
    '    Dim nFileSizeLow As Integer
    '    Dim dwReserved0 As Integer
    '    Dim dwReserved1 As Integer
    '    <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=MAX_PATH)> Public cFileName() As Char
    '    <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=14)> Public cAlternate() As Char
    'End Structure

    Private Structure OPENFILENAME
        Dim lStructSize As Integer
        Dim hwndOwner As Integer
        Dim hInstance As Integer
        Dim lpstrFilter As String
        Dim lpstrCustomFilter As String
        Dim nMaxCustFilter As Integer
        Dim nFilterIndex As Integer
        Dim lpstrFile As String
        Dim nMaxFile As Integer
        Dim lpstrFileTitle As String
        Dim nMaxFileTitle As Integer
        Dim lpstrInitialDir As String
        Dim lpstrTitle As String
        Dim flags As Integer
        Dim nFileOffset As Short
        Dim nFileExtension As Short
        Dim lpstrDefExt As String
        Dim lCustData As Integer
        Dim lpfnHook As Integer
        Dim lpTemplateName As String
    End Structure

    Public currentDir As String

    Function GetFile(ByRef startdir As String) As String

        Dim OpenFile As OPENFILENAME
        Dim lReturn As Integer
        Dim sFilter As String
        Dim s As String

        OpenFile.lStructSize = Len(OpenFile)
        sFilter = "Excel Worksheet(*.xls)" & Chr(0) & "*.xls" & Chr(0)
        OpenFile.lpstrFilter = sFilter
        OpenFile.nFilterIndex = 1
        OpenFile.lpstrFile = New String(Chr(0), 257)
        OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
        OpenFile.lpstrFileTitle = OpenFile.lpstrFile
        OpenFile.nMaxFileTitle = OpenFile.nMaxFile
        OpenFile.lpstrInitialDir = startdir
        OpenFile.lpstrTitle = "Select file to Open"
        lReturn = GetOpenFileName(OpenFile)

        If lReturn = 0 Then
            GetFile = ""
        Else
            s = Trim(OpenFile.lpstrFile)
            If InStr(s, Chr(0)) Then s = Left(s, InStr(s, Chr(0)) - 1)
            GetFile = s
            OldBulkBOMFile = GetFile
        End If
    End Function

    Function CopyBOMFile(ByVal OldFileNam As String, ByVal RevNo As String) As Object
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           Sometime before 4/2/2024
        '-------Description:    Copy information collected to Spreadsheet
        '-------
        '-------Updates:        Description:
        '-------12-18-2024      Looking for error Bill Sieg is talking about in Ticket-368442 Ship List/BBOM on Windows 10 machine   
        '-------                
        '-------                
        '------------------------------------------------------------------------------------------------
        Dim Worksheets As Object
        Dim FileDir, JobNo, BomListRev, BomListFileName, Test, SearchSlash, FirstPart, SecondPart As String
        Dim BOMMnu As ReadDwgs
        Dim SaveAsFilename As SaveAsFilename
        Dim BOMWrkSht As Worksheet
        Dim WorkSht As Worksheet
        Dim Workbooks As Microsoft.Office.Interop.Excel.Workbooks
        Dim SlashPos, PrevSlashPos As Integer

        BOMMnu = ReadDwgs
        PrgName = "CopyBOMFile-StartCopyFile"                            '-------DJL-12-18-2024

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

        On Error GoTo Err_CopyBOMFile

        PrgName = "CopyBOMFile-ExcelHookMade"                            '-------DJL-12-18-2024
        FileDir = OldFileNam
        Workbooks = ExcelApp.Workbooks
        WorkShtName = "BulK BOM"
        BOMWrkSht = Workbooks.Application.Worksheets(WorkShtName)
        WorkSht = Workbooks.Application.ActiveSheet
        WorkShtName = WorkSht.Name

        With BOMWrkSht
            JobNo = .Range("B3").Value
        End With

        PrgName = "CopyBOMFile-MakingSpdsht"                            '-------DJL-12-18-2024

        If JobNo = "Job No:" Then
            JobNo = OldFileNam
            Test = JobNo
            SearchSlash = "\"
            SlashPos = InStr(Test, SearchSlash)
            PrevSlashPos = 0

            While SlashPos > 0
                If PrevSlashPos > 0 Then
                    FirstPart = Mid(Test, 1, ((SlashPos - 1) + PrevSlashPos))
                    SecondPart = Mid(Test, (SlashPos + PrevSlashPos + 1), Len(Test))
                    PrevSlashPos = (PrevSlashPos + SlashPos)
                    SlashPos = InStr(SecondPart, SearchSlash)
                Else
                    FirstPart = Mid(Test, 1, (SlashPos - 1))
                    SecondPart = Mid(Test, (SlashPos + 1), (Len(Test) - (SlashPos + 2)))
                    PrevSlashPos = SlashPos
                    SlashPos = InStr(SecondPart, SearchSlash)
                End If
            End While

            JobNo = SecondPart
        End If

        PrgName = "CopyBOMFile-JobNoFound"                            '-------DJL-12-18-2024
        BomListRev = RevNo

        If JobNo = Nothing Then
            JobNo = InputBox("What is your Job Number?")
        End If

        BomListFileName = FileDir & JobNo & "-BULKBOM-R" & BomListRev & ".xls"
        PrgName = "CopyBOMFile-ProducingSpdSht"                            '-------DJL-12-18-2024
        BOMMnu.MainBOMFile.Worksheets.Copy()
        NewBulkBOM = BOMWrkSht

CheckFileName:
        Dim Style, Msg, Title As String
        Dim Response As Object

        If Dir(BomListFileName) <> vbNullString Then
            PrgName = "CopyBOMFile-GettingSpdShtNam"                            '-------DJL-12-18-2024

            Msg = "File " & BomListFileName & " already exists. Do you want to overwrite it?"
            Style = CStr(MsgBoxStyle.YesNo)
            Title = "Save Bulk BOM"

            'Next Look at "ReadyToContinue = False"
            'If SaveAsFilename.Focused = False Then                            '-------DJL-12-18-2024      'If SaveAsFilename.Enabled = False Then <> 1 Then    'If SaveAsFilename.AutoValidate = 1
            Response = MsgBox(Msg, CDbl(Style), Title)
            'End If

            If Response = MsgBoxResult.Yes Then
                MsgBox("After 30 seconds check your spreadsheet make sure it is not waiting on you to pick Save/Continue?")     '-------DJL-12-18-2024
                Kill((BomListFileName))
                Workbooks.Application.ActiveWorkbook.SaveAs(Filename:=BomListFileName, FileFormat:=XlFileFormat.xlWorkbookNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)
                PrgName = "CopyBOMFile-DelExistingSht"                            '-------DJL-12-18-2024
                ProgramFinished()                           '-------DJL-12-19-2024

            ElseIf Response = MsgBoxResult.No Then
                GenInfo3233.BomListFileName = JobNo & "-BULKBOM-R" & BomListRev & ".xls"            'BomListFileName
                GenInfo3233.FileDir = FileDir
                SaveAsFilename.Show()
                PrgName = "CopyBOMFile-ShowRenameForm"                            '-------DJL-12-18-2024
            End If
        Else
            PrgName = "CopyBOMFile-ShowRenForm3"                            '-------DJL-12-18-2024
            MsgBox("After 30 seconds check your spreadsheet make sure it is not waiting on you to pick Save/Continue?")     '-------DJL-12-18-2024
            Workbooks.Application.ActiveWorkbook.SaveAs(Filename:=BomListFileName, FileFormat:=XlFileFormat.xlWorkbookNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)

            ProgramFinished()                           '-------DJL-12-19-2024
        End If

Err_CopyBOMFile:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = "53" And InStr(ErrMsg, "No files found matching") > 0 Then
                Resume Next
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            BOMMnu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

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

    Public Function FinishCopyBOMFile(ByVal PassFileName As String)
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           12-18-2024
        '-------Description:    Finish program above when the user whats to put the file in a different location, or change the file name.
        '-------
        '-------Updates:        Description:
        '-------12-18-2024      Looking for error Bill Sieg is talking about in Ticket-368442 Ship List/BBOM on Windows 10 machine   
        '-------                
        '-------                'Dim Workbooks As Microsoft.Office.Interop.Excel.Workbooks
        '------------------------------------------------------------------------------------------------
        Dim BomListFileName, SearchSlash, FirstPart, SecondPart, Test, FileDir As String
        Dim SlashPos, PrevSlashPos As Integer
        Dim BOMMnu As ReadDwgs
        BOMMnu = ReadDwgs
        '-----------------------------------------------------------
        Dim Workbooks2 As Microsoft.Office.Interop.Excel.Workbooks      '-------DJL-12-19-2024

        Workbooks2 = ExcelApp.Workbooks      '-------DJL-12-19-2024

        '-------DJL-12-19-2024-Check to make sure user has a Valid directory srtuture.
        PrgName = "FinishCopyBOMFile-ChkDir"

        If Dir(PassFileName) = vbNullString Then
            Test = PassFileName
            SearchSlash = "\"
            SlashPos = InStr(Test, SearchSlash)
            PrevSlashPos = 0

            While SlashPos > 0
                If PrevSlashPos > 0 Then
                    FirstPart = Mid(Test, 1, ((SlashPos - 1) + PrevSlashPos))
                    SecondPart = Mid(Test, (SlashPos + PrevSlashPos + 1), Len(Test))
                    PrevSlashPos = (PrevSlashPos + SlashPos)
                    SlashPos = InStr(SecondPart, SearchSlash)
                Else
                    FirstPart = Mid(Test, 1, (SlashPos - 1))
                    SecondPart = Mid(Test, (SlashPos + 1), (Len(Test) - (SlashPos + 2)))
                    PrevSlashPos = SlashPos
                    SlashPos = InStr(SecondPart, SearchSlash)
                End If
            End While

            FileDir = PassFileName.Replace(SecondPart, "")

            If Dir(FileDir) = vbNullString Then

                If System.Environment.UserName = "bsieg" Then
                    PrgName = "FinishCopyBOMFile-DirIssueBillSieg"
                    MsgBox("Bill Sieg Directory does not exist Please create it in Windows exploder then click on the OK button.")
                Else
                    PrgName = "FinishCopyBOMFile-DirIssue"
                    MsgBox("User must create directory before trying to save file. Directory does not exist Please create it in Windows exploder then click on the OK button.")
                End If
            End If
        End If
        '-----------------------------------------------------------

        PrgName = "FinishCopyBOMFile-Gettingsave"

        On Error GoTo Err_FinishCopyBOMFile

        If PassFileName <> vbNullString And ReadyToContinue = True Then
            If Right(PassFileName, 4) = ".xls" Then
                MsgBox("After 30 seconds check your spreadsheet make sure it is not waiting on you to pick Save/Continue?")     '-------DJL-12-18-2024
                PrgName = "FinishCopyBOMFile-Gettingsave1"                             '-------DJL-12-18-2024
                Workbooks2.Application.ActiveWorkbook.SaveAs(Filename:=PassFileName, FileFormat:=XlFileFormat.xlWorkbookNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)            '-------DJL-12-19-2024
                ProgramFinished()                           '-------DJL-12-19-2024
            Else
                MsgBox("After 30 seconds check your spreadsheet make sure it is not waiting on you to pick Save/Continue?")     '-------DJL-12-18-2024
                PrgName = "FinishCopyBOMFile-Gettingsave2"                             '-------DJL-12-18-2024
                Workbooks2.Application.ActiveWorkbook.SaveAs(Filename:=PassFileName, FileFormat:=XlFileFormat.xlWorkbookNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False, AddToMru:=True)            '-------DJL-12-19-2024
                ProgramFinished()                           '-------DJL-12-19-2024
            End If
        ElseIf PassFileName = "CancelProgram" And ReadyToContinue = False Then
            PrgName = "CopyBOMFile-PrgCancelByUser"                            '-------DJL-12-18-2024
            Exit Function
        Else
            PrgName = "CopyBOMFile-CheckFileName"                            '-------DJL-12-18-2024
            'GoTo CheckFileName         'Do not do this when user wants to use anothwer file name.  '-------DJL-12-18-2024
        End If

Err_FinishCopyBOMFile:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = "53" And InStr(ErrMsg, "No files found matching") > 0 Then
                Resume Next
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            BOMMnu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

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

    Public Function ProgramFinished()
        '------------------------------------------------------------------------------------------------
        '-------Creator:        Dennis J. Long
        '-------Date:           12-19-2024
        '-------Description:    Finish program above when the user whats to put the file in a different location, or change the file name.
        '-------
        '-------Updates:        Description:
        '-------12-19-2024      Looking for error Bill Sieg is talking about in Ticket-368442 Ship List/BBOM on Windows 10 machine   
        '-------                
        '-------                'Had to move this part and create location to end for three diferent paths.
        '------------------------------------------------------------------------------------------------
        Dim BOMMnu As ReadDwgs
        BOMMnu = ReadDwgs
        PrgName = "ProgramFinished"

        On Error GoTo Err_ProgramFinished

        PrgName = "StartButton_Click-Part33"
        ExcelApp.Application.Visible = True

        If GenInfo3233.StartAdept = True Then
            OpenPrg("Adept")
        End If

        MsgBox("Your Bulk BOM has been Created.")
        ExcelApp.Application.Visible = True

        PrgName = "StartButton_Click-Part34"
        BOMMnu.MainBOMFile.Close(False)
        PrgName = "StartButton_Click-Part35"
        BOMMnu.Close()

Err_ProgramFinished:
        ErrNo = Err.Number

        If ErrNo <> 0 Then
            PriPrg = "BulkBOMFab3D-Intent"
            ErrMsg = Err.Description
            ErrSource = Err.Source
            ErrDll = Err.LastDllError
            ErrLastLineX = Err.Erl
            ErrException = Err.GetException

            If ErrNo = "53" And InStr(ErrMsg, "No files found matching") > 0 Then
                Resume Next
            End If

            Dim st As New StackTrace(Err.GetException, True)
            CntFrames = st.FrameCount
            GetFramesSrt = st.GetFrames
            PrgLineNo = GetFramesSrt(CntFrames - 1).ToString
            PrgLineNo = PrgLineNo.Replace("@", "at")
            PrgLineNo = PrgLineNo.Replace("VbCrlf", "")
            PrgLineNo = PrgLineNo.Replace(Chr(15), "")

            BOMMnu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

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

    Public Function OpenPrg(ByVal sender As System.Object)
        Dim BOMGen As Object

        Select Case sender
            Case "Adept"
                BOMGen = Shell("C:\Program Files (x86)\Synergis\Adept10\Client\Adept.exe", AppWinStyle.NormalFocus)
            Case Else
                MsgBox("Program needs to be added.")
        End Select

    End Function
End Module