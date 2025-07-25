Option Strict Off           'Option Strict On       '-------DJL-12-17-2024 turned it back off to remove errors Option Strict On does not allow for late binding.
Option Explicit On
Imports Microsoft.Office.Interop.Excel

Friend Class SaveAsFilename
    Inherits System.Windows.Forms.Form
    Dim WorkShtName, PriPrg, ErrNo, ErrMsg, ErrSource, ErrDll, ErrLastLineX, PrgName, ProbPart As String
    Dim ErrException As System.Exception
    Dim PurchaseProb As Boolean
    Public FileDir As String
    Public WorkBooks2 As Microsoft.Office.Interop.Excel.Workbooks

    Private Sub SaveButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim BOMMnu As ReadDwgs
        BOMMnu = ReadDwgs
        PrgName = "SaveButton_Click"

        On Error GoTo Err_SaveButton_Click          '-------DJL-12-17-2024--Bill Sieg is getting Errors that I am not seeing Let's do some additional error tracking.

        If TextBox1.Text <> vbNullString Then
            PassFilename = TextBox1.Text
            ReadyToContinue = True
        Else
            ReadyToContinue = False
        End If

        Me.Close()
        FinishCopyBOMFile(PassFilename)                    '-------DJL-12-19-2024

Err_SaveButton_Click:
        ErrNo = Err.Number      '-------DJL-12-17-2024--Bill Sieg is getting Errors that I am not seeing Let's do some additional error tracking.

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

            BOMMnu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

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

    End Sub
    Private Sub CancelButton_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim BOMMnu As ReadDwgs
        BOMMnu = ReadDwgs
        PrgName = "CancelButton_Renamed"

        On Error GoTo Err_CancelButton_Renamed          '-------DJL-12-17-2024--Bill Sieg is getting Errors that I am not seeing Let's do some additional error tracking.

        PassFilename = "CancelProgram"
        ReadyToContinue = False
        Me.Close()

        FinishCopyBOMFile(PassFilename)                    '-------DJL-12-19-2024

Err_CancelButton_Renamed:
        ErrNo = Err.Number      '-------DJL-12-17-2024--Bill Sieg is getting Errors that I am not seeing Let's do some additional error tracking.

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

            BOMMnu.HandleErrSQL(PrgName, ErrNo, ErrMsg, ErrSource, PriPrg, ErrDll, DwgItem, PrgLineNo)

            If IsNothing(GenInfo3233.UserName) = True Then
                GenInfo3233.UserName = System.Environment.UserName()
            End If

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
    End Sub

    Private Sub StartForm() Handles Me.Activated
        Me.TextBox1.Focus()
        Me.LblFileName.Text = GenInfo3233.FileDir & GenInfo3233.BomListFileName
        Me.TextBox1.Text = GenInfo3233.FileDir & GenInfo3233.BomListFileName
        Me.Refresh()
    End Sub
End Class