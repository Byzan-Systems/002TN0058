Imports System.IO
Imports System.Text
Imports System
Imports System.Data
Module GenrateOutput
    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As ClsShared
    Dim objFileValidate As ClsValidation
    Dim objBaseClass As ClsBase



    Public Function GenerateOutPutFile(ByRef dt As DataTable, ByVal strFileName As String) As Boolean
        Dim gstrA2Afile As String = String.Empty

        Try
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            objFileValidate = New ClsValidation("", My.Application.Info.DirectoryPath & "\settings.ini")

            FileCounter = objBaseClass.GetINISettings("General", "File Counter", My.Application.Info.DirectoryPath & "\settings.ini")

            If FileCounter <> "" Then
                FileCounter = FileCounter + 1
                If Len(FileCounter) < 3 Then
                    FileCounter = FileCounter.PadLeft(4, "0").Trim()
                    FileCounter = FileCounter.Substring(FileCounter.Length - 3, 3)
                End If

                strFileName = objFileValidate.IsJustAlpha(Path.GetFileNameWithoutExtension(strFileName), 10, "N")
                gstrOutputFile_EPAY = strFileName & Format(CDate(Now), "_ddMMyyhhmmss").ToString().Trim() & ".csv"

                If Generate_Output_CSV(dt, gstrOutputFile_EPAY) = True Then
                    GenerateOutPutFile = True
                Else
                    GenerateOutPutFile = False
                End If


                Call objBaseClass.SetINISettings("General", "File Counter", Val(FileCounter), My.Application.Info.DirectoryPath & "\settings.ini")

            End If
        Catch ex As Exception
            GenerateOutPutFile = True
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "GenerateOutPutFile")

        Finally
        End Try
    End Function


    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function

    Public Function Generate_Output_CSV(ByRef _dt As DataTable, ByVal strFileName As String) As Boolean

        Dim strData As String = ""
        Dim objStrmWriter As StreamWriter = Nothing


        If objBaseClass Is Nothing Then
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
        End If
        objFileValidate = New ClsValidation("", My.Application.Info.DirectoryPath & "\settings.ini")

        Try

            If (_dt IsNot Nothing) Then
                If (_dt.Rows.Count > 0) Then

                    objStrmWriter = New StreamWriter(strOutputFolderPath & "\" & strFileName)
                    objBaseClass.LogEntry("Generating Standard CSV Output File Started...")

                    strData = ""
                    For Each drCol As DataColumn In _dt.Columns
                        If (drCol.ColumnName.ToString().Trim() <> "TXN_NO" And drCol.ColumnName.ToString().Trim() <> "SUBTXN_NO" And drCol.ColumnName.ToString().Trim() <> "Reason") Then
                            strData = strData & """" & drCol.ColumnName.ToString().Trim() & """" & ","
                        End If
                    Next
                    strData = strData.Substring(0, strData.Length - 1)
                    objStrmWriter.WriteLine(strData, strFileName)

                    Dim Counter As Integer
                    Counter = 0
                    For Each drRow As DataRow In _dt.Rows
                        Counter += 1
                        strData = ""
                        For Inti As Int32 = 0 To drRow.ItemArray.Length - 4
                            'strData = strData & """" & Check_Comma(drRow.ItemArray(Inti).ToString().Trim() & """")
                            strData = strData & """" & drRow.ItemArray(Inti).ToString().Trim() & """" & ","
                        Next
                        strData = strData.Substring(0, strData.Length - 1)

                        If Counter < _dt.Rows.Count Then
                            objStrmWriter.WriteLine(strData, strFileName)
                        ElseIf Counter = _dt.Rows.Count Then
                            objStrmWriter.Write(strData, strFileName)
                        End If
                    Next

                    If Not objStrmWriter Is Nothing Then
                        objStrmWriter.Close()
                        objStrmWriter.Dispose()
                    End If
                    objBaseClass.LogEntry("Standard CSV Output File [" & strFileName & "] is  Generated Successfully")
                Else
                    objBaseClass.LogEntry("No Records Found to Generate CSV Output File.")
                End If
            End If
        Catch ex As Exception
            Generate_Output_CSV = False
            Call objBaseClass.Handle_Error(ex, "GenerateCSVOuput", Err.Number, "GenerateCSVOuput")
        End Try
    End Function

    'Public Function Generate_Output_CSV(ByRef _dt As DataTable, ByVal strFileName As String) As Boolean

    '    Dim strData As String = ""
    '    Dim objStrmWriter As StreamWriter = Nothing


    '    If objBaseClass Is Nothing Then
    '        objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
    '    End If
    '    objFileValidate = New ClsValidation("", My.Application.Info.DirectoryPath & "\settings.ini")


    '    Try

    '        If (_dt IsNot Nothing) Then
    '            If (_dt.Rows.Count > 0) Then

    '                objStrmWriter = New StreamWriter(strOutputFolderPath & "\" & strFileName)
    '                objBaseClass.LogEntry("Generating Standard CSV Output File Started...")

    '                strData = ""
    '                For Each drCol As DataColumn In _dt.Columns
    '                    If (drCol.ColumnName.ToString().Trim() <> "TXN_NO" And drCol.ColumnName.ToString().Trim() <> "SUBTXN_NO" And drCol.ColumnName.ToString().Trim() <> "Reason") Then
    '                        strData = strData & drCol.ColumnName.ToString().Trim() & ","
    '                    End If
    '                Next
    '                strData = strData.Substring(0, strData.Length - 1)
    '                objStrmWriter.WriteLine(strData, strFileName)

    '                Dim Counter As Integer
    '                Counter = 0
    '                For Each drRow As DataRow In _dt.Rows
    '                    Counter += 1
    '                    strData = ""
    '                    For Inti As Int32 = 0 To drRow.ItemArray.Length - 4
    '                        strData = strData & (drRow.ItemArray(Inti).ToString()) & ","
    '                    Next
    '                    strData = strData.Substring(0, strData.Length - 1)

    '                    If Counter < _dt.Rows.Count Then
    '                        objStrmWriter.WriteLine(strData, strFileName)
    '                    ElseIf Counter = _dt.Rows.Count Then
    '                        objStrmWriter.Write(strData, strFileName)
    '                    End If
    '                Next

    '                If Not objStrmWriter Is Nothing Then
    '                    objStrmWriter.Close()
    '                    objStrmWriter.Dispose()
    '                End If
    '                objBaseClass.LogEntry("Standard CSV Output File [" & strFileName & "] is  Generated Successfully")
    '            Else
    '                objBaseClass.LogEntry("No Records Found to Generate CSV Output File.")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Generate_Output_CSV = False
    '        Call objBaseClass.Handle_Error(ex, "GenrateOutput", Err.Number, "GenerateResponseISSUANCEOuput")
    '    End Try
    'End Function

    Public Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then

                Check_Comma = Chr(34) & strTemp & Chr(34) & ","
                ' Check_Comma = strTemp
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "Check_Comma")

        End Try
    End Function

End Module
