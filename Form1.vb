Imports System.IO
Imports System.Data
Public Class Form1
    Dim objBaseClass As ClsBase
    Dim objFileValidate As ClsValidation
    Dim objGetSetINI As ClsShared

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Timer1.Interval = 100
            Timer1.Enabled = True
            Generate_SettingFile()
        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form Load", "FrmLoad")
        End Try
    End Sub
    Private Sub Generate_SettingFile()

        Dim strConverterCaption As String = ""
        Dim strSettingsFilePath As String = My.Application.Info.DirectoryPath & "\settings.ini"

        Try
            objGetSetINI = New ClsShared
            '-Genereate Settings.ini File-
            If Not File.Exists(strSettingsFilePath) Then
                '-General Section-
                Call objGetSetINI.SetINISettings("General", "Date", Format(Now, "dd/MM/yyyy"), strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Audit Log", My.Application.Info.DirectoryPath & "\Audit", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Error Log", My.Application.Info.DirectoryPath & "\Error", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Input Folder", My.Application.Info.DirectoryPath & "\Input", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Output Folder", My.Application.Info.DirectoryPath & "\Output", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Report Folder", My.Application.Info.DirectoryPath & "\Report", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Validation", My.Application.Info.DirectoryPath & "\Validation\Validation.xls", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\UnSuccess", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Converter Caption", "PM Cares Converter ", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Process Output File Ignoring Invalid Transactions", "N", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "File Counter", "0", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "==", "==========================================", strSettingsFilePath) 'Separator

                '-Client Details Section-
                Call objGetSetINI.SetINISettings("Client Details", "Client Name", "PM CARES", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "Input Date Format", "dd/MM/yyyy", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("Client Details", "==", "====================================", strSettingsFilePath) 'Separator
            End If

            '-Get Converter Caption from Settings-
            If File.Exists(strSettingsFilePath) Then
                strConverterCaption = objGetSetINI.GetINISettings("General", "Converter Caption", strSettingsFilePath)
                If strConverterCaption <> "" Then
                    Text = strConverterCaption.ToString() & " - Version " & Mid(Application.ProductVersion.ToString(), 1, 3)
                Else
                    MsgBox("Either settings.ini file does not contains the key as [ Converter Caption ] or the key value is blank" & vbCrLf & "Please refer to " & strSettingsFilePath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End
                End If
            End If

        Catch ex As Exception
            MsgBox("Error" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error while Generating Settings File")
            End

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try

            Timer1.Interval = 1000
            Timer1.Enabled = False

            Conversion_Process()

            Timer1.Enabled = True

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form Load", "Timer1_Tick")
        End Try
    End Sub
    Private Sub Conversion_Process()
        Dim objfolderAll As DirectoryInfo

        Try
            If objBaseClass Is Nothing Then
                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            End If

            '-Get Settings-
            If GetAllSettings() = True Then
                MsgBox("Either file path is invalid or any key value is left blank in settings.ini file", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error In Settings")
                Exit Sub
            End If
            '-Process PAYMENT Input-
            objfolderAll = New DirectoryInfo(strInputFolderPath)

            If objfolderAll.GetFiles.Length = 0 Then
                objfolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started......")

                For Each file As FileInfo In objfolderAll.GetFiles("*")
                    If Mid(file.FullName, file.FullName.Length - 4, 5).ToString().ToUpper() = ".XLSX".ToUpper Or Mid(file.FullName, file.FullName.Length - 3, 4).ToString().ToUpper() = ".XLS" Then
                        objBaseClass.isCompleteFileAvailable(file.FullName)
                        gstrInputFile = file.Name
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("Input File [ " & file.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)
                        Process_Each(file.FullName)
                        objfolderAll.Refresh()
                    End If
                Next
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Conversion_Process")

        Finally
            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If
        End Try
    End Sub
    Private Function GetAllSettings() As Boolean

        Try
            GetAllSettings = False

            If Not File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                GetAllSettings = True
                MsgBox("Either settings.ini file does not exists or invalid file path" & vbCrLf & My.Application.Info.DirectoryPath & "\settings.ini", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            End If

            '-Audit Folder Path-
            If strAuditFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strAuditFolderPath) Then
                    Directory.CreateDirectory(strAuditFolderPath)
                    If Not Directory.Exists(strAuditFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Audit Log folder. Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", True)
                        End If
                        MsgBox("Invalid path for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                        Exit Function
                    End If
                End If
            End If

            '-Error Folder Path-
            If strErrorFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Error Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strErrorFolderPath) Then
                    Directory.CreateDirectory(strErrorFolderPath)
                    If Not Directory.Exists(strErrorFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Error Log folder. Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Error Log folder." & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Input Folder Path-
            If strInputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Input Folder " & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strInputFolderPath) Then
                    Directory.CreateDirectory(strInputFolderPath)
                    If Not Directory.Exists(strInputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Input Folder. Please check settings.ini file, the key as [ Input Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Input Folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "settings Error")
                    End If
                End If
            End If

            '-Output Folder Path-
            If strOutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Output folder" & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strOutputFolderPath) Then
                    Directory.CreateDirectory(strOutputFolderPath)
                    If Not Directory.Exists(strOutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Output Folder. Please check [ settings.ini ] file, the key as [ Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Archived Success Path-
            If strArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archived Success folder" & vbCrLf & "Please check settings.ini file, the key as [ Archived Success Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderSuc) Then
                    Directory.CreateDirectory(strArchivedFolderSuc)
                    If Not Directory.Exists(strArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archived Success Please check [ settings.ini ] file, the key as [ Archived Success Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archived Success Folder." & vbCrLf & "Please check settings.ini file, the key as [ Archived Success Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Archived Unsuccess Path-
            If strArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archived Unsuccess folder" & vbCrLf & "Please check settings.ini file, the key as [ Archived Unsuccess Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strArchivedFolderUnSuc)
                    If Not Directory.Exists(strArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archived Unsuccess Folder. Please check [ settings.ini ] file, the key as [ Archived Unsuccess Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archived Unsuccess Folder." & vbCrLf & "Please check settings.ini file, the key as [ Archived Unsuccess Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Validation File Path-
            If strValidationPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Validation file." & vbCrLf & "Please check settings.ini file, the key as [ Validation ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not File.Exists(strValidationPath) Then
                    GetAllSettings = True
                    If Not objBaseClass Is Nothing Then
                        objBaseClass.LogEntry("Error in settings.ini file, Validation file does not exist or invalid file path", True)
                    End If
                    MsgBox("Validation file does not exist or invalid file path" & vbCrLf & strValidationPath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                End If
            End If

            '-Report Folder Path-
            If strReportFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Report folder" & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strReportFolderPath) Then
                    Directory.CreateDirectory(strReportFolderPath)
                    If Not Directory.Exists(strReportFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Report Folder. Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Report Folder." & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If
        Catch ex As Exception
            GetAllSettings = True
            'MsgBox("Error - " & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Log Path from Settings.ini File")
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "GetAllSettings")

        End Try

    End Function
    Private Sub Process_Each(ByVal strInputFileName As String)
        Dim TrnProcSuc As Boolean
        Try
            gstrInputFolder = strInputFileName.Substring(0, strInputFileName.LastIndexOf("\"))
            gstrInputFile = strInputFileName.Substring(strInputFileName.LastIndexOf("\"))
            gstrInputFile = gstrInputFile.Replace("\", "")

            '-Conversion Process-

            objBaseClass.LogEntry("", False)
            objBaseClass.LogEntry("Process Started")
            objBaseClass.LogEntry("Reading Input File " & gstrInputFile, False)
            strEpayFileName = strInputFileName

            objFileValidate = New ClsValidation(strInputFileName, objBaseClass.gstrIniPath)

            If objFileValidate.CheckValidateFile(gstrInputFolder & "\" & gstrInputFile) = True Then

                objBaseClass.LogEntry("Input File Reading Completed Successfully", False)

                If (objFileValidate.DtUnSucInput.Rows.Count = 0) Or (strProceed.ToString().Trim().ToUpper() = "Y") Then
                    objBaseClass.LogEntry("Input File Validated Successfully", False)

                    If objFileValidate.DtInput.Rows.Count > 0 Then

                        objBaseClass.LogEntry("Output File Generation Process Started", False)

                        If GenerateOutPutFile(objFileValidate.DtInput, gstrInputFile) = True Then       ''Generating Output
                            TrnProcSuc = False
                            objBaseClass.LogEntry("Output File Generation process failed due to Error", True)
                            objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                        Else
                            TrnProcSuc = True
                            objBaseClass.LogEntry("Output Files is Generated Successfully", False)
                            objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderSuc & "\" & gstrInputFile)
                        End If

                    Else
                        TrnProcSuc = False
                        objBaseClass.LogEntry("No Valid Record present in Input File")
                        objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                    End If
                Else
                    TrnProcSuc = False
                    objBaseClass.LogEntry("No Valid Record present in Input File")
                    objBaseClass.FileMove(gstrInputFolder & "\" & gstrInputFile, strArchivedFolderUnSuc & "\" & gstrInputFile)
                End If

                If objFileValidate.DtUnSucInput.Rows.Count > 0 Then
                    objBaseClass.LogEntry("Input File contains following Discrepancies")
                    objBaseClass.LogEntry("Writing Instruction failed for Input File following ")

                    With objFileValidate.DtUnSucInput
                        For Each _dtRow As DataRow In .Rows
                            If _dtRow("Reason").ToString().Trim() <> "" Then
                            End If
                            objBaseClass.LogEntry(_dtRow("Reason").ToString)
                        Next
                    End With
                End If


                'Summary Report Writing
                Dim strSummaryFileName As String
                strSummaryFileName = Path.GetFileNameWithoutExtension(gstrInputFile)
                objBaseClass.LogEntry("[Writing Summary Report]")
                Call Summary_Report()
                objBaseClass.LogEntry("Summary Report File Generated Successfully")

            Else
                TrnProcSuc = False
                objBaseClass.LogEntry("Invalid Input File")
                objBaseClass.FileMove(gstrInputFolder & "\" & strRenameFile, strArchivedFolderUnSuc & "\" & strRenameFile)
            End If
            If TrnProcSuc <> False Then
                objBaseClass.LogEntry("Process Completed Successfully", False)
                objBaseClass.LogEntry("-------------------------------------------------------------------------------------", False)
            Else
                objBaseClass.LogEntry("Output File Generation Failed")
                objBaseClass.LogEntry("Proccess Terminated")
                objBaseClass.LogEntry("-------------------------------------------------------------------------------------", False)
            End If

        Catch ex As Exception
            objBaseClass.LogEntry("Proccess Terminated", True)
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "PM CARES Converter", "Process_Each")

        Finally
            If Not objFileValidate Is Nothing Then
                objBaseClass.ObjectDispose(objFileValidate.DtInput)
                objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)

                objFileValidate.Dispose()
                objFileValidate = Nothing
            End If
        End Try
    End Sub
    Private Sub Summary_Report()
        Dim strSumFileName As String
        Dim Count_SuccRec As Integer = 0
        Dim Count_UnSuccRec As Integer = 0
        Try
            strSumFileName = Path.GetFileNameWithoutExtension(gstrInputFile) & "_Summary_" & DateTime.Now.ToString("ddMMyyyyHHmmss") & ".txt"

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "[" & Format(Now, "dd-MM-yyyy hh:mm:ss") & "]")
            objBaseClass.WriteSummaryTxt(strSumFileName, "-----------------------------------------------------------")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Summary Report for Input File " & "[" & gstrInputFile & "]")
            objBaseClass.WriteSummaryTxt(strSumFileName, "-----------------------------------------------------------")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Total number of transactions : " & objFileValidate.DtInput.Rows.Count + objFileValidate.DtUnSucInput.Rows.Count & " ")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Successful Record Count : " & objFileValidate.DtInput.Rows.Count & " ")
            objBaseClass.WriteSummaryTxt(strSumFileName, "UnSuccessful Record Count : " & objFileValidate.DtUnSucInput.Rows.Count & " ")
            objBaseClass.WriteSummaryTxt(strSumFileName, "-----------------------------------------------------------")
        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "frm_Load", "Summary_Report")
        End Try

    End Sub
End Class

