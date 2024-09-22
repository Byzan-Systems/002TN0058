Imports System
Imports System.Data
Imports System.IO

Public Class ClsValidation

    Implements IDisposable

    Private ObjBaseClass As ClsBase         ''need to be dispose 
    Private DtValidation As DataTable       ''need to be dispose
    'Private DtSpCharValidation As DataTable       ''need to be dispose   ''Changed on 20-04-12
    Private DtMaster As DataTable

    Private DtTemp As DataTable             ''need to be dispose
    Public DtInput As DataTable             ''need to be dispose
    Public DtUnSucInput As DataTable        ''need to be dispose
    Public DtInputTemp As DataTable             ''need to be dispose
    Public DtSocietyMaster As DataTable

    Private StrFilePath As String
    Private ValidationPath As String
    'Private SpCharValidationPath As String   ''Changed on 20-04-12
    Public ErrorMessage As String
    Public DtInputResp As DataTable                     ''need to be dispose
    Public DtUnSucResp As DataTable                ''need to be dispose


    Public Sub New(ByVal _strFilePath As String, ByVal _SettINIPath As String)

        StrFilePath = _strFilePath

        Try
            ObjBaseClass = New ClsBase(_SettINIPath)
            ValidationPath = ObjBaseClass.GetINISettings("General", "Validation", _SettINIPath)

            DtInput = New DataTable("DtInput")
            DefineColumn(DtInput)
            DtUnSucInput = New DataTable("DtInput")
            DefineColumn(DtUnSucInput)


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Constructor")

        End Try

    End Sub
    'Private Sub DefineColumn(ByRef DtInput As DataTable)

    '    DtInput.Columns.Add(New DataColumn("Company Code"))   '0
    '    DtInput.Columns.Add(New DataColumn("Code"))  '1
    '    DtInput.Columns.Add(New DataColumn("Name"))  '2
    '    DtInput.Columns.Add(New DataColumn("Bank Account Number"))  '3
    '    DtInput.Columns.Add(New DataColumn("SWIFT Code"))  '4
    '    DtInput.Columns.Add(New DataColumn("Name of Bank"))  '5
    '    DtInput.Columns.Add(New DataColumn("City"))  '6
    '    DtInput.Columns.Add(New DataColumn("Postal Code"))  '7
    '    DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))   ''8
    '    DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   ''9
    '    DtInput.Columns.Add(New DataColumn("Reason"))   ''10
    'End Sub

    Private Sub DefineColumn(ByRef DtInput As DataTable)

        DtInput.Columns.Add(New DataColumn("txn_fullname"))   '0
        DtInput.Columns.Add(New DataColumn("txn_amount"))  '1
        DtInput.Columns.Add(New DataColumn("txn_email"))  '2
        DtInput.Columns.Add(New DataColumn("txn_mobile"))  '3
        DtInput.Columns.Add(New DataColumn("txn_pan"))  '4
        DtInput.Columns.Add(New DataColumn("txn_add1"))  '5
        DtInput.Columns.Add(New DataColumn("txn_zip"))  '6
        DtInput.Columns.Add(New DataColumn("txn_ret_bankTxnId"))  '7
        DtInput.Columns.Add(New DataColumn("txn_ret_transactionDate"))  '8
        DtInput.Columns.Add(New DataColumn("txn_ret_paymode"))  '9
        DtInput.Columns.Add(New DataColumn("txn_ret_bankRefNumber"))  '10
        DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))   ''11
        DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   ''12
        DtInput.Columns.Add(New DataColumn("Reason"))   ''13
    End Sub
    Public Function CheckValidateFile(ByVal strInputFileFolderPath As String) As Boolean

        Try
            If Not File.Exists(StrFilePath) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Input folder path is incorrect or File not found. [" & StrFilePath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            End If

            If File.Exists(strValidationPath) Then
                CheckValidateFile = Validate(strInputFileFolderPath)
            Else
                'Call ObjBaseClass.Handle_Error(New ApplicationException("Validation File path is incorrect. [" & ValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
                Call ObjBaseClass.Handle_Error(New ApplicationException("Check Validation,Mapping & Master File path is incorrect."), "ClsValidation", -123, "CheckValidateFile")
            End If

        Catch ex As Exception

            CheckValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckValidateFile")
        End Try

    End Function

    Public Function RemoveBlankRow(ByRef _DtTemp As DataTable)
        'To Remove Blank Row Exists in DataTable
        Dim blnRowBlank As Boolean

        Try
            For Each vRow As DataRow In _DtTemp.Rows
                blnRowBlank = True

                For intCol As Int32 = 0 To _DtTemp.Columns.Count - 1
                    If vRow.Item(intCol).ToString().Trim() <> "" Then
                        blnRowBlank = False
                        Exit For
                    End If
                Next

                If blnRowBlank = True Then
                    _DtTemp.Rows(vRow.Table.Rows.IndexOf(vRow)).Delete()
                End If

            Next
            _DtTemp.AcceptChanges()

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveBlankRow")

        End Try

    End Function
    Private Function Validate(ByVal strInputFileName As String) As Boolean
        Validate = False
        Dim DrValidOutputColumn() As DataRow = Nothing
        Dim StrDataRow(13) As String
        Dim ArrDataRow As Object
        Dim InputLineNumber As Int32
        Dim Amount As Double = 0
        Dim TXN_NO As Integer
        Dim SUBTXN_NO As Integer = 1
        Dim HardCode As Integer = 2
        Dim intPosField As Integer = 3
        Dim MandatoryPos As Integer = 4
        'Dim LengthPosMax As Integer = 5
        Dim CharType As Integer = 5

        Try
            DtValidation = ObjBaseClass.GetDataTable_ExcelSheet(strValidationPath, "")
            RemoveBlankRow(DtValidation)
            DrValidOutputColumn = DtValidation.Select("[SRNO] <> 0  ", "[SRNO]")

            DtTemp = New DataTable()
            DtTemp = ObjBaseClass.GetDataTable_ExcelSheet(gstrInputFolder & "\" & gstrInputFile, "")
            RemoveBlankRow(DtTemp)

            InputLineNumber = 1
            TXN_NO = 0
            SUBTXN_NO = 0

            If DtTemp.Rows.Count > 0 Then
                For Each dtRow In DtTemp.Rows
                    InputLineNumber += 1

                    ClearArray(StrDataRow)
                    ArrDataRow = dtRow.ItemArray()
                    TXN_NO += 1
                    SUBTXN_NO = 1

                    For intIndex As Int32 = 0 To DrValidOutputColumn.Length - 1

                        If Val(DrValidOutputColumn(intIndex)(intPosField).ToString().Trim()) <> 0 Then
                            StrDataRow(intIndex) = GetValueFormArray(ArrDataRow, DrValidOutputColumn(intIndex)(intPosField).ToString()).Trim()
                        Else
                            If StrDataRow(intIndex) = "~Error~" Then
                                StrDataRow(13) = "Input Line " & InputLineNumber & "  " & DrValidOutputColumn(intIndex)(3).ToString().Trim() & " Error in Input Position |"
                            Else
                                StrDataRow(intIndex) = ""
                            End If
                        End If



                    

                        ''Val(dtRow(3).ToString().Replace("-", "").Replace(",", ""))

                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "txn_amount".Trim().ToUpper() Then
                            If StrDataRow(intIndex).ToString() <> "" Then
                                Dim stramt As string=""
                                stramt = IsJustAlpha(StrDataRow(intIndex).ToString(), Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()), DrValidOutputColumn(intIndex)(CharType + 1).ToString().Trim()) 'Val(StrDataRow(intIndex).ToString().Replace("-", "").Replace(",", "")).ToString(".00")
                                StrDataRow(intIndex) = Val(stramt.ToString().Replace("-", "").Replace(",", "")).ToString(".00")
                            Else
                                StrDataRow(13) = StrDataRow(13) & "Input Line " & InputLineNumber & "  [" & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "] This is Mandatory Field & it is Blank |"
                            End If
                        End If


                        'txn_ret_paymode
                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "txn_ret_paymode".Trim().ToUpper() Then
                            Dim strtxn_ret_paymode As String = ""
                            strtxn_ret_paymode = StrDataRow(intIndex).ToString().Trim()
                            If strtxn_ret_paymode = "FT" Then
                                StrDataRow(intIndex) = "Funds Transfer"
                            ElseIf strtxn_ret_paymode = "IX" Then
                                StrDataRow(intIndex) = "InvoiceXpress"
                            Else
                                StrDataRow(intIndex) = StrDataRow(intIndex).ToString().Trim()
                            End If
                        End If

                        ''txn_ret_transactionDate
                        If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "txn_ret_transactionDate".Trim().ToUpper() Then
                            If StrDataRow(intIndex).ToString() <> "" Then
                                Dim SplitTime As String() = Nothing
                                'SplitTime = StrDataRow(intIndex).ToString().Trim().Split(" ")
                                Dim StrDATE As String = ""
                                Dim sDate As String = StrDataRow(intIndex).ToString()
                                Dim dDate As DateTime
                                If DateTime.TryParse(sDate, dDate) Then
                                    sDate = dDate.ToString("yyyy-MM-dd HH:mm:ss tt")
                                    ' Console.Write(sDate)
                                    SplitTime = sDate.ToString().Trim().Split(" ")
                                    If SplitTime.Length > 1 Then
                                        StrDataRow(intIndex) = SplitTime(0).ToString() & " " & SplitTime(1).ToString()
                                    Else
                                        StrDataRow(intIndex) = SplitTime(0).ToString() & " " & "00:00:00"
                                    End If
                                End If

                                'StrDATE = SplitTime(0).ToString.Trim
                                'If GetValidateDate(StrDATE) = True Then
                                '    If SplitTime.Length > 1 Then
                                '        StrDataRow(intIndex) = Format(CDate(StrDATE), "yyyy-MM-dd").ToString() & " " & SplitTime(1).ToString()
                                '    Else
                                '        StrDataRow(intIndex) = Format(CDate(StrDATE), "yyyy-MM-dd").ToString() & " " & "00:00:00"
                                '    End If
                                'End If
                            Else
                                StrDataRow(13) = StrDataRow(13) & "Input Line " & InputLineNumber & "  [" & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "] This is Mandatory Field & it is Blank |"
                            End If
                        End If
                        'If DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() = "txn_add1".Trim().ToUpper() Then
                        '    StrDataRow(intIndex) = StrDataRow(intIndex).Replace("""", "")
                        'End If

                        'If StrDataRow(intIndex).ToString().Trim() <> "" Then
                        '    StrDataRow(intIndex) = StrDataRow(intIndex).Replace("&", "And")
                        'End If

                        'HardCode Value
                        If DrValidOutputColumn(intIndex)(HardCode).ToString().Trim() <> "" Then
                            StrDataRow(intIndex) = DrValidOutputColumn(intIndex)(HardCode).ToString()
                        End If

                        '''''End here
                        ''Character Validation
                        If Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()) > 0 Then
                            StrDataRow(intIndex) = IsJustAlpha(StrDataRow(intIndex).Trim(), Val(DrValidOutputColumn(intIndex)(CharType).ToString().Trim()), DrValidOutputColumn(intIndex)(CharType + 1).ToString().Trim())
                        End If

                        '--------------Check mandatory 
                        If DrValidOutputColumn(intIndex)(MandatoryPos).ToString().Trim() = "M" And StrDataRow(intIndex).Trim() = "" And DrValidOutputColumn(intIndex)(1).ToString().Trim().ToUpper() <> "txn_amount".Trim().ToUpper() Then
                            StrDataRow(13) = StrDataRow(13) & "Input Line " & InputLineNumber & "  [" & DrValidOutputColumn(intIndex)(1).ToString().Trim() & "] This is Mandatory Field & it is Blank |"
                        End If

                    Next
                    StrDataRow(11) = TXN_NO
                    StrDataRow(12) = SUBTXN_NO
                    If StrDataRow(13).ToString().Trim() = "" Then
                        DtInput.Rows.Add(StrDataRow)
                    Else
                        DtUnSucInput.Rows.Add(StrDataRow)
                    End If
                Next
                Validate = True
            Else
                Call ObjBaseClass.Handle_Error(New ApplicationException("Validation is not maintained properly in " & Path.GetFileName(strValidationPath) & " validation file. It must be atleast 4 columns defination."), "ClsValidation", -123, "Validate")
            End If
            Validate = True
        Catch ex As Exception
            Validate = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Validate")
        Finally
            DrValidOutputColumn = Nothing
            ObjBaseClass.ObjectDispose(DtTemp)
            ObjBaseClass.ObjectDispose(DtValidation)
        End Try
    End Function
    Public Function GetValidateDate(ByRef pStrDate As String) As Boolean

        Try

            strInputDateFormat = strInputDateFormat.ToUpper()

            Dim TmpstrInputDateFormat() As String
            Dim TempStrDateValue() As String = pStrDate.Split(" ")

            If InStr(TempStrDateValue(0), "/") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("/")
                TmpstrInputDateFormat = strInputDateFormat.Split("/")
            ElseIf InStr(TempStrDateValue(0), "-") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("-")
                TmpstrInputDateFormat = strInputDateFormat.Split("-")
            End If

            Dim HsUserDate As New Hashtable
            Dim HsSystemDate As New Hashtable
            Dim StrFinalDate As String

            If TempStrDateValue.Length = 3 Then
                For IntStr As Integer = 0 To TempStrDateValue.Length - 1

                    HsUserDate.Add(GetShort(TmpstrInputDateFormat(IntStr).ToString().Trim()), TempStrDateValue(IntStr))
                Next
                Dim SysDate() As String
                Dim dtSys As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
                If InStr(dtSys, "/") > 0 Then
                    SysDate = dtSys.Split("/")
                ElseIf InStr(dtSys, "-") > 0 Then
                    SysDate = dtSys.Split("-")
                End If

                StrFinalDate = ""
                For IntStr As Integer = 0 To SysDate.Length - 1
                    If StrFinalDate = "" Then
                        StrFinalDate += HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    Else
                        StrFinalDate += "/" & HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    End If
                Next

                Try
                    pStrDate = CDate(StrFinalDate)

                    GetValidateDate = True

                Catch ex As Exception
                    GetValidateDate = False

                End Try
            Else
                GetValidateDate = False
            End If

        Catch ex As Exception
            GetValidateDate = False

        End Try

    End Function
    Private Function GetShort(ByVal pStr As String) As String

        pStr = pStr.ToUpper

        If InStr(pStr, "D") > 0 Then
            GetShort = "D"
        ElseIf InStr(pStr, "M") > 0 Then
            GetShort = "M"
        ElseIf InStr(pStr, "Y") > 0 Then
            GetShort = "Y"
        End If

    End Function

    Public Function IsJustAlpha(ByVal sText As String, ByVal num As Integer, ByVal ReplaceWithSpace As String) As String
        Try
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            'If sText <> "" Then
            For n = 1 To iTextLen
                sChar = Mid(sText, n, 1)
                If ChkText(sChar, num) Then
                    IsJustAlpha = IsJustAlpha + sChar
                Else
                    If (ReplaceWithSpace = "Y") Then
                        IsJustAlpha = IsJustAlpha + " "
                    End If

                End If
            Next
            'End If

            If Not IsJustAlpha Is Nothing Then
                Return IsJustAlpha
            Else
                Return ""
            End If


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "IsJustAlpha")
        End Try
    End Function

    Private Function ChkText(ByVal sChr As String, ByVal num As Integer) As Boolean

        Try
            Select Case num
                Case 1
                    '- name field 
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 2
                    '- amount field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[.]" 'Or sChr Like "[,]"
                    'ChkText = True
                Case 3
                    '- alhpa numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[ ]" Or sChr Like "[.]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[:]"
                    'ChkText = True
                Case 4
                    '- address field
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 5
                    '- number field
                    ChkText = sChr Like "[0-9]"
                    'ChkText = True
                Case 6
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 7
                    '- Date field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[:]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[-]" Or sChr Like "[.]"
                    'ChkText = True
                Case 8
                    '- alhpa numeric field & All Characters on Keyboard
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[_]" Or sChr Like "[&]" Or sChr Like "[$]" Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[\]" Or sChr Like "[[]" Or sChr Like "[]]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[']"
                    'ChkText = True
                Case 9
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                Case 10
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[-]" Or sChr Like "[ ]" Or sChr Like "[_]"

                Case 11
                    '- alhpa numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[ ]" Or sChr Like "[.]"
                Case 12
                    '- address field
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[|]" Or sChr Like "[!]" Or sChr Like "[#]" Or sChr Like "[@]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[%]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 13
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[:]" Or sChr Like "[.]" Or sChr Like "[\]" Or sChr Like "[/]"

                Case 14
                    '- Print Date field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[~]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[|]"
                    'ChkText = True
                Case 15
                    '- name field 
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]" Or sChr Like "[.]"
                    'ChkText = True
                Case 16
                    '- Date field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[:]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[-]" Or sChr Like "[.]" Or sChr Like "[ ]"
                    'ChkText = True
                Case Else
                    ChkText = False
            End Select

            Return ChkText

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "ChkText")
        End Try
    End Function


    Private Function ChkText1(ByVal sChr As String, ByVal num As Integer) As Boolean

        Try
            Select Case num
                Case 1
                    '- name field 
                    ChkText1 = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 2
                    '- amount field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[.]" Or sChr Like "[,]"
                    'ChkText = True
                Case 3
                    '- alhpa numeric field
                    ' ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[:]"
                    'ChkText = True
                    ChkText1 = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[`]" Or sChr Like "[!]" Or sChr Like "[#]" Or sChr Like "[@]" Or sChr Like "[$]" Or sChr Like "[%]" Or sChr Like "[*]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[_]" Or sChr Like "[-]" Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[[]" Or sChr Like "[]]" Or sChr Like "[|]" Or sChr Like "[\]" Or sChr Like "[:]" Or sChr Like "[;]" Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[,]" Or sChr Like "[.]" Or sChr Like "[']" Or sChr Like "[?]" Or sChr Like "[/]"
                    'ChkText = True

                Case 4
                    '- address field
                    ChkText1 = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[@]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 5
                    '- number field
                    ChkText1 = sChr Like "[0-9]"
                    'ChkText = True
                Case 6
                    '- alhpa and numeric field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 7
                    '- Date field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[:]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[-]" Or sChr Like "[.]"
                    'ChkText = True
                Case 8
                    '- alhpa numeric field & All Characters on Keyboard
                    'ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[_]" Or sChr Like "[&]" Or sChr Like "[$]" Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[\]" Or sChr Like "[[]" Or sChr Like "[]]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[']"
                    ChkText1 = IsAlpha(sChr)

                    'ChkText = True
                Case 9
                    '- alhpa and numeric field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[-]" Or sChr Like "[ ]" Or sChr Like "[_]"
                Case 10
                    '- alhpa and numeric field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[_]" Or sChr Like "[ ]"
                Case 11
                    '- alhpa and numeric field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[-]" Or sChr Like "[_]" Or sChr Like "[/] "
                    '  "a - z", "A - Z", "0 - 9", " / - _"
                Case 12
                    '- alhpa and numeric field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[_]"
                Case 13
                    '- alhpa and numeric field
                    ChkText1 = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                Case Else
                    ChkText1 = False
            End Select

            Return ChkText1

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "ChkText")
        End Try
    End Function

    Public Function RemoveJunk(ByVal sText As String) As String
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsAlpha(sChar) Then
                        RemoveJunk = RemoveJunk + sChar
                    End If
                Next
            End If

        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "RemoveJunk")

        End Try
    End Function

    Private Function IsAlpha(ByVal sChr As String) As Boolean
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters

        IsAlpha = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" _
        Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[;]" Or sChr Like "[:]" _
        Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[?]" Or sChr Like "[/]" _
        Or sChr Like "[']" Or sChr Like "[""]" Or sChr Like "[|]" Or sChr Like "[\]" _
        Or sChr Like "[{]" Or sChr Like "[[]" Or sChr Like "[}]" Or sChr Like "[]]" _
        Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[_]" Or sChr Like "[-]" _
        Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[*]" Or sChr Like "[&]" _
        Or sChr Like "[^]" Or sChr Like "[%]" Or sChr Like "[$]" Or sChr Like "[#]" _
        Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[`]" Or sChr Like "[~]" _
        Or sChr Like "[ ]" 'commented dtd 03-06-2011

    End Function
    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function
    Private Sub ClearArray(ByRef ArrRow() As String)
        Try
            For i As Integer = 0 To ArrRow.Length - 1
                ArrRow(i) = ""
            Next
        Catch ex As Exception

        End Try

    End Sub
    Private Function GetValueFormArray(ByRef pArray() As Object, ByVal pPosition As Int16) As String
        Try
            If pArray.Length >= pPosition Then
                GetValueFormArray = pArray(pPosition - 1).ToString()
            Else
                ErrorMessage = "~ERROR~"
            End If

        Catch ex As Exception
            ErrorMessage = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetValueFormArray")
        End Try
    End Function
    Private Function GetSubstring(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring = "~ERROR~"
                Else
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else

                        GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")

        End Try
    End Function


#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose

        If Not ObjBaseClass Is Nothing Then ObjBaseClass.Dispose()
        If Not DtValidation Is Nothing Then DtValidation.Dispose()
        If Not DtInput Is Nothing Then DtInput.Dispose()
        If Not DtUnSucInput Is Nothing Then DtUnSucInput.Dispose()
        If Not DtTemp Is Nothing Then DtTemp.Dispose()

        ObjBaseClass = Nothing
        DtValidation = Nothing
        DtInput = Nothing
        DtUnSucInput = Nothing
        DtTemp = Nothing

        GC.SuppressFinalize(Me)

    End Sub
#End Region

End Class
