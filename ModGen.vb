Option Explicit On

Module ModGen

    Public gstrOutputFile As String
    Public gstrInputFile As String
    Public gstrInputFolder As String
    Public InputDate As Date
    Public blnErrorLog As Boolean = False
    Public strSettingClientCode_B As String
    Public strSettingClientCode_P As String
    Public strSettingClientName As String
    Public strUploadFile As String


    Public gstrResponseInputFolder As String
    Public gstrResponseInputFile As String
    Public strAuditFolderPath As String             ' Audit folder path
    Public strErrorFolderPath As String             ' Error folder path
    Public strInputFolderPath As String             ' Input folder path
    Public strInput_BackupFolderPath As String             ' Inputbackup folder path
    Public strTempFolderPath As String             ' Input folder path
    Public strOutputFolderPath As String            ' Output folder path
    Public strResponseFolderPath As String             ' Response folder path
    Public strReverseResponseFolderPath As String            ' RevResponse folder path
    Public gstrOutputFile_EPAY As String

    Public strRejectedFolderPath As String            ' Rejected folder path
    Public strReportFolderPath As String            ' Report folder path
    Public strMappingFilePath As String ' Mapping Path
    Public strArchivedFolderUnSuc As String ' Success folder Path
    Public strArchivedFolderSuc As String ' UnSuccess folder Path
    Public strValidationPath As String ' Validation file path
    Public strProceed As String
    Public strMasterCustIdFilePath As String ' Master File Path
    Public strMaxTransactionAmount As Integer
    Public strInputFileFormat As String
    Public strTransStartLineNo As Integer
    Public strTypeOfConvertor As String
    Public strEncrypt As String

    Public TransactionRefNo As String

    Public RemitterName As String
    Public AddressLine1 As String
    Public AddressLine2 As String
    Public AddressLine3 As String
    Public BeneAddLine1 As String
    Public BeneAddLine2 As String
    Public BeneAddLine3 As String
    Public BeneAddLine4 As String
    Public AddInfo2 As String
    Public AddInfo3 As String
    Public AddInfo4 As String

    Public strConctPersonName As String
    Public strContactNo As String
    Public InputFilesCount As Integer

    'Public StrSplitOutput As String
    'Public NoOfRecords As Double
    Public strCompanyCode As String

    '-Client Details-
    Public strClientCode_B As String
    Public strClientCode_P As String
    Public strClientName As String
    Public strInputDateFormat As String
    Public strNoOfDays As String
    Public strOutputDateFormat As String
    Public OutputDateFormat As String
    Public IntOutPut As Integer
    Public strSpecialCharMaster As String
    Public strRenameFile As String

    Public FileCounter As String
    Public strCompanyName As String
    Public strMasterCode As String
    Public strBookingFileCounterStart As String
    Public strPaymentFileCounterStart As String
    Public Client_Code As String
    Public Upload_date As Date
    Public RTGSAmount As String
    Public TxnRefNo As Integer
    Public TxnRefLetter As String

    Public FlgRenameFile As Boolean = False
    '-Encryption-
    Public strYBLEncryptionEpayFile As String
    Public strYBLEncryptionAdviceFile As String
    Public strYBLBatchFilePath As String
    Public strYBLPICKDPath As String
    Public strYBLDROPDPath As String
    Public strYBLCRCPPath As String
    Public strEncryptionTime As String
    Public strBatchFilePathYBL As String
    Public strVendorMaster As String



    Public strDebitAccountNumber As String
    Public strRemitterName As String
    Public strAddressLine1 As String
    Public strAddressLine2 As String
    Public strAddressLine3 As String
    'Public strBeneAddLine1 As String
    'Public strBeneAddLine2 As String
    'Public strBeneAddLine3 As String
    'Public strBeneAddLine4 As String
    Public strAddInfo1 As String
    Public strAddInfo2 As String
    Public strAddInfo3 As String
    Public strAddInfo4 As String
    Public strBene_Addr1 As String
    Public strBene_Addr2 As String
    Public strBene_Addr3 As String
    Public strBene_Addr4 As String
    Public strEpayInputFolderPath As String
    Public strAdviceInputFolderPath As String


    Public gstrAdviceInputFolder As String
    Public gstrAdviceInputFile As String
    Public strEpayFileName As String
    Public strImageLogoPath As String
    Public strWCTAmount As String

End Module
