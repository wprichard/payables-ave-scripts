VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEmailSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************
' Module Name:  cEmailSummary
' Author:       Wes Prichard, Optika
' Date:         09/23/2003
' Description:
' This class contains the functionality to generate an email that contains a URL
' for each package that the user needs to review and approve. The script will fire
' with an execution context of one package but looks for all packages in the queue
' that have the same approver name. The script then creates the email containing
' a URL for each package and then updates the threshold of each package to the next day.
' The script is designed to run on threshold 1.
'
' Edit History:
' mm/dd/yyyy, name, company
' description of change
'
' 10/17/2004, Margaret Randolph, Stellent
' Added code to use Mutex to insure that only 1 email is sent daily per approver.
'
' 10/17/2004, Margaret Randolph, Stellent
' Added code to read Sender Address from the registry
'
' 05/24/2004, Margaret Randolph, Stellent
' Added code to enable banner in email link when generating approver email
' Added module level variable, mstrBanner, for banner parameter and added code to read the value from the registry
'
'
' Public Interface
' Methods:
' Execute

' Events:
' (none)

' Properties:
' (none)

' Note - An object of this class is instantiated automatically when this server
' is instantiated.  Since there is not code that explicitly creates this object and
' no object variable, there is no way to initialize properties when the object
' is created.  Therefore, the properties are initialized from global variables
' when the class initializes.

' Note on error handling - All errors are reported to the creator of this
' object through the ErrorEvent. Even though some
' errors are recoverable by design (a failed query can be treated like a query
' with no results), all errors are reported so that the user can investigate.
' Do not END this program after an error; let the creator destroy it gracefully.

' Note - If this DLL is to be referenced from multiple scripts, modify Initialize
' to disable logging. This is a workaround to avoid a
' permission denied error when multiple instances attempt to open the same log.
'
' Setup Notes:
' See code in ModScript to use int he Process script event.

' Dependencies:
' OTProcessLib, OTProcess.dll
' ADODB, Microsoft ActiveX Data Objects 2.6 Library
'*****************************************************************************************
Option Explicit

'Class Error enumeration
Private Const ERRORBASE = ErrorBase2
Public Enum errcEmailSummary
    errRegistry = ERRORBASE + 0
    errGeneric = ERRORBASE + 1
'    errUpdatePkgJournal = ERRORBASE + 2
'    errGetFieldMap = ERRORBASE + 3
    errExecute = ERRORBASE + 4
'    = ERRORBASE + 5
'    = ERRORBASE + 6
'    = ERRORBASE + 7
'    = ERRORBASE + 8
'    = ERRORBASE + 9
End Enum

Const CLASSNAME = "EmailSummary"
Const MutexTimeout = 1


'Module-scope object references:
Private mobjLog As cLog                         'trace logging object ref
Private WithEvents mobjProcess As cProcess      'Process interface obj ref
Attribute mobjProcess.VB_VarHelpID = -1

Dim mblnLogMultithreaded As Boolean 'Log threading
Dim mstrMutexName As String
'Dim mstrLogMutex As String          'Log Mutex name (must be unique to the component)
Dim mlngLogMutexTimeout As Long     'Log Mutex timeout (in milliseconds)

'Application configurable properties
Dim mstrLogPath As String            'path to log files
Dim mblnTraceLogEnabled As Boolean  'Logging enable
Dim mlngLogMin As Long              'Log carry-over record count
Dim mlngLogMax As Long              'Log reduction threshold
Dim mstrLogFileName As String       'Log file name (no path because app.path is used)
Dim mudtLogType As enuLogType       'Log type

Dim mstrWebServer As String         'name of web server for email URLs
Dim mstrIntervalUnit As String      'unit symbol for the threshold interval (see DateAdd)
Dim mlngIntervalNum As Long         'number of units that make up the interval (ex. 3 "d" = 3 days)
Dim mstrSenderEmail As String       'email address of the sender
Dim mintMailMode As String          'mail mode for email server (see OTEmailLib otMailMode enumeration)
Dim mstrBanner As String            'banner variable for hiding or displaying web client banner when accessing packages through email link


Private Sub Class_Initialize()
'This is the entry point for this ActiveX server.
'An object of this class automatically gets instantiated when the server is
'instantiated.  Therefore, this class must initialize itself.

'Enable error trap
    On Error GoTo errHandler

'Do custom initialization
    Call Initialize

'Update the trace log
    mobjLog.WriteLine CLASSNAME & " class initialized"
    Debug.Print CLASSNAME & " class initialized"
  
Exit Sub

errHandler:
    Err.Raise Err.Number, CLASSNAME & ".Class_Initialize", Err.Description

End Sub

Private Sub Class_Terminate()
'Clean up
    
'Ignore errors so that the class will terminate
    On Error Resume Next

'Close the log file
    mobjLog.WriteLine "Processor class terminated, closing log"
    mobjLog.CloseLogFile
    
'Destroy objects
    Set mobjLog = Nothing
    
    Debug.Print CLASSNAME & " class terminated"
    
End Sub

Public Sub Initialize()
'Perform custom initialization for this Processor.
'Get registry settings and login to databases.
'Server sits idle until the start method
'of the Processor class object is invoked.
      
Dim i As Integer        'loop counter
Dim strProvConn As String           'provider connection string

'Enable error handler
    On Error GoTo errHandler
    
'Read the registry settings
    Call GetRegSettings
    
'Start Trace Log
    Set mobjLog = New cLog    'explicitly instantiate object
    'Configure log object according to registry settings
    mobjLog.LogType = mudtLogType
    mobjLog.MaxLines = mlngLogMax   'used only for circular log
    mobjLog.MinLines = mlngLogMin   'used only for circular log
    'Disable loging to prevent permission denied error when multiple instances of
    'the class is instantiated
    mobjLog.LoggingEnabled = mblnTraceLogEnabled  'false
    mobjLog.Multithreaded = True
    mobjLog.MutexName = CLASSNAME & "_MUTEX"
    mobjLog.MutexTimeout = mlngLogMutexTimeout
    mobjLog.OpenLogFile mstrLogPath & "\" & mstrLogFileName
    mobjLog.WriteLine ""
    mobjLog.WriteLine "Log file opened for new instance of " & CLASSNAME & " - version " & _
        App.Major & "." & App.Minor & "." & App.Revision
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case 70 'permission denied
            'This can occur when trying to open the log file.
            'Disable logging and resume
            mobjLog.LoggingEnabled = False
            Resume Next
        Case Else
            mobjLog.WriteLine "Raising Error in Initialize: " & Err.Number & "-" & Err.Description
            Err.Raise Err.Number, Err.Source, Err.Description
'Resume  '@@@
    End Select

End Sub

Private Sub GetRegSettings()
'Get values from the registry that are needed to execute.
'Note - do not write to log from this procedure (since log object could not be initialized)

'The following requires a project reference to Optika Registry Tool (optikareg.dll)
Dim cReg As cRegistry
Dim strVal As String    'a temporary holder for a registry value

Const APPKEYOPTIKA = "Software\Optika\Common"
Const APPKEYCLASS = "Software\OptikaCustom\" & CLASSNAME
    
'Enable error trap
    On Error GoTo errHandler

'Instantiate registry object
    Set cReg = New cRegistry

'Trace log values
    'Put log in same location as other Acorde service log files
    mstrLogPath = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYOPTIKA, "EventLogPath", App.Path)
    'Get log values from the other key
    mblnTraceLogEnabled = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogEnabled", False)
    mlngLogMin = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogMin", 1000)
    mlngLogMax = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogMax", 5000)
    mstrLogFileName = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogFileName", CLASSNAME & "Log.txt")
    
    '    mblnLogMultithreaded = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogMultithreaded", True)
    mlngLogMutexTimeout = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogMutexTimeoutMillisec", 10000)

    strVal = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogType", "Daily")
    'If value is daily then...
    If LCase$(strVal) = "daily" Then
        mudtLogType = logDaily
    Else
        mudtLogType = logCircular
    End If
    
'Get other class specific values
    mstrWebServer = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "Web Server", "AcordePrdDWeb")
    mstrIntervalUnit = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "Threshold Interval Unit", "d")
    mlngIntervalNum = CLng(cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "Threshold Interval Number", "1"))
    mstrSenderEmail = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "Sender Email", "Process@Acorde.com")
    mstrMutexName = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "MutexName", "EmailSummary")
    strVal = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "Mail Mode", "SMTP")
    Select Case UCase(strVal)
        Case "SMTP"
            mintMailMode = otMailMode_SMTP
        Case "CLIENT"
            mintMailMode = otMailMode_Client
        Case "SERVER"
            mintMailMode = otMailMode_Server
        Case Else
            Err.Raise errRegistry, , "The value of the registry setting 'Mail Mode' " & _
                "is not recognized. The recognized values are 'SMTP', 'Client', 'Server'"
    End Select
'Insert new properties here
    mstrBanner = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "Banner", "0")
    
    
'Destroy the reg object reference
    Set cReg = Nothing

Exit Sub

errHandler:

    Select Case Err.Number
        Case Else
            Call RaiseError(Err.Number, CLASSNAME & ".GetRegSettings", _
                Err.Number & "-" & Err.Description)
        End Select
End Sub

Public Sub Execute(objExecutionContext As OTWFSCRLib.ScriptExecContext)
'This procedure is the main entry point for processing performed by this class.
'Query for the approver's email addresses and update the corresponding fields of
'the package.
        
Dim objPackage As OTPROCESSLib.Package  'the ref to the source package object
'Dim lngPkgID As Long            'unique ID of subject workflow package
'Dim blnSuccess As Boolean       'indicates if query was successful
'Dim udtApprovers As udtApprovers    'data pertaining to approvers for the package's invoice
'Dim strErr As String            'error message text
Dim objFieldValues As OTPROCESSLib.FieldValues
Dim strMsgTxt As String         'email message body text
Dim strAprvrUID As String        'user id of the approver
Dim strApproverField As String  'name of approver field for the current queue
Dim datNewThreshold As Date     'new value for the package thresholds
Dim objPkgs As OTPROCESSLib.Packages    'collection of packages to ref in email
Dim objPkg As OTPROCESSLib.Package  'one package from the collection
Dim strFields() As String       'a dynamic array of pkf fields for profile
Dim varVals() As Variant        'a dynamic array of pkg field values for profile
Dim i As Integer                'loop counter
Dim strApproverEmailField As String 'name of approver's email addres field
Dim strRecipientEmail As String 'email address of recipient
Dim datPkgThresh As Date        'modified threshold for individual package
Dim dblInvAmt As Double         'invoice amount
Dim strInvNum As String
Dim datInvDate As Date
Dim strSupplier As String
Dim strBU As String
Dim strInvCat As String
Dim hMutex As Long          'handle to mutex
Dim dwResult As Long        'result from WaitForSingleObject


'Process event names supported by this script
Const L1_QUEUE = "Approve Level 1"
Const L2_QUEUE = "Approve Level 2"
Const L3_QUEUE = "Approve Level 3"
Const EXCEPTION = "Resolve Exception"
'Package template field names used by the script
Const L1_UID = "Approver1 UID"
Const L2_UID = "Approver2 UID"
Const L3_UID = "Approver3 UID"
Const AP_UID = "APUID"
Const L1_EMAIL = "Approver1 EMAIL"
Const L2_EMAIL = "Approver2 EMAIL"
Const L3_EMAIL = "Approver3 EMAIL"
Const AP_EMAIL = "APEmail"
Const INV_AMT = "Invoice Amount"
Const INV_SUPPLIER = "JDE Supplier Name"
Const INV_NUM = "Invoice Number"
Const INV_DATE = "Invoice Date"
Const BU = "Business Unit"
Const INV_CAT = "Invoice Category"

'Enalbe error trap
    On Error GoTo errHandler
    
'Set default return status (in case an error occurs)
    objExecutionContext.ResultStatus = False
    
'Get the package object
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Getting package object from execution context..."
    Set objPackage = objExecutionContext.Package
    Set objFieldValues = objPackage.FieldValues
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Running in event '" & _
        objPackage.EventName & "' on package '" & _
        CStr(objPackage.PkgId) & "/" & CStr(objPackage.PkgRecId) & ":" & _
        objPackage.Title & "'"
    
'Start the email message body
    strMsgTxt = ""
    strMsgTxt = strMsgTxt & "The following invoice packages require your attention:" & vbCrLf
    strMsgTxt = strMsgTxt & "" & vbCrLf

'Get the user that will be notified based on the current event name
    Select Case objPackage.EventName
        Case L1_QUEUE
            strApproverField = L1_UID
            strApproverEmailField = L1_EMAIL
            
        Case L2_QUEUE
            strApproverField = L2_UID
            strApproverEmailField = L2_EMAIL
        
        Case L3_QUEUE
            strApproverField = L3_UID
            strApproverEmailField = L3_EMAIL
            
        Case EXCEPTION
            strApproverField = AP_UID
            strApproverEmailField = AP_EMAIL
            
        Case Else
            Err.Raise 1, CLASSNAME & ".Execute", "This script is not designed to be used in this event: " & objPackage.EventName
    End Select
    
'Get other package fields needed for processing
    strAprvrUID = objFieldValues.FindByKey(strApproverField, otKeyType_FieldName).Value
    If strAprvrUID = vbNullString Then
        Err.Raise errExecute, , "User ID field '" & strApproverField & "' is blank. Packages for the user cannot be found."
    End If
    
    strRecipientEmail = objPackage.FieldValues.FindByKey(strApproverEmailField, otKeyType_FieldName).Value
    If strRecipientEmail = vbNullString Then
        Err.Raise errExecute, , "Blank recipient email address: " & strApproverEmailField
    ElseIf Not (InStr(1, strRecipientEmail, "@") >= 1) Then
        Err.Raise errExecute, , "Invalid recipient SMTP email address: " & strRecipientEmail
    End If
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Generating email for user '" & _
        strApproverField & " = " & strAprvrUID & "' and sending to address '" & _
        strApproverEmailField & " = " & strRecipientEmail & "'"

'Get the new date/time for the threshold
    datNewThreshold = DateAdd("d", 1, Now)  'tomorrow this time
    
'Create a handle to the mutex (with no owner)
    hMutex = CreateMutex(ByVal 0&, (False), (strAprvrUID))
    If hMutex = 0 Then
    'Error creating handle
        mobjLog.WriteLine "Handle could not be created."
        objExecutionContext.ResultStatus = False
        Exit Sub
    'If Mutex is available, acquire it and finish processing
    End If
    dwResult = WaitForSingleObject(hMutex, 1)
    datPkgThresh = objPackage.FieldValues.FindByKey(otSystemFieldId_Threshold1, otKeyType_FieldId).Value
'Check to see if Mutex is in use by another thread
    If dwResult <> WAIT_OBJECT_0 Then
            mobjLog.WriteLine "Mutex timeout.  Assuming another package for this Approver is already processing. . ."
            objExecutionContext.ResultStatus = True
            Exit Sub
'Check to see if Threshold1 has been updated by another thread
    ElseIf datPkgThresh >= Now Then
            mobjLog.WriteLine "Threshold already reset.  Finished processing."
            objExecutionContext.ResultStatus = True
            Exit Sub
    End If


'Get a collection of packages in this queue for the user
    Set mobjProcess = New cProcess
    Set mobjProcess.objUserToken = objPackage.UserToken
    
    ReDim strFields(0)
    strFields(0) = strApproverField
    ReDim varVals(0)
    varVals(0) = strAprvrUID
    Set objPkgs = mobjProcess.FindPackagesByAttribsAndQueue(strFields, varVals, _
        objPackage.TemplateName, objPackage.ProcessId, objPackage.QueueId)
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) " & objPkgs.Count & _
        " packages found for this user in this queue"
    
'Iterate though each package and get the URL and build up an email message
'Also reset each package's theshold
    For i = 1 To objPkgs.Count
        'Get package fields for message text
        strInvNum = objPkgs(i).FieldValues.FindByKey(INV_NUM, otKeyType_FieldName).Value
        datInvDate = objPkgs(i).FieldValues.FindByKey(INV_DATE, otKeyType_FieldName).Value
        dblInvAmt = objPkgs(i).FieldValues.FindByKey(INV_AMT, otKeyType_FieldName).Value
        strSupplier = objPkgs(i).FieldValues.FindByKey(INV_SUPPLIER, otKeyType_FieldName).Value
        strBU = objPkgs(i).FieldValues.FindByKey(BU, otKeyType_FieldName).Value
        strInvCat = objPkgs(i).FieldValues.FindByKey(INV_CAT, otKeyType_FieldName).Value
        
        'Add a URL to this package to the message body
        strMsgTxt = strMsgTxt & " " & strBU & " , in the amount of " & dblInvAmt & " from " & strSupplier & ", " & Format(datInvDate, "mm-dd-yyyy") & ",  Invoice " & strInvNum & ", " & strInvCat & vbCrLf
        'strMsgTxt = strMsgTxt & "Invoice " & strInvNum & ", " & Format(datInvDate, "mm-dd-yyyy") & " in the amount of " & dblInvAmt & " from " & strSupplier & vbCrLf
        strMsgTxt = strMsgTxt & "http://" & mstrWebServer & "/IBPMWebWIS/default.asp?ToolName=PackageBarTool&PkgRecID1="
        strMsgTxt = strMsgTxt & objPkgs(i).PkgRecId
        strMsgTxt = strMsgTxt & "&HideBanner=" & mstrBanner & "&EOF=1" & vbCrLf
        strMsgTxt = strMsgTxt & vbCrLf
        mobjLog.WriteLine "(" & CLASSNAME & ".Execute) PkgID = " & objPkgs(i).PkgId & _
            ", PkgRecID = " & objPkgs(i).PkgRecId
        
        'Reset package's threshold adding another 1 minute for each successive package
        'so that they all do not expire at the same time
        datPkgThresh = DateAdd("n", i, datNewThreshold)
        objPkgs(i).FieldValues.FindByKey(otSystemFieldId_Threshold1, otKeyType_FieldId).Value = datPkgThresh
        objPkgs(i).FieldValues.FindByKey(otSystemFieldId_ThresholdFired, otKeyType_FieldId).Value = 0
        objPkgs(i).FieldValues.Commit
        mobjLog.WriteLine "(" & CLASSNAME & ".Execute) New threshold = " & datPkgThresh
        mobjLog.WriteLine "(" & CLASSNAME & ".Execute) "
        
    Next

'Send the email
    Call GenerateEmail(objExecutionContext.UserToken, strMsgTxt, strRecipientEmail)
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Email sent"

'Set status for script
    objExecutionContext.ResultStatus = True
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Processing complete, returning to script"
    
'Release Mutex, close Handle.
    If hMutex <> 0 Then
        mobjLog.WriteLine "Releasing Mutex"
        ReleaseMutex hMutex
        CloseHandle hMutex
    End If
    
    Exit Sub

errHandler:
    'Report error to the script but do not raise an error from here
    objExecutionContext.ResultStatus = False
    objExecutionContext.ErrorDescription = objExecutionContext.ErrorDescription & _
        "(" & CLASSNAME & " Script Failed: " & Err.Description & ")"
    Dim strError As String
    Dim lngNum As Long
    strError = Err.Description
    lngNum = Err.Number
    On Error Resume Next
    Call mobjProcess.UpdatePkgJournal(objPackage, CLASSNAME & " Script Failed: " & strError)
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) ERROR: " & lngNum & " - " & strError
Resume  '@@@
End Sub

Private Sub GenerateEmail(objUserToken As OTACORDELib.UserToken, _
    strMessageText As String, strRecipientEmail As String)
'Build the email using the provided message text and send it.
'(Code modified from stock email script.)
'Inputs:
'objUserToken - current user token
'strMessageText - the body of the email message
'strRecipientEmail - the email address of the recipient

Dim objMessage As OTEMAILLib.EMail

    On Error GoTo errHandler

'Get a ToolKit eMail object
    Set objMessage = New OTEMAILLib.EMail

'Set the Mail Mode
    objMessage.MailMode = mintMailMode

'Set the message's UserToken to that of our Context
    Set objMessage.UserToken = objUserToken

'Set the message's Subject and Body
    objMessage.MessageSubject = "Invoices to Approve"
    objMessage.MessageBody = strMessageText

'Add SMTP From Address
    objMessage.SenderMailAddress = mstrSenderEmail

'Add Recipients
    objMessage.AddRecipientMailAddress (strRecipientEmail) ' 2 = otKeyType_Name

'Send the message
    objMessage.Send

Exit Sub

errHandler:
    Select Case Err.Number
        Case 1 'generated errors
            Call RaiseError(errGeneric, CLASSNAME & ".GenerateEmail", _
                Err.Description)
        Case Else
            Call RaiseError(errGeneric, CLASSNAME & ".GenerateEmail", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub

Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error. This procedure is called from another procedures error handler
'to perform logging and raise the error.

'Log the error
    mobjLog.WriteLine "Error in " & Source & ": " & ErrorNumber & " - " & Description
    
'Raise an error back to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub


