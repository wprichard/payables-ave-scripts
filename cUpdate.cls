VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLookupApprovers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************
' Module Name:  cLookupApprovers
' Author:       Wes Prichard, Optika
' Date:         09/16/2003
' Description:
' This class contains the functionality to access the table of authorizations and find
' the record containing the approvers based on the business unit. Package fields are
' updated. The public methods in this class are designed to be called from an
' Acorde Process Script event. Developed for Acorde 3.1
'
' Edit History:
' mm/dd/yyyy, name, company
' description of change
'
' Public Interface:
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
Private Const ERRORBASE = ErrorBase1
Public Enum errcLookupApprovers
    errRegistry = ERRORBASE + 0
    errQueryForApprovers = ERRORBASE + 1
    errUpdatePkgJournal = ERRORBASE + 2
'    errGetFieldMap = ERRORBASE + 3
    errExecute = ERRORBASE + 4
'    = ERRORBASE + 5
'    = ERRORBASE + 6
'    = ERRORBASE + 7
'    = ERRORBASE + 8
'    = ERRORBASE + 9
End Enum

Const CLASSNAME = "cLookupApprovers"

'Module-scope object references:
Private mobjLog As cLog                         'trace logging object ref
Private WithEvents mobjApprovalDB As cDatabase  'database object reference
Attribute mobjApprovalDB.VB_VarHelpID = -1

'Application configurable properties
Dim mstrLogPath As String            'path to log files
Dim mblnTraceLogEnabled As Boolean  'Logging enable
Dim mlngLogMin As Long              'Log carry-over record count
Dim mlngLogMax As Long              'Log reduction threshold
Dim mstrLogFileName As String       'Log file name (no path because app.path is used)
Dim mudtLogType As enuLogType       'Log type

Dim mstrDBServer As String          'OneWorld DB data source name
Dim mstrDBName As String            'OneWorld library (database)
Dim mstrDBUser As String            'OneWorld DB user name
Dim mstrDBPassword As String        'OneWorld DB account password

Private Type udtApprovers   'structure for the approver data
    BusinessUnit As String  'JDE Business Unit ID - determines the approvers
    BizUnitName As String   'Business Unit name
    
    APUID As String         'user id
    APName As String        'user name
    APAmount As Double      'approval limit
    APEmail As String       'email address
    
    AP2UID As String        'user id
    AP2Name As String       'user name
    AP2Amount As Double     'approval limit
    AP2Email As String      'email address
    
    V1UID As String         'user id
    V1Name As String        'user name
    V1Amount As Double      'approval limit
    V1Email As String       'email address
    
    V2UID As String         'user id
    V2Name As String        'user name
    V2Amount As Double      'approval limit
    V2Email As String       'email address
    
    V3UID As String         'user id
    V3Name As String        'user name
    V3Amount As Double      'approval limit
    V3Email As String       'email address
End Type

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

'Logout of all databases
    mobjApprovalDB.Disconnect

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
    strVal = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogType", "Daily")
    'If value is daily then...
    If LCase$(strVal) = "daily" Then
        mudtLogType = logDaily
    Else
        mudtLogType = logCircular
    End If
    
    'Get database access values
    mstrDBServer = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "DB Server", "AcordeSQLServer")
    mstrDBName = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "DB Name", "Payables")
    mstrDBUser = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "DB User", "PADMIN")
    mstrDBPassword = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "DB Password", "PADMIN")
    
'Insert new properties here

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
        
Dim objPackage As OTPROCESSLib.Package
Dim lngPkgID As Long            'unique ID of subject workflow package
Dim blnSuccess As Boolean       'indicates if query was successful
Dim udtApprovers As udtApprovers    'data pertaining to approvers for the package's invoice
Dim strErr As String            'error message text
Dim objFieldValues As OTPROCESSLib.FieldValues
    
'Enalbe error trap
    On Error GoTo errHandler
    
'Set default return status (in case an error occurs)
    objExecutionContext.ResultStatus = False
    
'Get the package object
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Getting package object from execution context..."
    Set objPackage = objExecutionContext.Package
    Set objFieldValues = objPackage.FieldValues
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Got package object from execution context"
    
'Connect to database
    mobjLog.WriteLine "Connecting to Approver database via OLEDB..."
    Set mobjApprovalDB = New cDatabase
    If mobjApprovalDB.ConnectSQLServer(mstrDBServer, mstrDBName, mstrDBUser, mstrDBPassword) = False Then
        Err.Raise 1, CLASSNAME & ".Execute", "Connection to '" & _
            mstrDBServer & ":" & mstrDBName & "' failed"
    End If
    mobjLog.WriteLine "Connected to " & mstrDBServer & " database though OLEDB"
   
'Get the package ID
    lngPkgID = objPackage.PkgId
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Package ID = " & lngPkgID
        
'Query for the approver table record to get approvers
    'Note udtApprovers is updated and returned by the function
    udtApprovers.BusinessUnit = objFieldValues.FindByKey("Business Unit", otKeyType_FieldName).Value
    blnSuccess = QueryForApprovers(mobjApprovalDB, udtApprovers)
    
'If no record then...
    If blnSuccess = False Then
        strErr = "There are no approver records for Business unit '" & _
            udtApprovers.BusinessUnit & "'. Ensure the Table of Authorizations is up-to-date."

        'Raise a fatal error. The required info was not obtained from the database.
        Err.Raise 2, , strErr
    
    Else    'application table data obtained
        
        'Save the values to the package fields
        objFieldValues.FindByKey("APEmail", otKeyType_FieldName).Value = _
            udtApprovers.APEmail
        objFieldValues.FindByKey("APName", otKeyType_FieldName).Value = _
            udtApprovers.APName
        objFieldValues.FindByKey("APUID", otKeyType_FieldName).Value = _
            udtApprovers.APUID
        
        objFieldValues.FindByKey("AP2Email", otKeyType_FieldName).Value = _
            udtApprovers.AP2Email
        objFieldValues.FindByKey("AP2Name", otKeyType_FieldName).Value = _
            udtApprovers.AP2Name
        objFieldValues.FindByKey("AP2UID", otKeyType_FieldName).Value = _
            udtApprovers.AP2UID
        
        objFieldValues.FindByKey("Approver1 Amount", otKeyType_FieldName).Value = _
            udtApprovers.V1Amount
        objFieldValues.FindByKey("Approver1 Email", otKeyType_FieldName).Value = _
            udtApprovers.V1Email
        objFieldValues.FindByKey("Approver1 Name", otKeyType_FieldName).Value = _
            udtApprovers.V1Name
        objFieldValues.FindByKey("Approver1 UID", otKeyType_FieldName).Value = _
            udtApprovers.V1UID
        
        objFieldValues.FindByKey("Approver2 Amount", otKeyType_FieldName).Value = _
            udtApprovers.V2Amount
        objFieldValues.FindByKey("Approver2 Email", otKeyType_FieldName).Value = _
            udtApprovers.V2Email
        objFieldValues.FindByKey("Approver2 Name", otKeyType_FieldName).Value = _
            udtApprovers.V2Name
        objFieldValues.FindByKey("Approver2 UID", otKeyType_FieldName).Value = _
            udtApprovers.V2UID
        
        objFieldValues.FindByKey("Approver3 Amount", otKeyType_FieldName).Value = _
            udtApprovers.V3Amount
        objFieldValues.FindByKey("Approver3 Email", otKeyType_FieldName).Value = _
            udtApprovers.V3Email
        objFieldValues.FindByKey("Approver3 Name", otKeyType_FieldName).Value = _
            udtApprovers.V3Name
        objFieldValues.FindByKey("Approver3 UID", otKeyType_FieldName).Value = _
            udtApprovers.V3UID
        
        objFieldValues.Commit
        mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Approvers assigned for Business Unit: '" & udtApprovers.BusinessUnit & " - " & udtApprovers.BizUnitName & "'"
        
        'Update journal
        objPackage.AddJournalEntry "Approvers initialized by script for Business Unit: '" & udtApprovers.BusinessUnit & "'"
    
    End If
        
'Set status for script
    objExecutionContext.ResultStatus = True
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) Processing complete, returning to script"
    
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
    Call UpdatePkgJournal(objPackage, CLASSNAME & " Script Failed: " & strError)
    mobjLog.WriteLine "(" & CLASSNAME & ".Execute) ERROR: " & lngNum & " - " & strError
'Resume  '@@@
End Sub

Private Function QueryForApprovers(objDB As cDatabase, udtApprovers As udtApprovers) _
    As Boolean
'Query the OneWorld Approver table for the approvers and
'populate the data structure with values from the application table.
'Inputs:
'udtApprovers - a structure containing the Job Number, & all approver fields empty
'objDB - the database object to use for querying
'Outputs:
'udtApprovers - input structure updated from fields in the application table
'Boolean - True if query returned results

Dim recTable As ADODB.Recordset   'query results
Dim strSQL As String            'query string
Dim datStart As Date            'Time hack used to measure query response time
Dim lngQTime As Long            'query response time in seconds

'Enable error trap
    On Error GoTo errHandler

'Build the query statement
    'Get all fields from the specified application table where JOBNUM is the
    'specified value
    'Example query:
    'SELECT * FROM TOA_Approvers WHERE BusinessUnit = '1024'
    
    strSQL = "SELECT * from TOA_Approvers WHERE BusinessUnit = '" & _
        udtApprovers.BusinessUnit & "'"
    Debug.Print strSQL

'Run the query
    datStart = Now
    Set recTable = objDB.Query(strSQL, adOpenStatic, adLockReadOnly)
    lngQTime = DateDiff("s", datStart, Now)
    mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) Approver table query time = " & _
        lngQTime & " secs"

'If a recordset exists then...
    If Not (recTable Is Nothing) Then
        'If there are records then...
        If recTable.EOF = False Then
            Do While Not recTable.EOF
                udtApprovers.APAmount = recTable.Fields("APAmount").Value
                udtApprovers.APEmail = GetDBString(recTable.Fields("APEmail").Value)
                udtApprovers.APName = GetDBString(recTable.Fields("APName").Value)
                udtApprovers.APUID = GetDBString(recTable.Fields("APUID").Value)
                
                udtApprovers.AP2Amount = recTable.Fields("AP2Amount").Value
                udtApprovers.AP2Email = GetDBString(recTable.Fields("AP2Email").Value)
                udtApprovers.AP2Name = GetDBString(recTable.Fields("AP2Name").Value)
                udtApprovers.AP2UID = GetDBString(recTable.Fields("AP2UID").Value)
                
                udtApprovers.V1Amount = recTable.Fields("Approver1Amount").Value
                udtApprovers.V1Email = GetDBString(recTable.Fields("Approver1Email").Value)
                udtApprovers.V1Name = GetDBString(recTable.Fields("Approver1Name").Value)
                udtApprovers.V1UID = GetDBString(recTable.Fields("Approver1UID").Value)
                
                udtApprovers.V2Amount = recTable.Fields("Approver2Amount").Value
                udtApprovers.V2Email = GetDBString(recTable.Fields("Approver2Email").Value)
                udtApprovers.V2Name = GetDBString(recTable.Fields("Approver2Name").Value)
                udtApprovers.V2UID = GetDBString(recTable.Fields("Approver2UID").Value)
               
                udtApprovers.V3Amount = recTable.Fields("Approver3Amount").Value
                udtApprovers.V3Email = GetDBString(recTable.Fields("Approver3Email").Value)
                udtApprovers.V3Name = GetDBString(recTable.Fields("Approver3Name").Value)
                udtApprovers.V3UID = GetDBString(recTable.Fields("Approver3UID").Value)
                
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP Name = '" & udtApprovers.APName & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP UID = '" & udtApprovers.APUID & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP Email = '" & udtApprovers.APEmail & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP Amount = '" & udtApprovers.APAmount & "'"

                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP2 Name = '" & udtApprovers.AP2Name & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP2 UID = '" & udtApprovers.AP2UID & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP2 Email = '" & udtApprovers.AP2Email & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "AP2 Amount = '" & udtApprovers.AP2Amount & "'"

                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 1 Name = '" & udtApprovers.V1Name & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 1 UID = '" & udtApprovers.V1UID & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 1 Email = '" & udtApprovers.V1Email & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 1 Amount = '" & udtApprovers.V1Amount & "'"
                    
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 2 Name = '" & udtApprovers.V2Name & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 2 UID = '" & udtApprovers.V2UID & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 2 Email = '" & udtApprovers.V2Email & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 2 Amount = '" & udtApprovers.V2Amount & "'"
                    
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 3 Name = '" & udtApprovers.V3Name & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 3 UID = '" & udtApprovers.V3UID & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 3 Email = '" & udtApprovers.V3Email & "'"
                mobjLog.WriteLine "(" & CLASSNAME & ".QueryForApprovers) " & _
                    "Approver 3 Amount = '" & udtApprovers.V3Amount & "'"

                recTable.MoveNext
                'Note - if there is more than one record per business unit, the last
                'record found is used.
            Loop
            
            'Indicate success
            QueryForApprovers = True

        Else    'empty recordset
            'Indicate no records found for the business unit
            QueryForApprovers = False
        End If

    Else    'no recordset
        'Raise an error
        Err.Raise 1, , "Unable to obtain approver table record, " & _
            "no recordset was returned from the query"
    End If
    Set recTable = Nothing
    
Exit Function

errHandler:
    Select Case Err.Number
        Case 1  'no recordset returned
            Call RaiseError(errQueryForApprovers, CLASSNAME & ".QueryForApprovers", _
                Err.Description)
        Case Else
            Call RaiseError(errQueryForApprovers, CLASSNAME & ".QueryForApprovers", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

Private Sub UpdatePkgJournal(objPackage As OTPROCESSLib.Package, strJournal As String)
'Update the journal of the specified package with the specified string.
'Inputs:
'objPackage - the object representing of the desired package
'strJournal - the string to write to theh journal
'Outputs:
'none

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is not valid then...
    If objPackage Is Nothing Then
        Err.Raise 1, , "Unable to update package journal because the specified " & _
            "package object is nothing"
    Else
        'Add the journal entry
        objPackage.AddJournalEntry strJournal
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case 1 'generated errors
            Call RaiseError(errUpdatePkgJournal, "cWorkflowII.UpdatePkgJournal", _
                Err.Description)
        Case Else
            Call RaiseError(errUpdatePkgJournal, CLASSNAME & ".UpdatePkgJournal", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub

Private Sub mobjApprovalDB_LogMsg(Message As String)
'Log the message
    
'Enable error trap
    On Error GoTo errHandler
    
'Write the log message
    mobjLog.WriteLine Message
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(Err.Number, CLASSNAME & ".mobjApprovalDB_LogMsg", _
                Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Function GetDBString(varDBField As Variant) As String
'Check for a null string from a database and convert it to an empty string.

    If Not IsNull(varDBField) Then
        GetDBString = Trim(varDBField)
    Else
        GetDBString = vbNullString
    End If

End Function

Private Function GetDBDouble(varDBField As Variant) As Double
'Check for a null value from a database and convert it to a zero.

    If Not IsNull(varDBField) Then
        GetDBDouble = varDBField
    Else
        GetDBDouble = 0
    End If

End Function

Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error. This procedure is called from another procedures error handler
'to perform logging and raise the error.

'Log the error
    mobjLog.WriteLine "Error in " & Source & ": " & ErrorNumber & " - " & Description
    
'Raise an error back to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub


