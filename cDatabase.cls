VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'*****************************************************************************************
' Class Name:   cDatabase
' Author:       Wes Prichard, Optika
' Date:         08/28/2000
' Description:  Provides an interface to a database via ADO.
'
' Edit History:
' mm/dd/yyyy - Modified by name, company
'   Description of change
'
' Public Interface:
' Methods:
' ConnectSQLServer
' ConnectOracle
' Disconnect
' ExecuteCommand
' Query

' Events:
' LogMsg

' Properties:
' None

' Errors:
' see enum below

' References:
' Microsoft ActiveX Data Objects 2.6 Library (msado15.dll)
' Microsoft DAO 3.51 Object Library ()

'Design/Maintenance Notes:

'*****************************************************************************************
Option Explicit

Private Const ERRORBASE = ErrorBase7    'see modErrorHandling

'Class Error enumeration
Public Enum errCDatabase
    errConnect = ERRORBASE + 0
    errDBCommand = ERRORBASE + 1
    errDBQuery = ERRORBASE + 2
    errDBStartTrans = ERRORBASE + 3
    errDBCommitTrans = ERRORBASE + 4
    errDBRollbackTrans = ERRORBASE + 5
End Enum

'Public events
Public Event LogMsg(Message As String)  'message to the log

'Private class data
'The following requires a project reference to
'Microsoft ActiveX Data Objects 2.6 (formerly 2.1) Library (msado15.dll)
Private madoConn As ADODB.Connection    'ADO connection object reference
'Private mdaoDB As DAO.Database          'DAO database object

Private Sub Class_Initialize()
'Don't use this to initialize the class because events (such as LogMsg) will not work.
End Sub

Private Sub Class_Terminate()
    Set madoConn = Nothing
End Sub

Public Function ConnectSQLServer(strDBServer As String, strDBName As String, _
    strDBUserID As String, strDBPassword As String, Optional blnTrusted As Boolean) _
    As Boolean
'Establish an ADO connection to the specified database.
'Inputs:
'strDBServer - the network name of the database server
'strDBname - the name of the target database
'strDBUserID - the UserID for the login
'strDBPassword - the password for the login
'blnTrusted - optional, indicates if trusted connection to be used
'Outputs:
'True if connection established, False if not.

Dim strProvConn As String   'provider connection string

'Enable error trap
    On Error GoTo errHandler

    RaiseEvent LogMsg("(cDatabase.ConnectSQLServer) Connecting to " & strDBName & " database via SQLOLEDB...")

'Create the connection object
    Set madoConn = New ADODB.Connection
    
'Specify connection string on Open method.
    madoConn.Provider = "sqloledb"

''Trusted connection
'    strProvConn = "Server=" & m_strDBServer & ";Database=" & m_strDBname & _
'        ";Trusted_Connection=yes"
'    m_adoConn.Open strProvConn
    
'Login connection
    strProvConn = "Server=" & strDBServer & ";Database=" & strDBName
    madoConn.Open strProvConn, strDBUserID, strDBPassword
    
    'madoConn.CommandTimeout = 0    'no timeout on connection object
    
'If connection is open then...
    If madoConn.State = adStateOpen Then
        RaiseEvent LogMsg("(cDatabase.ConnectSQLServer) Connected to " & strDBName & " database though " & madoConn.Provider)
        ConnectSQLServer = True
    Else
        RaiseEvent LogMsg("(cDatabase.ConnectSQLServer) ADO connection is not open")
        ConnectSQLServer = False
    End If

Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            RaiseEvent LogMsg("(cDatabase.ConnectSQLServer) ADO error: " & madoConn.Errors(0).Description)
            Call RaiseError(errConnect, "cDatabase.Connect", Err.Number & "-" & Err.Description)
    End Select

End Function

Public Function ConnectOracle(strServiceName As String, _
    strDBUserID As String, strDBPassword As String) As Boolean
'Establish an ADO connection to the specified Oracle database.
'Inputs:
'strServiceName - the Oracle service name (not the DSN) (ODBC DSN not required?)
'strDBUserID - the UserID for the login
'strDBPassword - the password for the login
'Outputs:
'True if connection established, False if not.

Dim strProvConn As String   'provider connection string

'Enable error trap
    On Error GoTo errHandler

    RaiseEvent LogMsg("(cDatabase.ConnectOracle) Connecting to datasource " & strServiceName & "  via OraOLEDB.Oracle...")

'Create the connection object
    Set madoConn = New ADODB.Connection
    
'Login connection
'CTS example
'    oADOLookUp.ConnStr = "Provider=MSDAORA;Password=temp123;User ID=e23907;Data Source=ctsdev;Persist Security Info=True"
'    oADOLookUp.OpenConnection
    
'Method 1 - based on SQL code - works
    madoConn.Provider = "OraOLEDB.Oracle"
'    strProvConn = "Data Source=" & strServiceName  'original string
    strProvConn = "Data Source=" & strServiceName & ";DistribTx = 0"
    madoConn.Open strProvConn, strDBUserID, strDBPassword
    
'Method 2 - based on Oracle example - works
'    madoConn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=" & _
'        strServiceName & ";" & "User ID=" & strDBUserID & ";Password=" & _
'        strDBPassword & ";"
'    madoConn.Open

'Set command timeout
    'madoConn.CommandTimeout = 0    'no timeout on connection object
    
'If connection is open then...
    If madoConn.State = adStateOpen Then
        RaiseEvent LogMsg("(cDatabase.ConnectOracle) Connected to " & strServiceName & " though " & madoConn.Provider)
        ConnectOracle = True
    Else
        RaiseEvent LogMsg("(cDatabase.ConnectOracle) ADO connection is not open")
        ConnectOracle = False
    End If

Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            'RaiseEvent LogMsg("(cDatabase.ConnectOracle) ADO error: " & madoConn.Errors(0).Description)
            Call RaiseError(errConnect, "cDatabase.ConnectOracle", Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select

End Function

'Public Function ConnectDB2(strServiceName As String, _
'    strDBUserID As String, strDBPassword As String) As Boolean
''Establish an ADO connection to the specified DB2 database.
''Inputs:
''strServiceName - the DB2 service name (not the DSN) (using OLEDB)
''strDBUserID - the UserID for the login
''strDBPassword - the password for the login
''Outputs:
''True if connection established, False if not.
'
'Dim strProvConn As String   'provider connection string
'
''Enable error trap
'    On Error GoTo ErrHandler
'
'    RaiseEvent LogMsg("(cDatabase.ConnectDB2) Connecting to datasource " & strServiceName & "  via OraOLEDB.Oracle...")
'
''Create the connection object
'    Set madoConn = New ADODB.Connection
'
''Login connection
'
''From IBM sample code
'    'Open the connection
'    'examples:
'    'cnAS400.Open "Data Source=MySystem;", "USERID", "PWD"
'    'S1056967.Open "Provider=IBMDA400;Data Source=" & systemName & ";", "", ""
'
'    'madoConn.Open "Provider=IBMDA400;Data Source=" & strServiceName & ";", "prichardw", "E247M4F"
'    madoConn.Open "Provider=IBMDA400;Data Source=" & strServiceName & ";", "", ""
'    Debug.Print madoConn.State = adStateOpen
'
''Set command timeout
'    'madoConn.CommandTimeout = 0    'no timeout on connection object
'
''If connection is open then...
'    If madoConn.State = adStateOpen Then
'        'RaiseEvent LogMsg("(cDatabase.ConnectDB2) Connected to " & strServiceName & " though " & madoConn.Provider)
'        ConnectDB2 = True
'    Else
'        RaiseEvent LogMsg("(cDatabase.ConnectDB2) ADO connection is not open")
'        ConnectDB2 = False
'    End If
'
'Exit Function
'
'ErrHandler:
'    Select Case Err.Number
'        Case Else
'            'RaiseEvent LogMsg("(cDatabase.ConnectDB2) ADO error: " & madoConn.Errors(0).Description)
'            Call RaiseError(errConnect, "cDatabase.ConnectDB2", Err.Number & "-" & Err.Description)
''Resume  '@@@
'    End Select
'End Function

'Public Function ConnectODBC_ADO(strDBDSN As String, strDBName As String, _
'    strDBUserID As String, strDBPassword As String, Optional blnTrusted As Boolean) _
'    As Boolean
''Establish an ADO connection to the specified database.
''Inputs:
''strDBServer - the network name of the database server
''strDBname - the name of the target database
''strDBUserID - the UserID for the login
''strDBPassword - the password for the login
''blnTrusted - optional, indicates if trusted connection to be used
''Outputs:
''True if connection established, False if not.
'
'Dim strProvConn As String   'provider connection string
'
''Enable error trap
'    On Error GoTo ErrHandler
'
'    RaiseEvent LogMsg("(cDatabase.ConnectODBC_ADO) Connecting to " & strDBName & " database via SQLOLEDB...")
'
''Create the connection object
'    Set madoConn = New ADODB.Connection
'
''Specify connection string on Open method.
'    madoConn.Provider = "ODBC"
'
'''Trusted connection
''    strProvConn = "Server=" & m_strDBServer & ";Database=" & m_strDBname & _
''        ";Trusted_Connection=yes"
''    m_adoConn.Open strProvConn
'
''Login connection
'    strProvConn = "DSN=" & strDBDSN & ";PWD=" & strDBPassword & ";DATABASE=" & strDBName & _
'                  ";UID=" & strDBUserID & ";LOGINTIMEOUT=0;"
'    'madoConn.Open strProvConn, strDBUserID, strDBPassword
'    madoConn.Open strProvConn
'
'    'madoConn.CommandTimeout = 0    'no timeout on connection object
'
''If connection is open then...
'    If madoConn.State = adStateOpen Then
'        RaiseEvent LogMsg("(cDatabase.ConnectODBC_ADO) Connected to " & strDBName & " database though " & madoConn.Provider)
'        ConnectODBC_ADO = True
'    Else
'        RaiseEvent LogMsg("(cDatabase.ConnectODBC_ADO) ADO connection is not open")
'        ConnectODBC_ADO = False
'    End If
'
'Exit Function
'
'ErrHandler:
'    Select Case Err.Number
'        Case Else
'            RaiseEvent LogMsg("(cDatabase.ConnectODBC_ADO) ADO error: " & madoConn.Errors(0).Description)
'            Call RaiseError(errConnect, "cDatabase.Connect", Err.Number & "-" & Err.Description)
'    End Select
'
'End Function

Public Sub Disconnect()
'Close and release the database connection.
    
    If Not (madoConn Is Nothing) Then
        madoConn.Close
    End If
    
'    If Not (mdaoDB Is Nothing) Then
'        mdaoDB.Close
'    End If
        
End Sub

Public Function ExecuteCommand(strCommand As String, _
    Optional CommandType As CommandTypeEnum, Optional lngTimeout As Long) As Long
'Execute a SQL command using the ADO command object.
'Use this for statements such as Update, Delete, or Insert that do not return records.
'Inputs:
'strCommand - a SQL statement for an Update, Delete, or Insert
'CommandType - optional, specifies type command
'lngTimeout - optional, timeout period for command in seconds
'Outputs:
'The number of records affected by the command.

Dim adoCmd As ADODB.Command     'command object ref
Dim lngRecsAffected As Long     'number of records affected by command

'Ensure there is an ADO connection
    If madoConn.State <> adStateOpen Then
        Err.Raise 1, , "Command cannot be performed because there is no open connection"
    End If
        
'Initialize optional parameters if not specified
    If CommandType = 0 Then
        CommandType = adCmdText
    End If
    
    'The default value for lngTimeout is 0, no timeout

'Instantiate the command object
    Set adoCmd = New ADODB.Command
    
'Initialize the command object
    Set adoCmd.ActiveConnection = madoConn
    adoCmd.CommandType = CommandType
    RaiseEvent LogMsg("(cDatabase.ExecuteCommand) " & strCommand)
    adoCmd.CommandText = strCommand
    adoCmd.CommandTimeout = lngTimeout

'Execute the command
    adoCmd.Execute lngRecsAffected
    
'Return results
    ExecuteCommand = lngRecsAffected
    Set adoCmd = Nothing    'destroy the local object ref

Exit Function

errHandler:
    Select Case Err.Number
        Case 1  'no connection
            Call RaiseError(errDBCommand, "cDatabase.ExecuteCommand", Err.Description)
            
        Case Else
            Call RaiseError(errDBCommand, "cDatabase.ExecuteCommand", Err.Number & "-" & Err.Description)
    End Select
End Function

Public Function Query(strSQL As String, Optional CursorType As CursorTypeEnum, _
    Optional LockType As LockTypeEnum, Optional adoRec As ADODB.Recordset) As ADODB.Recordset
'Use ADO to perform the specified query and return the results.
'Use this mehtod only for queries that return records.
'Inputs:
'strSQL - a SQL statement that returns data
'CursorType - optional, specifies type of recordset (ForwardOnly, Keyset, Dynamic, Static)
'LockType - determines how records are locked (ReadOnly, Pessimistic, Optimistic, BatchOptimistic)
'Outputs:
'An ADO recordset.

'Dim adoRec As ADODB.Recordset

'Enable error trap
    On Error GoTo errHandler

'Ensure there is an ADO connection
    If madoConn.State <> adStateOpen Then
        Err.Raise 1, , "Query cannot be performed because there is no open connection"
    End If
        
'Initialize optional parameters if not specified
    If CursorType = 0 Then
        CursorType = adOpenStatic   'A static copy of a set of records that you
                                    'can use to find data or generate reports.
                                    'Additions, changes, or deletions by other
                                    'users are not visible.
    End If
    
    If LockType = 0 Then
        LockType = adLockReadOnly
    End If

'Instantiate the recordset object if not passed in
    If adoRec Is Nothing Then
        Set adoRec = New ADODB.Recordset
    End If

'Open the recordset object to execute the query
    RaiseEvent LogMsg("(cDatabase.Query) " & strSQL)
    'adoRec.Open strSQL, madoConn, adOpenStatic, adLockReadOnly
    adoRec.Open strSQL, madoConn, CursorType, LockType
    'If the connection object throws an error or the error collection is empty then
    'display the vb runtime error info

'Return the results
    Set Query = adoRec
    Set adoRec = Nothing    'destroy the local object ref

Exit Function

errHandler:
Dim vbErrInfo As String
Dim adoErrInfo As String
  
  'Save off the vb runtime err info.  It will be overwritten if another exception occurs.
  vbErrInfo = "Error Number: " & Err.Number & vbCr & _
               "Error Description: " & Err.Description & vbCr & _
               "Source: " & Err.Source
  
  'We must re-enable exception handling in case the ADO connection object
  'is not set.
  Resume Continue         'Only a resume or exit re-enables error handling
Continue:
  On Error Resume Next    'Handle an invalid connection object
    If madoConn.Errors.Count = 0 Then
       MsgBox vbErrInfo
    Else
       'we'll display the ADO error collection
       Dim cnErr As ADODB.Error
       For Each cnErr In madoConn.Errors
          adoErrInfo = "Error Number: " & cnErr.Number & vbCr & vbTab & _
                       "Error Description: " & cnErr.Description & vbCr & vbTab & _
                       "Source: " & cnErr.Source & vbCr & vbTab & _
                        "SQLState:  " & cnErr.SQLState & vbCr
       RaiseEvent LogMsg(adoErrInfo)
       Next cnErr
       'MsgBox adoErrInfo
    End If
    
    Select Case Err.Number
        Case 1  'no connection
            Call RaiseError(errDBQuery, "cDatabase.Query", Err.Description)
            
        Case Else
            Call RaiseError(errDBQuery, "cDatabase.Query", Err.Number & "-" & Err.Description)

    End Select
End Function

'Public Function QueryODBC_DAO(strSQL As String, Optional daoRec As DAO.Recordset) As DAO.Recordset
''Use DAO to perform the specified query and return the results.
''Use this method only for queries that return records.
''Inputs:
''strSQL - a SQL statement that returns data
''CursorType - optional, specifies type of recordset (ForwardOnly, Keyset, Dynamic, Static)
''LockType - determines how records are locked (ReadOnly, Pessimistic, Optimistic, BatchOptimistic)
''Outputs:
''An DAO recordset.
'
''Dim adoRec As DAO.Recordset
'
''Enable error trap
'    On Error GoTo errHandler
'
''Ensure there is an DAO connection
'    If mdaoDB.Connect = vbNullString Then
'        Err.Raise 1, , "Query cannot be performed because there is no open connection"
'    End If
'
'''Initialize optional parameters if not specified
''    If CursorType = 0 Then
''        CursorType = adOpenStatic   'A static copy of a set of records that you
''                                    'can use to find data or generate reports.
''                                    'Additions, changes, or deletions by other
''                                    'users are not visible.
''    End If
''
''    If LockType = 0 Then
''        LockType = adLockReadOnly
''    End If
'
'
''Open the recordset object to execute the query
'    RaiseEvent LogMsg("(cDatabase.QueryODBC_DAO) " & strSQL)
'    'adoRec.Open strSQL, madoConn, adOpenStatic, adLockReadOnly
''    daoRec.Open strSQL, madoConn, CursorType, LockType
'
'    Set daoRec = mdaoDB.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)  'dbo
'
''Return the results
'    Set QueryODBC_DAO = daoRec
'    Set daoRec = Nothing    'destroy the local object ref
'
'Exit Function
'
'errHandler:
'
'    Select Case Err.Number
'        Case 1  'no connection
'            Call RaiseError(errDBQuery, "cDatabase.QueryODBC_DAO", Err.Description)
'
'        Case Else
'            Call RaiseError(errDBQuery, "cDatabase.QueryODBC_DAO", Err.Number & "-" & Err.Description)
'
'    End Select
''Resume '@@@
'End Function

Public Sub StartTransaction()
'Start a database transaction.

'Enable error trap
    On Error GoTo errHandler

'Ensure there is an ADO connection
    If madoConn.State <> adStateOpen Then
        Err.Raise 1, , "Transaction cannot be started because there is no open connection"
    End If
        
'Begin the transaction
    madoConn.BeginTrans
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case 1  'no connection
            Call RaiseError(errDBStartTrans, "cDatabase.StartTransaction", Err.Description)
            
        Case Else
            Call RaiseError(errDBStartTrans, "cDatabase.StartTransaction", Err.Number & "-" & Err.Description)
    End Select
End Sub

Public Sub CommitTransaction()
'Commit a database transaction.

'Enable error trap
    On Error GoTo errHandler

'Ensure there is an ADO connection
    If madoConn.State <> adStateOpen Then
        Err.Raise 1, , "Transaction cannot be committed because there is no open connection"
    End If
        
'Commit the transaction
    madoConn.CommitTrans
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case 1  'no connection
            Call RaiseError(errDBCommitTrans, "cDatabase.CommitTransaction", Err.Description)
            
        Case Else
            Call RaiseError(errDBCommitTrans, "cDatabase.CommitTransaction", Err.Number & "-" & Err.Description)
    End Select
End Sub

Public Sub RollbackTransaction()
'Rollback a database transaction.

'Enable error trap
    On Error GoTo errHandler

'Ensure there is an ADO connection
    If madoConn.State <> adStateOpen Then
        Err.Raise 1, , "Transaction cannot be undone because there is no open connection"
    End If
        
'Begin the transaction
    madoConn.RollbackTrans
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case 1  'no connection
            Call RaiseError(errDBRollbackTrans, "cDatabase.RollbackTransaction", Err.Description)
            
        Case Else
            Call RaiseError(errDBRollbackTrans, "cDatabase.RollbackTransaction", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error.

'Log the error
    RaiseEvent LogMsg("Error in " & Source & ": " & ErrorNumber & " - " & Description)
    
'Raise an error up to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub

'Public Function ConnectODBC(DSN As String, Database As String, UID As String, _
'    PWD As String) As Boolean
''Use the specified existing ODBC DSN to connect to the specified database.
'
'Dim sConnectStr As String
'Dim dbODBC As DAO.Database
'
'    On Error GoTo errHandler
'
''Define the connection string
'    sConnectStr = "ODBC;DSN=" & DSN & ";PWD=" & PWD & ";DATABASE=" & Database & _
'                  ";UID=" & UID & ";LOGINTIMEOUT=0;"
''For DAO 3.51
'    Set mdaoDB = Workspaces(0).OpenDatabase("", False, False, sConnectStr)
''For DAO 3.6
'    'DBEngine.DefaultType = dbUseODBC    'make workspace ODBCDirect
'    'Workspaces(0).Type = dbUseODBC
'    'Set mdbOE = Workspaces(0).OpenDatabase(m_strDSN, False, False, sConnectStr)
'
'    If mdaoDB Is Nothing Then
'        'Err.Raise errWFODBCConnect, , "No database object created"
'        ConnectODBC = False
'    Else
'        ConnectODBC = True
'    End If
'
'Exit Function
'
'errHandler:
'    'Err.Raise errWFODBCConnect, "WFODBC_Connect", Err.Description
'
'
'
'
'End Function
