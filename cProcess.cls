VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************************************************************
' Class Name:   cProcess
' Author:       Wes Prichard, Optika
' Date:         09/24/2003
' Description:  Manages the interface to Acorde Process through the Process Toolkit
'               Components.
' SDK Version:  This code is written for the Acorde SDK Version 3.1 GA.
'               Upgraded from cWorkflowII (Acorde 2.1)
'
' Edit History:
' mm/dd/yyyy - Modified by name, company
'   Description of change
'
' Public Interface:
' Methods:
' AddPkgObject
' CreatePackage
' DeletePackage
' FindPackagesByAttribs
' FindPackageByID (only valid for CR2 SDK or later)
' GetPackageEvent
' GetPackageField
' GetPackageID
' Initialize
' InjectAtEvent
' ModifyPackageField
' MovePkgToEvent
' PackageExists (only valid for CR2 SDK or later)
' UpdatePackageJournal

' Events:
' LogMsg

' Properties:
' objUserToken

' Errors:
' see enum below

' References:
' OTACORDELib
' OTPROCESSLib

'Usage notes:
'Instantiate an object of the class then call the Initialize method. The objUserToken
'property must then be initialized with a valid user token.

''Prepare object for use by instantiating, initializing properties, then calling the
''connect method.
''Use Disconnect method before destroying object to release connections.

'Design/Maintenance Notes:
'Error convention -
'Procedures that do not need to return a value are coded as Subs - status return codes
'are not used.
'Any error that occurs in a procedure is raised using an error code unique to that
'procedure. Only one error code is raised per procedure so that it can be trapped easliy.
'The original error number is embedded in the error description text.
'A public error code enumeration is provided so that all errors can be externally
'referenced by name rather than number. The project that uses this class must assign the
'base error number and reserve the next 50 numbers for this class.

'*****************************************************************************************
Option Explicit

Const CLASSNAME = "cProcess"

'Class Error enumeration
Private Const ERRORBASE = ErrorBase4    'defined by project
Public Enum errCWorkflowII              'up to 50 errors
    errModifyPackageField = ERRORBASE + 0
    errAddPkgObject = ERRORBASE + 1
'    errGetExistingPackage = ERRORBASE + 2
    errCreatePackage = ERRORBASE + 3
    errInjectAtEvent = ERRORBASE + 4
    errFindPackagesByAttribs = ERRORBASE + 5
    errPackageExists = ERRORBASE + 6
    errInitialize = ERRORBASE + 7
    errGetPackageID = ERRORBASE + 8
    errGetPackageEvent = ERRORBASE + 9
    errDeletePackage = ERRORBASE + 10
    errGetPackageField = ERRORBASE + 11
    errUpdatePkgJournal = ERRORBASE + 12
    errGetPackageAttachments = ERRORBASE + 13
    errFindPackageByID = ERRORBASE + 14
    errGetPackagesInQueue = ERRORBASE + 15
    errGetQueues = ERRORBASE + 16
'    errDBQuery = ERRORBASE + 17
'    errDBQuery = ERRORBASE + 18
'    errDBQuery = ERRORBASE + 19
'    errDBQuery = ERRORBASE + 20
End Enum

'Enumeration for different types of Context document references
Public Enum enuObjectRefType
    refObjectID = 1    'full objectID (a.b.RECID.d.e.f.g.h.i.j.k.l.m.n.o.p.q.r.s.t)
    refRECID = 2       'nnnnnnnn
    refObjectKey = 3   'workflow subset of ObjectID (0.RECID.0.0)
End Enum

'Public events
Public Event LogMsg(Message As String)  'message to the log

'Class Property variables
'The following requires a project reference to: Acorde Base Object Library
'(OTAcorde.dll)
Private mobjUserToken As OTACORDELib.UserToken     'user security token for workflow access
'Private mstrProcess As String                       'name of process applied to this instance of the object

Public Property Set objUserToken(ByVal vData As OTACORDELib.UserToken)
Attribute objUserToken.VB_Description = "A valid user token object providing secure access to workflow"
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objUserToken = Form1
    Set mobjUserToken = vData
    
End Property

Public Property Get objUserToken() As OTACORDELib.UserToken
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objUserToken
    Set objUserToken = mobjUserToken
End Property

Private Sub Class_Initialize()
'Don't use this to initialize the class because events (such as LogMsg) will not work
'in this procedure.
End Sub

Public Sub Initialize()
'Perform any class initialization here rather than in the Class_Initialize.

Dim objDatabases As OTPROCESSLib.Databases
Dim objDatabase As OTPROCESSLib.Database

'Enable error trap
    On Error GoTo errHandler

'Initialize the database object with the user token
    Set objDatabases = New OTPROCESSLib.Databases
    objDatabases.Refresh
    
    If objDatabases.Count > 1 Then
        Err.Raise 1, , "There are multiple process databases to connect to. The desired database must be specified. (Code change required)"
    Else
        Set objDatabase = objDatabases.Item(1)
        RaiseEvent LogMsg("(cWorkflowII.Initialize) Connecting to database '" & objDatabases.Item(1).DatabaseName & "'")
        objDatabase.Connect mobjUserToken
        RaiseEvent LogMsg("(cWorkflowII.Initialize) Database connection successfull")
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errInitialize, CLASSNAME & ".Initialize", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub

'Public Sub XAddPkgObject(strObjectRef As String, ObjectRefType As enuObjectRefType, _
'    strObjectTitle As String, objPackage As OTPROCESSLib.Package, strAppTable As String, _
'    Optional strAttachmentType As String, Optional strMIMEType As String)
''Add the specified Context object/document reference to the specified package.
''Inputs:
''strObjectRef - a RECID or ObjectID for the document to be added
''Unique - 1) File Objects = (file name and path)
''         2) COLD = (full objectID)
''         3) Image & Univ = (RecID)
''strObjectTitle - the title string to be used for the attachment in the package
''ObjectRefType - enum indicating whether strObjectRef is a RECID or ObjectID
''objPackage - an object referencing the package the doc is to be appended to
''strAttachmentType - Optional, the name of the attachment type, not defined if not provided
''strMIMEType - Optional, the MIME type of the document (default is image/tiff)
'
''Outputs:
''none
'
'' MIMETypes - 1) Image = image/tiff
''             2) COLD = text/x-emedia.cold,
''             3) Universal = application/x-emedia.universal
''             4) File =  File Objects = application/x-emedia.wfuniversal
''Provider IDs - File Object = "{081016EE-F688-11D3-9E61-00C04F097031}"
''               Acorde Objects = "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
''Unique ID - 1) File Objects = (file name and path)
''            2) COLD = (full objectID)
''            3) Image & Univ = (RecID)
''IndexID is for Acorde objects only (leave blank for File Objects)
'    'For Imaging and Universal use the RecID
'    'For COLD use the DocumentID:Physical Cold Number (from the Optika Object ID - prepadded with zeros to make 10 digits each)
'    'COLD example - "0012456512:0000000090"
''IndexName is for Acorde objects only (leave blank for File Objects)
'    'LinkServerName.Instance.Catalog.Table
'    'example "IMG.Acorde.dbo.Image4Main" or "COLD..ColdAll.Main"
''Index Provider as follows:
'    'File Objects - "{081016EE-F688-11D3-9E61-00C04F097031}"
'    'Cold Objects - "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
'    'Acorde Objects - "{608FCB70-10BF-11d4-A931-00C04F94786A}"
'
'Dim objAttachments As OTPROCESSLib.Attachments    'Attachments collection
'Dim objAttachment As OTPROCESSLib.Attachment      'Attachment object
'Dim strRECID As String                          'RECID - unique ref to document
'Dim objAttachmentTypes As OTPROCESSLib.AttachmentTypes    'Attachment types collection
'Dim objAttachmentType As OTPROCESSLib.AttachmentType      'Attachment type object
'Dim objTemplate As OTPROCESSLib.Template                  'Template object
''Dim adCM As ADODB.Command
''Dim sSQL As String
'Dim intDot1 As Integer                  'position of first period in object key string
'Dim intDot2 As Integer                  'position of second period in object key string
'
'
''Enable error trap
'    On Error GoTo ErrHandler
'
''Get the existing attachments for the package
'    Set objAttachments = objPackage.Attachments
'    objAttachments.Refresh
'
''Get the RECID for the object (document)
'    Select Case ObjectRefType
'        Case enuObjectRefType.refObjectID   'eMedia ObjectID
'            'Get the RECID portion
'            strRECID = Mid(strObjectRef, 5, InStr(5, strObjectRef, ".") - 5)
'
'        Case enuObjectRefType.refRECID  'RECID
'            strRECID = strObjectRef
'
'        'This is probably not need for 2.x
'        Case enuObjectRefType.refObjectKey 'Object key
'            'Get the RECID portion
'            intDot1 = InStr(1, strObjectRef, ".")
'            intDot2 = InStr(intDot1 + 1, strObjectRef, ".")
'            strRECID = Mid(strObjectRef, intDot1 + 1, intDot2 - intDot1 - 1)
'
'        Case Else   'other
'            'Unexpected object reference input indicating
'            'the code is out of sync with the object type enum
'            Err.Raise 1, , "An unknown document object reference type was provided"
'    End Select
'
''Get the package template
'    Set objTemplate = objPackage.Template
'
''If an attachment type was specified then...
'    If strAttachmentType <> vbNullString Then
'        'Prepare to attach the document as the specified type attachment
'        Set objAttachmentTypes = objTemplate.AttachmentTypes
'        Set objAttachmentType = objAttachmentTypes.FindByKey(strAttachmentType, otKeyType_TypeName)
'        If objAttachmentType Is Nothing Then
'            Err.Raise 2, , "The specified package attachment type '" & _
'                strAttachmentType & "' is not defined in the process"
'        End If
'    End If
'
''Create an attachment object for the new object/document to be attached
'    Set objAttachment = objAttachments.CreateAttachment
'
''Initialize attachment object properties
'    'Unique ID (see notes above)
'    objAttachment.UniqueID = strRECID
'    'MIME type
'    If strMIMEType = vbNullString Then strMIMEType = "image/tiff"
'    objAttachment.MimeType = strMIMEType
'    'Provider ID for Acorde Objects
'    objAttachment.ProviderID = "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
'    'Title
'    objAttachment.Title = strObjectTitle
'    'Attachment Type
'    If Not (objAttachmentType Is Nothing) Then
'        objAttachment.AttachmentType = objAttachmentType
'    End If
'    objAttachment.IndexId.IndexId = strRECID    'use RECID for imaging
'    objAttachment.IndexId.IndexName = "Imaging.Context.dbo." & strAppTable
'
''Add the object/document
'    objAttachments.Commit
'
'''Jeff's WorkAround to set the IndexID, IndexName and AttachmentTypeID on the record
''    Set adCM = New ADODB.Command
''    Set adCM.ActiveConnection = objDataSource.adCN_Workflow
''
''    sSQL = "Update WF_PKGATTACHMENT set INDEXID = '" & strRECID & "' , INDEXNAME = 'Imaging.Context.dbo." & _
''        objDataSource.sEmediaApplication & objDataSource.sEmediaIndex & "', " & _
''        "ATTACHMENTTYPEID = " & objAttachmentType.TypeId & " Where PKGID = " & _
''        m_objPackage.PkgId
''
''    With adCM
''        .CommandType = adCmdText
''        .CommandText = sSQL
''        .Execute
''    End With
''
''    Set adCM = Nothing
'
''Destroy object references
'    Set objTemplate = Nothing
'    Set objAttachmentTypes = Nothing
'    Set objAttachmentType = Nothing
'    Set objAttachment = Nothing
'
'Exit Sub
'
'ErrHandler:
'    Select Case Err.Number
'        Case 1  'unexpected object reference type
'            Call RaiseError(errAddPkgObject, CLASSNAME & ".AddPkgObject", _
'                Err.Description)
'        Case Else
'            Call RaiseError(errAddPkgObject, CLASSNAME & ".AddPkgObject", _
'                Err.Number & "-" & Err.Description)
''Resume  '@@@
'    End Select
'End Sub

Public Sub AddPkgObject(strObjectRef As String, ObjectRefType As enuObjectRefType, _
    strObjectTitle As String, objPackage As OTPROCESSLib.Package, _
    strLinkServer As String, strDBInstance As String, strCatalog As String, _
    strAppTable As String, _
    Optional strAttachmentType As String, Optional strMIMEType As String = "image/tiff")
'Add the specified Context object/document reference to the specified package.
'Inputs:
'strObjectRef - a RECID or ObjectID for the document to be added
'ObjectRefType - enum indicating whether strObjectRef is a RECID or ObjectID
'strObjectTitle - the title string to be used for the attachment in the package
'objPackage - an object referencing the package the doc is to be appended to
'strLinkServer - the name of the Link Server (ex. Imaging)
'strDBInstance - the DB instance for the application from the link server configuration
'strCatalog - the DB catalog for the application table (ex. dbo)
'strAppTable - the name of the application table
'strAttachmentType - Optional, the name of the attachment type, not defined if not provided
'strMIMEType - Optional, the MIME type of the document (default is image/tiff)

'Outputs:
'none

Dim strRECID As String                          'RECID - unique ref to document
Dim objTemplate As OTPROCESSLib.Template                  'Template object
Dim objAttachmentTypes As OTPROCESSLib.AttachmentTypes    'Attachment types collection
Dim objAttachmentType As OTPROCESSLib.AttachmentType      'Attachment type object
Dim objAttachment As OTPROCESSLib.Attachment
Dim objAttachments As OTPROCESSLib.Attachments
'Dim objattachTypes As OTPROCESSLib.AttachmentTypes
Dim objAttachType As OTPROCESSLib.AttachmentType
'The following requires aproject reference to OTDocumentProvider.dll
Dim objDocID As OTACORDELib.DocumentId  ' OTDOCUMENTPROVIDERLib.DocumentId
Dim objIndexID As OTACORDELib.IndexId   ' OTDOCUMENTPROVIDERLib.IndexId

'Get the RECID for the object (document)
    Select Case ObjectRefType
        Case enuObjectRefType.refObjectID   'eMedia ObjectID
            'Get the RECID portion
            strRECID = Mid(strObjectRef, 5, InStr(5, strObjectRef, ".") - 5)
    
        Case enuObjectRefType.refRECID  'RECID
            strRECID = strObjectRef
    
        Case Else   'other
            'Unexpected object reference input indicating
            'the code is out of sync with the object type enum
            Err.Raise 1, , "An unknown document object reference type was provided"
    End Select

'Get the package template
    Set objTemplate = objPackage.Template
    
'If an attachment type was specified then...
    If strAttachmentType <> vbNullString Then
        'Prepare to attach the document as the specified type attachment
        Set objAttachmentTypes = objTemplate.AttachmentTypes
        Set objAttachmentType = objAttachmentTypes.FindByKey(strAttachmentType, otKeyType_TypeName)
        If objAttachmentType Is Nothing Then
            Err.Raise 2, , "The specified package attachment type '" & _
                strAttachmentType & "' is not defined in the process"
        End If
    End If
    
'From Optika Engineering Services sample:
'*****************************************************************************
'Here is the way to attach an Acorde image - this is not supported however.
    
'Create and init a documentID object
    Set objDocID = New OTACORDELib.DocumentId   ' OTDOCUMENTPROVIDERLib.DocumentId
    
    'Populate the DocumentID object - used Integrated sample to get this information
    ' MIMETypes - 1) Image = image/tiff
    '             2) COLD = text/x-emedia.cold,
    '             3) Universal = application/x-emedia.universal
    '             4) File =  File Objects = application/x-emedia.wfuniversal
    objDocID.MimeType = strMIMEType
    
    ' Provider IDs -File Object = "{081016EE-F688-11D3-9E61-00C04F097031}"
    '               Acorde Objects = "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
    objDocID.ProviderID = "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
    
    'Unique - 1) File Objects = (file name and path)
    '         2) COLD = (full objectID)
    '         3) Image & Univ = (RecID)
    objDocID.UniqueID = strRECID
    
'Create and Populate the IndexID object - used Integrated sample to get this information
    Set objIndexID = New OTACORDELib.IndexId    ' OTDOCUMENTPROVIDERLib.IndexId
    
    'IndexID is for Acorde objects only (leave blank for File Objects)
        'For Imaging and Universal use the RecID
        'For COLD use the DocumentID:Physical Cold Number (from the Optika Object ID - prepadded with zeros to make 10 digits each)
        'COLD example - "0012456512:0000000090"
    objIndexID.IndexId = strRECID
    
    'IndexName is for Acorde objects only (leave blank for File Objects)
        'LinkServerName.Instance.Catalog.Table
        'example "IMG.Acorde.dbo.Image4Main" or "COLD..ColdAll.Main"
    'objIndexID.IndexName = "IMG.Acorde21.dbo" & strAppTable
    objIndexID.IndexName = strLinkServer & "." & strDBInstance & "." & strCatalog & "." & strAppTable
    
    'Index Provider as follows:
        'File Objects - "{081016EE-F688-11D3-9E61-00C04F097031}"
        'Cold Objects - "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
        'Acorde Objects - "{608FCB70-10BF-11d4-A931-00C04F94786A}"
    objIndexID.IndexProvider = "{608FCB70-10BF-11d4-A931-00C04F94786A}"
    
'Create an attachment object, still need to commit to save (see below)
    Set objAttachments = objPackage.Attachments
    Set objAttachment = objAttachments.CreateAttachment
    Set objAttachmentTypes = objPackage.Template.AttachmentTypes
    
'Initialize attachment object
    objAttachment.IndexId = objIndexID
    objAttachment.DocumentId = objDocID
    
'### The following commented code is redundant because these properties are
'part of the DocumentID object. These properties were eliminated in 2.1 CR2
'    ' Provider IDs - 1) File Object = {081016EE-F688-11D3-9E61-00C04F097031}
'    '                2) Acorde Objects ={0BF3C340-4C13-11d3-8166-00C04F99E979}
'    objAttachment.ProviderID = "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
'
'    ' MIMETypes - 1) Image = image/tiff
'    '             2) COLD = text/x-emedia.cold,
'    '             3) Universal = application/x-emedia.universal
'    '             4) File =  File Objects = application/x-emedia.wfuniversal
'    objAttachment.MimeType = strMIMEType
'
'    'Unique - 1) File Objects = (file name and path)
'    '         2) COLD = (full objectID)
'    '         3) Image & Univ = (RecID)
'    objAttachment.UniqueID = strRECID
    
    'Title is the name of the attachment
    objAttachment.Title = strObjectTitle
    
'Call commit method to save attachment
    objAttachments.Commit
    
Exit Sub
    
errHandler:
    Select Case Err.Number
        Case 1  'unexpected object reference type
            Call RaiseError(errAddPkgObject, CLASSNAME & ".AddPkgObject", _
                Err.Description)
        Case Else
            Call RaiseError(errAddPkgObject, CLASSNAME & ".AddPkgObject", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub


Public Function PackageExists(strPkgID As String, strPkgTemplate As String) As Boolean
'Determine if a package exists by searching for the package ID in the package fields.
'Inputs:
'strPkgID - a package ID to search for
'strPkgTemplate - the type of package to search for
'Outputs:
'True if the package was found

'Dim strAtNames(0) As String
'Dim varAtVal(0) As Variant
Dim objPackage As OTPROCESSLib.Package    'package object

'Enable error trap
    On Error GoTo errHandler

'Search for the package
    Set objPackage = FindPackageByID(strPkgID, strPkgTemplate)

'If no package found then...
    If objPackage Is Nothing Then
        PackageExists = False
    Else
        'If PackageIDs don't match then...
        If objPackage.PkgId = strPkgID Then
            PackageExists = True
        Else
            PackageExists = False
        End If
    End If

'Cleanup references
    Set objPackage = Nothing

Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errPackageExists, CLASSNAME & ".PackageExists", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select

End Function

Public Sub MovePkgToEvent(objPackage As OTPROCESSLib.Package, strDestinationEvent As String, _
    strProcessName As String)
'Move the specified package to the specified event.
'objPackage - object representing package to be injected
'strDestinationEvent - the name of the event to move the package to
'strProcessName - the name of the process to inject into
'intPriority - Optional, the priority of the package when injected
'Outputs:
'none

Dim objProcesses As OTPROCESSLib.Processes
Dim objProcess As OTPROCESSLib.Process
Dim objEvents As OTPROCESSLib.Events
Dim objEvent As OTPROCESSLib.Event

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is not valid then...
    If (objPackage Is Nothing) Then
        Err.Raise 1, , "Unable to move package because the specified " & _
            "package object is nothing"
    End If

'Instantiate and initialize a new processes collection
    Set objProcesses = New OTPROCESSLib.Processes
    
'If user token is valid then...
    If Not (mobjUserToken Is Nothing) Then
        'Init collection with user token
        objProcesses.UserToken = mobjUserToken
        objProcesses.Refresh
    Else
        'Token object is nothing, raise an error
        Err.Raise 1, , "User token is invalid"
    End If
    
'Get an object for the specified process
    Set objProcess = objProcesses.FindByKey(strProcessName, otKeyType_ProcessName)
    If objProcess Is Nothing Then
        Err.Raise 2, , "The specified process name '" & strProcessName & _
            "' was not found in the database"
    End If
    
'Get the object for the specified event
    Set objEvents = objProcess.Events
    Set objEvent = objEvents.FindByKey(strDestinationEvent, otKeyType_EventName)
    If objEvent Is Nothing Then
        Err.Raise 3, , "The specified destination event name '" & strDestinationEvent & _
            "' was not found in process '" & strProcessName & "'"
    End If
    
'Route the package
    objPackage.RouteToEvent objEvent

Exit Sub

errHandler:
    Select Case Err.Number
        Case 1  'invalid package object
            Call RaiseError(errGetPackageEvent, CLASSNAME & ".GetPackageEvent", _
                Err.Description)
        Case Else
            Call RaiseError(errGetPackageEvent, CLASSNAME & ".GetPackageEvent", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub

Public Function GetPackageEvent(objPackage As OTPROCESSLib.Package) As String
'Return the name of the current event or events for the specified package.
'Inputs:
'objPackage - a valid package object representing the package to process

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is valid then...
    If Not (objPackage Is Nothing) Then
        'Get the event name
        GetPackageEvent = objPackage.EventName
    Else
        Err.Raise 1, , "Unable to get package event because the specified " & _
            "package object is nothing"
    End If

Exit Function

errHandler:
    Select Case Err.Number
        Case 1  'invalid package object
            Call RaiseError(errGetPackageEvent, CLASSNAME & ".GetPackageEvent", _
                Err.Description)
        Case Else
            Call RaiseError(errGetPackageEvent, CLASSNAME & ".GetPackageEvent", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

Public Sub DeletePackage(objPackage As OTPROCESSLib.Package)
'Delete the specified package from the process.
'Inputs:
'objPackage - a valid package object representing the package to process

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is valid then...
    If Not (objPackage Is Nothing) Then
        'Delete the package
        RaiseEvent LogMsg("(cWorkflowII.DeletePackage) CanDo for package deletion = " & objPackage.CanDoOperation(otPackageOperation_Delete))
        objPackage.Delete otPackageDelete_All
    Else
        Err.Raise 1, , "Unable to delete the package because the specified " & _
            "package object is nothing"
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case 1  'invalid package object
            Call RaiseError(errDeletePackage, CLASSNAME & ".DeletePackage", _
                Err.Description)
        Case Else
            Call RaiseError(errDeletePackage, CLASSNAME & ".DeletePackage", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub

Public Function CreatePackage(strPkgTemplateName As String) As OTPROCESSLib.Package
'Create a new package based on the specified package template.
'Inputs:
'strPkgTemplateName - name of package template on which to base nnew package
'Outputs:
'A package object representing the new package

Dim objTemplates As OTPROCESSLib.Templates    'templates collection
Dim objTemplate As OTPROCESSLib.Template      'object for specified template
Dim objPackage As OTPROCESSLib.Package        'new package object
         
Dim objProcesses As OTPROCESSLib.Processes
Dim objProcess As OTPROCESSLib.Process
    
'Enable error trap
    On Error GoTo errHandler

'Instantiate a new templates collection
    Set objTemplates = New OTPROCESSLib.Templates
    
'If user token is valid then...
    If Not (mobjUserToken Is Nothing) Then
        'Init collection with user token
        objTemplates.UserToken = mobjUserToken
        objTemplates.Refresh
    Else
        'Token object is nothing, raise an error
        Err.Raise 1, , "User token is invalid, unable to look for package template"
    End If
        
'Get the specified package template from the collection
    Set objTemplate = objTemplates.FindByKey(strPkgTemplateName, otKeyType_TemplateName)
  
'Create the Package and return the object reference
    If Not (objTemplate Is Nothing) Then
        Set objPackage = objTemplate.CreatePackage
        If objPackage Is Nothing Then
            'raise an error
            Err.Raise 3, , "A package was not created"
            'TODO get more error information
        Else
            Set CreatePackage = objPackage
        End If
    Else
        Err.Raise 2, , "The specified package template '" & strPkgTemplateName & _
            "' was not found in the process"
    End If
    
'Cleanup object references
    Set objTemplate = Nothing
    Set objTemplates = Nothing

Exit Function

errHandler:
    Select Case Err.Number
        Case 1, 2  'No user token, bad package template name
            Call RaiseError(errCreatePackage, CLASSNAME & ".CreatePackage", _
                Err.Description)
        Case Else
            Call RaiseError(errCreatePackage, CLASSNAME & ".CreatePackage", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

Public Sub InjectAtEvent(objPackage As OTPROCESSLib.Package, strProcessName As String, _
    strEventname As String, Optional intPriority As Integer = 0)
'Inject the specified package at the specified event in the specified process.
'Inputs:
'objPackage - object representing package to be injected
'strProcessName - the name of the process to inject into
'strEventName - the name of the event to inject at
'intPriority - Optional, the priority of the package when injected
'Outputs:
'none
    
Dim objProcesses As OTPROCESSLib.Processes
Dim objProcess As OTPROCESSLib.Process
Dim objEvents As OTPROCESSLib.Events
Dim objEvent As OTPROCESSLib.Event
         
'Enable error trap
    On Error GoTo errHandler

'Instantiate and initialize a new processes collection
    Set objProcesses = New OTPROCESSLib.Processes
    
'If user token is valid then...
    If Not (mobjUserToken Is Nothing) Then
        'Init collection with user token
        objProcesses.UserToken = mobjUserToken
        objProcesses.Refresh
    Else
        'Token object is nothing, raise an error
        Err.Raise 1, , "User token is invalid"
    End If
    
'Get an object for the specified process
    Set objProcess = objProcesses.FindByKey(strProcessName, otKeyType_ProcessName)
    If objProcess Is Nothing Then
        Err.Raise 2, , "The specified process name '" & strProcessName & _
            "' was not found in the database"
    End If
    
'Get the process's events collection
    Set objEvents = objProcess.Events
'Find the specified event in the collection
    Set objEvent = objEvents.FindByKey(strEventname, otKeyType_EventName)
    If objEvent Is Nothing Then
        Err.Raise 3, , "The specified event name '" & strEventname & _
            "' was not found in process '" & strProcessName & "'"
    End If
    
'Place the package in the flow at the event, with the specified priority
    objPackage.PlaceInFlow objEvent, intPriority, "Injected by " & App.ProductName
        
'Cleanup object references
    Set objEvent = Nothing
    Set objEvents = Nothing
    Set objProcess = Nothing
    Set objProcesses = Nothing

Exit Sub

errHandler:
    Select Case Err.Number
        Case 1, 2, 3  'No user token, bad process name, bad event name
            Call RaiseError(errInjectAtEvent, CLASSNAME & ".InjectAtEvent", _
                Err.Description)
        Case Else
            Call RaiseError(errInjectAtEvent, CLASSNAME & ".InjectAtEvent", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
        
End Sub

Public Sub ModifyPackageField(objPackage As OTPROCESSLib.Package, strFieldName As String, _
    varFieldValue As Variant)
'Modify the specified field of the specified package to the specified value.
'Inputs:
'objPackage - the package to be modified
'strFieldName - the name of the field to modify
'varFieldValue - the new value for the field
'Outputs:
'none

Dim objFieldValue As OTPROCESSLib.FieldValue
     
'Enable error trap
    On Error GoTo errHandler

'Check for a blank field name
     If strFieldName = vbNullString Then
            Err.Raise 1, , "The specified package field name " & _
                " is blank - the package field cannot be modifed"
     End If
     
'Check for a null input parameter
     If IsNull(varFieldValue) Then
            Err.Raise 1, , "The specified value to write to the package field " & _
                strFieldName & " is Null - the field cannot be modified"
     End If
     
'Get the current value of the specified field
     Set objFieldValue = objPackage.FieldValues.FindByKey(strFieldName, otKeyType_FieldName)
     
'If field value exists then...
    If Not (objFieldValue Is Nothing) Then
        'Modify the field to the new value
        objFieldValue.Value = varFieldValue
        objPackage.FieldValues.Commit
    Else
        Err.Raise 2, , "The specified package field '" & _
            strFieldName & "' was not found in the package. The field cannot be modified."
    End If
     
'Cleanup object references
     Set objFieldValue = Nothing
     
Exit Sub

errHandler:
    Select Case Err.Number
        Case 1, 2 'Null new field value, no specified field
            Call RaiseError(errModifyPackageField, CLASSNAME & ".ModifyPackageField", _
                Err.Description)
        Case Else
            Call RaiseError(errModifyPackageField, CLASSNAME & ".ModifyPackageField", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
        
End Sub

Public Function GetPackageField(objPackage As OTPROCESSLib.Package, strFieldName As String) _
    As Variant
'Return the content of the specified field in the specified package.
'Inputs:
'objPackage - a valid package object
'strFieldName - the name of the field to read
'Outputs:
'The value of the specified field

Dim objFieldValues As OTPROCESSLib.FieldValues
Dim objFieldValue As OTPROCESSLib.FieldValue

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is not valid then...
    If objPackage Is Nothing Then
        Err.Raise 1, , "Unable to get the package field because the specified " & _
            "package object is nothing"
    End If

'Get the field value object for the specified field
    Set objFieldValues = objPackage.FieldValues
    Set objFieldValue = objFieldValues.FindByKey(strFieldName, otKeyType_FieldName)
    
'If the specified field Value object was found then...
    If Not (objFieldValue Is Nothing) Then
        
        'Return the field value
        GetPackageField = objFieldValue.Value
       
    Else   'the specified field was not found
       'Raise an error
        Err.Raise 2, , "The specified field name '" & _
            strFieldName & "' was not found in the specified package"
    End If
    
Exit Function

errHandler:
    Select Case Err.Number
        Case 1, 2  'generated errors
            Call RaiseError(errGetPackageField, CLASSNAME & ".GetPackageField", _
                Err.Description)
        Case Else
            Call RaiseError(errGetPackageField, CLASSNAME & ".GetPackageField", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

Public Sub UpdatePkgJournal(objPackage As OTPROCESSLib.Package, strJournal As String)
'Update the journal of the specified package with the specified string.
'Inputs:
'objPackage - the object representing of the desired package
'strJournal - the string to write to theh journal
'Outpus:
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
            Call RaiseError(errUpdatePkgJournal, CLASSNAME & ".UpdatePkgJournal", _
                Err.Description)
        Case Else
            Call RaiseError(errUpdatePkgJournal, CLASSNAME & ".UpdatePkgJournal", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Sub

Public Function GetPackageID(objPackage As OTPROCESSLib.Package) As String
'Return the packageID for the specified package.
'Inputs:
'objPackage - the object representing of the desired package
'Outputs:
'The specified package's package ID

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is not valid then...
    If objPackage Is Nothing Then
        Err.Raise 1, , "Unable to update package journal because the specified " & _
            "package object is nothing"
    Else
        'Get the package ID
        GetPackageID = objPackage.PkgId
    End If

Exit Function

errHandler:
    Select Case Err.Number
        Case 1 'generated errors
            Call RaiseError(errGetPackageID, CLASSNAME & ".GetPackageID", _
                Err.Description)
        Case Else
            Call RaiseError(errGetPackageID, CLASSNAME & ".GetPackageID", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

Public Function FindPackagesByAttribs(strAttribNames() As String, varAttribValues() As Variant, _
    strPkgTemplate As String) As OTPROCESSLib.Packages
'Find packages that have matching field values to those that are specified.
'One or more field/value pairs may be specified.
'Build up a profile of the specified field names and values and use the profile
'to get a collection of packages fitting that profile.
'Inputs:
'strAttribNames() - a zero-based one-dimensional array of package field names
'varAttribValues() - a zero-based one-dimensional array of package field values
'                   corresponding to the strAttribNames array
'strPkgTemplate - the name of the package template to search

'Outputs:
'a packages collection containing packages with the specified field values

Dim objTemplates As OTPROCESSLib.Templates
Dim objTemplate As OTPROCESSLib.Template
Dim objFieldDefinitions As OTPROCESSLib.FieldDefinitions
Dim objFieldDefinition As OTPROCESSLib.FieldDefinition
Dim objProfile As OTPROCESSLib.Profile
Dim objProfileCriteria As OTPROCESSLib.ProfileCriteria
Dim objFieldCriteria As OTPROCESSLib.FieldCriteria
Dim objFieldCriterion As OTPROCESSLib.FieldCriterion
Dim objPackages As OTPROCESSLib.Packages
Dim intIndex As Integer                                 'input array index
    
'Enable error trap
    On Error GoTo errHandler

'Make sure the input arrays are the same size
    If UBound(strAttribNames) <> UBound(varAttribValues) Then
        'Raise an error
        Err.Raise 1, , "The number of specified field names does not match " & _
            "the number of specified field values"
    End If

'Instantiate and initialize a templates collection
    Set objTemplates = New OTPROCESSLib.Templates
    objTemplates.UserToken = mobjUserToken
    objTemplates.Refresh
     
'Find the specified package template in the templates collection
    Set objTemplate = objTemplates.FindByKey(strPkgTemplate, otKeyType_TemplateName)
     
'If the template was found then...
    If Not (objTemplate Is Nothing) Then
         
        'Get the fields collection the package template
        Set objFieldDefinitions = objTemplate.FieldDefinitions
        
        '@@@ list package field names
'        For intIndex = 1 To objFieldDefinitions.Count
'            Debug.Print objFieldDefinitions.Item(intIndex).FieldName, objFieldDefinitions.Item(intIndex).IsSystemField
'        Next intIndex
        '@@@
        
        'Instantiate a new profile object
        Set objProfile = New OTPROCESSLib.Profile
        
        'Instantiate a new profile criteria object and initialize it
        Set objProfileCriteria = New OTPROCESSLib.ProfileCriteria
        objProfileCriteria.TemplateId = objTemplate.TemplateId
        Set objProfileCriteria.UserToken = mobjUserToken
        
        'Get the field criteria collection object from the profile criteria
        Set objFieldCriteria = objProfileCriteria.FieldCriteria
                
        'For each field in the input arrays
        For intIndex = 0 To UBound(strAttribNames)
        
            'Find the specified field in the collection
            Set objFieldDefinition = objFieldDefinitions.FindByKey(strAttribNames(intIndex), _
                otKeyType_FieldName)
            
            'If the specified field was found then...
            If Not (objFieldDefinition Is Nothing) Then
                
                'Instantiate a new field criterion object
                Set objFieldCriterion = New OTPROCESSLib.FieldCriterion
                'Define criterion object as specified FieldName = specified FieldValue
                'objFieldCriterion.Set objFieldDefinition, varAttribValues(intIndex), 0, _
                    otOperator_EqualTo
                objFieldCriterion.SetOperation objFieldDefinition, otOperator_EqualTo
                objFieldCriterion.AddOperand varAttribValues(intIndex)
                'Add the criterion to the criteria collection
                objFieldCriteria.AddItem objFieldCriterion
             
            Else   'the specified field was not found
               'Raise an error
                Err.Raise 2, , "The specified field name '" & _
                    strAttribNames(intIndex) & "' was not found in the specified " & _
                    "package template '" & strPkgTemplate & "'"
            End If
        
        Next intIndex
    
        'Assign the criteria to the Profile
        Set objProfile.ProfileCriteria = objProfileCriteria
        
        'Instantiate a new packages collection
        Set objPackages = New OTPROCESSLib.Packages
        'Apply the profile to the packages collection and refresh to search
        Set objPackages.Profile = objProfile
        Set objPackages.UserToken = mobjUserToken
        objPackages.Refresh
        
        '@@@ list packages returned
'        For intIndex = 1 To objPackages.Count
'            Debug.Print objPackages(intIndex).PkgId, objPackages(intIndex).PkgRecId
'        Next intIndex
        '@@@
        
        'Return the packages object
        Set FindPackagesByAttribs = objPackages
    
    Else    'specified package template not found
        'raise an error
        Err.Raise 3, , "The specified package template '" & _
            strPkgTemplate & "' was not found"
    
    End If
    
'Cleanup
    Set objTemplates = Nothing
    Set objFieldDefinitions = Nothing
    Set objFieldDefinition = Nothing
    Set objProfile = Nothing
    Set objProfileCriteria = Nothing
    Set objFieldCriteria = Nothing
    Set objPackages = Nothing
    
Exit Function

errHandler:
    Select Case Err.Number
        Case 1  'mismatched input arrays
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribs", _
                Err.Description)
        Case 2  'field name not in package template
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribs", _
                Err.Description)
        Case 3  'package template not found
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribs", _
                Err.Description)
        Case Else
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribs", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

Public Function FindPackagesByAttribsAndQueue(strAttribNames() As String, varAttribValues() As Variant, _
    strPkgTemplate As String, lngProcessID As Long, lngQueueID As Long) As OTPROCESSLib.Packages
'Find packages that have matching field values to those that are specified.
'One or more field/value pairs may be specified.
'Build up a profile of the specified field names and values and use the profile
'to get a collection of packages fitting that profile.
'Inputs:
'strAttribNames() - a zero-based one-dimensional array of package field names
'varAttribValues() - a zero-based one-dimensional array of package field values
'                   corresponding to the strAttribNames array
'strPkgTemplate - the name of the package template to search
'lngProcessID - the id for the process containing the queue of interest
'lngQueueID - the ID for the queue of interest

'Outputs:
'a packages collection containing packages with the specified field values

Dim objTemplates As OTPROCESSLib.Templates
Dim objTemplate As OTPROCESSLib.Template
Dim objFieldDefinitions As OTPROCESSLib.FieldDefinitions
Dim objFieldDefinition As OTPROCESSLib.FieldDefinition
Dim objProfile As OTPROCESSLib.Profile
Dim objProfileCriteria As OTPROCESSLib.ProfileCriteria
Dim objFieldCriteria As OTPROCESSLib.FieldCriteria
Dim objFieldCriterion As OTPROCESSLib.FieldCriterion
Dim objPackages As OTPROCESSLib.Packages
Dim intIndex As Integer                                 'input array index
Dim objLocCriteria As OTPROCESSLib.LocationCriteria
Dim objLocCriterion As OTPROCESSLib.LocationCriterion

'Enable error trap
    On Error GoTo errHandler

'Make sure the input arrays are the same size
    If UBound(strAttribNames) <> UBound(varAttribValues) Then
        'Raise an error
        Err.Raise 1, , "The number of specified field names does not match " & _
            "the number of specified field values"
    End If

'Instantiate and initialize a templates collection
    Set objTemplates = New OTPROCESSLib.Templates
    objTemplates.UserToken = mobjUserToken
    objTemplates.Refresh
     
'Find the specified package template in the templates collection
    Set objTemplate = objTemplates.FindByKey(strPkgTemplate, otKeyType_TemplateName)
     
'If the template was found then...
    If Not (objTemplate Is Nothing) Then
         
        'Get the fields collection the package template
        Set objFieldDefinitions = objTemplate.FieldDefinitions
        
        '@@@ list package field names
'        For intIndex = 1 To objFieldDefinitions.Count
'            Debug.Print objFieldDefinitions.Item(intIndex).FieldName, objFieldDefinitions.Item(intIndex).IsSystemField
'        Next intIndex
        '@@@
        
        'Instantiate a new profile object
        Set objProfile = New OTPROCESSLib.Profile
        
        'Instantiate a new profile criteria object and initialize it
        Set objProfileCriteria = New OTPROCESSLib.ProfileCriteria
        objProfileCriteria.TemplateId = objTemplate.TemplateId
        Set objProfileCriteria.UserToken = mobjUserToken
        
        'Get the field criteria collection object from the profile criteria
        Set objFieldCriteria = objProfileCriteria.FieldCriteria
                
        'For each field in the input arrays
        For intIndex = 0 To UBound(strAttribNames)
        
            'Find the specified field in the collection
            Set objFieldDefinition = objFieldDefinitions.FindByKey(strAttribNames(intIndex), _
                otKeyType_FieldName)
            
            'If the specified field was found then...
            If Not (objFieldDefinition Is Nothing) Then
                
                'Instantiate a new field criterion object
                Set objFieldCriterion = New OTPROCESSLib.FieldCriterion
                'Define criterion object as specified FieldName = specified FieldValue
                'objFieldCriterion.Set objFieldDefinition, varAttribValues(intIndex), 0, _
                    otOperator_EqualTo
                objFieldCriterion.SetOperation objFieldDefinition, otOperator_EqualTo
                objFieldCriterion.AddOperand varAttribValues(intIndex)
                'Add the criterion to the criteria collection
                objFieldCriteria.AddItem objFieldCriterion
             
            Else   'the specified field was not found
               'Raise an error
                Err.Raise 2, , "The specified field name '" & _
                    strAttribNames(intIndex) & "' was not found in the specified " & _
                    "package template '" & strPkgTemplate & "'"
            End If
        
        Next intIndex
    
        'Add a location criterion for the event
        'Get the field criteria collection object from the profile criteria
        Set objLocCriteria = objProfileCriteria.LocationCriteria
                
        'Instantiate a new location criterion object
        Set objLocCriterion = New OTPROCESSLib.LocationCriterion
        'Define queue by queueid
        objLocCriterion.SetProcessById lngProcessID
        objLocCriterion.AddQueueById lngQueueID
        'Add the criterion to the criteria collection
        objLocCriteria.AddItem objLocCriterion
        
        'Assign the updated criteria to the Profile
        Set objProfile.ProfileCriteria = objProfileCriteria
        
        'Instantiate a new packages collection
        Set objPackages = New OTPROCESSLib.Packages
        'Apply the profile to the packages collection and refresh to search
        Set objPackages.Profile = objProfile
        Set objPackages.UserToken = mobjUserToken
        objPackages.Refresh
        
        '@@@ list packages returned
'        For intIndex = 1 To objPackages.Count
'            Debug.Print objPackages(intIndex).PkgId, objPackages(intIndex).PkgRecId
'        Next intIndex
        '@@@
        
        'Return the packages object
        Set FindPackagesByAttribsAndQueue = objPackages
    
    Else    'specified package template not found
        'raise an error
        Err.Raise 3, , "The specified package template '" & _
            strPkgTemplate & "' was not found"
    
    End If
    
'Cleanup
    Set objTemplates = Nothing
    Set objFieldDefinitions = Nothing
    Set objFieldDefinition = Nothing
    Set objProfile = Nothing
    Set objProfileCriteria = Nothing
    Set objFieldCriteria = Nothing
    Set objFieldCriterion = Nothing
    Set objLocCriteria = Nothing
    Set objLocCriterion = Nothing
    Set objPackages = Nothing
    
Exit Function

errHandler:
    Select Case Err.Number
        Case 1  'mismatched input arrays
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribsAndQueue", _
                Err.Description)
        Case 2  'field name not in package template
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribsAndQueue", _
                Err.Description)
        Case 3  'package template not found
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribsAndQueue", _
                Err.Description)
        Case Else
            Call RaiseError(errFindPackagesByAttribs, CLASSNAME & ".FindPackagesByAttribsAndQueue", _
                Err.Number & "-" & Err.Description)
Resume  '@@@
    End Select
End Function

Public Function GetPackageAttachments(objPackage As OTPROCESSLib.Package) As OTPROCESSLib.Attachments
'Get the collection of attachments for a package.
'Inputs:
'objPackage - the object representing of the desired package
'Outputs:
'Package's attachment collection

'Enable error trap
    On Error GoTo errHandler

'If the specified package object is not valid then...
    If objPackage Is Nothing Then
        Err.Raise 1, , "Unable to update package journal because the specified " & _
            "package object is nothing"
    Else
        'Get the package attachments collection
        Set GetPackageAttachments = objPackage.Attachments
    End If

Exit Function

errHandler:
    Select Case Err.Number
        Case 1 'generated errors
            Call RaiseError(errGetPackageAttachments, CLASSNAME & ".GetPackageAttachments", _
                Err.Description)
        Case Else
            Call RaiseError(errGetPackageAttachments, CLASSNAME & ".GetPackageAttachments", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function


Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error. This procedure is called from another procedures error handler
'to perform logging and raise the error.

'Log the error
    RaiseEvent LogMsg("Error in " & Source & ": " & ErrorNumber & " - " & Description)
    
'Raise an error back to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub

Public Function FindPackageByID(strPkgID As String, strPkgTemplate As String) _
    As OTPROCESSLib.Package
'Find the package with the matching package ID. Return the first one found.
'Build up a profile of the specified package ID and use the profile
'to get a collection of packages fitting that profile.
'Note this functionality became available in 2.1 CR2.
'Inputs:
'strPkgID() - the ID of the package to search for (NOT a PkgRecID)
'strPkgTemplate - the name of the package template to search

'Outputs:
'a package object for the specified package

Dim objTemplates As OTPROCESSLib.Templates
Dim objTemplate As OTPROCESSLib.Template
'Dim objFieldDefinitions As OTPROCESSLib.FieldDefinitions
'Dim objFieldDefinition As OTPROCESSLib.FieldDefinition
Dim objProfile As OTPROCESSLib.Profile
'Dim objProfileCriteria As OTPROCESSLib.ProfileCriteria
'Dim objFieldCriteria As OTPROCESSLib.FieldCriteria
'Dim objFieldCriterion As OTPROCESSLib.FieldCriterion
Dim objPackages As OTPROCESSLib.Packages
'Dim intIndex As Integer                                 'input array index
    
'Enable error trap
    On Error GoTo errHandler

'Make sure a PkgID is specified
    If strPkgID = vbNullString Then
        'Raise an error
        Err.Raise 1, , "Unable to search for an existing package because the " & _
            "specified PkgID is blank"
    End If

'Instantiate and initialize a templates collection
    Set objTemplates = New OTPROCESSLib.Templates
    objTemplates.UserToken = mobjUserToken
    objTemplates.Refresh
     
'Find the specified package template in the templates collection
    Set objTemplate = objTemplates.FindByKey(strPkgTemplate, otKeyType_TemplateName)
     
'If the template was found then...
    If Not (objTemplate Is Nothing) Then
         
'        'Get the fields collection the package template
'        Set objFieldDefinitions = objTemplate.FieldDefinitions
        
        '@@@ list package field names
'        For intIndex = 1 To objFieldDefinitions.Count
'            Debug.Print objFieldDefinitions.Item(intIndex).FieldName, objFieldDefinitions.Item(intIndex).IsSystemField
'        Next intIndex
        '@@@
        
        'Instantiate a new profile object
        Set objProfile = New OTPROCESSLib.Profile
        
        'Initialize profile to search on PkgID
        'objProfile.ProfileType = otProfileType_PkgId
        objProfile.PkgId = strPkgID
        
        
'        'Instantiate a new profile criteria object and initialize it
'        Set objProfileCriteria = New OTPROCESSLib.ProfileCriteria
'        objProfileCriteria.TemplateId = objTemplate.TemplateId
'        Set objProfileCriteria.UserToken = mobjUserToken
'
'        'Get the field criteria collection object from the profile criteria
'        Set objFieldCriteria = objProfileCriteria.FieldCriteria
'
'        'For each field in the input arrays
'        For intIndex = 0 To UBound(strAttribNames)
'
'            'Find the specified field in the collection
'            Set objFieldDefinition = objFieldDefinitions.FindByKey(strAttribNames(intIndex), _
'                otKeyType_FieldName)
'
'            'If the specified field was found then...
'            If Not (objFieldDefinition Is Nothing) Then
'
'                'Instantiate a new field criterion object
'                Set objFieldCriterion = New OTPROCESSLib.FieldCriterion
'                'Define criterion object as specified FieldName = specified FieldValue
'                objFieldCriterion.Set objFieldDefinition, varAttribValues(intIndex), 0, _
'                    otOperator_EqualTo
'                'Add the criterion to the criteria collection
'                objFieldCriteria.AddItem objFieldCriterion
'
'            Else   'the specified field was not found
'               'Raise an error
'                Err.Raise 2, , "The specified field name '" & _
'                    strAttribNames(intIndex) & "' was not found in the specified " & _
'                    "package template '" & strPkgTemplate & "'"
'            End If
'
'        Next intIndex
'
'        'Assign the criteria to the Profile
'        Set objProfile.ProfileCriteria = objProfileCriteria
        
        'Instantiate a new packages collection
        Set objPackages = New OTPROCESSLib.Packages
        'Apply the profile to the packages collection and refresh to search
        Set objPackages.Profile = objProfile
        Set objPackages.UserToken = mobjUserToken
        objPackages.Refresh
        
        '@@@ list packages returned
'        For intIndex = 1 To objPackages.Count
'            Debug.Print objPackages(intIndex).PkgId, objPackages(intIndex).PkgRecId
'        Next intIndex
        '@@@
        
        'Return the packages object
        If objPackages.Count > 0 Then
            RaiseEvent LogMsg("(cWorkflowII.FindPackageByID) " & objPackages.Count & " package(s) found for pkgid '" & strPkgID & "'")
            Set FindPackageByID = objPackages.Item(1)
        Else
            RaiseEvent LogMsg("(cWorkflowII.FindPackageByID) 0 packages found for pkgid '" & strPkgID & "'")
            Set FindPackageByID = Nothing
        End If
    
    Else    'specified package template not found
        'raise an error
        Err.Raise 3, , "The specified package template '" & _
            strPkgTemplate & "' was not found"
    
    End If
    
'Cleanup
    Set objTemplates = Nothing
'    Set objFieldDefinitions = Nothing
'    Set objFieldDefinition = Nothing
    Set objProfile = Nothing
'    Set objProfileCriteria = Nothing
'    Set objFieldCriteria = Nothing
    Set objPackages = Nothing
    
Exit Function

errHandler:
    Select Case Err.Number
        Case 1  'mismatched input arrays
            Call RaiseError(errFindPackageByID, CLASSNAME & ".FindPackageByID", _
                Err.Description)
'        Case 2  'field name not in package template
'            Call RaiseError(errFindPackageByID, CLASSNAME & ".FindPackageByID", _
'                Err.Description)
        Case 3  'package template not found
            Call RaiseError(errFindPackageByID, CLASSNAME & ".FindPackageByID", _
                Err.Description)
        Case Else
            Call RaiseError(errFindPackageByID, CLASSNAME & ".FindPackageByID", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select

End Function

Public Function GetPackagesInQueue(strQueue As String, strProcessName As String) As OTPROCESSLib.Packages
'Return a collection of package objects containing packages in the specified queue(s).
'Inputs:
'strQueue - the name of the queue of interest
'strProcessName - the name of the process contaiing the queue

Dim objQueues As OTPROCESSLib.Queues
Dim objQueue    As OTPROCESSLib.Queue
Dim objProfile  As OTPROCESSLib.Profile
Dim objProcesses As OTPROCESSLib.Processes
Dim objProcess As OTPROCESSLib.Process
Dim objPackages As OTPROCESSLib.Packages

'Enable error trap
    On Error GoTo errHandler

'Instantiate and initialize a new processes collection
    Set objProcesses = New OTPROCESSLib.Processes
    
'If user token is valid then...
    If Not (mobjUserToken Is Nothing) Then
        'Init collection with user token
        objProcesses.UserToken = mobjUserToken
        objProcesses.Refresh
    Else
        'Token object is nothing, raise an error
        Err.Raise 1, , "User token is invalid"
    End If
    
'Get an object for the specified process
    Set objProcess = objProcesses.FindByKey(strProcessName, otKeyType_ProcessName)
    If objProcess Is Nothing Then
        Err.Raise 2, , "The specified process name '" & strProcessName & _
            "' was not found in the database"
    End If
    
'Initialize the profile object to be a queue profile
    Set objProfile = New OTPROCESSLib.Profile
    'objProfile.ProfileType = otProfileType_Queue
    objProfile.QueueSelectType = otQSelectType_All
    
'Get an object for the specified queue
    Set objQueues = objProcess.Queues
    Set objQueue = objQueues.FindByKey(strQueue, otKeyType_QueueName)
    
'If the queue was not found then...
    If objQueue Is Nothing Then
        Err.Raise 2, , "The specified queue name '" & strQueue & _
            "' was not found in the process"
    Else
        'Initialize the profile's queue object
        objProfile.Queue = objQueue
        
        'Determine the queue's behavior
        Select Case objQueue.Behavior
            Case otQueueBehavior_FIFO   'only returns one package
                Debug.Print "Queue behavior is FIFO"
            Case otQueueBehavior_WIP    'returns all packages
                Debug.Print "Queue behavior is WIP"
        End Select
    End If
    
'Get the packages in the queue
    Set objPackages = New OTPROCESSLib.Packages
    Set objPackages.UserToken = objUserToken
    Set objPackages.Profile = objProfile
    objPackages.Refresh
     
    Debug.Print objPackages.Count
    
'Return the collection
    Set GetPackagesInQueue = objPackages
 
Exit Function
 
errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGetPackagesInQueue, CLASSNAME & ".GetPackagesInQueue", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select

End Function

Public Function GetProcesses() As OTPROCESSLib.Processes
'Get the processes in the database.

Dim objProcesses As OTPROCESSLib.Processes

    Set objProcesses = New OTPROCESSLib.Processes
    Set objProcesses.UserToken = mobjUserToken
    objProcesses.Refresh
    Set GetProcesses = objProcesses
    
End Function

Public Function GetQueues(strProcessName As String) As OTPROCESSLib.Queues
'Get the processes in the database.

Dim objProcesses As OTPROCESSLib.Processes
Dim objProcess As OTPROCESSLib.Process
Dim objQueues As OTPROCESSLib.Queues

'Get the processes collection for the initialized database
    Set objProcesses = New OTPROCESSLib.Processes
    Set objProcesses.UserToken = mobjUserToken
    objProcesses.Refresh
    
'Get an object for the specified process
    Set objProcess = objProcesses.FindByKey(strProcessName, otKeyType_ProcessName)
    If objProcess Is Nothing Then
        Err.Raise 2, , "The specified process name '" & strProcessName & _
            "' was not found in the database"
    End If
    
'Get the queues collection
    Set objQueues = objProcess.Queues
    Set GetQueues = objQueues
    
Exit Function
 
errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGetQueues, CLASSNAME & ".GetQueues", _
                Err.Number & "-" & Err.Description)
'Resume  '@@@
    End Select
End Function

