Attribute VB_Name = "modScript"
'*****************************************************************************************
' Module Name:  modScript
' Author:       Wes Prichard, Optika
' Date:         11/27/2001
' Description:  This is a documentation file that contains the custom script used
' in the process script events. This is the code that invokes the DLL in this project.
' The code in this module is not actually part of this project but is listed for
' reference. Copy this code into the process script event.
'
' Edit History:
' mm/dd/yyyy - Modified by name, company
'   Description of change
'
'*****************************************************************************************

Option Explicit

'UNCOMMENT AND USE THE FOLLOWING SCRIPTS IN THE PROCESS SCRIPT EVENT:

'Sub ProcessScriptingFramework1(objExecutionContext)
''Scripting Framework Entry Point
''This code is executed by Package Broker inside the script event.
'
'    ScriptExecute (objExecutionContext)    '(copy this line only into script)
'
'End Sub
'
'
'
'Sub ScriptExecute(objExecutionContext)
''This procedure invokes a custom DLL to perform the actual processing.
'
'Dim objUpdate
'
'    On Error Resume Next
'
'    'Instantiate an object from the custom scripts DLL
'    Set objUpdate = CreateObject("PayablesAVEScripts.cLookupApprovers")
'    'If an error occurred then...
'    If Eval("Err.Number <> 0") Then
'        objExecutionContext.ResultStatus = False
'        objExecutionContext.ErrorDescription = objExecutionContext.ErrorDescription & " (Error creating PayablesAVEScripts.cLookupApprover: " & Err.Description & ")"
'        Exit Sub
'    End If
'
'    'Execute the Update for the current package
'    objUpdate.Execute (objExecutionContext)
'    If Eval("Err.Number <> 0") Then
'        objExecutionContext.ResultStatus = False
'        objExecutionContext.ErrorDescription = objExecutionContext.ErrorDescription & " (Error in .Execute: " & Err.Description & ")"
'        Exit Sub
'    End If
'
'    objExecutionContext.ResultStatus = True
'
'End Sub


