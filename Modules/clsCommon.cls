VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsCommon.cls
'*      PURPOSE: Compilation of functions that deals with Common functions
'*      AUTHOR: Roland Ross Hadi
'*      CREATION DATE: May 06, 2009
'*      NOTES:
'*
'*******************************************************************************************************************************************************************************************
'*      MODIFICATION HISTORY
'*      NAME                                            DATE OF CHANGE                                          FUNCTION NAME
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*      *:                                              *:                                                      *:
'*******************************************************************************************************************************************************************************************

'*******************************************************************************************************************************************************************************************
'*      GENERAL DECLARATION

'*******************************************************************************************************************************************************************************************

' @Function HandleError
' -----------------------------
'@Author Roland Ross Hadi
'@Description Handle unexcpected error
'@Comments
' Parameter:
'       ErrHeader - Error header description
'       ErrWindowObj - Form that contains error
'       ScreenError - Error object
Public Sub HandleError(ErrHeader As String, ErrWindowObj As Form, ScreenError As ErrObject)
Dim tmpErrorNumber As String
Dim tmpErrorDescription As String
tmpErrorNumber = ScreenError.Number
tmpErrorDescription = ScreenError.Description
frmError.Caption = "Runtime Error"
frmError.lblErrorDescription.Caption = "Error Description: " & ErrHeader & vbCrLf & vbCrLf & _
                                       "Technical Information " & _
                                       "Current User: " & curUser & vbCrLf & _
                                       "Window Name: " & ErrWindowObj.Caption & vbCrLf & _
                                       "Error Date: " & Format(Now, "mmmm dd, yyyy") & vbCrLf & _
                                       "Error Time: " & Format(Now, "hh:mm:ss") & vbCrLf & _
                                       "Technical Error Number: " & tmpErrorNumber & vbCrLf & _
                                       "Technical Error Description: " & tmpErrorDescription & vbCrLf
WriteToEndOfFile App.Path & "\Client Error - " & Format(Now, "mmm-dd-yy hhmmss") & ".txt", frmError.lblErrorDescription.Caption
frmError.Show
End Sub
' Function HandleError
' -----------------------------

' @Function HandleErrorMessageBox
' -----------------------------
'@Author Roland Ross Hadi
'@Description Handle unexcpected error
'@Comments
' Parameter:
'       ErrHeader - Error header description
'       ErrWindowObj - Form that contains error
'       ScreenError - Error object
Public Sub HandleErrorMessageBox(ErrHeader As String, ErrWindowObj As Form, ScreenError As ErrObject)
Dim tmpErrorNumber As String
Dim tmpErrorDescription As String
Dim tmp As String
tmpErrorNumber = ScreenError.Number
tmpErrorDescription = ScreenError.Description
tmp = "Error Description: " & ErrHeader & vbCrLf & vbCrLf & _
                                       "Technical Information " & _
                                       "Window Name: " & ErrWindowObj.Caption & vbCrLf & _
                                       "Error Date: " & Format(Now, "mmmm dd, yyyy") & vbCrLf & _
                                       "Error Time: " & Format(Now, "hh:mm:ss") & vbCrLf & _
                                       "Technical Error Number: " & tmpErrorNumber & vbCrLf & _
                                       "Technical Error Description: " & tmpErrorDescription & vbCrLf
WriteToEndOfFile App.Path & "\Client Error - " & Format(Now, "mmm-dd-yy hhmmss") & ".txt", tmp
MsgBox tmp, vbCritical, "Runtime Error"
End Sub
' Function HandleErrorMessageBox
' -----------------------------

' Support Function for HandleError & HandleErrorMessageBox
' -----------------------------
Private Function WriteToEndOfFile(strPath As String, strValue As String)
    If FileExists(strPath) = True Then
      Open strPath For Append As #1
      Print #1, strValue
      Close #1
    Else
      WriteNewFile strPath, strValue
    End If
End Function
Private Function WriteNewFile(strPath As String, strValue As String)
    Open strPath For Output As #1
    Print #1, strValue
    Close #1
End Function
Private Function FileExists(ByVal strPathName As String) As Boolean
        If Dir$(strPathName) <> "" Then
            FileExists = True
        Else
            FileExists = False
        End If
End Function
' Support Function for HandleError & HandleErrorMessageBox
' -----------------------------
