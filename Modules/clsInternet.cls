VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************************************************************************************************
'*      LIBRARY NAME: clsInternet.cls
'*      PURPOSE: Compilation of functions that deals with Internet
'*      AUTHOR: Roland Ross Hadi
'*      CREATION DATE: April 07, 2009
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
        Option Explicit
        Private Declare Function InternetGetConnectedState _
        Lib "wininet.dll" (ByRef lpdwFlags As Long, _
        ByVal dwReserved As Long) As Long

        ' Local system uses a modem to connect to the Internet
        Private Const INTERNET_CONNECTION_MODEM As Long = &H1

        ' Local system uses a LAN to connect to the Internet.
        Private Const INTERNET_CONNECTION_LAN As Long = &H2

        ' Local system uses a proxy server to connect to the Internet.
        Private Const INTERNET_CONNECTION_PROXY As Long = &H4

        ' No longer used
        Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
        Private Const INTERNET_RAS_INSTALLED As Long = &H10
        Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
        Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
'*******************************************************************************************************************************************************************************************

' @Function IsNetConnectViaLAN
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is Net Connect Via LAN
'@Comments
' Parameter:
'
Public Function IsNetConnectViaLAN() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the API will
    'return the flags associated with the connection
    Call InternetGetConnectedState(dwFlags, 0&)

    'return True if the flags indicate a LAN connection
    IsNetConnectViaLAN = dwFlags And INTERNET_CONNECTION_LAN
End Function
' Function IsNetConnectViaLAN
' -----------------------------

' @Function IsNetConnectViaModem
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is Net Connect Via Modem
'@Comments
' Parameter:
'
Public Function IsNetConnectViaModem() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the API will
    'return the flags associated with the connection
    Call InternetGetConnectedState(dwFlags, 0&)

    'return True if the flags indicate a modem connection
    IsNetConnectViaModem = dwFlags And INTERNET_CONNECTION_MODEM
End Function
' Function IsNetConnectViaModem
' -----------------------------

' @Function IsNetConnectViaProxy
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is Net Connect Via Modem
'@Comments
' Parameter:
'
Public Function IsNetConnectViaProxy() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the API will
    'return the flags associated with the connection
    Call InternetGetConnectedState(dwFlags, 0&)

    'return True if the flags indicate a proxy connection
    IsNetConnectViaProxy = dwFlags And INTERNET_CONNECTION_PROXY
End Function
' Function IsNetConnectViaProxy
' -----------------------------

' @Function IsNetConnectOnline
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is Net Connect Online
'@Comments
' Parameter:
'
Public Function IsNetConnectOnline() As Boolean
    'no flags needed here - the API returns True
    'if there is a connection of any type
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function
' Function IsNetConnectOnline
' -----------------------------

' @Function IsNetRASInstalled
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is Net RAS Installed
'@Comments
' Parameter:
'
Public Function IsNetRASInstalled() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the API will
    'return the flags associated with the connection
    Call InternetGetConnectedState(dwFlags, 0&)

    'return True if the falgs include RAS installed
    IsNetRASInstalled = dwFlags And INTERNET_RAS_INSTALLED
End Function
' Function IsNetRASInstalled
' -----------------------------

' @Function GetNetConnectString
' -----------------------------
'@Author Roland Ross Hadi
'@Description Is Net RAS Installed
'@Comments
' Parameter:
'
Public Function GetNetConnectString() As String
    Dim dwFlags As Long
    Dim msg As String

    'build a string for display
    If InternetGetConnectedState(dwFlags, 0&) Then
        If dwFlags And INTERNET_CONNECTION_CONFIGURED Then
            msg = msg & "You have a network connection configured." & vbCrLf
        End If

        If dwFlags And INTERNET_CONNECTION_LAN Then
            msg = msg & "The local system connects to the Internet via a LAN"
        End If

        If dwFlags And INTERNET_CONNECTION_PROXY Then
            msg = msg & ", and uses a proxy server. "
        Else: msg = msg & "."
        End If

        If dwFlags And INTERNET_CONNECTION_MODEM Then
            msg = msg & "The local system uses a modem to connect to the Internet. "
        End If

        If dwFlags And INTERNET_CONNECTION_OFFLINE Then
            msg = msg & "The connection is currently offline. "
        End If

        If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then
            msg = msg & "The local system's modem is busy With a non-Internet connection. "
        End If

        If dwFlags And INTERNET_RAS_INSTALLED Then
            msg = msg & "Remote Access Services are installed On this system."
        End If
    Else
        msg = "Not connected to the internet now."
    End If

    GetNetConnectString = msg
End Function
' Function GetNetConnectString
' -----------------------------

' @Function ValidateEmail
' -----------------------------
'@Author Roland Ross Hadi
'@Description Validate a given e-mail address.
'@Comments
' Parameter:
'     strAddress - string to be validated
Public Function ValidateEmail(ByVal strAddress As String) As Boolean
    Dim lngIndex As Long            ' Position in strAddress
    Dim lngCountAt As Long          ' Number of "@"
    Dim lngLastDotPos As Long       ' Position of the previous dot in the string
    Dim strCurrentChar As String    ' Buffer that holds the contents of the string one char at a time.

    On Error GoTo Fail_Validation

    ValidateEmail = True            ' Prove me wrong!
    strAddress = Trim(strAddress)
    lngLastDotPos = 0
    lngCountAt = 0

    ' If the address isn't at least this (a@b.com) long,
    ' it's not a valid address.
    If Len(strAddress) < 7 Then GoTo Fail_Validation

    ' Check for certain generably allowable characters in the leading position.
    ' If found, it's not a valid address.
    strCurrentChar = left$(strAddress, 1)
    If strCurrentChar = "." Or strCurrentChar = "@" Or strCurrentChar = "_" Or _
        strCurrentChar = "-" Then GoTo Fail_Validation

    ' Check the string for non-allowable characters.
    For lngIndex = 1 To Len(strAddress)
        strCurrentChar = Mid$(strAddress, lngIndex, 1)

        ' Count the number of "@".
        If strCurrentChar = "@" Then lngCountAt = lngCountAt + 1

        ' If there are two consecutive dots, it's not a valid address.
        If strCurrentChar = "." Then
            If lngIndex = lngLastDotPos + 1 Then
                GoTo Fail_Validation
            Else
                lngLastDotPos = lngIndex
            End If
        End If

        Select Case Asc(strCurrentChar)
            ' These characters are not allowable in e-mail addresses.
            Case 1 To 44, 47, 58 To 63, 91 To 94, 96, 123 To 127, 128 To 255
                GoTo Fail_Validation
        End Select
    Next lngIndex

    ' If there isn't one, and only one "@", then it's not a valid address.
    If lngCountAt <> 1 Then GoTo Fail_Validation

    ' If the extension isn't a known one, it's not a valid address.
    Select Case right$(strAddress, 4)
        Case ".com", ".org", ".net", ".edu", ".mil", ".gov"
            ' Yes, it's valid.
        Case Else
            GoTo Fail_Validation
    End Select

ValidateEMail_Exit:
    Exit Function

Fail_Validation:
    ValidateEmail = False
    GoTo ValidateEMail_Exit
End Function
' Function ValidateEmail
' -----------------------------


