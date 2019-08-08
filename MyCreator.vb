'-----------------------------------------------------------------------------
'                                                                            |
'                   Softing Industrial Automation GmbH                       |
'                        Richard-Reitzner-Allee 6                            |
'                           85540 Haar, Germany                              |
'                                                                            |
'                 This is a part of the Softing OPC Toolkit                  |
'        Copyright (c) 2005-2018 Softing Industrial Automation GmbH          |
'                           All Rights Reserved                              |
'                                                                            |
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'                             OPC Toolkit NET                                |
'                                                                            |
'  Filename    : MyCreator.vb                                                |
'  Version     : 4.46                                                        |
'  Date        : 15-March-2018                                               |
'                                                                            |
'  Description : User's speciffic OPC Server's objects creator class         |
'                                                                            |
'-----------------------------------------------------------------------------

Imports System
Imports System.Collections
Imports System.Text
Imports System.Runtime.InteropServices
Imports Softing.OPCToolbox
Imports Softing.OPCToolbox.Server
Namespace OPCSimpleTrial1
  Public Class MyCreator
  Inherits Creator
#Region "Public Methods"
    '---------------------
    Public Overloads Overrides Function CreateInternalDaAddressSpaceElement(ByVal anItemId As String, ByVal aName As String, ByVal anUserData As System.UInt32, ByVal anObjectHandle As UInt32, ByVal aParentHandle As UInt32) As DaAddressSpaceElement
      Return CType(New MyDaAddressSpaceElement(anItemId, aName, anUserData, anObjectHandle, aParentHandle), DaAddressSpaceElement)
    End Function

    Public Overloads Overrides Function CreateDaAddressSpaceRoot() As DaAddressSpaceRoot
      Return CType(New MyDaAddressSpaceRoot, DaAddressSpaceRoot)
    End Function

    Public Overloads Overrides Function CreateTransaction(ByVal transactionType As EnumTransactionType, ByVal requestList As DaRequest(), ByVal sessionKey As UInt32) As DaTransaction
      Return CType(New MyTransaction(transactionType, requestList, sessionKey), DaTransaction)
    End Function

    Public Overridable Function CreateMyDaAddressSpaceElement() As DaAddressSpaceElement
      Return CType(New MyDaAddressSpaceElement, DaAddressSpaceElement)
    End Function

    Public Overloads Overrides Function CreateRequest(ByVal aTransactionType As EnumTransactionType, ByVal aSessionHandle As UInt32, ByVal anElement As DaAddressSpaceElement, ByVal aPropertyId As Integer, ByVal aRequestHandle As UInt32) As DaRequest
      Return New MyRequest(aTransactionType, aSessionHandle, anElement, aPropertyId, aRequestHandle)
    End Function
    '--
#End Region
  End Class
End Namespace
