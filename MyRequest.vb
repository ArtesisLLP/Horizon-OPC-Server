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
'  Filename    : MyRequest.vb                                                |
'  Version     : 4.46                                                        |
'  Date        : 15-March-2018                                               |
'                                                                            |
'  Description : User's speciffic DataAccess OPC server                      |
'                DaRequest definition                                        |
'                                                                            |
'-----------------------------------------------------------------------------
Imports System
Imports System.Collections
Imports System.Text
Imports Softing.OPCToolbox
Imports Softing.OPCToolbox.Server

Namespace OPCSimpleTrial1
  Public Class MyRequest
  Inherits DaRequest
#Region "Constructor"
    Public Sub New(ByVal transactionType As EnumTransactionType, ByVal sessionHandle As UInt32, ByVal aDaAddressSpaceElement As DaAddressSpaceElement, ByVal propertyID As Integer, ByVal requestHandle As UInt32)
      MyBase.New(transactionType, sessionHandle, aDaAddressSpaceElement, propertyID, requestHandle)
    End Sub
#End Region
  End Class
End Namespace
