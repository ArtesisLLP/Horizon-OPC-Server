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
'  Filename    : MyTransaction.vb                                            |
'  Version     : 4.46                                                        |
'  Date        : 15-March-2018                                               |
'                                                                            |
'  Description : User's speciffic DataAccess OPC server                      |
'                DaTransaction definition                                    |
'                                                                            |
'-----------------------------------------------------------------------------
Imports System
Imports System.Collections
Imports System.Text
Imports Softing.OPCToolbox
Imports Softing.OPCToolbox.Server

Namespace OPCSimpleTrial1
  Public Class MyTransaction
        Inherits DaTransaction
#Region "Public Methods"

        Public Sub New(ByVal aTransactionType As EnumTransactionType, ByVal requestList As DaRequest(), ByVal aSessionKey As UInt32)
            MyBase.New(aTransactionType, requestList, aSessionKey)
        End Sub

        Public Overloads Overrides Function HandleReadRequests() As System.Int32
            Dim count As Integer = Requests.Count
            Dim i As Integer = 0
            While i < count
                Dim request As DaRequest = CType(Requests(i), DaRequest)

                If request.ProgressRequestState(EnumRequestState.CREATED, EnumRequestState.INPROGRESS) = True Then
                    If request.PropertyId = 0 Then
                        ' get address space element value
                        ' take the toolkit cache value
                        Dim cacheValue As ValueQT = Nothing
                        request.AddressSpaceElement.GetCacheValue(cacheValue)
                        request.Value = cacheValue
                        request.Result = [Enum].ToObject(GetType(EnumResultCode), EnumResultCode.S_OK)
                    Else
                        Dim element As MyDaAddressSpaceElement = CType(request.AddressSpaceElement, MyDaAddressSpaceElement)

                        If Not (element Is Nothing) Then
                            element.GetPropertyValue(request)
                        Else
                            request.Result = [Enum].ToObject(GetType(EnumResultCode), EnumResultCode.E_FAIL)
                        End If
                    End If
                End If

                i += 1
            End While
            Return CompleteRequests
        End Function

        Public Overloads Overrides Function HandleWriteRequests() As System.Int32
            Dim count As Integer = Requests.Count
            Dim i As Integer = 0
            While i < count
                Dim request As DaRequest = CType(Requests(i), DaRequest)

                If request.ProgressRequestState(EnumRequestState.CREATED, EnumRequestState.INPROGRESS) = True Then
                    If Not request Is Nothing Then
                        Dim element As MyDaAddressSpaceElement = CType(request.AddressSpaceElement, MyDaAddressSpaceElement)

                        If Not element Is Nothing Then
                            request.Result = [Enum].ToObject(GetType(EnumResultCode), element.ValueChanged(request.Value))
                        End If
                    End If
                End If

                i += 1
            End While
            Return CompleteRequests
        End Function
#End Region
    End Class
End Namespace
