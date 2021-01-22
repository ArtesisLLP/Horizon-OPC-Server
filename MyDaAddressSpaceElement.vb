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
'  Filename    : MyDaAddressSpaceElement.vb                                  |
'  Version     : 4.46                                                        |
'  Date        : 15-March-2018                                               |
'                                                                            |
'  Description : User's speciffic DataAccess OPC Server                      |
'                address space element definition                            |
'                                                                            |
'-----------------------------------------------------------------------------
Imports System
Imports System.Collections
Imports System.Text
Imports Softing.OPCToolbox
Imports Softing.OPCToolbox.Server
Namespace HorizonOPCServer 'OPCSimpleTrial1
    Public Class MyDaAddressSpaceElement
        Inherits DaAddressSpaceElement
#Region "Constructors"
        '------------------------------
        Public Sub New(ByVal anItemID As String, ByVal aName As String, ByVal anUserData As System.UInt32, ByVal anObjectHandle As UInt32, ByVal aParentHandle As UInt32)
            MyBase.New(anItemID, aName, anUserData, anObjectHandle, aParentHandle)
        End Sub

        Public Sub New()
        End Sub
        '--
#End Region

#Region "Private Attributes"
        '-------------------------------
        Private m_properties As Hashtable = New Hashtable
        '--
#End Region

#Region "Public Methods"
        '---------------------
        ''' <summary>
        ''' Get elements property value data
        ''' </summary>
        Public Sub GetPropertyValue(ByVal aRequest As DaRequest)
            If aRequest.PropertyId = Convert.ToInt32(EnumPropertyId.ITEM_DESCRIPTION) Then
                aRequest.Value = New ValueQT("description", [Enum].ToObject(GetType(EnumQuality), EnumQuality.GOOD), DateTime.Now)
                aRequest.Result = [Enum].ToObject(GetType(EnumResultCode), EnumResultCode.S_OK)
            Else
                aRequest.Result = [Enum].ToObject(GetType(EnumResultCode), EnumResultCode.E_NOTFOUND)
            End If
        End Sub

        Public Overloads Overrides Function QueryProperties(ByRef aPropertyList As ArrayList) As System.Int32
            If m_properties.Count > 0 Then
                aPropertyList = New ArrayList
                aPropertyList.AddRange(m_properties.Values)
            Else
                aPropertyList = Nothing
            End If
            Return Convert.ToInt32(EnumResultCode.S_OK)
        End Function

        Public Function AddProperty(ByVal aProperty As DaProperty) As System.Int32
            If Not (aProperty Is Nothing) Then
                m_properties.Add(aProperty.Id, aProperty)
                Return Convert.ToInt32(EnumResultCode.S_OK)
            Else
                Return Convert.ToInt32(EnumResultCode.S_FALSE)
            End If
        End Function
        '--
#End Region
    End Class
End Namespace
