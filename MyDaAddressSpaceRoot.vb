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
'                address space root element definition                       |
'                                                                            |
'-----------------------------------------------------------------------------
Imports System
Imports System.Collections
Imports System.Text
Imports Softing.OPCToolbox
Imports Softing.OPCToolbox.Server
Namespace OPCSimpleTrial1
  Public Class MyDaAddressSpaceRoot
  Inherits DaAddressSpaceRoot
#Region "Public Methods"
    '--------------------
    Public Overloads Overrides Function QueryAddressSpaceElementData(ByVal elementId As String, ByRef anAddressSpaceElement As AddressSpaceElement) As System.Int32
      '  TODO: add string based address space validations
      anAddressSpaceElement = Nothing
      Return Convert.ToInt32(EnumResultCode.E_NOTIMPL)
    End Function
    '--
#End Region
  End Class
End Namespace
