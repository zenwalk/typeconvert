Option Strict Off
Option Explicit On 
Imports System.Runtime.InteropServices

<Guid("977A840F-6916-4534-AF66-ECA29C97A795")> _
Public Class GPTypeConvertFactory
    Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        GPFunctionFactories.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        GPFunctionFactories.Unregister(regKey)
    End Sub
#End Region

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunctionFactory_Alias() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.Alias
        Get
            Return "TypeConvert Tools"
        End Get
    End Property

    Private ReadOnly Property IGPFunctionFactory_Name() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.Name
        Get
            Return "TypeConvertTools"
        End Get
    End Property

    Private Function IGPFunctionFactory_GetFunction(ByVal name As String) As ESRI.ArcGIS.Geoprocessing.IGPFunction Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.GetFunction
        Select Case name
            Case "ConvertToPolyline"
                Return New ConvertToPolyline
            Case "ConvertToPolygon"
                Return New ConvertToPolygon
            Case "ConvertToPoint"
                Return New ConvertToPoint
            Case "ConvertToSegments"
                Return New ConvertToSegments
            Case "ConvertToCentroids"
                Return New ConvertToCentroids
            Case "ConvertToEnvelope"
                Return New ConvertToEnvelope
            Case "ConvertToConvexHull"
                Return New ConvertToConvexHull
            Case "RemoveDuplicates"
                Return New RemoveDuplicates
            Case "ConvertToBln"
                Return New ConvertToBln
            Case ("Stratification")
                Return New Stratification
            Case Else
                Return (Nothing)
        End Select
    End Function

    Private Function IGPFunctionFactory_GetFunctionEnvironments() As ESRI.ArcGIS.Geoprocessing.IEnumGPEnvironment Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.GetFunctionEnvironments
        Return (Nothing)
    End Function

    Private Function IGPFunctionFactory_GetFunctionName(ByVal name As String) As ESRI.ArcGIS.Geodatabase.IGPName Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.GetFunctionName
        Dim pGPFunction As ESRI.ArcGIS.Geoprocessing.IGPFunction
        pGPFunction = IGPFunctionFactory_GetFunction(name)

        IGPFunctionFactory_GetFunctionName = pGPFunction.FullName
    End Function

    Private Function IGPFunctionFactory_GetFunctionNames() As ESRI.ArcGIS.Geodatabase.IEnumGPName Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.GetFunctionNames
        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        pArray = New ESRI.ArcGIS.Geoprocessing.EnumGPName

        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToPolyline"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToPolygon"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToPoint"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToSegments"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToCentroids"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToEnvelope"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToConvexHull"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("RemoveDuplicates"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("ConvertToBln"))
        pArray.Add(IGPFunctionFactory_GetFunctionName("Stratification"))
        IGPFunctionFactory_GetFunctionNames = pArray
    End Function


    Public ReadOnly Property CLSID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.Geoprocessing.IGPFunctionFactory.CLSID
        Get
            Dim pUID As New ESRI.ArcGIS.esriSystem.UID
            pUID.Value = "GPTypeConvert.GPTypeConvertFactory"
            CLSID = pUID
        End Get

    End Property
End Class