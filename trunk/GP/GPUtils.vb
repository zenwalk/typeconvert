Option Strict Off
Option Explicit On
Module GPUtils

	
	Public Function CheckDatasetExists(ByRef pGPParameterIn As ESRI.ArcGIS.Geoprocessing.IGPParameter, ByRef pGPValueIn As ESRI.ArcGIS.Geodatabase.IGPValue, ByRef pGPMessages As ESRI.ArcGIS.Geodatabase.IGPMessages) As Boolean
		
		
		If pGPValueIn.IsEmpty Then
			CheckDatasetExists = False
			Exit Function
		End If
		
		Dim pGPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities
		Dim pGpParameter As ESRI.ArcGIS.Geoprocessing.IGPParameter
		Dim pGPVariable As ESRI.ArcGIS.Geodatabase.IGPVariable
		Dim pGPValue As ESRI.ArcGIS.Geodatabase.IGPValue
		Dim bDerived As Boolean
		'
		
		pGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
		If Not pGPUtilities.IsDatasetType(pGPValueIn) Then
			' Exit Function
		End If
		'
		bDerived = False
		'
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
        If TypeOf pGPParameterIn Is ESRI.ArcGIS.Geoprocessing.IGPParameter Then
            pGpParameter = pGPParameterIn
            pGPValue = pGpParameter.Value
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If TypeOf pGPValue Is ESRI.ArcGIS.Geodatabase.IGPVariable Then
                pGPVariable = pGPValue
                bDerived = pGPVariable.Derived
            End If
        End If
        '
        If Not bDerived And pGPUtilities.IsDatasetType(pGPValueIn) Then
            If Not pGPUtilities.Exists(pGPValueIn) Then
                pGPMessages.AddError(1, "Dataset " & pGPValueIn.GetAsText & " does not exists")
                CheckDatasetExists = False
            Else
                CheckDatasetExists = True
            End If
        End If
		
	End Function
	
	
	Public Function GetParameterValueByName(ByRef paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByRef sName As String) As ESRI.ArcGIS.Geodatabase.IGPValue
		
		Dim pGPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities
		pGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
		Dim pGpParameter As ESRI.ArcGIS.Geoprocessing.IGPParameter
		
		Dim lIndex As Integer
		For lIndex = 0 To (paramvalues.Count - 1)
			pGpParameter = paramvalues.Element(lIndex)
			
			If UCase(pGpParameter.name) = UCase(sName) Then
				GetParameterValueByName = pGPUtilities.UnpackGPValue(pGpParameter)
				Exit Function
			End If
		Next lIndex
		
        GetParameterValueByName = Nothing
		
	End Function
	
	Public Function GetParameterByName(ByRef paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByRef sName As String) As ESRI.ArcGIS.Geoprocessing.IGPParameter
		
		sName = UCase(sName)
		
		Dim n As Integer
		Dim pGpParameter As ESRI.ArcGIS.Geoprocessing.IGPParameter
		Dim pGPUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities
		pGPUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
		
		For n = 0 To paramvalues.Count - 1 Step 1
			pGpParameter = paramvalues.Element(n)
			
			If UCase(pGpParameter.name) = sName Then
				GetParameterByName = pGpParameter
				Exit Function
			End If
		Next n
        GetParameterByName = Nothing
		
	End Function
	
	
	Public Function CreateParameter(ByRef sName As String, ByRef sDisplayName As String, ByRef direction As ESRI.ArcGIS.Geoprocessing.esriGPParameterDirection, ByRef enum_type As ESRI.ArcGIS.Geoprocessing.esriGPParameterType, ByRef pGPValueType As ESRI.ArcGIS.Geodatabase.IGPDataType, Optional ByRef bEnabled As Boolean = True) As ESRI.ArcGIS.Geoprocessing.IGPParameter
		
		Dim pGPParameterEdit As ESRI.ArcGIS.Geoprocessing.IGPParameterEdit
        Dim pGPValue As ESRI.ArcGIS.Geodatabase.IGPValue
		
		pGPParameterEdit = New ESRI.ArcGIS.Geoprocessing.GPParameter
		pGPParameterEdit.DataType = pGPValueType
		pGPValue = pGPValueType.CreateValue("")
		
		pGPParameterEdit.Value = pGPValue
		pGPParameterEdit.ParameterType = enum_type
		pGPParameterEdit.direction = direction
		pGPParameterEdit.DisplayName = sDisplayName
		pGPParameterEdit.name = sName
		pGPParameterEdit.Enabled = bEnabled
		
		CreateParameter = pGPParameterEdit
		
	End Function
End Module