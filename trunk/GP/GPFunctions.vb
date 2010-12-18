Option Strict Off
Option Explicit On

<Guid("791E3856-4688-4c9d-871F-E4689F5971A6")> _
Public Class ConvertToPolyline
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToPolyline"
    Private Const sFunctionDispName As String = "To Polyline"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array

            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToPolyline(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "polyline", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return (Nothing)
        End Get
    End Property
End Class
<Guid("9E515FE7-2349-4b9d-B979-472AF09FA989")> _
Public Class ConvertToPolygon
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToPolygon"
    Private Const sFunctionDispName As String = "To Polygon"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array

            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToPolygon(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return Nothing
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "polygon", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return (Nothing)
        End Get
    End Property
End Class
<Guid("A6A8323D-D3ED-4185-B4A0-8DE888E22CFE")> _
Public Class ConvertToPoint
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToPoint"
    Private Const sFunctionDispName As String = "To Point"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array

            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToPoint(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "point", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return (Nothing)
        End Get
    End Property
End Class
<Guid("3E90D1BA-3EAA-42f5-A484-63C26C52468E")> _
Public Class ConvertToSegments
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToSegments"
    Private Const sFunctionDispName As String = "To Segments"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToSegments(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "seg", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

        Dim pInFeatureClass As IFeatureClass
        If Not pInputVal.IsEmpty Then
            pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pInputVal.GetAsText)
            If pInFeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                pGPMessages.AddError(1, "Could not convert a point feature class")
            End If
        End If

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return (Nothing)
        End Get
    End Property
End Class
<Guid("0A9957EA-94F6-421a-B02C-34EA277F0D66")> _
Public Class ConvertToCentroids
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToCentroids"
    Private Const sFunctionDispName As String = "To Centroids"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToCentroid(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "Centr", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return (Nothing)
        End Get
    End Property
End Class
<Guid("F651D2BF-27A1-4825-A88C-7AD0C814D977")> _
Public Class ConvertToEnvelope
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToEnvelope"
    Private Const sFunctionDispName As String = "To Envelope"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToEnvelope(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "env", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return (Nothing)
        End Get
    End Property
End Class
<Guid("E4DE6AAB-B526-4353-87F9-F5D483C00223")> _
Public Class ConvertToConvexHull
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToConvexHull"
    Private Const sFunctionDispName As String = "To ConvexHull"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.ToConvexHull(pInFeatureClass, Nothing)

        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        'IGPFunction_IsLicensed = Not (DateIsExpired() And IsDemo())
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "convhull", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return Nothing
        End Get
    End Property
End Class
<Guid("D08F2BD2-D9F3-41fe-BAAD-C1110678D1D4")> _
Public Class RemoveDuplicates
    Implements IGPFunction

    Private Const sFunctionName As String = "RemoveDuplicates"
    Private Const sFunctionDispName As String = "Remove Duplicates"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.RemoveDuplicates(pInFeatureClass, Nothing).FeatureClass(0)


        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "rd", pOutParam, pInputVal, "shp", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If

            pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
            pInputParam = GPUtils.GetParameterByName(paramvalues, "Input_FeatureClass")
            If pInputParam.Value.GetAsText <> "" And pOutParam.Value.GetAsText = pInputParam.Value.GetAsText Then
                pGPMessages.Clear()
                pGPMessages.AddError(2, "Unable to convert the object into itself")
            End If

            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return Nothing
        End Get
    End Property
End Class
<Guid("374C52C6-842D-4673-BE0B-A47D7349308E")> _
Public Class Stratification
    Implements IGPFunction

    Private Const sFunctionName As String = "Stratification"
    Private Const sFunctionDispName As String = "Stratification"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType


            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DELayerType)
            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)
            pArray.Add(GPUtils.CreateParameter("Output_FeatureClass", "Output data", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As IArray, ByVal TrackCancel As ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_FeatureClass")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim pOutFeatureClass As IFeatureClass
        pInFeatureLayer = pGPUtilities.OpenFeatureLayerFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        pFeatureConverter.DatasetName = pGPUtilities.CreateFeatureClassName(OutputFeatureClass.GetAsText)

        'Delete if exists
        If Not pGPUtilities.OpenDatasetFromLocation(OutputFeatureClass.GetAsText) Is Nothing Then
            pOutFeatureClass = pGPUtilities.OpenFeatureClassFromString(OutputFeatureClass.GetAsText)
            Dim pDataset As IDataset = pOutFeatureClass
            pDataset.Delete()
        End If
        pOutFeatureClass = pFeatureConverter.Stratify(pInFeatureLayer, Nothing).FeatureLayer(0).FeatureClass


        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pDefGPValue As IGPValue
        Dim pInputParam As IGPParameter
        Dim pOutParam As IGPParameter
        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)

            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If

            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                    On Error Resume Next
                    Err.Clear()
                    pGPUtilities.OpenFeatureLayerFromString(pGPValue.GetAsText)
                    If Err.Number <> 0 Then
                        pGPMessages.AddError(1, "Layer is not valid")
                    End If
                End If
            End If

            '-----------------------
            ' Check if Value Already Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput Then
                    If Not pGPUtilities.OpenDatasetFromLocation(pGPParameterIn.Value.GetAsText) Is Nothing Then
                        pGPMessages.Add(pGPMessage)
                        pGPMessages.AddWarning("The output object already exists. It will be replaced")
                    End If
                End If
            End If


            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

            '-----------------------
            ' Custom validation for output
            '-----------------------
            If pGPMessage.ErrorCode = 0 And updateValues = True Then
                pOutParam = GPUtils.GetParameterByName(paramvalues, "Output_FeatureClass")
                pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

                On Error Resume Next

                pDefGPValue = pGPUtilities.GenerateDefaultOutputValue(envMgr, "", pOutParam, pInputVal, "", 0)
                pGPUtilities.PackGPValue(pDefGPValue, pOutParam)

            End If
        Next i


        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As UID Implements IGPFunction.DialogCLSID
        Get
            Return Nothing
        End Get
    End Property
End Class
<Guid("7B91AF95-ED68-4535-84BD-D596E697F378")> _
Public Class ConvertToBln
    Implements IGPFunction

    Private Const sFunctionName As String = "ConvertToBln"
    Private Const sFunctionDispName As String = "To Bln"
    Private Const sFunctionCategory As String = "Convert"
    Private Const sFunctionDescription As String = ""
    Private m_sInput_Workspace As String

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private ReadOnly Property IGPFunction_DisplayName() As String Implements IGPFunction.DisplayName
        Get
            Return sFunctionDispName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_FullName() As ESRI.ArcGIS.esriSystem.IName Implements IGPFunction.FullName
        Get
            Dim pGPFunctionName As IGPName
            pGPFunctionName = New GPFunctionName

            With pGPFunctionName
                .Category = sFunctionCategory
                .Description = sFunctionDescription
                .DisplayName = sFunctionDispName
                .Name = sFunctionName
                .Factory = New GPTypeConvertFactory
            End With

            IGPFunction_FullName = pGPFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpContext() As Integer Implements IGPFunction.HelpContext
        Get
            '
        End Get
    End Property

    Private ReadOnly Property IGPFunction_HelpFile() As String Implements IGPFunction.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_MetadataFile() As String Implements IGPFunction.MetadataFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property IGPFunction_Name() As String Implements IGPFunction.Name
        Get
            Return sFunctionName
        End Get
    End Property

    Private ReadOnly Property IGPFunction_ParameterInfo() As ESRI.ArcGIS.esriSystem.IArray Implements IGPFunction.ParameterInfo
        Get

            Dim pArray As ESRI.ArcGIS.esriSystem.IArray
            pArray = New ESRI.ArcGIS.esriSystem.Array


            Dim pCompositeType As IGPCompositeDataType
            Dim pFileType As IDEFileType
            pFileType = New DEFileType

            pCompositeType = New GPCompositeDataType
            pCompositeType.AddDataType(New DEFeatureClassType)

            pArray.Add(GPUtils.CreateParameter("Input_FeatureClass", "Input data", esriGPParameterDirection.esriGPParameterDirectionInput, esriGPParameterType.esriGPParameterTypeRequired, pCompositeType))
            pArray.Add(GPUtils.CreateParameter("Output_BlnFile", "Output bln-file", esriGPParameterDirection.esriGPParameterDirectionOutput, esriGPParameterType.esriGPParameterTypeRequired, pFileType))


            IGPFunction_ParameterInfo = pArray

        End Get
    End Property

    Private Sub IGPFunction_Execute(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal TrackCancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction.Execute
        On Error GoTo errhan
        ' VALIDATE PARAMETERS
        message.AddMessage("  Validating...")
        System.Windows.Forms.Application.DoEvents()
        Dim pValidateMessages As IGPMessages
        pValidateMessages = IGPFunction_Validate(paramvalues, False, envMgr)
        If (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityAbort) Or (pValidateMessages.MaxSeverity = esriGPMessageSeverity.esriGPMessageSeverityError) Then
            message.AddMessages(pValidateMessages)
            Exit Sub
        End If

        'Get the source and target layer files
        message.AddMessage(("   Converting..."))

        Dim InputFeatureClass As IGPValue
        Dim OutputFeatureClass As IGPValue

        InputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")
        OutputFeatureClass = GPUtils.GetParameterValueByName(paramvalues, "Output_BlnFile")

        Dim pGPUtilities As IGPUtilities
        pGPUtilities = New GPUtilities

        Dim bSaveRefresh As Boolean
        bSaveRefresh = pGPUtilities.RefreshCatalogParent
        pGPUtilities.RefreshCatalogParent = True

        Dim pGxObject As ESRI.ArcGIS.Catalog.IGxObject
        pGxObject = pGPUtilities.GetGxObject(InputFeatureClass)
        Dim pInFeatureClass As IFeatureClass
        pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pGxObject.FullName)
        Dim pFeatureConverter As IFeatureTypeConverter = New FeatureTypeConverter
        Dim fileName As String = OutputFeatureClass.GetAsText
        If InStr(fileName, ".", CompareMethod.Text) = 0 Then
            fileName = fileName & ".bln"
        End If
        pFeatureConverter.ExportToBLN(pInFeatureClass, fileName, True)


        Exit Sub
errhan:
        message.AddError(0, Err.Description)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Function IGPFunction_GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction.GetRenderer
        Return (Nothing)
    End Function

    Private Function IGPFunction_IsLicensed() As Boolean Implements IGPFunction.IsLicensed
        Return True
    End Function

    Private Function IGPFunction_Validate(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction.Validate

        Dim pArray As ESRI.ArcGIS.esriSystem.IArray
        Dim pGPUtilities As IGPUtilities
        Dim pGPMessages As IGPMessages
        Dim pGPMessage As IGPMessage
        Dim pGpParameter As IGPParameter
        Dim pGPParameterIn As IGPParameter
        Dim pGPDataType As IGPDataType
        Dim pGPValue As IGPValue
        Dim pGPDomain As IGPDomain
        '
        Dim i As Integer
        '
        pGPMessages = New GPMessages
        pGPUtilities = New GPUtilities
        pArray = IGPFunction_ParameterInfo


        Dim pInputVal As IGPValue

        For i = 0 To pArray.Count - 1 Step 1
            pGpParameter = pArray.Element(i)
            pGPParameterIn = paramvalues.Element(i)
            '
            pGPDataType = pGpParameter.DataType
            pGPDomain = pGpParameter.Domain
            '
            pGPValue = pGPUtilities.UnpackGPValue(pGPParameterIn)
            pGPMessage = pGPDataType.ValidateValue(pGPValue, pGPDomain)
            '-----------------------
            ' Check for Empty Value
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired Then
                    If pGPValue.IsEmpty Then
                        pGPMessages.AddError(1, pGpParameter.DisplayName & " is Empty")
                    End If
                End If
            End If
            '-----------------------
            ' Check if Value Exists
            '-----------------------
            If pGPMessage.ErrorCode = 0 Then
                If pGpParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput Then
                    GPUtils.CheckDatasetExists(pGPParameterIn, pGPValue, pGPMessages)
                End If
            End If


            If pGPMessage.ErrorCode <> 0 Then
                pGPMessages.AddError(pGPMessage.ErrorCode, pGPMessage.Description)
            End If

        Next i

        pInputVal = GPUtils.GetParameterValueByName(paramvalues, "Input_FeatureClass")

        Dim pInFeatureClass As IFeatureClass
        If Not pInputVal.IsEmpty Then
            pInFeatureClass = pGPUtilities.OpenFeatureClassFromString(pInputVal.GetAsText)
            If pInFeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                pGPMessages.AddError(1, "Could not convert a point feature class")
            End If
        End If
        IGPFunction_Validate = pGPMessages

    End Function

    Public ReadOnly Property DialogCLSID() As ESRI.ArcGIS.esriSystem.UID Implements IGPFunction.DialogCLSID
        Get
            Return Nothing
        End Get
    End Property
End Class