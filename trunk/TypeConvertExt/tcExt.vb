Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices

<Guid("EF4E898F-0F81-462A-B622-AB4A53A433D3"), ComVisible(True)> _
Public Class tcExt
    Implements IExtension
    Implements IExtensionConfig
    'Allows extension to appear in the Extensions dialog

    Private m_pApp As IApplication
    Private m_extState As esriExtensionState
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxExtension.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxExtension.Unregister(regKey)
    End Sub
#End Region
#Region "IExtension Members"
    Private ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "TypeConvert"
        End Get
    End Property
    Private Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_pApp = initializationData
    End Sub

    Private Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        'UPGRADE_NOTE: Object m_pApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        m_pApp.Dispose()

    End Sub
#End Region
#Region "IExtensionConfig Members"
    Private ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
        Get
            Return "Converts one feature type feature class to another feature type"
        End Get
    End Property


    Private Property State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.State
        'Get
        '    Return m_extState
        'End Get
        'Set(ByVal Value As ESRI.ArcGIS.esriSystem.esriExtensionState)
        '    m_extState = Value
        'End Set
        Get
            Return esriExtensionState.esriESEnabled
        End Get
        Set(ByVal Value As ESRI.ArcGIS.esriSystem.esriExtensionState)
            m_extState = esriExtensionState.esriESEnabled
        End Set
    End Property

#End Region
End Class
<Guid("BFB73754-E370-40F8-B8B9-7AB2A79AF5C7"), ComVisible(True)> _
    Public Class tcExtToolBar
    Implements IToolBarDef
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommandBars.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommandBars.Unregister(regKey)
    End Sub
#End Region

#Region "IToolBarDef Members"

    Private ReadOnly Property ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IToolBarDef.ItemCount
        Get
            ' Two items on the toolbar
            Return 1
        End Get
    End Property

    Private ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.IToolBarDef.Name
        Get
            Return "TypeConvertToolbar"
        End Get
    End Property

    Private ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.IToolBarDef.Caption
        Get
            Return "Type convert"
        End Get
    End Property

    Private Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IToolBarDef.GetItemInfo
        ' Add the sample command and the built-in Add Data command to the toolbar
        Select Case pos
            Case 0
                itemDef.ID = "TypeConvertExt.tcExtMenu"
                itemDef.Group = False
        End Select
    End Sub
#End Region
End Class
<Guid("C729A9CB-0304-42F5-BC94-8CFEB34E76E6"), ComVisible(True)> _
Public Class tcExtMenu
    Implements IMenuDef

    Private Const MENUCAPTION As String = "Convert"
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxToolMenuCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxToolMenuCommands.Unregister(regKey)
    End Sub
#End Region
#Region "IMenuDef Members"

    Public Sub New()
        MyBase.New()

    End Sub

    Protected Overrides Sub Finalize()

        MyBase.Finalize()
    End Sub

    Private ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Caption
        Get
            Return MENUCAPTION
        End Get
    End Property

    Private ReadOnly Property ItemCount() As Integer Implements ESRI.ArcGIS.SystemUI.IMenuDef.ItemCount
        Get
            Return 14
        End Get
    End Property

    Private ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.IMenuDef.Name
        Get
            Return "TypeConvert_Convert"
        End Get
    End Property

    Private Sub GetItemInfo(ByVal pos As Integer, ByVal itemDef As ESRI.ArcGIS.SystemUI.IItemDef) Implements ESRI.ArcGIS.SystemUI.IMenuDef.GetItemInfo
        Select Case pos
            Case 0
                itemDef.ID = "TypeConvertExt.tcToPolyline"
                itemDef.Group = False
            Case 1
                itemDef.ID = "TypeConvertExt.tcToPolygon"
                itemDef.Group = False
            Case 2
                itemDef.ID = "TypeConvertExt.tcToPoint"
                itemDef.Group = False
            Case 3
                itemDef.ID = "TypeConvertExt.tcToSegments"
                itemDef.Group = True
            Case 4
                itemDef.ID = "TypeConvertExt.tcToConvexHull"
                itemDef.Group = False
            Case 5
                itemDef.ID = "TypeConvertExt.tcToEnvelope"
                itemDef.Group = False
            Case 6
                itemDef.ID = "TypeConvertExt.tcToCentroid"
                itemDef.Group = False
            Case 7
                itemDef.ID = "TypeConvertExt.tcFromGraphics"
                itemDef.Group = True
            Case 8
                itemDef.ID = "TypeConvertExt.tcRemoveDuplicates"
                itemDef.Group = True
            Case 9
                itemDef.ID = "TypeConvertExt.tcStratification"
                itemDef.Group = True
            Case 10
                itemDef.ID = "TypeConvertExt.tcDivideSegments"
                itemDef.Group = True
            Case 11
                itemDef.ID = "TypeConvertExt.tcBLN"
                itemDef.Group = True
            Case 12
                itemDef.ID = "TypeConvertExt.tcKML"
                itemDef.Group = True
            Case 13
                itemDef.ID = "TypeConvertExt.tcAbout"
                itemDef.Group = True
        End Select
    End Sub
#End Region
End Class
<Guid("4036E552-AA98-4213-BAD4-D7351AA27896"), ComVisible(True)> _
   Public Class tcToPolyline
    Implements ESRI.ArcGIS.SystemUI.ICommand

    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr

    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.polyline.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map
            'If DateIsExpired() And IsDemo() Then
            '    Return False
            '    Exit Property
            'End If

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap

            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToPolyline"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To polyline"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature class into a polyline feature class"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature class into a polyline feature class"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.TO_POLYLINE))

    End Sub
End Class
<Guid("4C6DA4A1-A3A0-437F-80F1-830A459A8FCF"), ComVisible(True)> _
   Public Class tcToPolygon
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.polygon.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map
            'If DateIsExpired() And IsDemo() Then
            '    Return False
            '    Exit Property
            'End If

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToPolygon"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To polygon"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature class into polygon feature class"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature class into polygon feature class"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.TO_POLYGON))

    End Sub
End Class
<Guid("797738DB-E0CB-4A43-B003-F86A183175AB"), ComVisible(True)> _
    Public Class tcToPoint
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.point.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map
            

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then ' And pFeatureLayer.FeatureClass.ShapeType <> esriGeometryPoint Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToPoint"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To point"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature class into a point feature class"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature class into a point feature class"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.TO_POINT))

    End Sub
End Class
<Guid("C0BB8892-32BB-4FB7-BC1D-A08C445E30A6"), ComVisible(True)> _
   Public Class tcToCentroid
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.Centroid.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map
            

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer And pFeatureLayer.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPoint Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToCentroid"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To centroid"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature selection into feature centroids"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature selection into feature centroids"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        pFeatureLayer = m_pDoc.SelectedLayer
        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.F_CENTROID))
    End Sub


End Class
<Guid("BBE07805-01F7-4DB1-8CC3-0C15BCC8698F"), ComVisible(True)> _
   Public Class tcToConvexHull
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.ConvexHull.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map
            

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToConvexHull"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To convex hull"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature selection into polygon convex hull"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature selection into polygon convex hull"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        pFeatureLayer = m_pDoc.SelectedLayer
        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.CONVEX_HULL))
    End Sub


End Class
<Guid("41557E05-5B09-4A7D-8B6A-0DB1457D707F"), ComVisible(True)> _
   Public Class tcToEnvelope
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.Envelope.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map
            

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToEnvelope"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To envelope"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature selection into envelope"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature selection into envelope"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        pFeatureLayer = m_pDoc.SelectedLayer
        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.ENVELOPE_RECTANGLE))
    End Sub


End Class
<Guid("8D2BD70F-0F8B-4D12-9782-F2849E8150F0"), ComVisible(True)> _
    Public Class tcToSegments
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.segment.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer And pFeatureLayer.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPoint Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToSegments"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To segments"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert feature class into segemnts"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert feature class into segemnts"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.SEGMENTS))

    End Sub
End Class
<Guid("ED3E0665-434C-4d72-A1BF-3E36D3E3E5EA"), ComVisible(True)> _
    Public Class tcDivideSegments
    Implements ESRI.ArcGIS.SystemUI.ICommand

    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr
    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.Div_segment.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap

            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer And pFeatureLayer.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPoint Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_DivideSegments"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "Divide segments"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Divide segments"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Divide segments"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Convert.DivideSegments(m_pDoc.SelectedLayer, m_pMap)

    End Sub
End Class
<Guid("07310669-3D1C-41B3-8CEC-84D0A15A0ACD"), ComVisible(True)> _
    Public Class tcFromGraphics
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.FromGraph.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get

            'Command is enabled only if the extension is turned on and there is data in the map
            Dim pGraphContainer As ESRI.ArcGIS.Carto.IGraphicsContainer
            Dim pElement As ESRI.ArcGIS.Carto.IElement
            On Error GoTo ErrorHandler

            m_pMap = m_pDoc.FocusMap
            pGraphContainer = m_pMap
            pGraphContainer.Reset()
            pElement = pGraphContainer.Next
            If (m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled) And (Not pElement Is Nothing) Then

                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_FromGraphics"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "From graphics"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Convert graphics into features"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Convert graphics into features"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Convert.ConvertGraphics(m_pMap)
    End Sub

End Class
<Guid("026BEA67-A757-47BE-9618-B42D0DC6E84E"), ComVisible(True)> _
    Public Class tcStratification
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As IMap
    Private pFeatureLayer As IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.Stratify.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()

        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get

            'Command is enabled only if the extension is turned on and there is data in the map
            Dim EnableFlag As Boolean
            Dim pGeoLayer As IGeoFeatureLayer
            Dim pUVRend As IUniqueValueRenderer
            EnableFlag = False
            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If TypeOf pFeatureLayer Is FeatureLayer Then
                pGeoLayer = pFeatureLayer
                If (TypeOf pGeoLayer.Renderer Is IUniqueValueRenderer) Then
                    pUVRend = pGeoLayer.Renderer
                    If pUVRend.FieldCount = 1 Then
                        EnableFlag = True
                    End If
                ElseIf TypeOf pGeoLayer.Renderer Is IClassBreaksRenderer Then
                    EnableFlag = True
                End If
            End If


            m_pMap = m_pDoc.FocusMap
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And EnableFlag Then

                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_Stratification"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "Stratification"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Stratify current layer to set of layers by legend classes"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Stratify current layer to set of layers by legend classes"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.STRATIFICATION))
    End Sub
End Class
<Guid("273AC0D6-02DC-41FF-AC4E-B737CB2E793F"), ComVisible(True)> _
    Public Class tcRemoveDuplicates
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr



    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.RemoveDup.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()
        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then
                'If TypeOf pFeatureLayer Is FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_RemoveDuplicates"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "Remove duplicates"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Remove duplicate features from the layer"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Remove duplicate features from the layer"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return (Nothing)
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        pFeatureLayer = m_pDoc.SelectedLayer
        'Convert.RemoveDuplicates pFeatureLayer.FeatureClass
        Convert.StartConvert(m_pMap, (m_pDoc.SelectedLayer), (Convert.ConvertType.REMOVE_DUPLICATES))
    End Sub
End Class
<Guid("E0A17824-2589-4D51-B6BE-B115BBE30AFE"), ComVisible(True)> _
    Public Class tcBLN
    Implements ESRI.ArcGIS.SystemUI.ICommand


    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr

    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.bln.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()
        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer And pFeatureLayer.FeatureClass.ShapeType <> esriGeometryType.esriGeometryPoint Then

                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToBln"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To *.bln"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "Export vertices of the feature class into bln-file"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Export vertices of the feature class into bln-file"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            ' TODO: Add your implementation here
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get

            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

        Dim ExportToBlnFile As New ExportToBlnFile

        ExportToBlnFile.InsideFlag.Enabled = (pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon)

        If ExportToBlnFile.ShowDialog() = DialogResult.OK Then
            Convert.ExportToBLN((m_pDoc.SelectedLayer), (ExportToBlnFile.BlnFileName).Text, ExportToBlnFile.InsideFlag.Checked)
        End If
        ExportToBlnFile.Close()
    End Sub

End Class
<Guid("064DBA2D-B1AF-4c37-82BE-55AF86BC0AD8"), ComVisible(True)> _
    Public Class tcKML
    Implements ESRI.ArcGIS.SystemUI.ICommand
    Private m_pApp As IApplication
    Private m_pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_pMap As ESRI.ArcGIS.Carto.IMap
    Private pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Private m_pExt As ESRI.ArcGIS.esriSystem.IExtensionConfig
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr

    Public Sub New()
        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.google_earth.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
    End Sub
#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        'm_pMap.Dispose()
        'm_pDoc.Dispose()
        'm_pExt.Dispose()
        'm_pApp.Dispose()
        MyBase.Finalize()
    End Sub
    Private Sub ICommand_OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        m_pApp = hook
        m_pDoc = m_pApp.Document

        'Get the extension
        Dim pUID As New ESRI.ArcGIS.esriSystem.UID
        pUID.Value = "TypeConvertExt.tcExt"
        m_pExt = m_pApp.FindExtensionByCLSID(pUID)

    End Sub
    Private ReadOnly Property ICommand_Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get
            'Command is enabled only if the extension is turned on and there is data in the map

            On Error GoTo ErrorHandler
            pFeatureLayer = m_pDoc.SelectedLayer
            m_pMap = m_pDoc.FocusMap
            'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
            If m_pExt.State = ESRI.ArcGIS.esriSystem.esriExtensionState.esriESEnabled And TypeOf pFeatureLayer Is ESRI.ArcGIS.Carto.FeatureLayer Then
                Return True
            Else
                Return False
            End If
            Exit Property

ErrorHandler:
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property ICommand_Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_ToKML"
        End Get
    End Property

    Private ReadOnly Property ICommand_Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return "To Google Earth"
        End Get
    End Property

    Private ReadOnly Property ICommand_Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return "EConverts ArcGIS layers into Google Earth KML files"
        End Get
    End Property

    Private ReadOnly Property ICommand_Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return "Converts ArcGIS layers into Google Earth KML files"
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Private ReadOnly Property ICommand_HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            ' TODO: Add your implementation here
        End Get
    End Property

    Private ReadOnly Property ICommand_Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get

            ' Get command bitmap from the image list on a form
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Private ReadOnly Property ICommand_Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Private Sub ICommand_OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        If CType(m_pDoc.SelectedLayer, IFeatureLayer).FeatureClass.FeatureCount(Nothing) > 600 Then
            MsgBox("Exported layer contains more than 600 objects." & vbNewLine & "Please, visit http://mi-perm.ru/gis/programs/kmler to get information about KMLer extension", MsgBoxStyle.Exclamation, "KMLer")
            Exit Sub
        End If
        Dim ConvertToGEForm As New ConvertToGE
        If Not m_pDoc.SelectedLayer.Valid Then
            MsgBox("Layer " & m_pDoc.SelectedLayer.Name & " is not valid")
            Exit Sub
        End If
        Dim pGeoLayer As IGeoFeatureLayer = m_pDoc.SelectedLayer
        Dim pUVRend As IUniqueValueRenderer

        With ConvertToGEForm
            Dim i As Integer
            Dim pTableFields As ITableFields = m_pDoc.SelectedLayer
            Dim pFeatureClass As IFeatureClass = DirectCast(m_pDoc.SelectedLayer, IFeatureLayer).FeatureClass
            'Dim pFields As IFields = pFeatureClass.Fields
            Dim pFields As ITableFields = m_pDoc.SelectedLayer
            If pFields.Field(pFeatureClass.FindField(pFeatureClass.ShapeFieldName)).GeometryDef.HasZ Then
                .cbAttributeValue.Items.Add("Z values")
            End If
            For i = 0 To pFields.FieldCount - 1
                If pTableFields.FieldInfo(i).Visible AndAlso Not pFields.Field(i) Is pFeatureClass.AreaField And Not pFields.Field(i) Is pFeatureClass.LengthField And pFields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry And (pFields.Field(i).Type = esriFieldType.esriFieldTypeDouble Or pFields.Field(i).Type = esriFieldType.esriFieldTypeInteger Or pFields.Field(i).Type = esriFieldType.esriFieldTypeSingle Or pFields.Field(i).Type = esriFieldType.esriFieldTypeSmallInteger) Then
                    .cbAttributeValue.Items.Add(pFields.Field(i).Name)
                End If

                If pTableFields.FieldInfo(i).Visible AndAlso pFields.Field(i).Type <> esriFieldType.esriFieldTypeGeometry Then
                    .LabelFieldList.Items.Add(pFields.Field(i).Name)
                End If
            Next
            If .cbAttributeValue.Items.Count = 0 Then
                .cbAttributeValue.Enabled = False
                .rbAttribute.Enabled = False
                .rbValue.Checked = True
            Else
                .cbAttributeValue.Text = .cbAttributeValue.GetItemText(.cbAttributeValue.Items.Item(0))
            End If
            If .LabelFieldList.Items.Count = 0 Then
                .LabelFieldList.Enabled = False
            Else
                .LabelFieldList.Enabled = True
                .LabelFieldList.Text = .LabelFieldList.GetItemText(.LabelFieldList.Items.Item(0))
            End If

            Dim blnExtrude, AltitudeMode As Integer
            Dim AttributeFieldName As String
            Dim SetValue As Double
            If pFeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
                .PlacemarkFlag.Enabled = False
            End If
            .SaveFileDialog.FileName = DirectCast(pFeatureClass, IDataset).Name
            If TypeOf pGeoLayer.Renderer Is IUniqueValueRenderer Then
                pUVRend = pGeoLayer.Renderer
                If pUVRend.FieldCount = 1 Then
                    .isUniqueValueRenderer = True
                Else
                    .isUniqueValueRenderer = False
                End If
            Else
                .isUniqueValueRenderer = False
            End If

            If ConvertToGEForm.ShowDialog = DialogResult.OK Then
                If .ExtrudeFlag.Checked Then
                    blnExtrude = 1
                    AltitudeMode = .cbAltitudeMode.SelectedIndex
                    If .rbAttribute.Checked Then
                        AttributeFieldName = .cbAttributeValue.Text
                        SetValue = 0
                    Else
                        AttributeFieldName = ""
                        SetValue = IIf(.SetValue.Text <> "", .SetValue.Text, 0)
                    End If
                Else
                    blnExtrude = 0
                    AttributeFieldName = ""
                    SetValue = 0
                End If

                Convert.ConvertToKML(m_pDoc.SelectedLayer, .kmlFileName.Text, .LabelFieldList.Text, .PlacemarkFlag.Checked, blnExtrude, AltitudeMode, AttributeFieldName, SetValue, (.DistributeFlag.Enabled And .DistributeFlag.Checked))
                If .cbOpenInGE.Checked Then
                    Try
                        Process.Start(.kmlFileName.Text)
                    Catch
                        MsgBox("Google Earth isn't installed. Look more at http://earth.google.com/", MsgBoxStyle.Exclamation, "Convert to Google Earth")
                    End Try
                End If
            End If
        End With
    End Sub

End Class
<Guid("12EB530A-5A45-4002-B024-9231190D6C32")> _
        Public Class tcAbout
    Implements ICommand
    Private m_bitmap As Bitmap
    Private m_hBitmap As IntPtr
    Private m_category As String
    Private m_caption As String
    Private m_message As String

#Region "Component Category Registration"
    <ComRegisterFunction()> _
    Shared Sub Reg(ByVal regKey As String)
        MxCommands.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Shared Sub Unreg(ByVal regKey As String)
        MxCommands.Unregister(regKey)
    End Sub
#End Region
    Public Sub New()

        m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("TypeConvertExt.information.bmp"))
        m_bitmap.MakeTransparent(m_bitmap.GetPixel(1, 1))
        m_hBitmap = m_bitmap.GetHbitmap()
        m_caption = "About"
        m_message = m_caption
    End Sub

    Public ReadOnly Property Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return m_caption
        End Get
    End Property

    Public ReadOnly Property Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return "TypeConvert"
        End Get
    End Property

    Public ReadOnly Property Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property Enabled() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get

            Return True
        End Get
    End Property

    Public ReadOnly Property HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get

        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return m_message
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return "TypeConvert_About"
        End Get
    End Property

    Public Sub OnClick() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick
        Dim pAboutForm As New AboutForm
        pAboutForm.ShowDialog()
    End Sub

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        
    End Sub

    Public ReadOnly Property Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return m_message
        End Get
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class