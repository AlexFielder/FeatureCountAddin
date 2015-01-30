Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports System.Drawing

Namespace RunFeatureCountOnAssembly
    <ProgIdAttribute("RunFeatureCountOnAssembly.StandardAddInServer"), _
    GuidAttribute("4c631cfe-22db-45dd-b9b5-00c441ca29e6")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Private m_inventorApplication As Inventor.Application

#Region "ApplicationAddInServer Members"
        ''' <summary>
        ''' Bulk of the activation code copied from here: http://adndevblog.typepad.com/manufacturing/2013/07/creating-a-ribbon-item-for-inventor.html
        ''' </summary>
        ''' <param name="addInSiteObject"></param>
        ''' <param name="firstTime"></param>
        ''' <remarks></remarks>
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application

            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.
            ' get the command manager control definition
            Dim conDefs As Inventor.ControlDefinitions = _
                m_inventorApplication. _
                CommandManager.ControlDefinitions

            ' our custom command ID
            Dim idCommand1 As String = "ID_COMAND_1"

            Try
                ' try get the existing command definition
                _defComando1 = conDefs.Item(idCommand1)
            Catch ex As Exception
                ' or create it
                _defComando1 = conDefs.AddButtonDefinition( _
                    "Feature Count", idCommand1, _
                    CommandTypesEnum.kEditMaskCmdType, _
                    Guid.NewGuid().ToString(), _
                    "Counts Features, Parameters, Occurrences and Constraints for use in Vault metrics!", _
                    "1, ah ah ah, 2, ah ah ah...", _
                    GetICOResource("RunFeatureCountOnAssembly.Pino-Sesame-Street-The-Count16x16.ico"), _
                    GetICOResource("RunFeatureCountOnAssembly.Pino-Sesame-Street-The-Count.ico"))
            End Try

            If (firstTime) Then
                If (m_inventorApplication.UserInterfaceManager.
                    InterfaceStyle =
                    InterfaceStyleEnum.kRibbonInterface) Then

                    ' 1. access the Part ribbon
                    Dim ribbonPart As Inventor.Ribbon = m_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")
                    Dim tabPartFeatureCount As Inventor.RibbonTab = ribbonPart.RibbonTabs.Add("FeatureCount", " PART_TAB_FEATURE_COUNT", Guid.NewGuid().ToString())
                    Dim MyPartCommandsPanel As Inventor.RibbonPanel = tabPartFeatureCount.RibbonPanels.Add("Part Feature Count", "PART_PNL_RUN_FEATURE_COUNT", Guid.NewGuid().ToString())
                    MyPartCommandsPanel.CommandControls.AddButton(_defComando1, True)

                    '2. access thr Assembly ribbon
                    Dim ribbonAssembly As Inventor.Ribbon = m_inventorApplication.UserInterfaceManager.Ribbons.Item("Assembly")
                    Dim tabAssemblyFeatureCount As Inventor.RibbonTab = ribbonAssembly.RibbonTabs.Add("FeatureCount", "ASSY_TAB_FEATURE_COUNT", Guid.NewGuid().ToString())
                    Dim MyAssemblyCommandsPanel As Inventor.RibbonPanel = tabAssemblyFeatureCount.RibbonPanels.Add("Assembly Feature Count", "ASSY_PNL_RUN_FEATURE_COUNT", Guid.NewGuid().ToString())
                    MyAssemblyCommandsPanel.CommandControls.AddButton(_defComando1, True)

                End If
            End If

            ' register the method that will be executed
            AddHandler _defComando1.OnExecute, AddressOf Command1Method

        End Sub


        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub

#End Region
        ' must be declared at the class level
        Private _defComando1 As Inventor.ButtonDefinition
        Private Sub Command1Method()
            If TypeOf m_inventorApplication.ActiveDocument Is PartDocument Then
                FeatureCount(m_inventorApplication.ActiveDocument)
            Else
                RunFeatureCount(m_inventorApplication.ActiveDocument)
            End If
            MessageBox.Show("Done!")
        End Sub

        ''' <summary>
        ''' Returns the relevant resource from the compiled .dll file
        ''' Means you don't need to Copy local .ico (or any other resource files)!
        ''' </summary>
        ''' <param name="icoResourceName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetICOResource( _
                  ByVal icoResourceName As String) As Object
            Dim assemblyNet As System.Reflection.Assembly = _
              System.Reflection.Assembly.GetExecutingAssembly()
            Dim stream As System.IO.Stream = _
              assemblyNet.GetManifestResourceStream(icoResourceName)
            Dim ico As System.Drawing.Icon = _
              New System.Drawing.Icon(stream)
            Return PictureDispConverter.ToIPictureDisp(ico)
        End Function

        ''' <summary>
        ''' Runs FeatureCount on Assembly or Part files.
        ''' </summary>
        ''' <param name="oDoc"></param>
        ''' <remarks></remarks>
        Public Sub RunFeatureCount(ByVal oDoc As Inventor.Document)
            Dim oAssy As Inventor.AssemblyDocument
            If oDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                oAssy = CType(oDoc, AssemblyDocument)
                FeatureCount(oAssy)
                For Each oComp In oAssy.ReferencedDocuments
                    If oComp.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
                        FeatureCount(oComp)
                        RunFeatureCount(oComp)
                    Else
                        FeatureCount(oComp)
                    End If
                Next
            ElseIf oDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
                FeatureCount(oDoc)
            End If
        End Sub

        ''' <summary>
        ''' Returns a bunch of Counts to the custom iProperties of whichever file(s) you run it on
        ''' </summary>
        ''' <param name="oDoc"></param>
        ''' <remarks></remarks>
        Sub FeatureCount(ByVal oDoc As Inventor.Document)
            Dim SaveRequired As Boolean = False
            If FileIsReadOnly(oDoc.FullFileName) Then
                Exit Sub
            End If
            Dim invCustomiPropSet As PropertySet = oDoc.PropertySets.Item("User Defined Properties")
            If TypeOf oDoc Is PartDocument Then
                Dim partDoc As PartDocument = oDoc
                Dim oFeats As Inventor.PartFeatures = partDoc.ComponentDefinition.Features
                Dim oParams As Inventor.Parameters = partDoc.ComponentDefinition.Parameters
                If Not oDoc.File.FullFileName.Contains("Content") Then 'skip CC and FACILITY files
                    If Not oDoc.File.FullFileName.Contains("FACILITY") Then
                        Try
                            Dim FeatureCountProp As Inventor.Property = (From a As Inventor.Property In invCustomiPropSet
                                                                        Where a.Name = "FEATURECOUNT"
                                                                        Select a).FirstOrDefault()
                            If Not FeatureCountProp Is Nothing Then
                                If Not FeatureCountProp.Value = oFeats.Count Then
                                    FeatureCountProp.Value = oFeats.Count
                                    SaveRequired = True
                                End If
                            Else
                                invCustomiPropSet.Add(oFeats.Count, "FEATURECOUNT")
                                SaveRequired = True
                            End If
                            Dim ParamCountProp As Inventor.Property = (From a As Inventor.Property In invCustomiPropSet
                                                                       Where a.Name = "PARAMETERCOUNT"
                                                                       Select a).FirstOrDefault()
                            If Not ParamCountProp Is Nothing Then
                                If Not ParamCountProp.Value = oParams.Count Then
                                    ParamCountProp.Value = oParams.Count
                                    SaveRequired = True
                                End If
                            Else
                                invCustomiPropSet.Add(oParams.Count, "PARAMETERCOUNT")
                                SaveRequired = True
                            End If
                            If SaveRequired Then
                                oDoc.Save() 'try to save the file.
                            End If
                            Exit Sub
                        Catch ex As Exception
                            MessageBox.Show("The exception was: " & ex.Message)
                        End Try
                    End If
                End If
            ElseIf TypeOf oDoc Is AssemblyDocument Then
                If Not oDoc.File.FullFileName.Contains("Content") Then 'skip CC and FACILITY files
                    If Not oDoc.File.FullFileName.Contains("FACILITY") Then
                        Dim AssyDoc As AssemblyDocument = oDoc
                        Dim oFeats As Inventor.Features = AssyDoc.ComponentDefinition.Features
                        Dim oParams As Inventor.Parameters = AssyDoc.ComponentDefinition.Parameters
                        Dim oConstraints As Inventor.AssemblyConstraints = AssyDoc.ComponentDefinition.Constraints
                        Dim Occs As Inventor.ComponentOccurrences = AssyDoc.ComponentDefinition.Occurrences
                        Dim FeatureCountNeedsUpdating As Boolean = True
                        Dim ParamCountNeedsUpdating As Boolean = True
                        Dim OccsCountNeedsUpdating As Boolean = True
                        Dim ConstraintCountNeedsUpdating As Boolean = True
                        Try
                            Dim FeatureCountProp As Inventor.Property = (From a As Inventor.Property In invCustomiPropSet
                                                                         Where a.Name = "FEATURECOUNT"
                                                                         Select a).FirstOrDefault()
                            If Not FeatureCountProp Is Nothing Then
                                If Not FeatureCountProp.Value = oFeats.Count Then
                                    FeatureCountProp.Value = oFeats.Count
                                    SaveRequired = True
                                End If
                            Else
                                invCustomiPropSet.Add(oFeats.Count, "FEATURECOUNT")
                                SaveRequired = True
                            End If
                            Dim ParamCountProp As Inventor.Property = (From a As Inventor.Property In invCustomiPropSet
                                                                       Where a.Name = "PARAMETERCOUNT"
                                                                       Select a).FirstOrDefault()
                            If Not ParamCountProp Is Nothing Then
                                If Not ParamCountProp.Value = oParams.Count Then
                                    ParamCountProp.Value = oParams.Count
                                    SaveRequired = True
                                End If
                            Else
                                invCustomiPropSet.Add(oParams.Count, "PARAMETERCOUNT")
                                SaveRequired = True
                            End If
                            Dim OccurrenceCountProp As Inventor.Property = (From a As Inventor.Property In invCustomiPropSet
                                                                            Where a.Name = "OCCURRENCECOUNT"
                                                                            Select a).FirstOrDefault()
                            If Not OccurrenceCountProp Is Nothing Then
                                If Not OccurrenceCountProp.Value = Occs.Count Then
                                    OccurrenceCountProp.Value = Occs.Count
                                    SaveRequired = True
                                End If
                            Else
                                invCustomiPropSet.Add(Occs.Count, "OCCURRENCECOUNT")
                                SaveRequired = True
                            End If
                            Dim ConstraintCountProp As Inventor.Property = (From a As Inventor.Property In invCustomiPropSet
                                                                            Where a.Name = "CONSTRAINTCOUNT"
                                                                            Select a).FirstOrDefault()
                            If Not ConstraintCountProp Is Nothing Then
                                If Not ConstraintCountProp.Value = oConstraints.Count Then
                                    ConstraintCountProp.Value = oConstraints.Count
                                    SaveRequired = True
                                End If
                            Else
                                invCustomiPropSet.Add(oConstraints.Count, "CONSTRAINTCOUNT")
                                SaveRequired = True
                            End If
                            If SaveRequired Then
                                oDoc.Save() 'try to save the file.
                            End If
                            Exit Sub
                        Catch ex As Exception
                            MessageBox.Show("The exception was: " & ex.Message)
                        End Try
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' Returns True if file is read-only
        ''' </summary>
        ''' <param name="p1">File to check.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function FileIsReadOnly(p1 As String) As Boolean
            Dim File_Attr As Long
            File_Attr = System.IO.File.GetAttributes(p1)
            If File_Attr And 1 Then
                Return True
            Else
                Return False
            End If
        End Function

    End Class

#Region "PictureDispConverter"
    ''' <summary>
    ''' Class that converts our icons into something Inventor can use.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll",
          EntryPoint:="OleCreatePictureIndirect",
          ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
    <MarshalAs(UnmanagedType.AsAny)> picdesc As Object, _
    ByRef iid As Guid, _
    <MarshalAs(UnmanagedType.Bool)> fOwn As Boolean _
    ) As IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType( _
          IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub
            'Picture Types
            Public Const PICTYPE_UNINITIALIZED As Short = -1
            Public Const PICTYPE_NONE As Short = 0
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_METAFILE As Short = 2
            Public Const PICTYPE_ICON As Short = 3
            Public Const PICTYPE_ENHMETAFILE As Short = 4

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf( _
                  GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(icon__1 As System.Drawing.Icon)
                    Me.hicon = icon__1.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf( _
                  GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(bitmap__1 As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap__1.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp( _
                          icon As System.Drawing.Icon _
                          ) As IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, _
                                            iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp( _
                          bmp As System.Drawing.Bitmap _
                          ) As IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, _
                                            iPictureDispGuid, True)
        End Function
    End Class
#End Region
End Namespace

