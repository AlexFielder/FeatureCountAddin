Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.IO
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
                    "Command 1", idCommand1, _
                    CommandTypesEnum.kEditMaskCmdType, _
                    Guid.NewGuid().ToString(), _
                    "Command 1 description", _
                    "Command 1 Tooltip", _
                    GetICOResource("RunFeatureCountOnAssembly.Pino-Sesame-Street-The-Count.ico"), _
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
            ' ToDo: do your code here
            MessageBox.Show("Hello World!")
        End Sub
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
        'Public Sub Main()
        '    If TypeOf ThisDoc.Document Is PartDocument Then
        '        PartFeatureCount(ThisDoc)
        '    Else
        '        RunFeatureCount(ThisDoc.Document)
        '    End If
        '    MessageBox.Show("Done!")
        'End Sub

        'Public Sub RunFeatureCount(ByVal oDoc As Inventor.Document)
        '    Dim oAssy As Inventor.AssemblyDocument
        '    Dim oSubDoc As Inventor.Document
        '    If oDoc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
        '        oAssy = CType(oDoc, AssemblyDocument)
        '        AssemblyFeatureCount(oAssy)
        '        For Each oComp In oAssy.ReferencedDocuments
        '            'oSubDoc = CType(oComp.Definition.Document,Document)
        '            'MessageBox.Show(oSubDoc.File.FullFileName)
        '            If oComp.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
        '                'run FeatureCount and then call RunFeatureCount to recurse the assembly structure
        '                'iLogicVb.RunExternalRule("FEATURECOUNT")
        '                FeatureCount(oComp)
        '                RunFeatureCount(oComp)
        '            Else
        '                'run FeatureCount
        '                'iLogicVb.RunExternalRule("FEATURECOUNT")
        '                FeatureCount(oComp)
        '            End If
        '        Next
        '    ElseIf oDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
        '        FeatureCount(oDoc)
        '    End If
        'End Sub

        'Sub FeatureCount(ByVal oDoc As Inventor.Document)
        '    Dim SaveRequired As Boolean = False
        '    Dim DocName As String = System.IO.Path.GetFileNameWithoutExtension(oDoc.DisplayName) & ":1"
        '    'uncomment for debugging purposes!
        '    'MessageBox.Show(DocName)
        '    If TypeOf oDoc Is PartDocument Then
        '        If Not oDoc.File.FullFileName.Contains("Content") Then 'skip CC and FACILITY files
        '            If Not oDoc.File.FullFileName.Contains("FACILITY") Then
        '                Dim oFeats As PartFeatures = oDoc.ComponentDefinition.Features
        '                Dim oParams As Parameters = oDoc.ComponentDefinition.Parameters
        '                Try
        '                    If Not iProperties.Value(DocName, "Custom", "FEATURECOUNT") = oFeats.Count Then 'or update it
        '                        iProperties.Value(DocName, "Custom", "FEATURECOUNT") = oFeats.Count
        '                        SaveRequired = True
        '                    End If
        '                    'MessageBox.Show("Feature Count for this part is: " & oFeats.Count, "FEATURECOUNT")
        '                    If Not iProperties.Value(DocName, "Custom", "PARAMETERCOUNT") = oParams.Count Then 'or update it
        '                        iProperties.Value(DocName, "Custom", "PARAMETERCOUNT") = oParams.Count
        '                        SaveRequired = True
        '                    End If
        '                    'MessageBox.Show("Parameter Count for " & oDoc.File.fullfilename &" is: " & oParams.Count, "PARAMETERCOUNT")
        '                    If SaveRequired Then
        '                        oDoc.Save() 'try to save the file.
        '                    End If
        '                Catch
        '                    iProperties.Value(DocName, "Custom", "FEATURECOUNT") = oFeats.Count
        '                    iProperties.Value(DocName, "Custom", "PARAMETERCOUNT") = oParams.Count
        '                    oDoc.Save() 'try to save the file.
        '                    Exit Sub
        '                End Try
        '            End If
        '        End If
        '    ElseIf TypeOf oDoc Is AssemblyDocument Then
        '        If Not oDoc.File.FullFileName.Contains("Content") Then 'skip CC and FACILITY files
        '            If Not oDoc.File.FullFileName.Contains("FACILITY") Then
        '                Dim oFeats As Features = oDoc.ComponentDefinition.Features
        '                Dim Occs As ComponentOccurrences = oDoc.ComponentDefinition.Occurrences
        '                Dim oParams As Parameters = oDoc.ComponentDefinition.Parameters
        '                'Dim oConstraints as Constraints = oDoc.ComponentDefinition.Constraints
        '                Try
        '                    If Not iProperties.Value(DocName, "Custom", "FEATURECOUNT") = oFeats.Count Then
        '                        iProperties.Value(DocName, "Custom", "FEATURECOUNT") = oFeats.Count
        '                        SaveRequired = True
        '                    End If
        '                    'MessageBox.Show("Feature Count for this assembly is: " & oFeats.Count, "FEATURECOUNT")
        '                    If Not iProperties.Value(DocName, "Custom", "OCCURRENCECOUNT") = Occs.Count Then
        '                        iProperties.Value(DocName, "Custom", "OCCURRENCECOUNT") = Occs.Count
        '                        SaveRequired = True
        '                    End If
        '                    'MessageBox.Show("Occurrence Count for " & oDoc.File.fullfilename & " is: " & Occs.Count, "OCCURRENCECOUNT")
        '                    If Not iProperties.Value(DocName, "Custom", "PARAMETERCOUNT") = oParams.Count Then
        '                        iProperties.Value(DocName, "Custom", "PARAMETERCOUNT") = oParams.Count
        '                        SaveRequired = True
        '                    End If
        '                    'MessageBox.Show("Parameter Count for this part is: " & oParams.Count, "PARAMETERCOUNT")
        '                    If Not iProperties.Value(DocName, "Custom", "CONSTRAINTCOUNT") = oDoc.ComponentDefinition.Constraints.Count Then
        '                        iProperties.Value(DocName, "Custom", "CONSTRAINTCOUNT") = oDoc.ComponentDefinition.Constraints.Count
        '                        SaveRequired = True
        '                    End If
        '                    'MessageBox.Show("Constraint Count for Assembly " & DocName & " is: " & oDoc.ComponentDefinition.Constraints.Count, "CONSTRAINTCOUNT")
        '                    If SaveRequired Then
        '                        oDoc.Save() 'try to save the file.
        '                    End If
        '                Catch
        '                    'creates any missing iProperties.
        '                    iProperties.Value(DocName, "Custom", "FEATURECOUNT") = oFeats.Count
        '                    iProperties.Value(DocName, "Custom", "OCCURRENCECOUNT") = Occs.Count
        '                    iProperties.Value(DocName, "Custom", "PARAMETERCOUNT") = oParams.Count
        '                    iProperties.Value(DocName, "Custom", "CONSTRAINTCOUNT") = oDoc.ComponentDefinition.Constraints.Count
        '                    oDoc.Save() 'saves the assembly
        '                    Exit Sub
        '                End Try
        '            End If
        '        End If
        '    End If
        'End Sub
        'Sub PartFeatureCount(ByVal oDoc)
        '    Dim SaveRequired As Boolean = False
        '    If Not oDoc.Document.File.FullFileName.Contains("Content") Then 'skip CC and FACILITY files
        '        If Not oDoc.Document.File.fullfilename.contains("FACILITY") Then
        '            Dim oFeats As PartFeatures = oDoc.Document.ComponentDefinition.Features
        '            Dim oParams As Parameters = oDoc.Document.ComponentDefinition.Parameters
        '            Try
        '                'may need to save the file when we're done, hence the boolean check
        '                If Not iProperties.Value("Custom", "FEATURECOUNT") = oFeats.Count Then
        '                    iProperties.Value("Custom", "FEATURECOUNT") = oFeats.Count
        '                    SaveRequired = True
        '                End If
        '                'MessageBox.Show("Feature Count for this part is: " & oFeats.Count, "FEATURECOUNT")

        '                If Not iProperties.Value("Custom", "PARAMETERCOUNT") = oParams.Count Then
        '                    iProperties.Value("Custom", "PARAMETERCOUNT") = oParams.Count
        '                    SaveRequired = True
        '                End If
        '                'MessageBox.Show("Parameter Count for " & oDoc.Document.File.fullfilename &" is: " & oParams.Count, "PARAMETERCOUNT")
        '                If SaveRequired Then
        '                    oDoc.Save() 'try to save the file.
        '                End If
        '            Catch
        '                'definitely need to save the file!
        '                iProperties.Value("Custom", "FEATURECOUNT") = oFeats.Count
        '                iProperties.Value("Custom", "PARAMETERCOUNT") = oParams.Count
        '                oDoc.Save() 'try to save the file.
        '                'oDoc.Close 'try to close the file - on a vaulted file this will fire the check-in dialogue.
        '            End Try
        '        End If
        '    End If
        'End Sub

        'Sub AssemblyFeatureCount(ByVal oDoc)
        '    Dim SaveRequired As Boolean = False
        '    Dim DocName As String = oDoc.DisplayName
        '    If Not oDoc.File.FullFileName.Contains("Content") Then 'skip CC and FACILITY files
        '        If Not oDoc.file.fullfilename.contains("FACILITY") Then
        '            Dim oFeats As Features = oDoc.ComponentDefinition.Features
        '            Dim Occs As ComponentOccurrences = oDoc.ComponentDefinition.Occurrences
        '            Dim oParams As Parameters = oDoc.ComponentDefinition.Parameters
        '            'Dim oConstraints as Constraints = oDoc.ComponentDefinition.Constraints
        '            Try
        '                If Not iProperties.Value("Custom", "FEATURECOUNT") = oFeats.Count Then
        '                    iProperties.Value("Custom", "FEATURECOUNT") = oFeats.Count
        '                    SaveRequired = True
        '                End If
        '                'MessageBox.Show("Feature Count for this assembly is: " & oFeats.Count, "FEATURECOUNT")
        '                If Not iProperties.Value("Custom", "OCCURRENCECOUNT") = Occs.Count Then
        '                    iProperties.Value("Custom", "OCCURRENCECOUNT") = Occs.Count
        '                    SaveRequired = True
        '                End If
        '                'MessageBox.Show("Occurrence Count for " & oDoc.File.fullfilename & " is: " & Occs.Count, "OCCURRENCECOUNT")
        '                If Not iProperties.Value("Custom", "PARAMETERCOUNT") = oParams.Count Then
        '                    iProperties.Value("Custom", "PARAMETERCOUNT") = oParams.Count
        '                    SaveRequired = True
        '                End If
        '                'MessageBox.Show("Parameter Count for this part is: " & oParams.Count, "PARAMETERCOUNT")
        '                If Not iProperties.Value("Custom", "CONSTRAINTCOUNT") = oDoc.ComponentDefinition.Constraints.Count Then
        '                    iProperties.Value("Custom", "CONSTRAINTCOUNT") = oDoc.ComponentDefinition.Constraints.Count
        '                    SaveRequired = True
        '                End If
        '                'MessageBox.Show("Constraint Count for Assembly " & DocName & " is: " & oDoc.ComponentDefinition.Constraints.Count, "CONSTRAINTCOUNT")
        '                If SaveRequired Then
        '                    oDoc.Save() 'try to save the file.
        '                End If
        '            Catch
        '                'creates any missing iProperties.
        '                iProperties.Value("Custom", "FEATURECOUNT") = oFeats.Count
        '                iProperties.Value("Custom", "OCCURRENCECOUNT") = Occs.Count
        '                iProperties.Value("Custom", "PARAMETERCOUNT") = oParams.Count
        '                iProperties.Value("Custom", "CONSTRAINTCOUNT") = oDoc.ComponentDefinition.Constraints.Count
        '                oDoc.Save() 'saves the assembly
        '                Exit Sub
        '            End Try
        '        End If
        '    End If
        'End Sub

    End Class
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
End Namespace

