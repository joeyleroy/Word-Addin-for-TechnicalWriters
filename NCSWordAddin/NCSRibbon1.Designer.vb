Partial Class NCSRibbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase
    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(NCSRibbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.MenuHeadings1 = Me.Factory.CreateRibbonMenu
        Me.ButHeading1 = Me.Factory.CreateRibbonButton
        Me.ButHeading2 = Me.Factory.CreateRibbonButton
        Me.ButHeading3 = Me.Factory.CreateRibbonButton
        Me.ButHeading4 = Me.Factory.CreateRibbonButton
        Me.MenuBody1 = Me.Factory.CreateRibbonMenu
        Me.ButBodyText1 = Me.Factory.CreateRibbonButton
        Me.ButBodyText2 = Me.Factory.CreateRibbonButton
        Me.ButBodyText3 = Me.Factory.CreateRibbonButton
        Me.ButTableSpace1 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButPageBreak = Me.Factory.CreateRibbonButton
        Me.ButKeepWithNext = Me.Factory.CreateRibbonButton
        Me.ButInsertGraphic = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ButSnipitz = Me.Factory.CreateRibbonButton
        Me.ButXREF = Me.Factory.CreateRibbonButton
        Me.ButTemplates = Me.Factory.CreateRibbonButton
        Me.ButUnitConverter = Me.Factory.CreateRibbonButton
        Me.ButUpdateFields = Me.Factory.CreateRibbonButton
        Me.ButToggleDocProps = Me.Factory.CreateRibbonButton
        Me.ButFixHeadings = Me.Factory.CreateRibbonButton
        Me.ButFormatNotes = Me.Factory.CreateRibbonButton
        Me.FFM = Me.Factory.CreateRibbonMenu
        Me.LMFB1 = Me.Factory.CreateRibbonButton
        Me.LMFB2 = Me.Factory.CreateRibbonButton
        Me.LMFB3 = Me.Factory.CreateRibbonButton
        Me.Configuration = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.KeyTip = "X"
        Me.Tab1.Label = "NCS"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.MenuHeadings1)
        Me.Group1.Items.Add(Me.MenuBody1)
        Me.Group1.Items.Add(Me.ButTableSpace1)
        Me.Group1.Label = "Styles"
        Me.Group1.Name = "Group1"
        '
        'MenuHeadings1
        '
        Me.MenuHeadings1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MenuHeadings1.Image = CType(resources.GetObject("MenuHeadings1.Image"), System.Drawing.Image)
        Me.MenuHeadings1.Items.Add(Me.ButHeading1)
        Me.MenuHeadings1.Items.Add(Me.ButHeading2)
        Me.MenuHeadings1.Items.Add(Me.ButHeading3)
        Me.MenuHeadings1.Items.Add(Me.ButHeading4)
        Me.MenuHeadings1.KeyTip = "H"
        Me.MenuHeadings1.Label = "Headings"
        Me.MenuHeadings1.Name = "MenuHeadings1"
        Me.MenuHeadings1.ScreenTip = "Heading Styles"
        Me.MenuHeadings1.ShowImage = True
        Me.MenuHeadings1.SuperTip = "Contains the standard heading styles. Use Headings for content organization."
        '
        'ButHeading1
        '
        Me.ButHeading1.Image = CType(resources.GetObject("ButHeading1.Image"), System.Drawing.Image)
        Me.ButHeading1.KeyTip = "H1"
        Me.ButHeading1.Label = "Heading 1"
        Me.ButHeading1.Name = "ButHeading1"
        Me.ButHeading1.ShowImage = True
        Me.ButHeading1.ShowLabel = False
        '
        'ButHeading2
        '
        Me.ButHeading2.Image = CType(resources.GetObject("ButHeading2.Image"), System.Drawing.Image)
        Me.ButHeading2.KeyTip = "H2"
        Me.ButHeading2.Label = "Heading 2"
        Me.ButHeading2.Name = "ButHeading2"
        Me.ButHeading2.ShowImage = True
        Me.ButHeading2.ShowLabel = False
        '
        'ButHeading3
        '
        Me.ButHeading3.Image = CType(resources.GetObject("ButHeading3.Image"), System.Drawing.Image)
        Me.ButHeading3.KeyTip = "H3"
        Me.ButHeading3.Label = "Heading 3"
        Me.ButHeading3.Name = "ButHeading3"
        Me.ButHeading3.ShowImage = True
        Me.ButHeading3.ShowLabel = False
        '
        'ButHeading4
        '
        Me.ButHeading4.Image = CType(resources.GetObject("ButHeading4.Image"), System.Drawing.Image)
        Me.ButHeading4.KeyTip = "H4"
        Me.ButHeading4.Label = "Heading 4"
        Me.ButHeading4.Name = "ButHeading4"
        Me.ButHeading4.ShowImage = True
        Me.ButHeading4.ShowLabel = False
        '
        'MenuBody1
        '
        Me.MenuBody1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MenuBody1.Image = CType(resources.GetObject("MenuBody1.Image"), System.Drawing.Image)
        Me.MenuBody1.Items.Add(Me.ButBodyText1)
        Me.MenuBody1.Items.Add(Me.ButBodyText2)
        Me.MenuBody1.Items.Add(Me.ButBodyText3)
        Me.MenuBody1.KeyTip = "B"
        Me.MenuBody1.Label = "Body Text"
        Me.MenuBody1.Name = "MenuBody1"
        Me.MenuBody1.ScreenTip = "BodyText Styles"
        Me.MenuBody1.ShowImage = True
        Me.MenuBody1.SuperTip = "Contains the standard body styles. Use BodyText for text content that is in parag" &
    "raph form."
        '
        'ButBodyText1
        '
        Me.ButBodyText1.Image = CType(resources.GetObject("ButBodyText1.Image"), System.Drawing.Image)
        Me.ButBodyText1.KeyTip = "B1"
        Me.ButBodyText1.Label = "Body Text 1"
        Me.ButBodyText1.Name = "ButBodyText1"
        Me.ButBodyText1.ShowImage = True
        Me.ButBodyText1.ShowLabel = False
        '
        'ButBodyText2
        '
        Me.ButBodyText2.Image = CType(resources.GetObject("ButBodyText2.Image"), System.Drawing.Image)
        Me.ButBodyText2.KeyTip = "B2"
        Me.ButBodyText2.Label = "Body Text 2"
        Me.ButBodyText2.Name = "ButBodyText2"
        Me.ButBodyText2.ShowImage = True
        Me.ButBodyText2.ShowLabel = False
        '
        'ButBodyText3
        '
        Me.ButBodyText3.Image = CType(resources.GetObject("ButBodyText3.Image"), System.Drawing.Image)
        Me.ButBodyText3.KeyTip = "B3"
        Me.ButBodyText3.Label = "Body Text 3"
        Me.ButBodyText3.Name = "ButBodyText3"
        Me.ButBodyText3.ShowImage = True
        Me.ButBodyText3.ShowLabel = False
        '
        'ButTableSpace1
        '
        Me.ButTableSpace1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButTableSpace1.Image = CType(resources.GetObject("ButTableSpace1.Image"), System.Drawing.Image)
        Me.ButTableSpace1.KeyTip = "TS"
        Me.ButTableSpace1.Label = "Table Space"
        Me.ButTableSpace1.Name = "ButTableSpace1"
        Me.ButTableSpace1.ScreenTip = "Format Selection to Table Space"
        Me.ButTableSpace1.ShowImage = True
        Me.ButTableSpace1.SuperTip = resources.GetString("ButTableSpace1.SuperTip")
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButPageBreak)
        Me.Group2.Items.Add(Me.ButKeepWithNext)
        Me.Group2.Items.Add(Me.ButInsertGraphic)
        Me.Group2.Label = "Controls"
        Me.Group2.Name = "Group2"
        '
        'ButPageBreak
        '
        Me.ButPageBreak.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButPageBreak.Image = CType(resources.GetObject("ButPageBreak.Image"), System.Drawing.Image)
        Me.ButPageBreak.KeyTip = "PB"
        Me.ButPageBreak.Label = "Page Break"
        Me.ButPageBreak.Name = "ButPageBreak"
        Me.ButPageBreak.ScreenTip = "Page Break Before the Selection"
        Me.ButPageBreak.ShowImage = True
        Me.ButPageBreak.SuperTip = resources.GetString("ButPageBreak.SuperTip")
        '
        'ButKeepWithNext
        '
        Me.ButKeepWithNext.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButKeepWithNext.Image = CType(resources.GetObject("ButKeepWithNext.Image"), System.Drawing.Image)
        Me.ButKeepWithNext.KeyTip = "KWN"
        Me.ButKeepWithNext.Label = "Keep w/Next"
        Me.ButKeepWithNext.Name = "ButKeepWithNext"
        Me.ButKeepWithNext.ScreenTip = "Keep Selection with the Following (Next) Content"
        Me.ButKeepWithNext.ShowImage = True
        Me.ButKeepWithNext.SuperTip = resources.GetString("ButKeepWithNext.SuperTip")
        '
        'ButInsertGraphic
        '
        Me.ButInsertGraphic.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButInsertGraphic.Image = CType(resources.GetObject("ButInsertGraphic.Image"), System.Drawing.Image)
        Me.ButInsertGraphic.KeyTip = "IG"
        Me.ButInsertGraphic.Label = "Insert Graphic"
        Me.ButInsertGraphic.Name = "ButInsertGraphic"
        Me.ButInsertGraphic.ScreenTip = resources.GetString("ButInsertGraphic.ScreenTip")
        Me.ButInsertGraphic.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ButSnipitz)
        Me.Group3.Items.Add(Me.ButXREF)
        Me.Group3.Items.Add(Me.ButTemplates)
        Me.Group3.Items.Add(Me.ButUnitConverter)
        Me.Group3.Items.Add(Me.ButUpdateFields)
        Me.Group3.Items.Add(Me.ButToggleDocProps)
        Me.Group3.Items.Add(Me.ButFixHeadings)
        Me.Group3.Items.Add(Me.ButFormatNotes)
        Me.Group3.Items.Add(Me.FFM)
        Me.Group3.Items.Add(Me.Configuration)
        Me.Group3.Label = "Tools"
        Me.Group3.Name = "Group3"
        '
        'ButSnipitz
        '
        Me.ButSnipitz.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButSnipitz.Image = CType(resources.GetObject("ButSnipitz.Image"), System.Drawing.Image)
        Me.ButSnipitz.KeyTip = "S"
        Me.ButSnipitz.Label = "Snipits"
        Me.ButSnipitz.Name = "ButSnipitz"
        Me.ButSnipitz.ScreenTip = "Snipits Manager"
        Me.ButSnipitz.ShowImage = True
        Me.ButSnipitz.SuperTip = resources.GetString("ButSnipitz.SuperTip")
        '
        'ButXREF
        '
        Me.ButXREF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButXREF.Image = CType(resources.GetObject("ButXREF.Image"), System.Drawing.Image)
        Me.ButXREF.KeyTip = "X"
        Me.ButXREF.Label = "XRef"
        Me.ButXREF.Name = "ButXREF"
        Me.ButXREF.ScreenTip = "Cross-Reference Manager"
        Me.ButXREF.ShowImage = True
        Me.ButXREF.SuperTip = resources.GetString("ButXREF.SuperTip")
        '
        'ButTemplates
        '
        Me.ButTemplates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButTemplates.Image = CType(resources.GetObject("ButTemplates.Image"), System.Drawing.Image)
        Me.ButTemplates.KeyTip = "TP"
        Me.ButTemplates.Label = "Templates"
        Me.ButTemplates.Name = "ButTemplates"
        Me.ButTemplates.ScreenTip = "Document Template Manager"
        Me.ButTemplates.ShowImage = True
        Me.ButTemplates.SuperTip = resources.GetString("ButTemplates.SuperTip")
        '
        'ButUnitConverter
        '
        Me.ButUnitConverter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButUnitConverter.Image = CType(resources.GetObject("ButUnitConverter.Image"), System.Drawing.Image)
        Me.ButUnitConverter.KeyTip = "UC"
        Me.ButUnitConverter.Label = "Unit Converter"
        Me.ButUnitConverter.Name = "ButUnitConverter"
        Me.ButUnitConverter.ScreenTip = "Converts the selected unit of measure."
        Me.ButUnitConverter.ShowImage = True
        Me.ButUnitConverter.SuperTip = resources.GetString("ButUnitConverter.SuperTip")
        '
        'ButUpdateFields
        '
        Me.ButUpdateFields.Image = CType(resources.GetObject("ButUpdateFields.Image"), System.Drawing.Image)
        Me.ButUpdateFields.KeyTip = "UF"
        Me.ButUpdateFields.Label = "Update Fields"
        Me.ButUpdateFields.Name = "ButUpdateFields"
        Me.ButUpdateFields.ScreenTip = "Update Document Fields"
        Me.ButUpdateFields.ShowImage = True
        Me.ButUpdateFields.SuperTip = resources.GetString("ButUpdateFields.SuperTip")
        '
        'ButToggleDocProps
        '
        Me.ButToggleDocProps.Image = CType(resources.GetObject("ButToggleDocProps.Image"), System.Drawing.Image)
        Me.ButToggleDocProps.KeyTip = "DP"
        Me.ButToggleDocProps.Label = "Document Properties"
        Me.ButToggleDocProps.Name = "ButToggleDocProps"
        Me.ButToggleDocProps.ScreenTip = "Toggle the Document Properties"
        Me.ButToggleDocProps.ShowImage = True
        Me.ButToggleDocProps.SuperTip = resources.GetString("ButToggleDocProps.SuperTip")
        '
        'ButFixHeadings
        '
        Me.ButFixHeadings.Image = CType(resources.GetObject("ButFixHeadings.Image"), System.Drawing.Image)
        Me.ButFixHeadings.KeyTip = "FH"
        Me.ButFixHeadings.Label = "Fix Headings"
        Me.ButFixHeadings.Name = "ButFixHeadings"
        Me.ButFixHeadings.ScreenTip = "Fix headings in the document"
        Me.ButFixHeadings.ShowImage = True
        Me.ButFixHeadings.SuperTip = "Use when headings appear incorrect, misaligned, or out of numbering sequence."
        '
        'ButFormatNotes
        '
        Me.ButFormatNotes.Image = CType(resources.GetObject("ButFormatNotes.Image"), System.Drawing.Image)
        Me.ButFormatNotes.KeyTip = "FN"
        Me.ButFormatNotes.Label = "Format Note"
        Me.ButFormatNotes.Name = "ButFormatNotes"
        Me.ButFormatNotes.ScreenTip = "Format the selected table as a note"
        Me.ButFormatNotes.ShowImage = True
        Me.ButFormatNotes.SuperTip = resources.GetString("ButFormatNotes.SuperTip")
        '
        'FFM
        '
        Me.FFM.Description = "Can't find a form? Click and select a form."
        Me.FFM.Image = CType(resources.GetObject("FFM.Image"), System.Drawing.Image)
        Me.FFM.Items.Add(Me.LMFB1)
        Me.FFM.Items.Add(Me.LMFB2)
        Me.FFM.Items.Add(Me.LMFB3)
        Me.FFM.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.FFM.KeyTip = "FF"
        Me.FFM.Label = "Find Form"
        Me.FFM.Name = "FFM"
        Me.FFM.ScreenTip = "Find a Lost Form"
        Me.FFM.ShowImage = True
        Me.FFM.SuperTip = "Lost your form? These things happen... Just click the button of the lost form to " &
    "find it. Magic!"
        '
        'LMFB1
        '
        Me.LMFB1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.LMFB1.Image = CType(resources.GetObject("LMFB1.Image"), System.Drawing.Image)
        Me.LMFB1.KeyTip = "FFS"
        Me.LMFB1.Label = "Find Snipits"
        Me.LMFB1.Name = "LMFB1"
        Me.LMFB1.ScreenTip = "Click to find a lost Snipits Form"
        Me.LMFB1.ShowImage = True
        '
        'LMFB2
        '
        Me.LMFB2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.LMFB2.Image = CType(resources.GetObject("LMFB2.Image"), System.Drawing.Image)
        Me.LMFB2.KeyTip = "FFX"
        Me.LMFB2.Label = "Find XRef"
        Me.LMFB2.Name = "LMFB2"
        Me.LMFB2.ScreenTip = "Click to find a lost XRef Form"
        Me.LMFB2.ShowImage = True
        '
        'LMFB3
        '
        Me.LMFB3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.LMFB3.Image = CType(resources.GetObject("LMFB3.Image"), System.Drawing.Image)
        Me.LMFB3.KeyTip = "FFT"
        Me.LMFB3.Label = "Find Templates"
        Me.LMFB3.Name = "LMFB3"
        Me.LMFB3.ScreenTip = "Click to find a lost Templates Form"
        Me.LMFB3.ShowImage = True
        '
        'Configuration
        '
        Me.Configuration.Label = "Configuration"
        Me.Configuration.Name = "Configuration"
        '
        'NCSRibbon1
        '
        Me.Name = "NCSRibbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents MenuHeadings1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents ButHeading1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButHeading2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButHeading3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButHeading4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MenuBody1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents ButBodyText1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButBodyText2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButBodyText3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButTableSpace1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButInsertGraphic As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButKeepWithNext As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButPageBreak As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButUnitConverter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButToggleDocProps As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButSnipitz As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButUpdateFields As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButFormatNotes As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButFixHeadings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButXREF As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButTemplates As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FFM As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents LMFB1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents LMFB2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents LMFB3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Configuration As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property NCSRibbon1() As NCSRibbon1
        Get
            Return Me.GetRibbon(Of NCSRibbon1)()
        End Get
    End Property
End Class
