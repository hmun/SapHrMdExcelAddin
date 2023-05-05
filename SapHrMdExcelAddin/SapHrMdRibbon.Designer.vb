Partial Class SapHrMdRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapHrMdRibbon))
        Me.SapMmMd = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ButtonSapHrMasterSave = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapMmMd.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapMmMd
        '
        Me.SapMmMd.Groups.Add(Me.Group1)
        Me.SapMmMd.Groups.Add(Me.Group2)
        Me.SapMmMd.Label = "SAP HR Md"
        Me.SapMmMd.Name = "SapMmMd"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonSapHrMasterSave)
        Me.Group1.Label = "HR Master"
        Me.Group1.Name = "Group1"
        '
        'ButtonSapHrMasterSave
        '
        Me.ButtonSapHrMasterSave.Image = CType(resources.GetObject("ButtonSapHrMasterSave.Image"), System.Drawing.Image)
        Me.ButtonSapHrMasterSave.Label = "Save HR Master"
        Me.ButtonSapHrMasterSave.Name = "ButtonSapHrMasterSave"
        Me.ButtonSapHrMasterSave.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButtonLogon)
        Me.Group2.Items.Add(Me.ButtonLogoff)
        Me.Group2.Label = "SAP Logon"
        Me.Group2.Name = "Group2"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'SapHrMdRibbon
        '
        Me.Name = "SapHrMdRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapMmMd)
        Me.SapMmMd.ResumeLayout(False)
        Me.SapMmMd.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapMmMd As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapHrMasterSave As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapHrMdRibbon
        Get
            Return Me.GetRibbon(Of SapHrMdRibbon)()
        End Get
    End Property
End Class
