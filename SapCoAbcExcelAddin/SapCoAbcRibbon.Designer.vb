Partial Class SapCoAbcRibbon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapCoAbcRibbon))
        Me.SapCoAbc = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButtonSapTLRead = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.ButtonSapTLDeleteCheck = Me.Factory.CreateRibbonButton
        Me.ButtonSapTLDeletePost = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.ButtonSapTLCreateCheck = Me.Factory.CreateRibbonButton
        Me.ButtonSapTLCreatePost = Me.Factory.CreateRibbonButton
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.ButtonSapTLGenerate = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapCoAbc.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCoAbc
        '
        Me.SapCoAbc.Groups.Add(Me.Group2)
        Me.SapCoAbc.Groups.Add(Me.Group3)
        Me.SapCoAbc.Label = "SAP CO-ABC"
        Me.SapCoAbc.Name = "SapCoAbc"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButtonSapTLRead)
        Me.Group2.Items.Add(Me.Separator1)
        Me.Group2.Items.Add(Me.ButtonSapTLDeleteCheck)
        Me.Group2.Items.Add(Me.ButtonSapTLDeletePost)
        Me.Group2.Items.Add(Me.Separator2)
        Me.Group2.Items.Add(Me.ButtonSapTLCreateCheck)
        Me.Group2.Items.Add(Me.ButtonSapTLCreatePost)
        Me.Group2.Items.Add(Me.Separator3)
        Me.Group2.Items.Add(Me.ButtonSapTLGenerate)
        Me.Group2.Label = "SAP Template"
        Me.Group2.Name = "Group2"
        '
        'ButtonSapTLRead
        '
        Me.ButtonSapTLRead.Image = CType(resources.GetObject("ButtonSapTLRead.Image"), System.Drawing.Image)
        Me.ButtonSapTLRead.Label = "Read Templates"
        Me.ButtonSapTLRead.Name = "ButtonSapTLRead"
        Me.ButtonSapTLRead.ShowImage = True
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'ButtonSapTLDeleteCheck
        '
        Me.ButtonSapTLDeleteCheck.Image = CType(resources.GetObject("ButtonSapTLDeleteCheck.Image"), System.Drawing.Image)
        Me.ButtonSapTLDeleteCheck.Label = "Check Delete Templates"
        Me.ButtonSapTLDeleteCheck.Name = "ButtonSapTLDeleteCheck"
        Me.ButtonSapTLDeleteCheck.ScreenTip = "Check Delete Templates"
        Me.ButtonSapTLDeleteCheck.ShowImage = True
        '
        'ButtonSapTLDeletePost
        '
        Me.ButtonSapTLDeletePost.Image = CType(resources.GetObject("ButtonSapTLDeletePost.Image"), System.Drawing.Image)
        Me.ButtonSapTLDeletePost.Label = "Post Delete Templates"
        Me.ButtonSapTLDeletePost.Name = "ButtonSapTLDeletePost"
        Me.ButtonSapTLDeletePost.ScreenTip = "Post Delete Templates"
        Me.ButtonSapTLDeletePost.ShowImage = True
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'ButtonSapTLCreateCheck
        '
        Me.ButtonSapTLCreateCheck.Image = CType(resources.GetObject("ButtonSapTLCreateCheck.Image"), System.Drawing.Image)
        Me.ButtonSapTLCreateCheck.Label = "Check Create Templates"
        Me.ButtonSapTLCreateCheck.Name = "ButtonSapTLCreateCheck"
        Me.ButtonSapTLCreateCheck.ShowImage = True
        '
        'ButtonSapTLCreatePost
        '
        Me.ButtonSapTLCreatePost.Image = CType(resources.GetObject("ButtonSapTLCreatePost.Image"), System.Drawing.Image)
        Me.ButtonSapTLCreatePost.Label = "Post Create Templates"
        Me.ButtonSapTLCreatePost.Name = "ButtonSapTLCreatePost"
        Me.ButtonSapTLCreatePost.ScreenTip = "Post Create Templates"
        Me.ButtonSapTLCreatePost.ShowImage = True
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'ButtonSapTLGenerate
        '
        Me.ButtonSapTLGenerate.Image = CType(resources.GetObject("ButtonSapTLGenerate.Image"), System.Drawing.Image)
        Me.ButtonSapTLGenerate.Label = "Generate Templates"
        Me.ButtonSapTLGenerate.Name = "ButtonSapTLGenerate"
        Me.ButtonSapTLGenerate.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ButtonLogon)
        Me.Group3.Items.Add(Me.ButtonLogoff)
        Me.Group3.Label = "SAP Logon"
        Me.Group3.Name = "Group3"
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
        'SapCoAbcRibbon
        '
        Me.Name = "SapCoAbcRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapCoAbc)
        Me.SapCoAbc.ResumeLayout(False)
        Me.SapCoAbc.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapCoAbc As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapTLRead As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ButtonSapTLDeleteCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapTLDeletePost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ButtonSapTLCreateCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapTLCreatePost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ButtonSapTLGenerate As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapCoAbcRibbon
        Get
            Return Me.GetRibbon(Of SapCoAbcRibbon)()
        End Get
    End Property
End Class
