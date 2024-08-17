<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAccessBackup : Inherits Form
#Region " Code généré par le Concepteur Windows Form "

    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        'Ajoutez une initialisation quelconque après l'appel InitializeComponent()

    End Sub

    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée en utilisant le Concepteur Windows Form.  
    'Ne la modifiez pas en utilisant l'éditeur de code.
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents timerDebut As System.Windows.Forms.Timer
    Friend WithEvents timerFin As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container
Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAccessBackup))
Me.lblInfo = New System.Windows.Forms.Label
Me.timerDebut = New System.Windows.Forms.Timer(Me.components)
Me.timerFin = New System.Windows.Forms.Timer(Me.components)
Me.SuspendLayout()
'
'lblInfo
'
Me.lblInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
Me.lblInfo.Location = New System.Drawing.Point(10, 10)
Me.lblInfo.Name = "lblInfo"
Me.lblInfo.Size = New System.Drawing.Size(532, 89)
Me.lblInfo.TabIndex = 0
Me.lblInfo.Text = "AccessBackup"
'
'timerDebut
'
Me.timerDebut.Interval = 500
'
'timerFin
'
Me.timerFin.Interval = 4000
'
'frmAccessBackup
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(552, 109)
Me.Controls.Add(Me.lblInfo)
Me.DockPadding.All = 10
Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
Me.Name = "frmAccessBackup"
Me.Text = "AccessBackup - Gestionnaire de sauvegarde"
Me.ResumeLayout(False)

    End Sub

#End Region
End Class
