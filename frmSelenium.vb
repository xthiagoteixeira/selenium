Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid.Views.Grid
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.IE
Imports OpenQA.Selenium.Support.UI
Imports System.ComponentModel

Public Class frmTRSEDOperador
    Inherits XtraForm

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents tAno As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbAno As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents tBimestre As System.Windows.Forms.NumericUpDown
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cbAF As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btContinuar As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents gridBoletim As DevExpress.XtraGrid.GridControl
    Friend WithEvents viewBoletim As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents txtP As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtT As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cbDisciplina As DevExpress.XtraEditors.LookUpEdit
    Friend WithEvents btTransferir As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents cbNavegador As System.Windows.Forms.ComboBox
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents barra As DevExpress.XtraEditors.ProgressBarControl
    Friend WithEvents bkTransferencia As System.ComponentModel.BackgroundWorker
    Friend WithEvents cbTurma As DevExpress.XtraEditors.LookUpEdit

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTRSEDOperador))
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl()
        Me.barra = New DevExpress.XtraEditors.ProgressBarControl()
        Me.btTransferir = New DevExpress.XtraEditors.SimpleButton()
        Me.cbDisciplina = New DevExpress.XtraEditors.LookUpEdit()
        Me.cbTurma = New DevExpress.XtraEditors.LookUpEdit()
        Me.txtP = New DevExpress.XtraEditors.LabelControl()
        Me.txtT = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.btContinuar = New DevExpress.XtraEditors.SimpleButton()
        Me.gridBoletim = New DevExpress.XtraGrid.GridControl()
        Me.viewBoletim = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.cbAF = New DevExpress.XtraEditors.CheckEdit()
        Me.tBimestre = New System.Windows.Forms.NumericUpDown()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.tAno = New System.Windows.Forms.NumericUpDown()
        Me.lbAno = New DevExpress.XtraEditors.LabelControl()
        Me.cbNavegador = New System.Windows.Forms.ComboBox()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.bkTransferencia = New System.ComponentModel.BackgroundWorker()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        CType(Me.barra.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbDisciplina.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbTurma.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridBoletim, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.viewBoletim, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbAF.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tBimestre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tAno, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupControl1
        '
        Me.GroupControl1.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupControl1.Appearance.Options.UseFont = True
        Me.GroupControl1.Controls.Add(Me.barra)
        Me.GroupControl1.Controls.Add(Me.btTransferir)
        Me.GroupControl1.Controls.Add(Me.cbDisciplina)
        Me.GroupControl1.Controls.Add(Me.cbTurma)
        Me.GroupControl1.Controls.Add(Me.txtP)
        Me.GroupControl1.Controls.Add(Me.txtT)
        Me.GroupControl1.Controls.Add(Me.LabelControl5)
        Me.GroupControl1.Controls.Add(Me.LabelControl4)
        Me.GroupControl1.Controls.Add(Me.btContinuar)
        Me.GroupControl1.Controls.Add(Me.gridBoletim)
        Me.GroupControl1.Controls.Add(Me.cbAF)
        Me.GroupControl1.Controls.Add(Me.tBimestre)
        Me.GroupControl1.Controls.Add(Me.LabelControl3)
        Me.GroupControl1.Controls.Add(Me.LabelControl2)
        Me.GroupControl1.Controls.Add(Me.LabelControl1)
        Me.GroupControl1.Controls.Add(Me.tAno)
        Me.GroupControl1.Controls.Add(Me.lbAno)
        Me.GroupControl1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupControl1.Location = New System.Drawing.Point(0, 36)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(220, 684)
        Me.GroupControl1.TabIndex = 18
        Me.GroupControl1.Text = "Escolher Boletim"
        '
        'barra
        '
        Me.barra.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barra.Location = New System.Drawing.Point(2, 657)
        Me.barra.Name = "barra"
        Me.barra.Properties.DisplayFormat.FormatString = "..."
        Me.barra.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom
        Me.barra.Properties.ShowTitle = True
        Me.barra.ShowProgressInTaskBar = True
        Me.barra.Size = New System.Drawing.Size(216, 25)
        Me.barra.TabIndex = 34
        '
        'btTransferir
        '
        Me.btTransferir.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btTransferir.Appearance.Options.UseFont = True
        Me.btTransferir.Enabled = False
        Me.btTransferir.Location = New System.Drawing.Point(7, 612)
        Me.btTransferir.Name = "btTransferir"
        Me.btTransferir.Size = New System.Drawing.Size(205, 33)
        Me.btTransferir.TabIndex = 33
        Me.btTransferir.Text = "Transferir!"
        '
        'cbDisciplina
        '
        Me.cbDisciplina.EditValue = ""
        Me.cbDisciplina.Enabled = False
        Me.cbDisciplina.Location = New System.Drawing.Point(67, 53)
        Me.cbDisciplina.Name = "cbDisciplina"
        Me.cbDisciplina.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cbDisciplina.Properties.NullText = ""
        Me.cbDisciplina.Size = New System.Drawing.Size(144, 20)
        Me.cbDisciplina.TabIndex = 32
        '
        'cbTurma
        '
        Me.cbTurma.EditValue = ""
        Me.cbTurma.Enabled = False
        Me.cbTurma.Location = New System.Drawing.Point(67, 27)
        Me.cbTurma.Name = "cbTurma"
        Me.cbTurma.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cbTurma.Properties.NullText = ""
        Me.cbTurma.Size = New System.Drawing.Size(144, 20)
        Me.cbTurma.TabIndex = 31
        '
        'txtP
        '
        Me.txtP.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtP.Appearance.Options.UseFont = True
        Me.txtP.Location = New System.Drawing.Point(161, 585)
        Me.txtP.Name = "txtP"
        Me.txtP.Size = New System.Drawing.Size(16, 16)
        Me.txtP.TabIndex = 30
        Me.txtP.Text = "00"
        '
        'txtT
        '
        Me.txtT.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtT.Appearance.Options.UseFont = True
        Me.txtT.Location = New System.Drawing.Point(84, 585)
        Me.txtT.Name = "txtT"
        Me.txtT.Size = New System.Drawing.Size(16, 16)
        Me.txtT.TabIndex = 29
        Me.txtT.Text = "00"
        '
        'LabelControl5
        '
        Me.LabelControl5.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl5.Appearance.Options.UseFont = True
        Me.LabelControl5.Location = New System.Drawing.Point(104, 585)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(56, 16)
        Me.LabelControl5.TabIndex = 28
        Me.LabelControl5.Text = "Previstas:"
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl4.Appearance.Options.UseFont = True
        Me.LabelControl4.Location = New System.Drawing.Point(49, 585)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(34, 16)
        Me.LabelControl4.TabIndex = 27
        Me.LabelControl4.Text = "Total:"
        '
        'btContinuar
        '
        Me.btContinuar.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btContinuar.Appearance.Options.UseFont = True
        Me.btContinuar.Location = New System.Drawing.Point(137, 106)
        Me.btContinuar.Name = "btContinuar"
        Me.btContinuar.Size = New System.Drawing.Size(75, 23)
        Me.btContinuar.TabIndex = 25
        Me.btContinuar.Text = "Continuar"
        '
        'gridBoletim
        '
        Me.gridBoletim.Cursor = System.Windows.Forms.Cursors.Default
        Me.gridBoletim.Location = New System.Drawing.Point(5, 158)
        Me.gridBoletim.MainView = Me.viewBoletim
        Me.gridBoletim.Name = "gridBoletim"
        Me.gridBoletim.Size = New System.Drawing.Size(207, 416)
        Me.gridBoletim.TabIndex = 24
        Me.gridBoletim.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.viewBoletim})
        '
        'viewBoletim
        '
        Me.viewBoletim.GridControl = Me.gridBoletim
        Me.viewBoletim.Name = "viewBoletim"
        Me.viewBoletim.OptionsBehavior.Editable = False
        Me.viewBoletim.OptionsSelection.MultiSelect = True
        Me.viewBoletim.OptionsView.ShowGroupPanel = False
        '
        'cbAF
        '
        Me.cbAF.Location = New System.Drawing.Point(107, 80)
        Me.cbAF.Name = "cbAF"
        Me.cbAF.Properties.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbAF.Properties.Appearance.Options.UseFont = True
        Me.cbAF.Properties.Caption = "Avaliação Final"
        Me.cbAF.Size = New System.Drawing.Size(104, 20)
        Me.cbAF.TabIndex = 23
        '
        'tBimestre
        '
        Me.tBimestre.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tBimestre.Location = New System.Drawing.Point(67, 78)
        Me.tBimestre.Maximum = New Decimal(New Integer() {4, 0, 0, 0})
        Me.tBimestre.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.tBimestre.Name = "tBimestre"
        Me.tBimestre.Size = New System.Drawing.Size(37, 23)
        Me.tBimestre.TabIndex = 22
        Me.tBimestre.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.tBimestre.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl3.Appearance.Options.UseFont = True
        Me.LabelControl3.Location = New System.Drawing.Point(7, 80)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(55, 16)
        Me.LabelControl3.TabIndex = 21
        Me.LabelControl3.Text = "Bimestre:"
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl2.Appearance.Options.UseFont = True
        Me.LabelControl2.Location = New System.Drawing.Point(4, 54)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(58, 16)
        Me.LabelControl2.TabIndex = 19
        Me.LabelControl2.Text = "Disciplina:"
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Location = New System.Drawing.Point(19, 28)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(43, 16)
        Me.LabelControl1.TabIndex = 17
        Me.LabelControl1.Text = "Turma:"
        '
        'tAno
        '
        Me.tAno.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tAno.Location = New System.Drawing.Point(67, 106)
        Me.tAno.Maximum = New Decimal(New Integer() {2500, 0, 0, 0})
        Me.tAno.Minimum = New Decimal(New Integer() {2000, 0, 0, 0})
        Me.tAno.Name = "tAno"
        Me.tAno.Size = New System.Drawing.Size(65, 23)
        Me.tAno.TabIndex = 16
        Me.tAno.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.tAno.Value = New Decimal(New Integer() {2018, 0, 0, 0})
        '
        'lbAno
        '
        Me.lbAno.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbAno.Appearance.Options.UseFont = True
        Me.lbAno.Location = New System.Drawing.Point(35, 108)
        Me.lbAno.Name = "lbAno"
        Me.lbAno.Size = New System.Drawing.Size(27, 16)
        Me.lbAno.TabIndex = 15
        Me.lbAno.Text = "Ano:"
        '
        'cbNavegador
        '
        Me.cbNavegador.FormattingEnabled = True
        Me.cbNavegador.Items.AddRange(New Object() {"Firefox", "Chrome", "Internet Explorer"})
        Me.cbNavegador.Location = New System.Drawing.Point(92, 6)
        Me.cbNavegador.Name = "cbNavegador"
        Me.cbNavegador.Size = New System.Drawing.Size(121, 24)
        Me.cbNavegador.TabIndex = 35
        '
        'LabelControl6
        '
        Me.LabelControl6.Appearance.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl6.Appearance.Options.UseFont = True
        Me.LabelControl6.Location = New System.Drawing.Point(12, 9)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(75, 16)
        Me.LabelControl6.TabIndex = 34
        Me.LabelControl6.Text = "Navegador:"
        '
        'bkTransferencia
        '
        Me.bkTransferencia.WorkerReportsProgress = True
        '
        'frmTRSEDOperador
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(220, 720)
        Me.Controls.Add(Me.cbNavegador)
        Me.Controls.Add(Me.GroupControl1)
        Me.Controls.Add(Me.LabelControl6)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTRSEDOperador"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transferir Notas! - SED"
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
        Me.GroupControl1.PerformLayout()
        CType(Me.barra.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbDisciplina.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbTurma.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridBoletim, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.viewBoletim, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbAF.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tBimestre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tAno, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim navegadorEscolhido

    Dim bimestre = String.Empty
    Dim turma, disciplina
    Dim trava = 0

    Dim descobreControlador = False

    Dim nroTotalAlunos_Caixa = 0
    Dim nroTotalAlunos_CaixaTT = False
    Dim aulasRealizadas
    Dim aulasPlanejadas
    Dim tabControlador = "0"

    Dim NroTransferencia = 1

    Dim NroAluno_Global = 0

    ReadOnly iMedia(0 To 100) As String
    ReadOnly iFalta(0 To 100) As Integer
    ReadOnly iAC(0 To 100) As Integer
    ReadOnly iPR(0 To 100) As String
    ReadOnly iTr(0 To 100) As Boolean

    ReadOnly iManual(0 To 100) As Boolean

    ReadOnly chk As New DataGridViewCheckBoxColumn

    Public Sub ZerarDados()

        NroTransferencia = 0

        contadorTransf = 1
        aulasPlanejadas = 0
        aulasRealizadas = 0

        'nroTotalAlunos = 0

        '  ifolha = 0
        '  inicio = 1
        '  final = 0
        '  trava = 0

        travaProcessos = 99

        Array.Clear(iMedia, 0, 100)
        Array.Clear(iFalta, 0, 100)
        Array.Clear(iAC, 0, 100)
        Array.Clear(iPR, 0, 100)
        Array.Clear(iTr, 0, 100)

        Array.Clear(iAV, 0, 100)
        Array.Clear(iFT, 0, 100)
        Array.Clear(iACT, 0, 100)
        Array.Clear(iPRT, 0, 100)
        Array.Clear(iTRT, 0, 100)

        btContinuar.Enabled = True
        btTransferir.Enabled = False
    End Sub

    Sub inicializar(meuNavegador As String, meuEndereco As String)

        Try
            If meuNavegador = "Chrome" Then
                navegadorEscolhido = New ChromeDriver()
                navegadorEscolhido.Navigate().GoToUrl(meuEndereco)

            ElseIf meuNavegador = "Firefox" Then
                navegadorEscolhido = New FirefoxDriver
                navegadorEscolhido.Navigate().GoToUrl(meuEndereco)

            ElseIf meuNavegador = "Internet Explorer" Then
                Dim options = New InternetExplorerOptions()
                options.IntroduceInstabilityByIgnoringProtectedModeSettings = True
                options.ElementScrollBehavior = InternetExplorerElementScrollBehavior.Bottom

                navegadorEscolhido = New InternetExplorerDriver(options)
                navegadorEscolhido.Navigate().GoToUrl(meuEndereco)
            End If
        Catch ex As Exception
            MsgBox("Navegador não encontrado!" & vbCrLf & "Erro: " & ex.Message, MsgBoxStyle.Information, "Informação")
        End Try

    End Sub

    Sub preencherBoletim(NroAluno As Integer, AV As String, Faltas As String, AC As String, Justificativa As String)

        Try
            'Dim NroAluno1 As IWebElement = navegadorEscolhido.FindElement(By.XPath(String.Format("(//table[@id='tabLancamentoFechamentoAvaliacoes']/tbody/tr/td[6])[{0}]", NroAluno)))
            Dim Aluno_Avaliacao As IWebElement =
                    navegadorEscolhido.FindElement(By.XPath(String.Format("(//input[@id='txtNota'])[{0}]", NroAluno)))
            Aluno_Avaliacao.Clear()
            Aluno_Avaliacao.SendKeys(AV)
        Catch ex As Exception
            '  MsgBox(ex.Message)
        End Try

        Try
            Dim Aluno_Faltas As IWebElement = navegadorEscolhido.FindElement(By.XPath(String.Format("(//input[@id='txtNumeroFaltas'])[{0}]", (NroAluno * 2) - 1)))
            Aluno_Faltas.Clear()
            Aluno_Faltas.SendKeys(Faltas)
        Catch ex As Exception
            '  MsgBox(ex.Message)
        End Try

        Try
            Dim Aluno_AC As IWebElement =
                    navegadorEscolhido.FindElement(
                        By.XPath(String.Format("(//input[@id='fechamento_FaltasCompensadas'])[{0}]", NroAluno)))
            Aluno_AC.Clear()
            Aluno_AC.SendKeys(AC)
        Catch ex As Exception
            '   MsgBox(ex.Message)
        End Try

        'Try
        '    If Faltas > 0 Then
        '        ''//Faltas...
        '        '//FUNCIONA 1 TAB  Dim myFaltas As IWebElement = navegadorEscolhido.FindElement(By.XPath(String.Format("//table[@id='tabLancamentoFechamentoAvaliacoes']/tbody/tr[{0}]/td[7]/input", NroAluno)))
        '        'Dim myFaltas As IWebElement = navegadorEscolhido.FindElement(By.XPath(String.Format("//table[@id='tabLancamentoFechamentoAvaliacoes']/tbody/tr[{0}]/td[7]/input", NroAluno)))


        '        If myFaltas.Enabled = True Then
        '            myFaltas.Clear()
        '            myFaltas.SendKeys(Faltas)
        '        End If
        '    End If

        'Catch ex As Exception
        '    ' MsgBox(ex.Message)
        'End Try

        'Try
        '    '//AC...

        '    If myAC.Enabled = True Then
        '        myAC.Clear()
        '        myAC.SendKeys(AC)
        '    End If

        'Catch ex As Exception
        'End Try


        ''/Justificativa...
        'navegadorEscolhido.FindElement(By.XPath("(//input[@id='fechamento_Justificativa'])[" & NroAluno & "]")).Clear()
        'navegadorEscolhido.FindElement(By.XPath("(//input[@id='fechamento_Justificativa'])[" & NroAluno & "]")).SendKeys(Justificativa)

        '' // Versão Anterior
        ' // Avaliação...
        ' Dim NroAluno1 = New SelectElement(navegadorEscolhido.FindElement(By.XPath(String.Format("//table[@id='tabLancamentoFechamentoAvaliacoes']/tbody/tr[{0}]/td[6]/select", NroAluno))))
        ' NroAluno1.SelectByText(AV)

    End Sub

    Sub preencherBoletimFinal(NroAluno As Integer, AV As String, ST As String)

        If descobreControlador = False Then

            While (descobreControlador = False)

                Try
                    '// Avaliação...
                    Dim NroAluno1 As IWebElement =
                            navegadorEscolhido.FindElement(By.XPath(String.Format("(//input[@id='txtNotaMediaFinal'])[{0}]",
                                                                                  NroAluno)))
                    NroAluno1.Clear()
                    NroAluno1.SendKeys(AV)
                    descobreControlador = True
                Catch ex As Exception
                End Try
                NroAluno = NroAluno + 1

            End While
            NroAluno_Global = NroAluno - 2

        Else

            Try
                '// Avaliação...
                Dim NroAluno1 As IWebElement =
                        navegadorEscolhido.FindElement(By.XPath(String.Format("(//input[@id='txtNotaMediaFinal'])[{0}]",
                                                                              NroAluno)))
                NroAluno1.Clear()
                NroAluno1.SendKeys(AV)
            Catch ex As Exception
            End Try

        End If


        If AV <> "S/N" Then
            Try
                '// SITUACOES
                If ST = 3 Then
                    'Retido(a) por frequência Insuficiente na disciplina
                    Dim NroAluno1 =
                            New SelectElement(
                                navegadorEscolhido.FindElement(
                                    By.XPath(String.Format("(//select[@id='slcSituacaoAlunoFechto'])[{0}]", NroAluno))))
                    NroAluno1.SelectByText("Retido(a) por frequência Insuficiente")

                ElseIf ST = 4 Then
                    'Retido(a) por rendimento Insuficiente na disciplina
                    Dim NroAluno1 =
                            New SelectElement(
                                navegadorEscolhido.FindElement(
                                    By.XPath(String.Format("(//select[@id='slcSituacaoAlunoFechto'])[{0}]", NroAluno))))
                    NroAluno1.SelectByText("Retido(a) por rendimento Insuficiente")

                ElseIf ST = 1 Or ST = "N" Then
                    'Aprovado(a)
                    Dim NroAluno1 =
                            New SelectElement(
                                navegadorEscolhido.FindElement(
                                    By.XPath(String.Format("(//select[@id='slcSituacaoAlunoFechto'])[{0}]", NroAluno))))
                    NroAluno1.SelectByText("Aprovado(a)")

                End If

            Catch ex As Exception
                '  MsgBox(ex.Message)
            End Try
        End If

        ' End If

        ''Try
        ''    '// Avaliação...
        ''    Dim NroAluno1 = New SelectElement(navegadorEscolhido.FindElement(By.XPath(String.Format("//table[@id='tabLancamentoFechamentoAvaliacoes']/tbody/tr[{0}]/td[6]/select", NroAluno))))
        ''    NroAluno1.SelectByText(AV)
        ''Catch ex As Exception
        ''End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.TopMost = "True"

        tBimestre.Value = consultaBimestre()
        tAno.Value = AnoVigente

        Dim SQL = "SELECT disciplina FROM disciplinas WHERE bloqueado='0' AND disciplina<>'-' ORDER BY disciplina;"
        cbDisciplina.Properties.DataSource = MySQL_combobox(SQL)
        cbDisciplina.Properties.DisplayMember = "disciplina"

        SQL = "SELECT classe FROM turma WHERE bloqueado='0' ORDER BY classe;"
        cbTurma.Properties.DataSource = MySQL_combobox(SQL)
        cbTurma.Properties.DisplayMember = "classe"
    End Sub

    Private Sub btTransferir_Click(sender As Object, e As EventArgs) Handles btTransferir.Click

        Dim DataHoje As String = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")
        Dim sql2 =
                String.Format(
                    "INSERT INTO logs (Descricao, idUsuario, idLogCat, DataLancamento) VALUES ('Transferiu {0}o Bimestre :{1} - {2}', '{3}', '5', '{4}');",
                    bimestre, cbTurma.Text, cbDisciplina.Text, idUsuarioDSK, DataHoje)
        MySQL_atualiza(sql2)

        cbNavegador.Enabled = False
        btTransferir.Enabled = False
        cbDisciplina.Enabled = False
        cbTurma.Enabled = False

        '...ProgressBar - Carregando...
        barra.EditValue = 0
        barra.Properties.Minimum = 1
        barra.Properties.Maximum = nroTotalAlunos - 1
        barra.Properties.DisplayFormat.FormatString = "Carregando"

        bkTransferencia.RunWorkerAsync()
    End Sub

    Private Sub cbAF_CheckedChanged(sender As Object, e As EventArgs) Handles cbAF.CheckedChanged

        If cbAF.Checked = False Then
            tBimestre.Enabled = True
        Else
            tBimestre.Enabled = False
        End If
    End Sub

    Private Sub nBimestre_ValueChanged(sender As Object, e As EventArgs) Handles tBimestre.ValueChanged

        If tBimestre.Value = 2 Or tBimestre.Value = 4 Then
            cbAF.Enabled = True
        Else
            cbAF.Enabled = False
        End If
    End Sub

    Private Sub cbTurma_Click(sender As Object, e As EventArgs) Handles cbTurma.Click
        ZerarDados()
    End Sub

    Private Sub cbDisciplina_Click(sender As Object, e As EventArgs) Handles cbDisciplina.Click
        ZerarDados()
        descobreControlador = False
        NroAluno_Global = 0

    End Sub

    Private Sub btContinuar_Click(sender As Object, e As EventArgs) Handles btContinuar.Click


        If Not (cbTurma.Text = String.Empty Or cbDisciplina.Text = String.Empty) Then

            nroTotalAlunos = 1
            ifolha = 0
            inicio = 1
            final = 0
            trava = 0

            travaProcessos = 99

            Array.Clear(iMedia, 0, 100)
            Array.Clear(iFalta, 0, 100)
            Array.Clear(iAC, 0, 100)
            Array.Clear(iPR, 0, 100)
            Array.Clear(iTr, 0, 100)

            If (cbAF.Checked = True And tBimestre.Value = 2) Then
                bimestre = "2AF"
            ElseIf (cbAF.Checked = True And tBimestre.Value = 4) Then
                bimestre = "4AF"
            ElseIf (cbAF.Checked = False) Then
                bimestre = tBimestre.Value
            End If

            turma = Consulta_Turma(cbTurma.Text)
            disciplina = Consulta_Disciplina(cbDisciplina.Text)

            Dim Sql =
                    String.Format(
                        "SELECT * FROM notasfreq WHERE disciplina='{0}' AND turma='{1}' AND cod_bimestre='{2}' AND anovigente='{3}';",
                        disciplina, turma, bimestre, tAno.Value)
            meuboletim = MySQL_consulta_campo(Sql, "cod_nf")
            txtP.Text = MySQL_consulta_campo(Sql, "previsaoaulas")
            txtT.Text = MySQL_consulta_campo(Sql, "qtdadeaulas")
            aulasRealizadas = txtT.Text
            aulasPlanejadas = txtP.Text

            If meuboletim = 0 Then
                MsgBox("Boletim não encontrado!", MsgBoxStyle.Information, "Informação!")
                travaProcessos = 99
                Exit Sub
            Else

                Sql =
                    String.Format(
                        "SELECT DISTINCT nro_aluno AS Nro, M AS AV, F AS FT, AC AS AC, S AS PR FROM boletim WHERE cod_boletim='{0}' ORDER BY nro_aluno ASC",
                        meuboletim)
                Dim MinhaDataGrid = MySQL_consulta_datagrid(Sql)
                gridBoletim.DataSource = Nothing
                gridBoletim.DataSource = MinhaDataGrid

                '/// INDICE DE ALUNOS
                For Each r In MinhaDataGrid.Rows

                    If r("AV") > 10 Then
                        iMedia(nroTotalAlunos) = String.Empty
                    Else
                        iMedia(nroTotalAlunos) = r("AV")
                    End If

                    iFalta(nroTotalAlunos) = r("FT")
                    iAC(nroTotalAlunos) = r("AC")
                    iPR(nroTotalAlunos) = r("PR")

                    nroTotalAlunos = nroTotalAlunos + 1
                Next

                btContinuar.Enabled = False
                btTransferir.Enabled = True

            End If

        End If
    End Sub

    Private Sub viewBoletim_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewBoletim.RowStyle

        Try
            Dim View As GridView = sender
            Dim NotasVermelhas As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("AV"))

            If NotasVermelhas > 4 Then
                e.Appearance.BackColor = Color.LightGreen
                e.Appearance.BackColor2 = Color.White
            Else
                e.Appearance.BackColor = Color.LightSalmon
                e.Appearance.BackColor2 = Color.White
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub cbNavegador_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbNavegador.KeyPress
        e.Handled = True
    End Sub

    Private Sub cbNavegador_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles cbNavegador.SelectedIndexChanged
        cbTurma.Enabled = True
    End Sub

    Private Sub cbTurma_EditValueChanged(sender As Object, e As EventArgs) Handles cbTurma.EditValueChanged
        cbDisciplina.Enabled = True
        nroTotalAlunos_Caixa = 0
        nroTotalAlunos_CaixaTT = False

    End Sub

    Private Sub cbNavegador_TextChanged(sender As Object, e As EventArgs) Handles cbNavegador.TextChanged

        If cbNavegador.Text <> String.Empty Then
            inicializar(cbNavegador.Text, "https://sed.educacao.sp.gov.br")
        End If
    End Sub

    Private Sub bkTransferencia_DoWork(sender As Object, e As DoWorkEventArgs) Handles bkTransferencia.DoWork

        ' Try

        Dim Controlador = 0
        Dim NroAluno = String.Empty

        'Verifica o nome da tab (Disciplina)
        ' NroAluno = navegadorEscolhido.findElement(By.Id("tabLancamentoFechamentoAvaliacoes")).Text()
        '  NroAluno = navegadorEscolhido.findElement(By.XPath("//table[@id=contains(.,'tabLancamentoFechamentoAvaliacoes')]")).Text()

        ' Controlador
        '//  NroAluno = navegadorEscolhido.findElement(By.XPath("//table[@id='tabLancamentoFechamentoAvaliacoes'])[2].10.5")).Text()
        '//  MsgBox(NroAluno)
        '// Finalizado

        'NroAluno = navegadorEscolhido.findElement(
        'By.XPath("//table[@id='tabLancamentoFechamentoAvaliacoes']/tbody/tr/td[2]")).Text()

        '// *************** Funciona com WebDriver
        '
        'Se for a primeira TAB... primeiro ALUNO
        'NroAluno = navegadorEscolhido.findElement(By.Id("tabLancamentoFechamentoAvaliacoes")).Text()
        ' PROCURA O NUMERO DO ALUNO
        Dim x = 1

        While (NroAluno = String.Empty)

            Select Case x
                Case 1
                    tabControlador = "3100"
                Case 2
                    tabControlador = "2600"
                Case 3
                    tabControlador = "1813"
                Case 4
                    tabControlador = "2800"
                Case 5
                    tabControlador = "2200"
                Case 6
                    tabControlador = "2300"
                Case 7
                    tabControlador = "1400"
                Case 8
                    tabControlador = "2400"
                Case 9
                    tabControlador = "2100"
                Case 10
                    tabControlador = "2700"
                Case 11
                    tabControlador = "1900"
                Case 12
                    tabControlador = "1111"
                Case 13
                    tabControlador = "2500"
                Case 14
                    tabControlador = "1100"
                Case 15
                    tabControlador = "2707"
                Case 16
                    tabControlador = "2105"
                Case 17
                    tabControlador = "1814"
                Case 18
                    tabControlador = "1401"
                Case 19
                    tabControlador = "2208"
                Case 20
                    tabControlador = "2306"
                Case 21
                    tabControlador = "3105"
                Case 22
                    tabControlador = "2605"
                Case 23
                    tabControlador = "2812"
                Case 24
                    tabControlador = "2413"
                Case 25
                    tabControlador = "1119"
                Case 26
                    tabControlador = "1903"
                Case 27
                    tabControlador = "1118"

                    '///...... Oficina ...
                Case 28
                    tabControlador = "8427"
                Case 29
                    tabControlador = "8442"
                Case 30
                    tabControlador = "8443"
                Case 31
                    tabControlador = "8446"
                Case 32
                    tabControlador = "8447"
                Case 33
                    tabControlador = "8441"
                    '// Fim
                Case 34
                    tabControlador = "2504"
                Case 35
                    tabControlador = "6400"

                Case 36
                    Dim tabControlador2 = 1100

                    While (NroAluno = String.Empty)
                        Try
                            NroAluno = navegadorEscolhido.findElement(By.XPath(String.Format("//table[contains(@id,'tabLancamentoFechamentoAvaliacoes{0}')]/tbody/tr/td[2]", tabControlador2))).Text()
                            tabControlador2 = tabControlador2 + 1
                            tabControlador = tabControlador2
                        Catch ex As Exception
                        End Try
                    End While

            End Select
            Try
                NroAluno = navegadorEscolhido.findElement(By.XPath(String.Format("//table[contains(@id,'tabLancamentoFechamentoAvaliacoes{0}')]/tbody/tr/td[2]", tabControlador))).Text()
            Catch ex As Exception
                NroAluno = String.Empty
            End Try
            x = x + 1
        End While

        ' *****************************************
        ' *********** OUTROS TAB ******************
        ' *****************************************

        If NroAluno = String.Empty Then

            Dim AulasJaEnviadas = False

            'Procura Indice da TAB
            For t = 2 To 20 Step 1

                If AulasJaEnviadas = False Then
                    Try
                        'AulasPlanejadas
                        Dim AulasPlanejadas As IWebElement =
                                navegadorEscolhido.FindElement(
                                    By.XPath(String.Format("(//input[contains(@id,'txtAulasPlanejadas')])[{0}]", t)))
                        AulasPlanejadas.Clear()
                        AulasPlanejadas.SendKeys(AulasPlanejadas)

                        'AulasRealizadas
                        Dim AulasRealizadas As IWebElement =
                                navegadorEscolhido.FindElement(
                                    By.XPath(String.Format("(//input[contains(@id,'txtAulasRealizadas')])[{0}]", t)))
                        AulasRealizadas.Clear()
                        AulasRealizadas.SendKeys(AulasRealizadas)
                        AulasJaEnviadas = True

                        Controlador = t

                        Exit For

                    Catch ex As Exception
                    End Try

                End If

            Next
            '\\ Fim da Indice da TAB

            If nroTotalAlunos_Caixa = 0 Then
                nroTotalAlunos_Caixa = nroTotalAlunos * (Controlador - 1)
            Else
                nroTotalAlunos_Caixa = nroTotalAlunos_Caixa * (Controlador - 1)
            End If

            ' ******************************************* Envia os dados OUTRAS Tabs ***************************
            For i = 1 To nroTotalAlunos - 1 Step 1

                bkTransferencia.ReportProgress(i + 1)

                Try
                    NroAluno = navegadorEscolhido.findElement(
                          By.XPath(String.Format("(//table[contains(@id,'tabLancamentoFechamentoAvaliacoes" & tabControlador & "')]/tbody/tr/td[2])[{0}]", i))).Text()

                    If iMedia(NroAluno) <> Nothing Then
                        If cbAF.Checked = True Or tBimestre.Value = 5 Then
                            preencherBoletimFinal(nroTotalAlunos_Caixa + i, iMedia(NroAluno), iPR(NroAluno))
                        Else
                            preencherBoletim(nroTotalAlunos_Caixa + i, iMedia(NroAluno), iFalta(NroAluno), iAC(NroAluno), "-")
                        End If
                    End If

                Catch ex As Exception
                End Try

            Next

        Else
            '******************************************
            ' *************************  TAB **********
            '******************************************

            Dim AulasJaEnviadas = False

            'Procura Indice da TAB
            '
            For r = 1 To 20 Step 1

                If AulasJaEnviadas = False Then
                    Try

                        'AulasPlanejadas
                        Dim AulasPlanejadas As IWebElement =
                                    navegadorEscolhido.FindElement(
                                        By.XPath(String.Format("(//input[@id='txtAulasPlanejadas'])[{0}]", r)))
                        AulasPlanejadas.Clear()
                        AulasPlanejadas.SendKeys(txtP.Text)

                        'AulasRealizadas
                        Dim AulasRealizadas As IWebElement =
                                        navegadorEscolhido.FindElement(
                                        By.XPath(String.Format("(//input[@id='txtAulasRealizadas'])[{0}]", r)))
                        AulasRealizadas.Clear()
                        AulasRealizadas.SendKeys(txtT.Text)
                        AulasJaEnviadas = True


                        Controlador = r
                        Exit For

                    Catch ex As Exception
                    End Try

                End If

            Next

            Dim t = 0
            ' LOCALIZAR A CAIXA...
            If nroTotalAlunos_Caixa = 0 And nroTotalAlunos_CaixaTT = False And Controlador = 1 Then
                t = 0
            ElseIf nroTotalAlunos_Caixa = 0 And nroTotalAlunos_CaixaTT = False And Controlador = 2 Then
                't = nroTotalAlunos - 1
                MsgBox("Transferir a primeira disciplina, para liberar esta!", MsgBoxStyle.Information, "Informação!")
                Exit Sub
            ElseIf nroTotalAlunos_Caixa = 0 And nroTotalAlunos_CaixaTT = False And Controlador > 2 Then
                't = (nroTotalAlunos - 1) * (Controlador - 1)
                MsgBox("Transferir a primeira disciplina, para liberar esta!", MsgBoxStyle.Information, "Informação!")
                Exit Sub
            ElseIf Controlador = 2 Then
                t = nroTotalAlunos_Caixa - 1
            ElseIf Controlador > 2 Then
                t = (nroTotalAlunos_Caixa - 1) * (Controlador - 1)
            End If

            ' Envia os dados Tab ...
            Try
                For i = 1 To nroTotalAlunos - 1 Step 1
                    Try
                        bkTransferencia.ReportProgress(i)
                        NroAluno = navegadorEscolhido.findElement(
                                    By.XPath(String.Format("//table[contains(@id,'tabLancamentoFechamentoAvaliacoes" & tabControlador & "')]/tbody/tr[{0}]/td[2]", i))).Text()
                        If cbAF.Checked = True Or tBimestre.Value = 5 Then
                            preencherBoletimFinal(NroAluno_Global + i, iMedia(NroAluno), iPR(NroAluno))
                        Else
                            preencherBoletim(t + i, iMedia(NroAluno), iFalta(NroAluno), iAC(NroAluno), "-")
                        End If

                        If nroTotalAlunos_CaixaTT = False Then
                            nroTotalAlunos_Caixa = i + 1
                        End If

                    Catch ex As Exception
                    End Try
                Next

                ' Para alunos que estão na lista mas sem nota (Mais Escola)
                If nroTotalAlunos_CaixaTT = False Then
                    nroTotalAlunos_CaixaTT = True
                    Dim ProximaCaixa = nroTotalAlunos_Caixa

                    For u = 0 To 20 Step 1
                        Try
                            NroAluno = navegadorEscolhido.findElement(
                            By.XPath(String.Format("//table[contains(@id,'tabLancamentoFechamentoAvaliacoes" & tabControlador & "')]/tbody/tr[{0}]/td[2]", ProximaCaixa + u))).Text()
                        Catch ex As Exception
                            NroAluno = String.Empty
                        End Try

                        If NroAluno <> String.Empty Then
                            nroTotalAlunos_Caixa = nroTotalAlunos_Caixa + 1
                        End If

                    Next

                End If

            Catch ex As Exception
            End Try

            ''Envia as AULAS DADAS E PREVISTAS...
            'Try
            '    'AulasPlanejadas
            '    Dim AulasPlanejadas As IWebElement =
            '            navegadorEscolhido.FindElement(
            '                By.XPath(String.Format("(//input[contains(@id,'txtAulasPlanejadas')])[1]")))
            '    AulasPlanejadas.Clear()
            '    AulasPlanejadas.SendKeys(txtP.Text)

            '    'AulasRealizadas
            '    Dim AulasRealizadas As IWebElement =
            '            navegadorEscolhido.FindElement(
            '                By.XPath(String.Format("(//input[contains(@id,'txtAulasRealizadas')])[1]")))
            '    AulasRealizadas.Clear()
            '    AulasRealizadas.SendKeys(txtT.Text)

            'Catch ex As Exception
            '    ' MsgBox(ex.Message)
            'End Try

            '' Não permite incrementar TEXTBOX 
            'nroTotalAlunos_CaixaTT = True

        End If

        'Catch ex As Exception
        '    ' MsgBox(ex.Message)
        'End Try

        ' Preenche o primeiro aluno a ser transferido.
        ''// Avaliação...
        'Dim NroAluno1 = navegadorEscolhido.FindElement(By.XPath("//table[@id='" & encontrada_DataTable & "']/tbody/tr[" & NroTransferencia & "]/td[6]/select"))
        'If NroAluno1.Enabled = True Then
        '    NroAluno1.SendKeys(iMedia(NroTransferencia))
        'End If
    End Sub

    Private Sub bkTransferencia_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) _
        Handles bkTransferencia.ProgressChanged

        Invoke(Sub()
                   Me.barra.Properties.DisplayFormat.FormatString = "Enviando: [ " & e.ProgressPercentage & " / " &
                                                                    nroTotalAlunos - 1 & " ]"
                   Me.barra.EditValue = e.ProgressPercentage
               End Sub)
    End Sub

    Private Sub bkTransferencia_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) _
        Handles bkTransferencia.RunWorkerCompleted
        cbDisciplina.Enabled = True
        cbTurma.Enabled = True

        barra.Properties.DisplayFormat.FormatString = "Finalizado!"
    End Sub

    Private Sub GroupControl1_Paint(sender As Object, e As PaintEventArgs) Handles GroupControl1.Paint
    End Sub
End Class
