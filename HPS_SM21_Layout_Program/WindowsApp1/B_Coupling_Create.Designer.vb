<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class B_Coupling_Create
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(B_Coupling_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Check2 = New System.Windows.Forms.CheckBox()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GBX_Manufacturer = New System.Windows.Forms.GroupBox()
        Me.CBX_CHN = New System.Windows.Forms.RadioButton()
        Me.CBX_KOR = New System.Windows.Forms.RadioButton()
        Me.CBX_DEFAULT = New System.Windows.Forms.RadioButton()
        Me.Coupling_Create = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Dummy_Loading_Check = New System.Windows.Forms.CheckBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Coupling_Clash_Check = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GBX_Manufacturer.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Check2)
        Me.GroupBox1.Controls.Add(Me.Check1)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 19)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(356, 47)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Type"
        '
        'Check2
        '
        Me.Check2.AutoSize = True
        Me.Check2.Location = New System.Drawing.Point(127, 20)
        Me.Check2.Name = "Check2"
        Me.Check2.Size = New System.Drawing.Size(57, 16)
        Me.Check2.TabIndex = 0
        Me.Check2.Text = "GEAR"
        Me.Check2.UseVisualStyleBackColor = True
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Location = New System.Drawing.Point(15, 20)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(51, 16)
        Me.Check1.TabIndex = 0
        Me.Check1.Text = "DISK"
        Me.Check1.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.GBX_Manufacturer)
        Me.GroupBox3.Controls.Add(Me.GroupBox1)
        Me.GroupBox3.Controls.Add(Me.Coupling_Create)
        Me.GroupBox3.Controls.Add(Me.GroupBox2)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 15)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(381, 218)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Coupling 실적 형상 검색"
        '
        'GBX_Manufacturer
        '
        Me.GBX_Manufacturer.Controls.Add(Me.CBX_CHN)
        Me.GBX_Manufacturer.Controls.Add(Me.CBX_KOR)
        Me.GBX_Manufacturer.Controls.Add(Me.CBX_DEFAULT)
        Me.GBX_Manufacturer.Location = New System.Drawing.Point(15, 72)
        Me.GBX_Manufacturer.Name = "GBX_Manufacturer"
        Me.GBX_Manufacturer.Size = New System.Drawing.Size(356, 47)
        Me.GBX_Manufacturer.TabIndex = 0
        Me.GBX_Manufacturer.TabStop = False
        Me.GBX_Manufacturer.Text = "제조사"
        '
        'CBX_CHN
        '
        Me.CBX_CHN.AutoSize = True
        Me.CBX_CHN.Location = New System.Drawing.Point(239, 20)
        Me.CBX_CHN.Name = "CBX_CHN"
        Me.CBX_CHN.Size = New System.Drawing.Size(49, 16)
        Me.CBX_CHN.TabIndex = 0
        Me.CBX_CHN.Text = "CHN"
        Me.CBX_CHN.UseVisualStyleBackColor = True
        '
        'CBX_KOR
        '
        Me.CBX_KOR.AutoSize = True
        Me.CBX_KOR.Location = New System.Drawing.Point(128, 20)
        Me.CBX_KOR.Name = "CBX_KOR"
        Me.CBX_KOR.Size = New System.Drawing.Size(48, 16)
        Me.CBX_KOR.TabIndex = 0
        Me.CBX_KOR.Text = "KOR"
        Me.CBX_KOR.UseVisualStyleBackColor = True
        '
        'CBX_DEFAULT
        '
        Me.CBX_DEFAULT.AutoSize = True
        Me.CBX_DEFAULT.Checked = True
        Me.CBX_DEFAULT.Location = New System.Drawing.Point(15, 20)
        Me.CBX_DEFAULT.Name = "CBX_DEFAULT"
        Me.CBX_DEFAULT.Size = New System.Drawing.Size(61, 16)
        Me.CBX_DEFAULT.TabIndex = 0
        Me.CBX_DEFAULT.TabStop = True
        Me.CBX_DEFAULT.Text = "Default"
        Me.CBX_DEFAULT.UseVisualStyleBackColor = True
        '
        'Coupling_Create
        '
        Me.Coupling_Create.Location = New System.Drawing.Point(233, 178)
        Me.Coupling_Create.Name = "Coupling_Create"
        Me.Coupling_Create.Size = New System.Drawing.Size(138, 31)
        Me.Coupling_Create.TabIndex = 1
        Me.Coupling_Create.Text = "실적 형상 검색"
        Me.Coupling_Create.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Dummy_Loading_Check)
        Me.GroupBox2.Location = New System.Drawing.Point(15, 125)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(356, 47)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'Dummy_Loading_Check
        '
        Me.Dummy_Loading_Check.AutoSize = True
        Me.Dummy_Loading_Check.Checked = True
        Me.Dummy_Loading_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Dummy_Loading_Check.Location = New System.Drawing.Point(15, 20)
        Me.Dummy_Loading_Check.Name = "Dummy_Loading_Check"
        Me.Dummy_Loading_Check.Size = New System.Drawing.Size(218, 16)
        Me.Dummy_Loading_Check.TabIndex = 1
        Me.Dummy_Loading_Check.Text = "Dummy 적용(Enclosure, Oil Tank)"
        Me.Dummy_Loading_Check.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.ProgressBar1)
        Me.GroupBox5.Controls.Add(Me.Message_Label)
        Me.GroupBox5.Controls.Add(Me.cancel_button)
        Me.GroupBox5.Controls.Add(Me.Coupling_Clash_Check)
        Me.GroupBox5.Controls.Add(Me.DataGridView1)
        Me.GroupBox5.Controls.Add(Me.GroupBox3)
        Me.GroupBox5.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(399, 512)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 480)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(381, 23)
        Me.ProgressBar1.TabIndex = 11
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(13, 445)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(288, 28)
        Me.Message_Label.TabIndex = 5
        Me.Message_Label.Text = "-"
        '
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(310, 445)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(80, 30)
        Me.cancel_button.TabIndex = 1
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Coupling_Clash_Check
        '
        Me.Coupling_Clash_Check.Location = New System.Drawing.Point(252, 411)
        Me.Coupling_Clash_Check.Name = "Coupling_Clash_Check"
        Me.Coupling_Clash_Check.Size = New System.Drawing.Size(138, 30)
        Me.Coupling_Clash_Check.TabIndex = 1
        Me.Coupling_Clash_Check.Text = "간섭 Check"
        Me.Coupling_Clash_Check.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(9, 239)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(381, 167)
        Me.DataGridView1.TabIndex = 4
        '
        'Col1
        '
        Me.Col1.HeaderText = ""
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        '
        'Col2
        '
        Me.Col2.HeaderText = ""
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        '
        'Col3
        '
        Me.Col3.HeaderText = ""
        Me.Col3.Name = "Col3"
        Me.Col3.ReadOnly = True
        '
        'Col4
        '
        Me.Col4.HeaderText = ""
        Me.Col4.Name = "Col4"
        Me.Col4.ReadOnly = True
        '
        'Col5
        '
        Me.Col5.HeaderText = ""
        Me.Col5.Name = "Col5"
        Me.Col5.ReadOnly = True
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(418, 57)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(260, 88)
        Me.Data_Path_List.TabIndex = 1
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(418, 35)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(51, 16)
        Me.VVIP_Check.TabIndex = 2
        Me.VVIP_Check.Text = "VVIP"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'B_Coupling_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(694, 551)
        Me.Controls.Add(Me.VVIP_Check)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Controls.Add(Me.GroupBox5)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "B_Coupling_Create"
        Me.Text = "B_Coupling_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GBX_Manufacturer.ResumeLayout(False)
        Me.GBX_Manufacturer.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Coupling_Create As Button
    Public WithEvents DataGridView1 As DataGridView
    Public WithEvents Col1 As DataGridViewTextBoxColumn
    Public WithEvents Col2 As DataGridViewTextBoxColumn
    Public WithEvents Col3 As DataGridViewTextBoxColumn
    Public WithEvents Col4 As DataGridViewTextBoxColumn
    Public WithEvents Col5 As DataGridViewTextBoxColumn
    Friend WithEvents cancel_button As Button
    Friend WithEvents Coupling_Clash_Check As Button
    Public WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents Check2 As System.Windows.Forms.CheckBox
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
    Public WithEvents Dummy_Loading_Check As System.Windows.Forms.CheckBox
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Friend WithEvents GBX_Manufacturer As System.Windows.Forms.GroupBox
    Friend WithEvents CBX_CHN As System.Windows.Forms.RadioButton
    Friend WithEvents CBX_KOR As System.Windows.Forms.RadioButton
    Friend WithEvents CBX_DEFAULT As System.Windows.Forms.RadioButton
End Class
