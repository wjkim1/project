<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DA_Oil_Tank_Template_Create
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
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DA_Oil_Tank_Template_Create))
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Oil_Tank_Parameter_Control = New System.Windows.Forms.Button()
        Me.Oil_Tank_Result_Part_Create = New System.Windows.Forms.Button()
        Me.Oil_Tank_Drawing_Create = New System.Windows.Forms.Button()
        Me.Oil_Tank_Clash_Check = New System.Windows.Forms.Button()
        Me.Oil_Tank_Template_Create = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Oil_Tank_Measure_Y = New System.Windows.Forms.Label()
        Me.Oil_Tank_Measure_X = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FRM_TankType = New System.Windows.Forms.GroupBox()
        Me.Option_CHN = New System.Windows.Forms.RadioButton()
        Me.Option_KOR = New System.Windows.Forms.RadioButton()
        Me.Option_Default = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Template_Part_Number_Text = New System.Windows.Forms.TextBox()
        Me.Oil_Tank_Parameter_Update = New System.Windows.Forms.Button()
        Me.Auto_Clash_Check = New System.Windows.Forms.CheckBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.FRM_TankType.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.GroupBox5)
        Me.GroupBox6.Controls.Add(Me.GroupBox4)
        Me.GroupBox6.Location = New System.Drawing.Point(0, -7)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(423, 625)
        Me.GroupBox6.TabIndex = 5
        Me.GroupBox6.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.ProgressBar1)
        Me.GroupBox5.Controls.Add(Me.Message_Label)
        Me.GroupBox5.Controls.Add(Me.DataGridView1)
        Me.GroupBox5.Controls.Add(Me.Oil_Tank_Parameter_Control)
        Me.GroupBox5.Controls.Add(Me.Oil_Tank_Result_Part_Create)
        Me.GroupBox5.Controls.Add(Me.Oil_Tank_Drawing_Create)
        Me.GroupBox5.Controls.Add(Me.Oil_Tank_Clash_Check)
        Me.GroupBox5.Controls.Add(Me.Oil_Tank_Template_Create)
        Me.GroupBox5.Location = New System.Drawing.Point(6, 256)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(410, 357)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(8, 326)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(396, 17)
        Me.ProgressBar1.TabIndex = 8
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(6, 306)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(398, 17)
        Me.Message_Label.TabIndex = 7
        Me.Message_Label.Text = "-"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(6, 56)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(398, 136)
        Me.DataGridView1.TabIndex = 6
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
        'Oil_Tank_Parameter_Control
        '
        Me.Oil_Tank_Parameter_Control.Location = New System.Drawing.Point(6, 198)
        Me.Oil_Tank_Parameter_Control.Name = "Oil_Tank_Parameter_Control"
        Me.Oil_Tank_Parameter_Control.Size = New System.Drawing.Size(127, 30)
        Me.Oil_Tank_Parameter_Control.TabIndex = 0
        Me.Oil_Tank_Parameter_Control.Text = "Parameter Control"
        Me.Oil_Tank_Parameter_Control.UseVisualStyleBackColor = True
        '
        'Oil_Tank_Result_Part_Create
        '
        Me.Oil_Tank_Result_Part_Create.Location = New System.Drawing.Point(277, 270)
        Me.Oil_Tank_Result_Part_Create.Name = "Oil_Tank_Result_Part_Create"
        Me.Oil_Tank_Result_Part_Create.Size = New System.Drawing.Size(127, 30)
        Me.Oil_Tank_Result_Part_Create.TabIndex = 0
        Me.Oil_Tank_Result_Part_Create.Text = "3. PART 확정"
        Me.Oil_Tank_Result_Part_Create.UseVisualStyleBackColor = True
        '
        'Oil_Tank_Drawing_Create
        '
        Me.Oil_Tank_Drawing_Create.Location = New System.Drawing.Point(277, 234)
        Me.Oil_Tank_Drawing_Create.Name = "Oil_Tank_Drawing_Create"
        Me.Oil_Tank_Drawing_Create.Size = New System.Drawing.Size(127, 30)
        Me.Oil_Tank_Drawing_Create.TabIndex = 0
        Me.Oil_Tank_Drawing_Create.Text = "2. 도면 생성"
        Me.Oil_Tank_Drawing_Create.UseVisualStyleBackColor = True
        '
        'Oil_Tank_Clash_Check
        '
        Me.Oil_Tank_Clash_Check.Location = New System.Drawing.Point(277, 198)
        Me.Oil_Tank_Clash_Check.Name = "Oil_Tank_Clash_Check"
        Me.Oil_Tank_Clash_Check.Size = New System.Drawing.Size(127, 30)
        Me.Oil_Tank_Clash_Check.TabIndex = 0
        Me.Oil_Tank_Clash_Check.Text = "간섭 Check"
        Me.Oil_Tank_Clash_Check.UseVisualStyleBackColor = True
        '
        'Oil_Tank_Template_Create
        '
        Me.Oil_Tank_Template_Create.Location = New System.Drawing.Point(277, 20)
        Me.Oil_Tank_Template_Create.Name = "Oil_Tank_Template_Create"
        Me.Oil_Tank_Template_Create.Size = New System.Drawing.Size(127, 30)
        Me.Oil_Tank_Template_Create.TabIndex = 0
        Me.Oil_Tank_Template_Create.Text = "1. Template 실행"
        Me.Oil_Tank_Template_Create.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.GroupBox3)
        Me.GroupBox4.Controls.Add(Me.FRM_TankType)
        Me.GroupBox4.Controls.Add(Me.GroupBox1)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 14)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(410, 236)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Oil_Tank_Measure_Y)
        Me.GroupBox3.Controls.Add(Me.Oil_Tank_Measure_X)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Location = New System.Drawing.Point(15, 139)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(380, 85)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Body Weld 선정 기준"
        '
        'Oil_Tank_Measure_Y
        '
        Me.Oil_Tank_Measure_Y.Location = New System.Drawing.Point(286, 55)
        Me.Oil_Tank_Measure_Y.Name = "Oil_Tank_Measure_Y"
        Me.Oil_Tank_Measure_Y.Size = New System.Drawing.Size(56, 14)
        Me.Oil_Tank_Measure_Y.TabIndex = 1
        Me.Oil_Tank_Measure_Y.Text = "-"
        Me.Oil_Tank_Measure_Y.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Oil_Tank_Measure_X
        '
        Me.Oil_Tank_Measure_X.Location = New System.Drawing.Point(286, 31)
        Me.Oil_Tank_Measure_X.Name = "Oil_Tank_Measure_X"
        Me.Oil_Tank_Measure_X.Size = New System.Drawing.Size(56, 14)
        Me.Oil_Tank_Measure_X.TabIndex = 1
        Me.Oil_Tank_Measure_X.Text = "-"
        Me.Oil_Tank_Measure_X.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(32, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(248, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Bull Gear ~ Motor 좌측 발 자리 끝 (Y방향) :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Bull Gear ~ Motor 좌측 발 자리 끝 (X방향) :"
        '
        'FRM_TankType
        '
        Me.FRM_TankType.Controls.Add(Me.Option_CHN)
        Me.FRM_TankType.Controls.Add(Me.Option_KOR)
        Me.FRM_TankType.Controls.Add(Me.Option_Default)
        Me.FRM_TankType.Location = New System.Drawing.Point(15, 83)
        Me.FRM_TankType.Name = "FRM_TankType"
        Me.FRM_TankType.Size = New System.Drawing.Size(380, 50)
        Me.FRM_TankType.TabIndex = 4
        Me.FRM_TankType.TabStop = False
        Me.FRM_TankType.Text = "Oil Tank TYPE"
        '
        'Option_CHN
        '
        Me.Option_CHN.AutoSize = True
        Me.Option_CHN.Location = New System.Drawing.Point(311, 20)
        Me.Option_CHN.Name = "Option_CHN"
        Me.Option_CHN.Size = New System.Drawing.Size(49, 16)
        Me.Option_CHN.TabIndex = 0
        Me.Option_CHN.Text = "CHN"
        Me.Option_CHN.UseVisualStyleBackColor = True
        '
        'Option_KOR
        '
        Me.Option_KOR.AutoSize = True
        Me.Option_KOR.Location = New System.Drawing.Point(165, 20)
        Me.Option_KOR.Name = "Option_KOR"
        Me.Option_KOR.Size = New System.Drawing.Size(48, 16)
        Me.Option_KOR.TabIndex = 0
        Me.Option_KOR.Text = "KOR"
        Me.Option_KOR.UseVisualStyleBackColor = True
        '
        'Option_Default
        '
        Me.Option_Default.AutoSize = True
        Me.Option_Default.Checked = True
        Me.Option_Default.Location = New System.Drawing.Point(19, 20)
        Me.Option_Default.Name = "Option_Default"
        Me.Option_Default.Size = New System.Drawing.Size(61, 16)
        Me.Option_Default.TabIndex = 0
        Me.Option_Default.TabStop = True
        Me.Option_Default.Text = "Default"
        Me.Option_Default.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Template_Part_Number_Text)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 20)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(380, 57)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Oil Tank PART No"
        '
        'Template_Part_Number_Text
        '
        Me.Template_Part_Number_Text.Location = New System.Drawing.Point(19, 20)
        Me.Template_Part_Number_Text.Name = "Template_Part_Number_Text"
        Me.Template_Part_Number_Text.Size = New System.Drawing.Size(341, 21)
        Me.Template_Part_Number_Text.TabIndex = 0
        '
        'Oil_Tank_Parameter_Update
        '
        Me.Oil_Tank_Parameter_Update.Location = New System.Drawing.Point(488, 49)
        Me.Oil_Tank_Parameter_Update.Name = "Oil_Tank_Parameter_Update"
        Me.Oil_Tank_Parameter_Update.Size = New System.Drawing.Size(127, 30)
        Me.Oil_Tank_Parameter_Update.TabIndex = 9
        Me.Oil_Tank_Parameter_Update.Text = "UPDATE"
        Me.Oil_Tank_Parameter_Update.UseVisualStyleBackColor = True
        '
        'Auto_Clash_Check
        '
        Me.Auto_Clash_Check.AutoSize = True
        Me.Auto_Clash_Check.Checked = True
        Me.Auto_Clash_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Auto_Clash_Check.Location = New System.Drawing.Point(445, 146)
        Me.Auto_Clash_Check.Name = "Auto_Clash_Check"
        Me.Auto_Clash_Check.Size = New System.Drawing.Size(100, 16)
        Me.Auto_Clash_Check.TabIndex = 6
        Me.Auto_Clash_Check.Text = "간섭체크 자동"
        Me.Auto_Clash_Check.UseVisualStyleBackColor = True
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(445, 114)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 7
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Checked = True
        Me.Check1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check1.Location = New System.Drawing.Point(445, 27)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(139, 16)
        Me.Check1.TabIndex = 8
        Me.Check1.Text = "TRANSMITTER 적용"
        Me.Check1.UseVisualStyleBackColor = True
        '
        'DA_Oil_Tank_Template_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(637, 704)
        Me.Controls.Add(Me.Oil_Tank_Parameter_Update)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.Auto_Clash_Check)
        Me.Controls.Add(Me.Check1)
        Me.Controls.Add(Me.VVIP_Check)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DA_Oil_Tank_Template_Create"
        Me.Text = "DA_Oil_Tank_Template_Create"
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.FRM_TankType.ResumeLayout(False)
        Me.FRM_TankType.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents Oil_Tank_Parameter_Control As System.Windows.Forms.Button
    Public WithEvents Oil_Tank_Result_Part_Create As System.Windows.Forms.Button
    Public WithEvents Oil_Tank_Drawing_Create As System.Windows.Forms.Button
    Public WithEvents Oil_Tank_Clash_Check As System.Windows.Forms.Button
    Public WithEvents Oil_Tank_Template_Create As System.Windows.Forms.Button
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents Oil_Tank_Measure_Y As System.Windows.Forms.Label
    Public WithEvents Oil_Tank_Measure_X As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents FRM_TankType As System.Windows.Forms.GroupBox
    Public WithEvents Option_CHN As System.Windows.Forms.RadioButton
    Public WithEvents Option_KOR As System.Windows.Forms.RadioButton
    Public WithEvents Option_Default As System.Windows.Forms.RadioButton
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Template_Part_Number_Text As System.Windows.Forms.TextBox
    Public WithEvents Oil_Tank_Parameter_Update As System.Windows.Forms.Button
    Public WithEvents Auto_Clash_Check As System.Windows.Forms.CheckBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
End Class
