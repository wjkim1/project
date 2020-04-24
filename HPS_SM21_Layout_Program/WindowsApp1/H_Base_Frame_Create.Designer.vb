<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class H_Base_Frame_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(H_Base_Frame_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.UI_MotorStarter_NO = New System.Windows.Forms.RadioButton()
        Me.UI_MotorStarter_YES = New System.Windows.Forms.RadioButton()
        Me.Option_Default_MotorStarter = New System.Windows.Forms.RadioButton()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.Base_Frame_Create = New System.Windows.Forms.Button()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.Option_Over_8 = New System.Windows.Forms.RadioButton()
        Me.Option_Under_8 = New System.Windows.Forms.RadioButton()
        Me.Option_Motor_Default = New System.Windows.Forms.RadioButton()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Auto_Trap_Option2 = New System.Windows.Forms.RadioButton()
        Me.Auto_Trap_Option1 = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.UI_ENCLOSURE_TYPE_2 = New System.Windows.Forms.RadioButton()
        Me.UI_ENCLOSURE_TYPE = New System.Windows.Forms.RadioButton()
        Me.Option_Default = New System.Windows.Forms.RadioButton()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Base_Frame_Clash_Check = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.VVIP_Frame = New System.Windows.Forms.GroupBox()
        Me.UI_PPN18_D06 = New System.Windows.Forms.CheckBox()
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.VVIP_Frame.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.VVIP_Check)
        Me.GroupBox1.Controls.Add(Me.Base_Frame_Create)
        Me.GroupBox1.Controls.Add(Me.Frame4)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(374, 192)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Base Frame 실적 형상 검색"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.UI_MotorStarter_NO)
        Me.GroupBox2.Controls.Add(Me.UI_MotorStarter_YES)
        Me.GroupBox2.Controls.Add(Me.Option_Default_MotorStarter)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 82)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(362, 56)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Motor Starter 유무"
        '
        'UI_MotorStarter_NO
        '
        Me.UI_MotorStarter_NO.AutoSize = True
        Me.UI_MotorStarter_NO.Location = New System.Drawing.Point(258, 24)
        Me.UI_MotorStarter_NO.Name = "UI_MotorStarter_NO"
        Me.UI_MotorStarter_NO.Size = New System.Drawing.Size(39, 16)
        Me.UI_MotorStarter_NO.TabIndex = 1
        Me.UI_MotorStarter_NO.Text = "No"
        Me.UI_MotorStarter_NO.UseVisualStyleBackColor = True
        '
        'UI_MotorStarter_YES
        '
        Me.UI_MotorStarter_YES.AutoSize = True
        Me.UI_MotorStarter_YES.Location = New System.Drawing.Point(128, 24)
        Me.UI_MotorStarter_YES.Name = "UI_MotorStarter_YES"
        Me.UI_MotorStarter_YES.Size = New System.Drawing.Size(45, 16)
        Me.UI_MotorStarter_YES.TabIndex = 2
        Me.UI_MotorStarter_YES.Text = "Yes"
        Me.UI_MotorStarter_YES.UseVisualStyleBackColor = True
        '
        'Option_Default_MotorStarter
        '
        Me.Option_Default_MotorStarter.AutoSize = True
        Me.Option_Default_MotorStarter.Checked = True
        Me.Option_Default_MotorStarter.Location = New System.Drawing.Point(20, 24)
        Me.Option_Default_MotorStarter.Name = "Option_Default_MotorStarter"
        Me.Option_Default_MotorStarter.Size = New System.Drawing.Size(61, 16)
        Me.Option_Default_MotorStarter.TabIndex = 3
        Me.Option_Default_MotorStarter.TabStop = True
        Me.Option_Default_MotorStarter.Text = "Default"
        Me.Option_Default_MotorStarter.UseVisualStyleBackColor = True
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(25, 158)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 2
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'Base_Frame_Create
        '
        Me.Base_Frame_Create.Location = New System.Drawing.Point(218, 147)
        Me.Base_Frame_Create.Name = "Base_Frame_Create"
        Me.Base_Frame_Create.Size = New System.Drawing.Size(150, 37)
        Me.Base_Frame_Create.TabIndex = 1
        Me.Base_Frame_Create.Text = "실적 형상 검색"
        Me.Base_Frame_Create.UseVisualStyleBackColor = True
        '
        'Frame4
        '
        Me.Frame4.Controls.Add(Me.Option_Over_8)
        Me.Frame4.Controls.Add(Me.Option_Under_8)
        Me.Frame4.Controls.Add(Me.Option_Motor_Default)
        Me.Frame4.Location = New System.Drawing.Point(160, 0)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Size = New System.Drawing.Size(362, 56)
        Me.Frame4.TabIndex = 0
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Motor 무게"
        Me.Frame4.Visible = False
        '
        'Option_Over_8
        '
        Me.Option_Over_8.AutoSize = True
        Me.Option_Over_8.Location = New System.Drawing.Point(257, 25)
        Me.Option_Over_8.Name = "Option_Over_8"
        Me.Option_Over_8.Size = New System.Drawing.Size(92, 16)
        Me.Option_Over_8.TabIndex = 0
        Me.Option_Over_8.Text = "8,000kg 이상"
        Me.Option_Over_8.UseVisualStyleBackColor = True
        '
        'Option_Under_8
        '
        Me.Option_Under_8.AutoSize = True
        Me.Option_Under_8.Location = New System.Drawing.Point(127, 25)
        Me.Option_Under_8.Name = "Option_Under_8"
        Me.Option_Under_8.Size = New System.Drawing.Size(92, 16)
        Me.Option_Under_8.TabIndex = 0
        Me.Option_Under_8.Text = "8,000kg 미만"
        Me.Option_Under_8.UseVisualStyleBackColor = True
        '
        'Option_Motor_Default
        '
        Me.Option_Motor_Default.AutoSize = True
        Me.Option_Motor_Default.Checked = True
        Me.Option_Motor_Default.Location = New System.Drawing.Point(19, 25)
        Me.Option_Motor_Default.Name = "Option_Motor_Default"
        Me.Option_Motor_Default.Size = New System.Drawing.Size(61, 16)
        Me.Option_Motor_Default.TabIndex = 0
        Me.Option_Motor_Default.TabStop = True
        Me.Option_Motor_Default.Text = "Default"
        Me.Option_Motor_Default.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Auto_Trap_Option2)
        Me.GroupBox4.Controls.Add(Me.Auto_Trap_Option1)
        Me.GroupBox4.Location = New System.Drawing.Point(320, 158)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(362, 56)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Auto Trap 사양"
        Me.GroupBox4.UseCompatibleTextRendering = True
        Me.GroupBox4.Visible = False
        '
        'Auto_Trap_Option2
        '
        Me.Auto_Trap_Option2.AutoSize = True
        Me.Auto_Trap_Option2.Location = New System.Drawing.Point(127, 25)
        Me.Auto_Trap_Option2.Name = "Auto_Trap_Option2"
        Me.Auto_Trap_Option2.Size = New System.Drawing.Size(59, 16)
        Me.Auto_Trap_Option2.TabIndex = 0
        Me.Auto_Trap_Option2.Text = "미적용"
        Me.Auto_Trap_Option2.UseVisualStyleBackColor = True
        '
        'Auto_Trap_Option1
        '
        Me.Auto_Trap_Option1.AutoSize = True
        Me.Auto_Trap_Option1.Checked = True
        Me.Auto_Trap_Option1.Location = New System.Drawing.Point(19, 25)
        Me.Auto_Trap_Option1.Name = "Auto_Trap_Option1"
        Me.Auto_Trap_Option1.Size = New System.Drawing.Size(47, 16)
        Me.Auto_Trap_Option1.TabIndex = 0
        Me.Auto_Trap_Option1.TabStop = True
        Me.Auto_Trap_Option1.Text = "적용"
        Me.Auto_Trap_Option1.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.UI_ENCLOSURE_TYPE_2)
        Me.GroupBox3.Controls.Add(Me.UI_ENCLOSURE_TYPE)
        Me.GroupBox3.Controls.Add(Me.Option_Default)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 20)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(362, 56)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Enclosure 유무"
        '
        'UI_ENCLOSURE_TYPE_2
        '
        Me.UI_ENCLOSURE_TYPE_2.AutoSize = True
        Me.UI_ENCLOSURE_TYPE_2.Location = New System.Drawing.Point(257, 25)
        Me.UI_ENCLOSURE_TYPE_2.Name = "UI_ENCLOSURE_TYPE_2"
        Me.UI_ENCLOSURE_TYPE_2.Size = New System.Drawing.Size(39, 16)
        Me.UI_ENCLOSURE_TYPE_2.TabIndex = 0
        Me.UI_ENCLOSURE_TYPE_2.Text = "No"
        Me.UI_ENCLOSURE_TYPE_2.UseVisualStyleBackColor = True
        '
        'UI_ENCLOSURE_TYPE
        '
        Me.UI_ENCLOSURE_TYPE.AutoSize = True
        Me.UI_ENCLOSURE_TYPE.Location = New System.Drawing.Point(127, 25)
        Me.UI_ENCLOSURE_TYPE.Name = "UI_ENCLOSURE_TYPE"
        Me.UI_ENCLOSURE_TYPE.Size = New System.Drawing.Size(45, 16)
        Me.UI_ENCLOSURE_TYPE.TabIndex = 0
        Me.UI_ENCLOSURE_TYPE.Text = "Yes"
        Me.UI_ENCLOSURE_TYPE.UseVisualStyleBackColor = True
        '
        'Option_Default
        '
        Me.Option_Default.AutoSize = True
        Me.Option_Default.Checked = True
        Me.Option_Default.Location = New System.Drawing.Point(19, 25)
        Me.Option_Default.Name = "Option_Default"
        Me.Option_Default.Size = New System.Drawing.Size(61, 16)
        Me.Option_Default.TabIndex = 0
        Me.Option_Default.TabStop = True
        Me.Option_Default.Text = "Default"
        Me.Option_Default.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.ProgressBar1)
        Me.GroupBox6.Controls.Add(Me.Message_Label)
        Me.GroupBox6.Controls.Add(Me.cancel_button)
        Me.GroupBox6.Controls.Add(Me.Base_Frame_Clash_Check)
        Me.GroupBox6.Controls.Add(Me.DataGridView1)
        Me.GroupBox6.Location = New System.Drawing.Point(12, 210)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(374, 315)
        Me.GroupBox6.TabIndex = 1
        Me.GroupBox6.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 294)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(362, 12)
        Me.ProgressBar1.TabIndex = 12
        '
        'Message_Label
        '
        Me.Message_Label.AutoSize = True
        Me.Message_Label.Location = New System.Drawing.Point(6, 276)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(11, 12)
        Me.Message_Label.TabIndex = 11
        Me.Message_Label.Text = "-"
        '
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(218, 251)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(150, 37)
        Me.cancel_button.TabIndex = 7
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Base_Frame_Clash_Check
        '
        Me.Base_Frame_Clash_Check.Location = New System.Drawing.Point(218, 208)
        Me.Base_Frame_Clash_Check.Name = "Base_Frame_Clash_Check"
        Me.Base_Frame_Clash_Check.Size = New System.Drawing.Size(150, 37)
        Me.Base_Frame_Clash_Check.TabIndex = 7
        Me.Base_Frame_Clash_Check.Text = "간섭 Check"
        Me.Base_Frame_Clash_Check.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(6, 20)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(362, 182)
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
        'VVIP_Frame
        '
        Me.VVIP_Frame.Controls.Add(Me.UI_PPN18_D06)
        Me.VVIP_Frame.Location = New System.Drawing.Point(409, 12)
        Me.VVIP_Frame.Name = "VVIP_Frame"
        Me.VVIP_Frame.Size = New System.Drawing.Size(197, 513)
        Me.VVIP_Frame.TabIndex = 2
        Me.VVIP_Frame.TabStop = False
        Me.VVIP_Frame.Text = "VVIP OPTION"
        '
        'UI_PPN18_D06
        '
        Me.UI_PPN18_D06.AutoSize = True
        Me.UI_PPN18_D06.Location = New System.Drawing.Point(23, 45)
        Me.UI_PPN18_D06.Name = "UI_PPN18_D06"
        Me.UI_PPN18_D06.Size = New System.Drawing.Size(159, 16)
        Me.UI_PPN18_D06.TabIndex = 3
        Me.UI_PPN18_D06.Text = "CONTROL_PANEL_외장"
        Me.UI_PPN18_D06.UseVisualStyleBackColor = True
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(12, 609)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(374, 40)
        Me.Data_Path_List.TabIndex = 3
        '
        'H_Base_Frame_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(616, 528)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Controls.Add(Me.VVIP_Frame)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "H_Base_Frame_Create"
        Me.Text = "H_Base_Frame_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.VVIP_Frame.ResumeLayout(False)
        Me.VVIP_Frame.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Base_Frame_Create As System.Windows.Forms.Button
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents Auto_Trap_Option2 As System.Windows.Forms.RadioButton
    Public WithEvents Auto_Trap_Option1 As System.Windows.Forms.RadioButton
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents UI_ENCLOSURE_TYPE_2 As System.Windows.Forms.RadioButton
    Public WithEvents UI_ENCLOSURE_TYPE As System.Windows.Forms.RadioButton
    Public WithEvents Option_Default As System.Windows.Forms.RadioButton
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents Option_Over_8 As System.Windows.Forms.RadioButton
    Public WithEvents Option_Under_8 As System.Windows.Forms.RadioButton
    Public WithEvents Option_Motor_Default As System.Windows.Forms.RadioButton
    Public WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents cancel_button As System.Windows.Forms.Button
    Public WithEvents Base_Frame_Clash_Check As System.Windows.Forms.Button
    Public WithEvents VVIP_Frame As System.Windows.Forms.GroupBox
    Public WithEvents UI_PPN18_D06 As System.Windows.Forms.CheckBox
    Public WithEvents GroupBox2 As GroupBox
    Public WithEvents UI_MotorStarter_NO As RadioButton
    Public WithEvents UI_MotorStarter_YES As RadioButton
    Public WithEvents Option_Default_MotorStarter As RadioButton
End Class
