<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CA_Coupling_Cover_Template_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CA_Coupling_Cover_Template_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.UI_STAGE_COMBO = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Template_Part_Number_Text = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.Coupling_Cover_Result_Part_Create = New System.Windows.Forms.Button()
        Me.Coupling_Cover_Drawing_Create = New System.Windows.Forms.Button()
        Me.Coupling_Cover_Clash_Check = New System.Windows.Forms.Button()
        Me.Coupling_Cover_Template_Create = New System.Windows.Forms.Button()
        Me.Coupling_Cover_Parameter_Control = New System.Windows.Forms.Button()
        Me.Coupling_Cover_Parameter_Update = New System.Windows.Forms.Button()
        Me.LENG_COVER_HOLE_SPACING_Text = New System.Windows.Forms.TextBox()
        Me.DIA_COVER_HOLE_Text = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.HOLE_DIA_MAX_Label = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Auto_Clash_Check = New System.Windows.Forms.CheckBox()
        Me.UI_PUNCH_HOLE = New System.Windows.Forms.CheckBox()
        Me.UI_COVER_EXT = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(295, 208)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.VVIP_Check)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 141)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(269, 55)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "VVIP"
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(16, 23)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 0
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.UI_STAGE_COMBO)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 80)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(269, 55)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "STAGE"
        '
        'UI_STAGE_COMBO
        '
        Me.UI_STAGE_COMBO.FormattingEnabled = True
        Me.UI_STAGE_COMBO.Location = New System.Drawing.Point(16, 20)
        Me.UI_STAGE_COMBO.Name = "UI_STAGE_COMBO"
        Me.UI_STAGE_COMBO.Size = New System.Drawing.Size(237, 20)
        Me.UI_STAGE_COMBO.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Template_Part_Number_Text)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 19)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(269, 55)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Coupling Cover Part No."
        '
        'Template_Part_Number_Text
        '
        Me.Template_Part_Number_Text.Location = New System.Drawing.Point(16, 20)
        Me.Template_Part_Number_Text.Name = "Template_Part_Number_Text"
        Me.Template_Part_Number_Text.Size = New System.Drawing.Size(237, 21)
        Me.Template_Part_Number_Text.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.ProgressBar1)
        Me.GroupBox5.Controls.Add(Me.Message_Label)
        Me.GroupBox5.Controls.Add(Me.Coupling_Cover_Result_Part_Create)
        Me.GroupBox5.Controls.Add(Me.Coupling_Cover_Drawing_Create)
        Me.GroupBox5.Controls.Add(Me.Coupling_Cover_Clash_Check)
        Me.GroupBox5.Controls.Add(Me.Coupling_Cover_Template_Create)
        Me.GroupBox5.Controls.Add(Me.Coupling_Cover_Parameter_Control)
        Me.GroupBox5.Location = New System.Drawing.Point(0, 207)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(295, 230)
        Me.GroupBox5.TabIndex = 1
        Me.GroupBox5.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 203)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(269, 15)
        Me.ProgressBar1.TabIndex = 2
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(10, 177)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(269, 23)
        Me.Message_Label.TabIndex = 1
        Me.Message_Label.Text = "-"
        '
        'Coupling_Cover_Result_Part_Create
        '
        Me.Coupling_Cover_Result_Part_Create.Location = New System.Drawing.Point(160, 140)
        Me.Coupling_Cover_Result_Part_Create.Name = "Coupling_Cover_Result_Part_Create"
        Me.Coupling_Cover_Result_Part_Create.Size = New System.Drawing.Size(121, 34)
        Me.Coupling_Cover_Result_Part_Create.TabIndex = 0
        Me.Coupling_Cover_Result_Part_Create.Text = "3. PART 확정"
        Me.Coupling_Cover_Result_Part_Create.UseVisualStyleBackColor = True
        '
        'Coupling_Cover_Drawing_Create
        '
        Me.Coupling_Cover_Drawing_Create.Location = New System.Drawing.Point(160, 100)
        Me.Coupling_Cover_Drawing_Create.Name = "Coupling_Cover_Drawing_Create"
        Me.Coupling_Cover_Drawing_Create.Size = New System.Drawing.Size(121, 34)
        Me.Coupling_Cover_Drawing_Create.TabIndex = 0
        Me.Coupling_Cover_Drawing_Create.Text = "2. 도면 생성"
        Me.Coupling_Cover_Drawing_Create.UseVisualStyleBackColor = True
        '
        'Coupling_Cover_Clash_Check
        '
        Me.Coupling_Cover_Clash_Check.Location = New System.Drawing.Point(160, 60)
        Me.Coupling_Cover_Clash_Check.Name = "Coupling_Cover_Clash_Check"
        Me.Coupling_Cover_Clash_Check.Size = New System.Drawing.Size(121, 34)
        Me.Coupling_Cover_Clash_Check.TabIndex = 0
        Me.Coupling_Cover_Clash_Check.Text = "간섭 Check"
        Me.Coupling_Cover_Clash_Check.UseVisualStyleBackColor = True
        '
        'Coupling_Cover_Template_Create
        '
        Me.Coupling_Cover_Template_Create.Location = New System.Drawing.Point(160, 20)
        Me.Coupling_Cover_Template_Create.Name = "Coupling_Cover_Template_Create"
        Me.Coupling_Cover_Template_Create.Size = New System.Drawing.Size(121, 34)
        Me.Coupling_Cover_Template_Create.TabIndex = 0
        Me.Coupling_Cover_Template_Create.Text = "1. Template 실행"
        Me.Coupling_Cover_Template_Create.UseVisualStyleBackColor = True
        '
        'Coupling_Cover_Parameter_Control
        '
        Me.Coupling_Cover_Parameter_Control.Location = New System.Drawing.Point(12, 20)
        Me.Coupling_Cover_Parameter_Control.Name = "Coupling_Cover_Parameter_Control"
        Me.Coupling_Cover_Parameter_Control.Size = New System.Drawing.Size(121, 34)
        Me.Coupling_Cover_Parameter_Control.TabIndex = 0
        Me.Coupling_Cover_Parameter_Control.Text = "Parameter Control"
        Me.Coupling_Cover_Parameter_Control.UseVisualStyleBackColor = True
        '
        'Coupling_Cover_Parameter_Update
        '
        Me.Coupling_Cover_Parameter_Update.Location = New System.Drawing.Point(387, 197)
        Me.Coupling_Cover_Parameter_Update.Name = "Coupling_Cover_Parameter_Update"
        Me.Coupling_Cover_Parameter_Update.Size = New System.Drawing.Size(121, 34)
        Me.Coupling_Cover_Parameter_Update.TabIndex = 30
        Me.Coupling_Cover_Parameter_Update.Text = "Update"
        Me.Coupling_Cover_Parameter_Update.UseVisualStyleBackColor = True
        '
        'LENG_COVER_HOLE_SPACING_Text
        '
        Me.LENG_COVER_HOLE_SPACING_Text.Location = New System.Drawing.Point(411, 132)
        Me.LENG_COVER_HOLE_SPACING_Text.Name = "LENG_COVER_HOLE_SPACING_Text"
        Me.LENG_COVER_HOLE_SPACING_Text.Size = New System.Drawing.Size(64, 21)
        Me.LENG_COVER_HOLE_SPACING_Text.TabIndex = 28
        Me.LENG_COVER_HOLE_SPACING_Text.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DIA_COVER_HOLE_Text
        '
        Me.DIA_COVER_HOLE_Text.Location = New System.Drawing.Point(411, 69)
        Me.DIA_COVER_HOLE_Text.Name = "DIA_COVER_HOLE_Text"
        Me.DIA_COVER_HOLE_Text.Size = New System.Drawing.Size(64, 21)
        Me.DIA_COVER_HOLE_Text.TabIndex = 29
        Me.DIA_COVER_HOLE_Text.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(481, 135)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(27, 12)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "mm"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(312, 137)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 12)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "HOLE Spacing :"
        '
        'HOLE_DIA_MAX_Label
        '
        Me.HOLE_DIA_MAX_Label.Location = New System.Drawing.Point(417, 102)
        Me.HOLE_DIA_MAX_Label.Name = "HOLE_DIA_MAX_Label"
        Me.HOLE_DIA_MAX_Label.Size = New System.Drawing.Size(58, 14)
        Me.HOLE_DIA_MAX_Label.TabIndex = 23
        Me.HOLE_DIA_MAX_Label.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(481, 102)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(27, 12)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "mm"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(312, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 12)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "HOLE DIA MAX :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(481, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(27, 12)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "mm"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(312, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 12)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "HOLE DIA :"
        '
        'Auto_Clash_Check
        '
        Me.Auto_Clash_Check.AutoSize = True
        Me.Auto_Clash_Check.Checked = True
        Me.Auto_Clash_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Auto_Clash_Check.Location = New System.Drawing.Point(314, 327)
        Me.Auto_Clash_Check.Name = "Auto_Clash_Check"
        Me.Auto_Clash_Check.Size = New System.Drawing.Size(128, 16)
        Me.Auto_Clash_Check.TabIndex = 18
        Me.Auto_Clash_Check.Text = "간섭체크 자동 확인"
        Me.Auto_Clash_Check.UseVisualStyleBackColor = True
        '
        'UI_PUNCH_HOLE
        '
        Me.UI_PUNCH_HOLE.AutoSize = True
        Me.UI_PUNCH_HOLE.Checked = True
        Me.UI_PUNCH_HOLE.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_PUNCH_HOLE.Location = New System.Drawing.Point(314, 42)
        Me.UI_PUNCH_HOLE.Name = "UI_PUNCH_HOLE"
        Me.UI_PUNCH_HOLE.Size = New System.Drawing.Size(76, 16)
        Me.UI_PUNCH_HOLE.TabIndex = 19
        Me.UI_PUNCH_HOLE.Text = "타공 생성"
        Me.UI_PUNCH_HOLE.UseVisualStyleBackColor = True
        '
        'UI_COVER_EXT
        '
        Me.UI_COVER_EXT.AutoSize = True
        Me.UI_COVER_EXT.Checked = True
        Me.UI_COVER_EXT.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_COVER_EXT.Location = New System.Drawing.Point(314, 12)
        Me.UI_COVER_EXT.Name = "UI_COVER_EXT"
        Me.UI_COVER_EXT.Size = New System.Drawing.Size(138, 16)
        Me.UI_COVER_EXT.TabIndex = 20
        Me.UI_COVER_EXT.Text = "UPPER COVER 연장"
        Me.UI_COVER_EXT.UseVisualStyleBackColor = True
        '
        'CA_Coupling_Cover_Template_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(525, 535)
        Me.Controls.Add(Me.Coupling_Cover_Parameter_Update)
        Me.Controls.Add(Me.LENG_COVER_HOLE_SPACING_Text)
        Me.Controls.Add(Me.DIA_COVER_HOLE_Text)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.HOLE_DIA_MAX_Label)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Auto_Clash_Check)
        Me.Controls.Add(Me.UI_PUNCH_HOLE)
        Me.Controls.Add(Me.UI_COVER_EXT)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CA_Coupling_Cover_Template_Create"
        Me.Text = "CA_Coupling_Cover_Template_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents UI_STAGE_COMBO As System.Windows.Forms.ComboBox
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents Template_Part_Number_Text As System.Windows.Forms.TextBox
    Public WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents Coupling_Cover_Result_Part_Create As System.Windows.Forms.Button
    Public WithEvents Coupling_Cover_Drawing_Create As System.Windows.Forms.Button
    Public WithEvents Coupling_Cover_Clash_Check As System.Windows.Forms.Button
    Public WithEvents Coupling_Cover_Template_Create As System.Windows.Forms.Button
    Public WithEvents Coupling_Cover_Parameter_Control As System.Windows.Forms.Button
    Public WithEvents Coupling_Cover_Parameter_Update As System.Windows.Forms.Button
    Public WithEvents LENG_COVER_HOLE_SPACING_Text As System.Windows.Forms.TextBox
    Public WithEvents DIA_COVER_HOLE_Text As System.Windows.Forms.TextBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents HOLE_DIA_MAX_Label As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Auto_Clash_Check As System.Windows.Forms.CheckBox
    Public WithEvents UI_PUNCH_HOLE As System.Windows.Forms.CheckBox
    Public WithEvents UI_COVER_EXT As System.Windows.Forms.CheckBox
End Class
