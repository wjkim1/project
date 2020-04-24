<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EA_Oil_Piping_Template_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EA_Oil_Piping_Template_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Oil_Piping_Part_Convert = New System.Windows.Forms.Button()
        Me.Oil_Piping_Drawing_Create = New System.Windows.Forms.Button()
        Me.Oil_Piping_Clash_Check = New System.Windows.Forms.Button()
        Me.Oil_Piping_Parameter_Control = New System.Windows.Forms.Button()
        Me.Oil_Piping_Template_Create = New System.Windows.Forms.Button()
        Me.Template_Part_Number_Text = New System.Windows.Forms.TextBox()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.FILTER_MOVE_V_LENGTH_Update = New System.Windows.Forms.Button()
        Me.FILTER_SIDE_DIRECTION_MOVE_Update = New System.Windows.Forms.Button()
        Me.FILTER_MOVE_H_LENGTH_Update = New System.Windows.Forms.Button()
        Me.filter_move_v_length_text = New System.Windows.Forms.TextBox()
        Me.filter_side_direction_move_text = New System.Windows.Forms.TextBox()
        Me.filter_move_h_length_text = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OilPipe_Delay_Text = New System.Windows.Forms.TextBox()
        Me.Oil_Piping_Parameter_Update = New System.Windows.Forms.Button()
        Me.Auto_Clash_Check = New System.Windows.Forms.CheckBox()
        Me.UI_SIGHT_GLASS = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Oil_Piping_Part_Convert)
        Me.GroupBox1.Controls.Add(Me.Oil_Piping_Drawing_Create)
        Me.GroupBox1.Controls.Add(Me.Oil_Piping_Clash_Check)
        Me.GroupBox1.Controls.Add(Me.Oil_Piping_Parameter_Control)
        Me.GroupBox1.Controls.Add(Me.Oil_Piping_Template_Create)
        Me.GroupBox1.Controls.Add(Me.Template_Part_Number_Text)
        Me.GroupBox1.Controls.Add(Me.Message_Label)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(346, 423)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 393)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(333, 17)
        Me.ProgressBar1.TabIndex = 4
        '
        'Oil_Piping_Part_Convert
        '
        Me.Oil_Piping_Part_Convert.Location = New System.Drawing.Point(190, 351)
        Me.Oil_Piping_Part_Convert.Name = "Oil_Piping_Part_Convert"
        Me.Oil_Piping_Part_Convert.Size = New System.Drawing.Size(149, 36)
        Me.Oil_Piping_Part_Convert.TabIndex = 3
        Me.Oil_Piping_Part_Convert.Text = "3. PART 확정"
        Me.Oil_Piping_Part_Convert.UseVisualStyleBackColor = True
        '
        'Oil_Piping_Drawing_Create
        '
        Me.Oil_Piping_Drawing_Create.Location = New System.Drawing.Point(190, 309)
        Me.Oil_Piping_Drawing_Create.Name = "Oil_Piping_Drawing_Create"
        Me.Oil_Piping_Drawing_Create.Size = New System.Drawing.Size(149, 36)
        Me.Oil_Piping_Drawing_Create.TabIndex = 3
        Me.Oil_Piping_Drawing_Create.Text = "2. 도면 생성"
        Me.Oil_Piping_Drawing_Create.UseVisualStyleBackColor = True
        '
        'Oil_Piping_Clash_Check
        '
        Me.Oil_Piping_Clash_Check.Location = New System.Drawing.Point(190, 267)
        Me.Oil_Piping_Clash_Check.Name = "Oil_Piping_Clash_Check"
        Me.Oil_Piping_Clash_Check.Size = New System.Drawing.Size(149, 36)
        Me.Oil_Piping_Clash_Check.TabIndex = 3
        Me.Oil_Piping_Clash_Check.Text = "간섭 Check"
        Me.Oil_Piping_Clash_Check.UseVisualStyleBackColor = True
        '
        'Oil_Piping_Parameter_Control
        '
        Me.Oil_Piping_Parameter_Control.Location = New System.Drawing.Point(6, 267)
        Me.Oil_Piping_Parameter_Control.Name = "Oil_Piping_Parameter_Control"
        Me.Oil_Piping_Parameter_Control.Size = New System.Drawing.Size(149, 36)
        Me.Oil_Piping_Parameter_Control.TabIndex = 3
        Me.Oil_Piping_Parameter_Control.Text = "Parameter Control"
        Me.Oil_Piping_Parameter_Control.UseVisualStyleBackColor = True
        '
        'Oil_Piping_Template_Create
        '
        Me.Oil_Piping_Template_Create.Location = New System.Drawing.Point(190, 108)
        Me.Oil_Piping_Template_Create.Name = "Oil_Piping_Template_Create"
        Me.Oil_Piping_Template_Create.Size = New System.Drawing.Size(149, 36)
        Me.Oil_Piping_Template_Create.TabIndex = 3
        Me.Oil_Piping_Template_Create.Text = "1. Template 실행"
        Me.Oil_Piping_Template_Create.UseVisualStyleBackColor = True
        '
        'Template_Part_Number_Text
        '
        Me.Template_Part_Number_Text.Location = New System.Drawing.Point(145, 19)
        Me.Template_Part_Number_Text.Name = "Template_Part_Number_Text"
        Me.Template_Part_Number_Text.Size = New System.Drawing.Size(185, 21)
        Me.Template_Part_Number_Text.TabIndex = 2
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(6, 363)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(178, 15)
        Me.Message_Label.TabIndex = 1
        Me.Message_Label.Text = "-"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(10, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(129, 12)
        Me.Label11.TabIndex = 1
        Me.Label11.Text = "Oil Piping PART NO. :"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.VVIP_Check)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 55)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(333, 47)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "VVIP"
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(25, 20)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 0
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.FILTER_MOVE_V_LENGTH_Update)
        Me.GroupBox2.Controls.Add(Me.FILTER_SIDE_DIRECTION_MOVE_Update)
        Me.GroupBox2.Controls.Add(Me.FILTER_MOVE_H_LENGTH_Update)
        Me.GroupBox2.Controls.Add(Me.filter_move_v_length_text)
        Me.GroupBox2.Controls.Add(Me.filter_side_direction_move_text)
        Me.GroupBox2.Controls.Add(Me.filter_move_h_length_text)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 150)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(333, 111)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "FILTER 이동"
        '
        'FILTER_MOVE_V_LENGTH_Update
        '
        Me.FILTER_MOVE_V_LENGTH_Update.Location = New System.Drawing.Point(246, 74)
        Me.FILTER_MOVE_V_LENGTH_Update.Name = "FILTER_MOVE_V_LENGTH_Update"
        Me.FILTER_MOVE_V_LENGTH_Update.Size = New System.Drawing.Size(75, 23)
        Me.FILTER_MOVE_V_LENGTH_Update.TabIndex = 2
        Me.FILTER_MOVE_V_LENGTH_Update.Text = "UPDATE"
        Me.FILTER_MOVE_V_LENGTH_Update.UseVisualStyleBackColor = True
        '
        'FILTER_SIDE_DIRECTION_MOVE_Update
        '
        Me.FILTER_SIDE_DIRECTION_MOVE_Update.Location = New System.Drawing.Point(246, 47)
        Me.FILTER_SIDE_DIRECTION_MOVE_Update.Name = "FILTER_SIDE_DIRECTION_MOVE_Update"
        Me.FILTER_SIDE_DIRECTION_MOVE_Update.Size = New System.Drawing.Size(75, 23)
        Me.FILTER_SIDE_DIRECTION_MOVE_Update.TabIndex = 2
        Me.FILTER_SIDE_DIRECTION_MOVE_Update.Text = "UPDATE"
        Me.FILTER_SIDE_DIRECTION_MOVE_Update.UseVisualStyleBackColor = True
        '
        'FILTER_MOVE_H_LENGTH_Update
        '
        Me.FILTER_MOVE_H_LENGTH_Update.Location = New System.Drawing.Point(246, 20)
        Me.FILTER_MOVE_H_LENGTH_Update.Name = "FILTER_MOVE_H_LENGTH_Update"
        Me.FILTER_MOVE_H_LENGTH_Update.Size = New System.Drawing.Size(75, 23)
        Me.FILTER_MOVE_H_LENGTH_Update.TabIndex = 2
        Me.FILTER_MOVE_H_LENGTH_Update.Text = "UPDATE"
        Me.FILTER_MOVE_H_LENGTH_Update.UseVisualStyleBackColor = True
        '
        'filter_move_v_length_text
        '
        Me.filter_move_v_length_text.Location = New System.Drawing.Point(180, 75)
        Me.filter_move_v_length_text.Name = "filter_move_v_length_text"
        Me.filter_move_v_length_text.Size = New System.Drawing.Size(60, 21)
        Me.filter_move_v_length_text.TabIndex = 1
        Me.filter_move_v_length_text.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'filter_side_direction_move_text
        '
        Me.filter_side_direction_move_text.Location = New System.Drawing.Point(180, 48)
        Me.filter_side_direction_move_text.Name = "filter_side_direction_move_text"
        Me.filter_side_direction_move_text.Size = New System.Drawing.Size(60, 21)
        Me.filter_side_direction_move_text.TabIndex = 1
        Me.filter_side_direction_move_text.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'filter_move_h_length_text
        '
        Me.filter_move_h_length_text.Location = New System.Drawing.Point(180, 21)
        Me.filter_move_h_length_text.Name = "filter_move_h_length_text"
        Me.filter_move_h_length_text.Size = New System.Drawing.Size(60, 21)
        Me.filter_move_h_length_text.TabIndex = 1
        Me.filter_move_h_length_text.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(119, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 12)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "상하(Z) :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(119, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "좌우(Y) :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(119, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "전후(X) :"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.OilPipe_Delay_Text)
        Me.GroupBox4.Controls.Add(Me.Oil_Piping_Parameter_Update)
        Me.GroupBox4.Controls.Add(Me.Auto_Clash_Check)
        Me.GroupBox4.Controls.Add(Me.UI_SIGHT_GLASS)
        Me.GroupBox4.Location = New System.Drawing.Point(361, 12)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(156, 404)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "YYIP OPTION"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 194)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 36)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Oil Piping" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "템플릿 생성 Delay" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "1000=1초"
        '
        'OilPipe_Delay_Text
        '
        Me.OilPipe_Delay_Text.Location = New System.Drawing.Point(6, 235)
        Me.OilPipe_Delay_Text.Name = "OilPipe_Delay_Text"
        Me.OilPipe_Delay_Text.Size = New System.Drawing.Size(100, 21)
        Me.OilPipe_Delay_Text.TabIndex = 4
        Me.OilPipe_Delay_Text.Text = "0"
        '
        'Oil_Piping_Parameter_Update
        '
        Me.Oil_Piping_Parameter_Update.Location = New System.Drawing.Point(75, 52)
        Me.Oil_Piping_Parameter_Update.Name = "Oil_Piping_Parameter_Update"
        Me.Oil_Piping_Parameter_Update.Size = New System.Drawing.Size(75, 23)
        Me.Oil_Piping_Parameter_Update.TabIndex = 3
        Me.Oil_Piping_Parameter_Update.Text = "UPDATE"
        Me.Oil_Piping_Parameter_Update.UseVisualStyleBackColor = True
        '
        'Auto_Clash_Check
        '
        Me.Auto_Clash_Check.AutoSize = True
        Me.Auto_Clash_Check.Checked = True
        Me.Auto_Clash_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Auto_Clash_Check.Location = New System.Drawing.Point(6, 140)
        Me.Auto_Clash_Check.Name = "Auto_Clash_Check"
        Me.Auto_Clash_Check.Size = New System.Drawing.Size(128, 16)
        Me.Auto_Clash_Check.TabIndex = 1
        Me.Auto_Clash_Check.Text = "간섭체크 자동 확인"
        Me.Auto_Clash_Check.UseVisualStyleBackColor = True
        '
        'UI_SIGHT_GLASS
        '
        Me.UI_SIGHT_GLASS.AutoSize = True
        Me.UI_SIGHT_GLASS.Checked = True
        Me.UI_SIGHT_GLASS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_SIGHT_GLASS.Location = New System.Drawing.Point(6, 29)
        Me.UI_SIGHT_GLASS.Name = "UI_SIGHT_GLASS"
        Me.UI_SIGHT_GLASS.Size = New System.Drawing.Size(132, 16)
        Me.UI_SIGHT_GLASS.TabIndex = 1
        Me.UI_SIGHT_GLASS.Text = "SIGHT GLASS 적용"
        Me.UI_SIGHT_GLASS.UseVisualStyleBackColor = True
        '
        'EA_Oil_Piping_Template_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(668, 481)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EA_Oil_Piping_Template_Create"
        Me.Text = "EA_Oil_Piping_Template_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Oil_Piping_Part_Convert As System.Windows.Forms.Button
    Public WithEvents Oil_Piping_Drawing_Create As System.Windows.Forms.Button
    Public WithEvents Oil_Piping_Clash_Check As System.Windows.Forms.Button
    Public WithEvents Oil_Piping_Parameter_Control As System.Windows.Forms.Button
    Public WithEvents Oil_Piping_Template_Create As System.Windows.Forms.Button
    Public WithEvents Template_Part_Number_Text As System.Windows.Forms.TextBox
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents FILTER_MOVE_V_LENGTH_Update As System.Windows.Forms.Button
    Public WithEvents FILTER_SIDE_DIRECTION_MOVE_Update As System.Windows.Forms.Button
    Public WithEvents FILTER_MOVE_H_LENGTH_Update As System.Windows.Forms.Button
    Public WithEvents filter_move_v_length_text As System.Windows.Forms.TextBox
    Public WithEvents filter_side_direction_move_text As System.Windows.Forms.TextBox
    Public WithEvents filter_move_h_length_text As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents Oil_Piping_Parameter_Update As System.Windows.Forms.Button
    Public WithEvents Auto_Clash_Check As System.Windows.Forms.CheckBox
    Public WithEvents UI_SIGHT_GLASS As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As Label
    Friend WithEvents OilPipe_Delay_Text As TextBox
End Class
