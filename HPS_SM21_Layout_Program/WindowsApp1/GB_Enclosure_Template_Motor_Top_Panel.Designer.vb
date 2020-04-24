<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GB_Enclosure_Template_Motor_Top_Panel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GB_Enclosure_Template_Motor_Top_Panel))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.Enclosure_Motor_Top_Panel_Update = New System.Windows.Forms.Button()
        Me.Enclosure_Extrusion_Option = New System.Windows.Forms.RadioButton()
        Me.Enclosure_Remove_Option = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Message_Label)
        Me.GroupBox1.Controls.Add(Me.Enclosure_Motor_Top_Panel_Update)
        Me.GroupBox1.Controls.Add(Me.Enclosure_Extrusion_Option)
        Me.GroupBox1.Controls.Add(Me.Enclosure_Remove_Option)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(308, 136)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "모터 상판 형상 정의"
        '
        'Message_Label
        '
        Me.Message_Label.AutoSize = True
        Me.Message_Label.Location = New System.Drawing.Point(20, 91)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(11, 12)
        Me.Message_Label.TabIndex = 2
        Me.Message_Label.Text = "-"
        '
        'Enclosure_Motor_Top_Panel_Update
        '
        Me.Enclosure_Motor_Top_Panel_Update.Location = New System.Drawing.Point(174, 81)
        Me.Enclosure_Motor_Top_Panel_Update.Name = "Enclosure_Motor_Top_Panel_Update"
        Me.Enclosure_Motor_Top_Panel_Update.Size = New System.Drawing.Size(111, 32)
        Me.Enclosure_Motor_Top_Panel_Update.TabIndex = 1
        Me.Enclosure_Motor_Top_Panel_Update.Text = "Update"
        Me.Enclosure_Motor_Top_Panel_Update.UseVisualStyleBackColor = True
        '
        'Enclosure_Extrusion_Option
        '
        Me.Enclosure_Extrusion_Option.AutoSize = True
        Me.Enclosure_Extrusion_Option.Location = New System.Drawing.Point(174, 41)
        Me.Enclosure_Extrusion_Option.Name = "Enclosure_Extrusion_Option"
        Me.Enclosure_Extrusion_Option.Size = New System.Drawing.Size(87, 16)
        Me.Enclosure_Extrusion_Option.TabIndex = 0
        Me.Enclosure_Extrusion_Option.TabStop = True
        Me.Enclosure_Extrusion_Option.Text = "상판 돌출형"
        Me.Enclosure_Extrusion_Option.UseVisualStyleBackColor = True
        '
        'Enclosure_Remove_Option
        '
        Me.Enclosure_Remove_Option.AutoSize = True
        Me.Enclosure_Remove_Option.Location = New System.Drawing.Point(38, 41)
        Me.Enclosure_Remove_Option.Name = "Enclosure_Remove_Option"
        Me.Enclosure_Remove_Option.Size = New System.Drawing.Size(87, 16)
        Me.Enclosure_Remove_Option.TabIndex = 0
        Me.Enclosure_Remove_Option.TabStop = True
        Me.Enclosure_Remove_Option.Text = "흡음재 제거"
        Me.Enclosure_Remove_Option.UseVisualStyleBackColor = True
        '
        'GB_Enclosure_Template_Motor_Top_Panel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(332, 158)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "GB_Enclosure_Template_Motor_Top_Panel"
        Me.Text = "GB_Enclosure_Template_Motor_Top_Panel"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Enclosure_Motor_Top_Panel_Update As System.Windows.Forms.Button
    Public WithEvents Enclosure_Extrusion_Option As System.Windows.Forms.RadioButton
    Public WithEvents Enclosure_Remove_Option As System.Windows.Forms.RadioButton
    Public WithEvents Message_Label As System.Windows.Forms.Label
End Class
