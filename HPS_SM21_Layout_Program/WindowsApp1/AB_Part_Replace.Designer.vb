<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AB_Part_Replace
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AB_Part_Replace))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Replace_Data_Selection_Text = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Replace_Data_Selection = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.File_numbering = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Replace_Data_Selection_Text)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Replace_Data_Selection)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 21)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(408, 61)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = " Replace Data Selection "
        '
        'Replace_Data_Selection_Text
        '
        Me.Replace_Data_Selection_Text.Location = New System.Drawing.Point(116, 29)
        Me.Replace_Data_Selection_Text.Name = "Replace_Data_Selection_Text"
        Me.Replace_Data_Selection_Text.Size = New System.Drawing.Size(185, 23)
        Me.Replace_Data_Selection_Text.TabIndex = 2
        Me.Replace_Data_Selection_Text.Text = "-"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Selection Data :"
        '
        'Replace_Data_Selection
        '
        Me.Replace_Data_Selection.Location = New System.Drawing.Point(307, 29)
        Me.Replace_Data_Selection.Name = "Replace_Data_Selection"
        Me.Replace_Data_Selection.Size = New System.Drawing.Size(95, 23)
        Me.Replace_Data_Selection.TabIndex = 0
        Me.Replace_Data_Selection.Text = "Data Selection "
        Me.Replace_Data_Selection.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DataGridView1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 99)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(408, 259)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = " Data Replace"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(17, 26)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(371, 217)
        Me.DataGridView1.TabIndex = 3
        '
        'Col1
        '
        Me.Col1.HeaderText = ""
        Me.Col1.Name = "Col1"
        '
        'Col2
        '
        Me.Col2.HeaderText = ""
        Me.Col2.Name = "Col2"
        '
        'Col3
        '
        Me.Col3.HeaderText = ""
        Me.Col3.Name = "Col3"
        '
        'Col4
        '
        Me.Col4.HeaderText = ""
        Me.Col4.Name = "Col4"
        '
        'Col5
        '
        Me.Col5.HeaderText = ""
        Me.Col5.Name = "Col5"
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(18, 365)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(402, 28)
        Me.Message_Label.TabIndex = 1
        Me.Message_Label.Text = "-"
        '
        'File_numbering
        '
        Me.File_numbering.FormattingEnabled = True
        Me.File_numbering.Location = New System.Drawing.Point(443, 125)
        Me.File_numbering.Name = "File_numbering"
        Me.File_numbering.Pattern = "*.*"
        Me.File_numbering.Size = New System.Drawing.Size(228, 100)
        Me.File_numbering.TabIndex = 16
        '
        'AB_Part_Replace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(746, 427)
        Me.Controls.Add(Me.File_numbering)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Message_Label)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AB_Part_Replace"
        Me.Text = "AB_Part"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Replace_Data_Selection As Button
    Friend WithEvents Replace_Data_Selection_Text As Label
    Friend WithEvents GroupBox2 As GroupBox
    Public WithEvents DataGridView1 As DataGridView
    Public WithEvents Col1 As DataGridViewTextBoxColumn
    Public WithEvents Col2 As DataGridViewTextBoxColumn
    Public WithEvents Col3 As DataGridViewTextBoxColumn
    Public WithEvents Col4 As DataGridViewTextBoxColumn
    Public WithEvents Col5 As DataGridViewTextBoxColumn
    Friend WithEvents Message_Label As Label
    Public WithEvents File_numbering As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
End Class
