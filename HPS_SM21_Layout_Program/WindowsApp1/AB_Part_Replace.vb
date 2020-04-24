Public Class AB_Part_Replace
    Public Selection_Replace_Value


    Private Sub AB_Part_Replace_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Top = 150
        Me.Left = 50
        Me.Width = 448
        Me.Height = 423
        Me.TopMost = True

        Call Form_Position_Initial(Me)
    End Sub


    Private Sub Replace_Data_Selection_Click(sender As Object, e As EventArgs) Handles Replace_Data_Selection.Click
        Call CATIA_Basic02()
        Message_Label.Text = "Replace 대상 Part 선택하세요"
        '--------------------------------------------------------------------------------------------------- 선택된 개체가 없을때...
        If selection.Count = "0" Then
            Call SHOW_MESSAGE_BOX("Tree에서 Replace Part를 선택하세요.", "", 1, 1)
            Exit Sub
        End If

        Selection_Replace_Value = selection.Item(1).Value
        Replace_Data_Selection_Text.Text = Selection_Replace_Value.Name

        '--------------------------------------------------------------------------------------------------- MSFlexGrid1에 Value 넣기
        DataGridView1.RowHeadersVisible = False
        DataGridView1.Rows.Clear()
        DataGridView1.ColumnHeadersVisible = True
        DataGridView1.Columns(0).HeaderText = "No"
        DataGridView1.Columns(1).HeaderText = "Replace Data Name"
        '--------------------------------------------------------------------------------------------------- MSFlexGrid1 Width
        DataGridView1.Columns(0).Width = 30
        DataGridView1.Columns(1).Width = 510
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Refresh()

        If Dir(sDirectory_Path_Text & "\" & assy_value_Name & "\" & Me.Replace_Data_Selection_Text.Text, vbDirectory) = "" Then
            Call SHOW_MESSAGE_BOX("올바른 EPN No.가 아닙니다.", "", 1, 1)
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------- 폴더 Path 가져오기
        File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Me.Replace_Data_Selection_Text.Text

        '--------------------------------------------------------------------------------------------------- 폴더내 파일 이름 정의
        For i = 0 To File_numbering.Items.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Item(0, i).Value = (i + 1).ToString()
            DataGridView1.Item(1, i).Value = File_numbering.Items(i).ToString()
        Next

        '--------------------------------------------------------------------------------------------------- 마지막 빈 칸 삭제
        DataGridView1.AllowUserToAddRows = False

        '--------------------------------------------------------------------------------------------------- NO열 위치 정렬
        Message_Label.Text = "교체 대상 Part를 Double 클릭 하세요 "
    End Sub


    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        If Me.DataGridView1.CurrentCell.RowIndex < 0 Then
            Exit Sub
        End If

        Me.Message_Label.Text = "Data Replace 중입니다..."
        Dim msflexgrid_dbclick = Me.DataGridView1.CurrentCell.RowIndex
        Try
            '--------------------------------------------------------------------------------------------------- Data Replace
            Dim replace_product = CATIA.ActiveDocument.Product.products.Item(Selection_Replace_Value.Name)
            Dim data_replace = CATIA.ActiveDocument.Product.products.ReplaceComponent(replace_product, sDirectory_Path_Text & "\" & assy_value_Name & "\" & Me.Replace_Data_Selection_Text.Text & "\" & Me.File_numbering.Items(msflexgrid_dbclick).ToString(), False)
            Threading.Thread.Sleep(500)

            CATIA.ActiveDocument.Product.Update()

            ''--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Dim file_name_org = Split(Me.File_numbering.Items(msflexgrid_dbclick).ToString(), ".CATPart")
            Call O_BOM_Excel_Writing(Me.Replace_Data_Selection_Text.Text, file_name_org(0))
            Call Btn_Refresh_Click()


            Dim fileNames
            Dim fileName
            fileName = Replace(Me.File_numbering.Items(msflexgrid_dbclick).ToString(), ".CATPart", "")
            replace_product.PartNumber = fileName

            '--------------------------------------------------------------------------------------------------- Comp Product 저장
            CATIA.ActiveDocument.Save()

            '--------------------------------------------------------------------------------------------------- Comp Part List 변경
            For i = 0 To A_Main_Form.DataGridView1.Rows.Count - 1
                If A_Main_Form.DataGridView1.Item(1, i).Value = Me.Replace_Data_Selection_Text.Text Then
                    A_Main_Form.DataGridView1.Item(4, i).Value = file_name_org(0)
                    Exit For
                End If
            Next

            Me.Message_Label.Text = "Data Replace 완료."
            Call SHOW_MESSAGE_BOX("Data Replcae가 완료되었습니다.", "", 1, 1)
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Data Replcae 오류", "", 1, 1)
        End Try
    End Sub

End Class