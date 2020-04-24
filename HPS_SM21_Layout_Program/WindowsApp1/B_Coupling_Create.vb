Public Class B_Coupling_Create
    Dim Coupling_Create_Close = False




    Private Sub B_Coupling_Create_Load(sender As Object, e As EventArgs) Handles Me.Load
        '--------------------------------------------------------------------------------------------------- Form Position
        Me.Top = 150
        Me.Left = 50
        Me.Width = 415
        Me.Height = 541
        Me.TopMost = True
        GBX_Manufacturer.Enabled = False        ' 제조사 비활성화
        Template_Data_Loading_Form.ProgressBar1_Change(20)
        '--------------------------------------------------------------------------------------------------- Coupling_Excel_Open
        If Not Coupling_Excel_Open() = True Then
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- DataGridView Setting
        Call DataGridView_Inicial_Setting(Me.DataGridView1, Coupling_Selection_Parameter_Count, Coupling_Selection_Parameter)

        Template_Data_Loading_Form.ProgressBar1_Change(90)

        Result_Template_Item_Name = STW_Template_Infomation(Form_Click_Index_Value, 1)
        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Result_Template_Data_Path = STW_Template_Infomation(Form_Click_Index_Value, 3)
        Template_Data_Loading_Form.ProgressBar1_Change(100)
        Call SetControlBox(Me, Me.cancel_button, True)
    End Sub




    Private Sub Coupling_Create_Click(sender As Object, e As EventArgs) Handles Coupling_Create.Click
        Call SetControlBox(Me, Me.cancel_button, False)
        ProgressBar1.Value = 0

        '--------------------------------------------------------------------------------------------------- Coupling  Type Check 가 없을때...
        If Me.Check1.Checked = False And Me.Check2.Checked = False Then
            Call SHOW_MESSAGE_BOX("Coupling Type 을 선택하세요.", "", 1, 1)
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- 제조사 선택이 되어 있지 않을떄..
        If Check1.Checked = True Then
            If CBX_DEFAULT.Checked = True Then
                Call SHOW_MESSAGE_BOX("Coupling 제조사를 선택하세요.", "", 1, 1)
                Exit Sub
            End If
        End If

        Me.Message_Label.Text = "Coupling 선정 중..."
        ProgressBar1.Value = 10

        Me.DataGridView1.RowCount = 0
        Me.DataGridView1.RowCount = 1
        Me.DataGridView1.RowHeadersVisible = False

        '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
        Call Clash_Check_Folder_Create()
        ProgressBar1.Value = 20

        '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
        Call Data_Constraint_Delete(Coupling_Initial_Constraint, Coupling_Initial_Constraint_Count, Result_Template_Item_Value)
        ProgressBar1.Value = 30

        '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
        If Get_Parameter_Data_Comparison(Coupling_Initial_Parameter, Coupling_Initial_Parameter_Count, Me) = False Then
            Exit Sub
        End If
        ProgressBar1.Value = 40

        '--------------------------------------------------------------------------------------------------- Data 비교
        Dim Part_Search_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart"
        Dim Part_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\"
        Call Data_Parameter_Comparison(Part_Search_Folder, Part_Folder, Me, Result_Template_Item_Value, Coupling_Selection_Parameter,
                                       Coupling_Initial_Parameter, Coupling_Selection_Parameter_Count)
        ProgressBar1.Value = 55

        '--------------------------------------------------------------------------------------------------- Dummy Loading
        If Me.Dummy_Loading_Check.Checked = True Then
            If Dummy_Data_open() = False Then
                If SHOW_MESSAGE_BOX("더미 데이터 오픈 실패!" & vbCrLf & "계속 진행 하시겠습니까?", "", 2, 1) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        ProgressBar1.Value = 75

        If Result_Check_Count > 1 Then
            '--------------------------------------------------------------------------------------------------- 구속조건 생성
            For i = 1 To Coupling_Selection_Constraint_Count - 1
                Call Constraint_Delete(Coupling_Selection_Constraint(i, 1), Coupling_Selection_Constraint(i, 2), Coupling_Selection_Constraint(i, 3))
            Next
            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Dim file_name_org = Split(Me.DataGridView1.Rows(0).Cells(2).Value, ".CATPart")
            Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))

            If Btn_Refresh_Click() = False Then
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            '--------------------------------------------------------------------------------------------------- 간섭체크 실행 유무 판단
            If DataGridView1.RowCount >= 1 Then
                If Not (DataGridView1.Rows(0).Cells(1).Value.ToString = "") Or (DataGridView1.Rows(0).Cells(2).Value.ToString = "") Then
                    '--------------------------------------------------------------------------------------------------- 폼 X 버튼 비활성화
                    'Call REMOVED_CLOSE_BUTTON(Me.Handle)
                End If
                Call SetControlBox(Me, Me.cancel_button, False)
            Else
                Call SetControlBox(Me, Me.cancel_button, True)
            End If

            ProgressBar1.Value = 100
            Me.Message_Label.Text = "Coupling 선정 완료"
        Else
            SetControlBox(Me, Me.cancel_button, True)
        End If
    End Sub


    Private Sub Coupling_Clash_Check_Click(sender As Object, e As EventArgs) Handles Coupling_Clash_Check.Click
        Message_Label.Text = "Coupling 간섭 체크중.."
        Call SetControlBox(Me, Me.cancel_button, True)
        Call Clash_Check_Folder_Create()

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(COMP_Product_File_Name & ".CATProduct", 0, 0)

        '--------------------------------------------------------------------------------------------------- Excel Value 가져오기
        Call Get_Clash_Check_Excel_Value(Result_Template_Item_Value)

        '--------------------------------------------------------------------------------------------------- Count Check
        If Check_Count = 0 Then Exit Sub
        '--------------------------------------------------------------------------------------------------- Clash 계산 및 결과값 입력
        Call Clash_Check_Execute(Result_Template_Item_Value)
        Message_Label.Text = "Coupling 간섭체크 완료"
    End Sub


    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        'cancel_button.Enabled = False
        Call SetControlBox(Me, Me.cancel_button, False)

        Call RE_SELECTION_PART(Me, Coupling_Selection_Constraint_Count, Coupling_Selection_Constraint, Coupling_Initial_Constraint, Coupling_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub


    Private Sub cancel_button_Click(sender As Object, e As EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub


    Private Sub Check1_CheckedChanged(sender As Object, e As EventArgs) Handles Check1.CheckedChanged
        If Check1.Checked = True Then
            GBX_Manufacturer.Enabled = True
        Else
            GBX_Manufacturer.Enabled = False
        End If
    End Sub

End Class