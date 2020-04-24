Public Class D_Oil_Tank_Create


    Private Sub D_Oil_Tank_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 711
        Me.Height = 479
        Me.TopMost = True
        '--------------------------------------------------------------------------------------------------- Oil_Tank_Excel_Open
        Template_Data_Loading_Form.ProgressBar1_Change(50)
        Call Oil_Tank_Excel_Open()

        '--------------------------------------------------------------------------------------------------- MSFlexGrid Setting
        Call DataGridView_Inicial_Setting(DataGridView1, Oil_Tank_Selection_Parameter_Count, Oil_Tank_Selection_Parameter)

        Template_Data_Loading_Form.ProgressBar1_Change(75)

        Result_Template_Item_Name = STW_Template_Infomation(Form_Click_Index_Value, 1)
        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Result_Template_Data_Path = STW_Template_Infomation(Form_Click_Index_Value, 3)

        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        ReDim Data_Version_Infomation(0 To 1)
        Call Get_Excel_Value_Function(sTemplate_Data_Path_text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & _
                                      "\DT_Design_Rule_Version.xlsx", "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                      Data_Version_Infomation_Count, Data_Version_Infomation, "0")

        Template_Data_Loading_Form.ProgressBar1_Change(90)

        '--------------------------------------------------------------------------------------------------- Type 활성화
        If Strings.Left(sLbl_ModelNumber, 2) = "SM" And Not Strings.Mid(sLbl_ModelNumber, 4, 1) = "1" Then
            Option_Default.Enabled = True
            Option_KOR.Enabled = True
            Option_CHN.Enabled = True
        Else
            Option_Default.Enabled = True
            Option_KOR.Enabled = True
            Option_CHN.Enabled = True
        End If
        Call SetControlBox(Me, Me.cancel_button, True)
    End Sub


    Private Sub Oil_Tank_Create_Click(sender As Object, e As EventArgs) Handles Oil_Tank_Create.Click
        Call SetControlBox(Me, Me.cancel_button, False)
        Try
            '--------------------------------------------------------------------------------------------------- 배열의 값 삭제
            For i = 0 To get_parameter_first_value.Length - 1
                get_parameter_first_value(i) = ""
            Next

            For i = 0 To get_parameter_second_value.Length - 1
                get_parameter_second_value(i) = ""
            Next

            ProgressBar1.Value = 0
            '--------------------------------------------------------------------------------------------------- Oil Tank Type 이 없을때
            If Option_Default.Checked = True And Strings.Left(sLbl_ModelNumber, 2) = "SM" Then
                Call SHOW_MESSAGE_BOX("Oil Tank Type을 선택하세요.", "", 1, 1)
                Exit Sub
            End If

            Message_Label.Text = "Oil Tank 선정 중..."
            DataGridView1.RowCount = 0
            DataGridView1.RowCount = 1

            ProgressBar1.Value = 10

            '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
            Call Clash_Check_Folder_Create()
            ProgressBar1.Value = 20

            '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
            Call Data_Constraint_Delete(Oil_Tank_Initial_Constraint, Oil_Tank_Initial_Constraint_Count, Result_Template_Item_Value)
            ProgressBar1.Value = 30

            '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
            Call Get_Parameter_Data_Comparison(Oil_Tank_Initial_Parameter, Oil_Tank_Initial_Parameter_Count, Me)
            ProgressBar1.Value = 40

            '--------------------------------------------------------------------------------------------------- Data 비교
            Dim Part_Search_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart"
            Dim Part_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\"

            Call Data_Parameter_Comparison(Part_Search_Folder, Part_Folder, Me, Result_Template_Item_Value, Oil_Tank_Selection_Parameter,
                                           Oil_Tank_Initial_Parameter, Oil_Tank_Selection_Parameter_Count)
            ProgressBar1.Value = 60

            If Result_Check_Count > 1 Then
                '--------------------------------------------------------------------------------------------------- 구속조건 생성
                For i = 1 To Oil_Tank_Selection_Constraint_Count - 1
                    Call Constraint_Delete(Oil_Tank_Selection_Constraint(i, 1), Oil_Tank_Selection_Constraint(i, 2), Oil_Tank_Selection_Constraint(i, 3))
                Next
                '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
                Dim file_name_org = Split(DataGridView1.Rows(0).Cells(2).Value, ".CATPart")
                ProgressBar1.Value = 80

                Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))
                Call Btn_Refresh_Click()
            Else
                Call SetControlBox(Me, Me.cancel_button, True)
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
            Message_Label.Text = "Oil Tank 선정 완료"
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            ProgressBar1.Value = 100
            Message_Label.Text = "Oil Tank 선정 오류..."
        End Try
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Call SetControlBox(Me, Me.cancel_button, False)
        Call RE_SELECTION_PART(Me, Oil_Tank_Selection_Constraint_Count, Oil_Tank_Selection_Constraint, Oil_Tank_Initial_Constraint, Oil_Tank_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub


    Private Sub Oil_Tank_Clash_Check_Click(sender As Object, e As EventArgs) Handles Oil_Tank_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Oil Tank 간섭 체크중.."
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
        Message_Label.Text = "Oil Tank 간섭체크 완료"
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs)
        Call VVIP_Check_Form_Size(711, 884, Me)
    End Sub


    Private Sub cancel_button_Click(sender As Object, e As EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub

End Class