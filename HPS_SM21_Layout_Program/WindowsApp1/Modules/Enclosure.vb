Module Enclosure
    '--------------------------------------------------------------------------------------------------- Enclosure 선택 초기화 Constraint
    Public Enclosure_Initial_Constraint, Enclosure_Initial_Constraint_Count
    '--------------------------------------------------------------------------------------------------- Enclosure 선택 초기화 Parameter
    Public Enclosure_Initial_Parameter, Enclosure_Initial_Parameter_Count
    '--------------------------------------------------------------------------------------------------- Enclosure 선택 Constraint
    Public Enclosure_Selection_Constraint, Enclosure_Selection_Constraint_Count
    '--------------------------------------------------------------------------------------------------- Enclosure 선택 Parameter
    Public Enclosure_Selection_Parameter, Enclosure_Selection_Parameter_Count
    '--------------------------------------------------------------------------------------------------- Enclosure 선택 편집설계
    Public Enclosure_Replace_Data, Enclosure_Replace_Data_Count
    '--------------------------------------------------------------------------------------------------- Enclosure DT
    Public Enclosure_DT_Model_No(8), Enclosure_DT_Motor_MAX(8), Enclosure_DT_W_ENC(8), Enclosure_DT_L_ENC(8), Enclosure_DT_H_ENC(8)
    '--------------------------------------------------------------------------------------------------- Enclosure Tempate Folder 저장 이름
    Public Enclosure_folder_name
    '--------------------------------------------------------------------------------------------------- Enclosure 크기 조절
    Public LEN_ENC_L_Expand, DIST_ENCLOSURE_L
    '--------------------------------------------------------------------------------------------------- Enclosure 모터 상판 형상 변경
    Public TOP_MOTOR_PNL_OPTION, TOP_MOTOR_PNL_ABSORBER_SELECT


    Public Function Enclosure_Excel_Open()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        If EXISTS_FILE_CHECK(sExcelFilePath) = False Then
            Return False
        Else
            If Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
                '--------------------------------------------------------------------------------------------------- Enclosure선택초기화 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(15, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Enclosure_Initial_Constraint_Count, Enclosure_Initial_Constraint, "0")
                Template_Data_Loading_Form.ProgressBar1_Change(30)
                '--------------------------------------------------------------------------------------------------- Enclosure선택초기화 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(15, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Enclosure_Initial_Parameter_Count, Enclosure_Initial_Parameter, "4")
                Template_Data_Loading_Form.ProgressBar1_Change(40)
                '--------------------------------------------------------------------------------------------------- Enclosure선택 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(16, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Enclosure_Selection_Constraint_Count, Enclosure_Selection_Constraint, "0")
                Template_Data_Loading_Form.ProgressBar1_Change(50)
                '--------------------------------------------------------------------------------------------------- Enclosure선택 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(16, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Enclosure_Selection_Parameter_Count, Enclosure_Selection_Parameter, "4")
                Template_Data_Loading_Form.ProgressBar1_Change(60)
                '--------------------------------------------------------------------------------------------------- Enclosure선택 편집설계
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(17, 1), 2, "B", "C", "", "", "", "", "", "", "", "", _
                                              Enclosure_Replace_Data_Count, Enclosure_Replace_Data, "0")
            Else
                '--------------------------------------------------------------------------------------------------- Enclosure선택초기화 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(17, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Enclosure_Initial_Constraint_Count, Enclosure_Initial_Constraint, "0")
                Template_Data_Loading_Form.ProgressBar1_Change(30)
                '--------------------------------------------------------------------------------------------------- Enclosure선택초기화 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(17, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Enclosure_Initial_Parameter_Count, Enclosure_Initial_Parameter, "4")
                Template_Data_Loading_Form.ProgressBar1_Change(40)
                '--------------------------------------------------------------------------------------------------- Enclosure선택 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(18, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Enclosure_Selection_Constraint_Count, Enclosure_Selection_Constraint, "0")
                Template_Data_Loading_Form.ProgressBar1_Change(50)
                '--------------------------------------------------------------------------------------------------- Enclosure선택 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(18, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Enclosure_Selection_Parameter_Count, Enclosure_Selection_Parameter, "4")
                Template_Data_Loading_Form.ProgressBar1_Change(60)
            End If

            '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
            sExcelFilePath = sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & "\DT_Design_Rule_Version.xlsx"
            If EXISTS_FILE_CHECK(sExcelFilePath) Then
                ReDim Data_Version_Infomation(0 To 1)
                Call Get_Excel_Value_Function(sExcelFilePath, "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                              Data_Version_Infomation_Count, Data_Version_Infomation, "0")
            Else : Return False
            End If
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(70)

        Return True
    End Function


    Public Sub Put_Result_Data()
        Try
            '--------------------------------------------------------------------------------------------------- 기종에 따른 Design Table Value 가져오기
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Enclosure_folder_name & "\DT_ENCLOSURE.xls")

            Dim Enclosure_DT_count = 1
            Dim Check = True
            Do   'Outer loop
                '--------------------------------------------------------------------------------------------------- Publish_Org_Name 유무 판별
                If Exc.Worksheets.Item(1).Range("A" & Enclosure_DT_count + 1).Value = "" Then        '조건이 True이면
                    Check = False                                                                             '플래그 값을 False로 설정합니다.
                    Exit Do                                                                                       '내부 루프를 종료합니다.
                End If

                '--------------------------------------------------------------------------------------------------- Design Table Model Number Value
                Enclosure_DT_Model_No(Enclosure_DT_count) = Exc.Worksheets.Item(1).Range("A" & Enclosure_DT_count + 1).Value
                '---------------------------------------------------------------------------------------------------
                Enclosure_DT_Motor_MAX(Enclosure_DT_count) = Exc.Worksheets.Item(1).Range("B" & Enclosure_DT_count + 1).Value
                '---------------------------------------------------------------------------------------------------
                Enclosure_DT_W_ENC(Enclosure_DT_count) = Exc.Worksheets.Item(1).Range("C" & Enclosure_DT_count + 1).Value
                '---------------------------------------------------------------------------------------------------
                Enclosure_DT_L_ENC(Enclosure_DT_count) = Exc.Worksheets.Item(1).Range("D" & Enclosure_DT_count + 1).Value
                '---------------------------------------------------------------------------------------------------
                Enclosure_DT_H_ENC(Enclosure_DT_count) = Exc.Worksheets.Item(1).Range("E" & Enclosure_DT_count + 1).Value
                Enclosure_DT_count = Enclosure_DT_count + 1
            Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.
            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()

            '--------------------------------------------------------------------------------------------------- MSFlexGrid1에 Value 넣기
            GA_Enclosure_Template_Create.DataGridView1.ColumnCount = 5
            GA_Enclosure_Template_Create.DataGridView1.Columns(0).HeaderText = ""
            GA_Enclosure_Template_Create.DataGridView1.Columns(1).HeaderText = "Model"
            GA_Enclosure_Template_Create.DataGridView1.Columns(2).HeaderText = "표준폭"
            GA_Enclosure_Template_Create.DataGridView1.Columns(3).HeaderText = "표준길이"
            GA_Enclosure_Template_Create.DataGridView1.Columns(4).HeaderText = "표준높이"
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            '--------------------------------------------------------------------------------------------------- MSFlexGrid1 Width
            GA_Enclosure_Template_Create.DataGridView1.Columns(0).Width = 23
            GA_Enclosure_Template_Create.DataGridView1.Columns(1).Width = 100
            GA_Enclosure_Template_Create.DataGridView1.Columns(2).Width = 100
            GA_Enclosure_Template_Create.DataGridView1.Columns(3).Width = 100
            GA_Enclosure_Template_Create.DataGridView1.Columns(4).Width = 100
            '--------------------------------------------------------------------------------------------------- 가운데 정렬
            GA_Enclosure_Template_Create.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Dim Enclosure_DT_no = 1
            Dim first_base_frame_value = Nothing

            For i = 1 To Enclosure_DT_count - 1
                If assy_value_Name = Left(Enclosure_DT_Model_No(i), 4) Then
                    If Enclosure_DT_Motor_MAX(i) > get_parameter_first_value(2) Then
                        GA_Enclosure_Template_Create.DataGridView1.RowCount = GA_Enclosure_Template_Create.DataGridView1.RowCount + 1
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(Enclosure_DT_no).Cells(1).Value = Enclosure_DT_Model_No(i)
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(Enclosure_DT_no).Cells(2).Value = Enclosure_DT_W_ENC(i)
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(Enclosure_DT_no).Cells(3).Value = Enclosure_DT_L_ENC(i)
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(Enclosure_DT_no).Cells(4).Value = Enclosure_DT_H_ENC(i)

                        If Enclosure_DT_no = 1 Then
                            first_base_frame_value = i
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(Enclosure_DT_no).Cells(0).Value = ">"
                        End If
                        Enclosure_DT_no = Enclosure_DT_no + 1
                    End If
                End If
            Next

            '--------------------------------------------------------------------------------------------------- 선택되는 Data가 없을때....
            If Enclosure_DT_no = 1 Then
                For i = 1 To Enclosure_DT_count - 1
                    '---------------------------------------------------------------------------------------------------
                    If assy_value_Name = Left(Enclosure_DT_Model_No(i), 4) Then
                        first_base_frame_value = i
                    End If
                Next
                Call SHOW_MESSAGE_BOX("일치하는 표준형Data가 없습니다. 기종의 제일 큰 표준형이 선택됩니다.", "", 1, 1)
            End If

            '--------------------------------------------------------------------------------------------------- Design Table Search
            '--------------------------------------------------------------------------------------------------- SMx100 일때 MSFlexGrid1 Value
            selection.clear
            selection.search("Name='DT_ENCLOSURE_CONFIG',all")
            If Not selection.Count = 0 Then
                selection.Item(1).Value.Value = first_base_frame_value
                ActiveDocument.Product.Update()
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Put_Result_Data() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Public Sub Enclosure_Option_Setting()
        Try
            Call CATIA_BASIC02()
            selection.clear
            '--------------------------------------------------------------------------------------------------- 기종 정보 변경
            selection.search("Name='MODEL',all")
            If Not selection.Count = 0 Then
                selection.Item(1).Value.Value = STW_Drawing_Infomation(1, 1)
            End If
            selection.clear
            '--------------------------------------------------------------------------------------------------- LEN_ENC_L_Expand 변경
            selection.search("Name='LEN_ENC_L_Expand',all")
            If Not selection.Count = 0 Then
                LEN_ENC_L_Expand = selection.Item(1).Value
                GA_Enclosure_Template_Create.LEN_ENC_L_Expand_Text.Text = LEN_ENC_L_Expand.Value
            End If
            selection.clear
            '--------------------------------------------------------------------------------------------------- Template Update
            selection.search("Name='TR_PLN_DIST_ENC',all")
            If Not selection.Count = 0 Then
                selection.Item(1).Value.Value = 400
            End If
            selection.clear
            selection.search("Name='TR_PLN_DIST_ENC',all")
            If Not selection.Count = 0 Then
                selection.Item(1).Value.Value = 0
            End If
            selection.clear
            '--------------------------------------------------------------------------------------------------- NOW_ENC_L 보여주기
            selection.search("Name='DIST_ENCLOSURE_L',all")
            If Not selection.Count = 0 Then
                DIST_ENCLOSURE_L = selection.Item(1).Value
                GA_Enclosure_Template_Create.NOW_ENC_L.Text = DIST_ENCLOSURE_L.Value
            End If

            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Product.Update()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Enclosure_Option_Setting() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Public Sub Enclosure_Motor_Top_Panel_Function()
        Try
            selection.clear
            '--------------------------------------------------------------------------------------------------- TOP_MOTOR_PNL_OPTION 가져오기
            selection.search("Name='TOP_MOTOR_PNL_OPTION',all")
            If Not selection.Count = 0 Then
                TOP_MOTOR_PNL_OPTION = selection.Item(1).Value
            End If
            selection.clear
            '--------------------------------------------------------------------------------------------------- TOP_MOTOR_PNL_ABSORBER_SELECT 가져오기
            selection.search("Name='TOP_MOTOR_PNL_ABSORBER_SELECT',all")
            If Not selection.Count = 0 Then
                TOP_MOTOR_PNL_ABSORBER_SELECT = selection.Item(1).Value
            End If

            '--------------------------------------------------------------------------------------------------- 상판이 형상 변경
            If TOP_MOTOR_PNL_ABSORBER_SELECT.Value = "OFF" Then
                GA_Enclosure_Template_Create.Enclosure_Standard_Option.Enabled = False
                GA_Enclosure_Template_Create.Enclosure_Remove_Option.Enabled = False
                GA_Enclosure_Template_Create.Enclosure_Extrusion_Option.Enabled = False
                GA_Enclosure_Template_Create.Enclosure_Motor_Top_Panel_Update.Enabled = False
            Else
                GA_Enclosure_Template_Create.Enclosure_Standard_Option.Enabled = False
                GA_Enclosure_Template_Create.Enclosure_Remove_Option.Enabled = True
                GA_Enclosure_Template_Create.Enclosure_Extrusion_Option.Enabled = True
                GA_Enclosure_Template_Create.Enclosure_Motor_Top_Panel_Update.Enabled = True
                GB_Enclosure_Template_Motor_Top_Panel.Show()
            End If

            If TOP_MOTOR_PNL_OPTION.Value = "STANDARD" Then
                GA_Enclosure_Template_Create.Enclosure_Standard_Option.Checked = True
            ElseIf TOP_MOTOR_PNL_OPTION.Value = "STLIP" Then
                GA_Enclosure_Template_Create.Enclosure_Remove_Option.Checked = True
            Else
                GA_Enclosure_Template_Create.Enclosure_Extrusion_Option.Checked = True
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Enclosure_Motor_Top_Panel_Function() 함수 오류!!", "", 1, 1)
        End Try
    End Sub

End Module
