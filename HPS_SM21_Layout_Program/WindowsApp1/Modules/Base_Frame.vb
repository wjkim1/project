Module Base_Frame
    '--------------------------------------------------------------------------------------------------- Base_Frame 선택 초기화 Constraint
    Public Base_Frame_Initial_Constraint, Base_Frame_Initial_Constraint_Count
    '--------------------------------------------------------------------------------------------------- Base_Frame 선택 초기화 Parameter
    Public Base_Frame_Initial_Parameter, Base_Frame_Initial_Parameter_Count
    '--------------------------------------------------------------------------------------------------- Base Frame 선택 Constraint
    Public Base_Frame_Selection_Constraint, Base_Frame_Selection_Constraint_Count
    '--------------------------------------------------------------------------------------------------- Base Frame 선택 Parameter
    Public Base_Frame_Selection_Parameter, Base_Frame_Selection_Parameter_Count
    '--------------------------------------------------------------------------------------------------- Base Frame Replace Parameter
    Public Base_Frame_Replace_Data, Base_Frame_Replace_Data_Count
    '--------------------------------------------------------------------------------------------------- Base Frame 선택 비교 Excel
    Public Base_Frame_Comparison, Base_Frame_Comparison_Count
    '--------------------------------------------------------------------------------------------------- Base Frame Tempate Folder 저장 이름
    Public Base_Frame_folder_name
    '--------------------------------------------------------------------------------------------------- 오일탱크 DT
    Public Base_Frame_DT_Model_No(8), Base_Frame_DT_Motor_MAX(8), Base_Frame_DT_W_BASE(8), Base_Frame_DT_L_BASE(8)
    '--------------------------------------------------------------------------------------------------- Base Frame 파일 수량 정의
    Public File_Name(1000)
    '--------------------------------------------------------------------------------------------------- Base Frame 도면 사용 Parameter
    Public KIM1, KIM2, KIM3, KIM4, KIM5, KIMa, KIMb, KIMc
    '--------------------------------------------------------------------------------------------------- Base Frame 길이 확장에 사용
    Public LEN_BASE_L_Expand, LEN_BASE_L_NOW
    '--------------------------------------------------------------------------------------------------- Base Frame OT_MT_Type 검색
    Public OT_MT_count

    Public Auto_Trap_Value, EPN05_D12


    Public Function Base_Frame_Excel_Open()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            '--------------------------------------------------------------------------------------------------- 기종 구분
            If Left(A_Main_Form.Lbl_ModelNumber.Text, 2) = "SM" And Not Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
                '--------------------------------------------------------------------------------------------------- Base Frame선택초기화 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(15, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Base_Frame_Initial_Constraint_Count, Base_Frame_Initial_Constraint, "0")

                Template_Data_Loading_Form.ProgressBar1_Change(30)
                '--------------------------------------------------------------------------------------------------- Base Frame선택초기화 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(15, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Base_Frame_Initial_Parameter_Count, Base_Frame_Initial_Parameter, "4")

                Template_Data_Loading_Form.ProgressBar1_Change(40)
                '--------------------------------------------------------------------------------------------------- Base Frame선택 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(16, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Base_Frame_Selection_Constraint_Count, Base_Frame_Selection_Constraint, "0")

                Template_Data_Loading_Form.ProgressBar1_Change(50)
                '--------------------------------------------------------------------------------------------------- Base Frame선택 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(16, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Base_Frame_Selection_Parameter_Count, Base_Frame_Selection_Parameter, "4")

                Template_Data_Loading_Form.ProgressBar1_Change(60)
                '--------------------------------------------------------------------------------------------------- Base Frame 편집설계
                Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(14, 1), 2, "B", "C", "", "", "", "", "", "", "", "", _
                                              Base_Frame_Replace_Data_Count, Base_Frame_Replace_Data, "0")

                Template_Data_Loading_Form.ProgressBar1_Change(70)
            Else
                '--------------------------------------------------------------------------------------------------- Base Frame선택초기화 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, "베이스프레임_선택초기화", 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Base_Frame_Initial_Constraint_Count, Base_Frame_Initial_Constraint, "0")

                Template_Data_Loading_Form.ProgressBar1_Change(30)
                '--------------------------------------------------------------------------------------------------- Base Frame선택초기화 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, "베이스프레임_선택초기화", 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Base_Frame_Initial_Parameter_Count, Base_Frame_Initial_Parameter, "4")

                Template_Data_Loading_Form.ProgressBar1_Change(40)
                '--------------------------------------------------------------------------------------------------- Base Frame선택 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, "베이스프레임_선택", 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                              Base_Frame_Selection_Constraint_Count, Base_Frame_Selection_Constraint, "0")

                Template_Data_Loading_Form.ProgressBar1_Change(50)
                '--------------------------------------------------------------------------------------------------- Base Frame선택 Parameter
                Call Get_Excel_Value_Function(sExcelFilePath, "베이스프레임_선택", 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                              Base_Frame_Selection_Parameter_Count, Base_Frame_Selection_Parameter, "4")

                Template_Data_Loading_Form.ProgressBar1_Change(60)
                '--------------------------------------------------------------------------------------------------- Base Frame 편집설계
                Call Get_Excel_Value_Function(sExcelFilePath, "베이스프레임_편집설계", 2, "B", "C", "", "", "", "", "", "", "", "", _
                                              Base_Frame_Replace_Data_Count, Base_Frame_Replace_Data, "0")

                Template_Data_Loading_Form.ProgressBar1_Change(70)
            End If
        Else : Return False
        End If

        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        sExcelFilePath = sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & "\DT_Design_Rule_Version.xlsx"
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            ReDim Data_Version_Infomation(0 To 1)
            Call Get_Excel_Value_Function(sExcelFilePath, "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                          Data_Version_Infomation_Count, Data_Version_Infomation, "0")
        Else : Return False
        End If

        Return True
    End Function


    Public Sub Base_Frame_Data_Open()
        Try
            Call CATIA_BASIC02()
            H_Base_Frame_Create.DataGridView1.Columns(0).Visible = False
            Dim Base_Frame_Addcomponents_Count = 1
            Dim Data_Selection_Value = 1

            '--------------------------------------------------------------------------------------------------- Enclosure 유무 정의
            If H_Base_Frame_Create.UI_ENCLOSURE_TYPE.Checked = True Then
                Enclosure_Value = "O"
            Else
                Enclosure_Value = "X"
            End If

            Dim MotorStarter_Value
            If H_Base_Frame_Create.UI_MotorStarter_YES.Checked = True Then
                MotorStarter_Value = "O"
            Else
                MotorStarter_Value = "X"
            End If
            '--------------------------------------------------------------------------------------------------- 기종 구분(SMx100 이면...)
            If Strings.Right(A_Main_Form.Lbl_ModelNumber.Text, 1) = "1" Then
                '--------------------------------------------------------------------------------------------------- Data 비교
                For i = 1 To Base_Frame_Comparison_Count - 1
                    '--------------------------------------------------------------------------------------------------- 같은 기종인지, MOTOR_MAX 값보다 큰지, Enclosure 유무가 맞는지 비교
                    If assy_value_Name & "00" = Base_Frame_Comparison(i, 1) And
                        MotorStarter_Value = Base_Frame_Comparison(i, 2) And
                        Enclosure_Value = Base_Frame_Comparison(i, 3) Then
                        ' And _
                        'axis_axis_measure_x < Base_Frame_Comparison(i, 2) And _
                        'Enclosure_Value = Base_Frame_Comparison(i, 3) Then

                        Base_Frame_Addcomponents_Count = 2
                        H_Base_Frame_Create.DataGridView1.RowCount = Data_Selection_Value '+ 1

                        H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(1).Value = Data_Selection_Value
                        H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(2).Value = Base_Frame_Comparison(i, 4)
                        H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(3).Value = Base_Frame_Comparison(i, 5)
                        H_Base_Frame_Create.Data_Path_List.Items.Add(sDirectory_Path_Text & "\" & assy_value_Name & "\" & Result_Template_Item_Value & "\")

                        If Data_Selection_Value = 1 Then
                            H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(0).Value = ">"
                        End If

                        '--------------------------------------------------------------------------------------------------- 선택한 행정보 데이터로 변경
                        If Data_Selection_Value < 2 Then
                            Call Result_Grid_Change_Base_Frame(H_Base_Frame_Create.GroupBox6, Data_Selection_Value - 1)
                        End If
                        'Debug.Print (H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(1).Value)
                        Base_Frame_Addcomponents_Count = 0                  '리스트에 추가된 데이터가 있는지 없는지 체크
                        Data_Selection_Value = Data_Selection_Value + 1     '리스트에 추가된 데이터 갯수 체크
                    End If
                Next i
                '--------------------------------------------------------------------------------------------------- 기종 구분(SMx100 아니면...)
            Else
                '--------------------------------------------------------------------------------------------------- Auto_Trap 유무
                If H_Base_Frame_Create.Auto_Trap_Option1.Checked = True Then
                    Auto_Trap_Value = "O"
                Else
                    Auto_Trap_Value = "X"
                End If

                '--------------------------------------------------------------------------------------------------- 오일탱크 크기 가져오기
                For i = 1 To Base_Frame_Selection_Parameter_Count
                    selection.clear
                    selection.search("Name='" & Base_Frame_Selection_Parameter(i, 1) & "',all")

                    If Not selection.Count = 0 Then
                        EPN05_D12 = selection.Item(1).Value
                        Exit For
                    End If
                Next i

                '--------------------------------------------------------------------------------------------------- Data 비교
                For i = 1 To Base_Frame_Comparison_Count - 1
                    Dim Len_assy_value_Name = Len(assy_value_Name)

                    '--------------------------------------------------------------------------------------------------- Data 비교
                    If Left(Base_Frame_Comparison(i, 1), Len_assy_value_Name) = assy_value_Name And
                        Base_Frame_Comparison(i, 2) = CStr(EPN05_D12.Value) And
                        Base_Frame_Comparison(i, 4) = Enclosure_Value And
                        Base_Frame_Comparison(i, 5) = Auto_Trap_Value Then

                        If Not EPN05_D12 = Nothing Then
                            Base_Frame_Addcomponents_Count = 2
                            H_Base_Frame_Create.DataGridView1.RowCount = Data_Selection_Value ' + 1
                            H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(1).Value = Data_Selection_Value
                            H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(2).Value = Base_Frame_Comparison(i, 3)
                            H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(3).Value = Base_Frame_Comparison(i, 2)
                            H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(4).Value = Base_Frame_Comparison(i, 4)
                            H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(5).Value = Base_Frame_Comparison(i, 5)

                            H_Base_Frame_Create.Data_Path_List.Items.Add(sDirectory_Path_Text & "\" & assy_value_Name & "\" & Result_Template_Item_Value & "\")
                            H_Base_Frame_Create.Data_Path_List.Items.Add(sDirectory_Path_Text & "\" & assy_value_Name & "\" & Result_Template_Item_Value & "\")

                            If Data_Selection_Value = 1 Then
                                H_Base_Frame_Create.DataGridView1.Rows(H_Base_Frame_Create.DataGridView1.RowCount - 1).Cells(0).Value = ">"
                            End If

                            '--------------------------------------------------------------------------------------------------- 선택한 행정보 데이터로 변경
                            If Data_Selection_Value < 2 Then
                                'GroupBox6
                                'DataGridView1
                                Call Result_Grid_Change(H_Base_Frame_Create.GroupBox6, Data_Selection_Value)
                                'Call Result_Grid_Change(H_Base_Frame_Create, Data_Selection_Value)
                            End If

                            Base_Frame_Addcomponents_Count = 0                  '리스트에 추가된 데이터가 있는지 없는지 체크
                            Data_Selection_Value = Data_Selection_Value + 1     '리스트에 추가된 데이터 갯수 체크
                        End If
                    End If
                Next
            End If
            '--------------------------------------------------------------------------------------------------- 엑셀에 조건에 맞는 데이터가 없으면...
            If Base_Frame_Addcomponents_Count = 1 Then
                H_Base_Frame_Create.DataGridView1.Rows(1).Cells(1).Value = ""
                H_Base_Frame_Create.DataGridView1.Rows(1).Cells(2).Value = ""
                H_Base_Frame_Create.DataGridView1.Rows(1).Cells(3).Value = ""
                MsgBox("선정정보와 일치하는 품번이 없습니다. (Excel)")
                '--------------------------------------------------------------------------------------------------- 조건에 맞는 Part가 없으면...
            ElseIf Base_Frame_Addcomponents_Count = 2 Then
                MsgBox("품번과 일치하는 실적 데이터가 없습니다. (Part)")
            Else
                '--------------------------------------------------------------------------------------------------- 구속조건 생성
                For i = 1 To Base_Frame_Selection_Constraint_Count - 1
                    Call Constraint_Delete(Base_Frame_Selection_Constraint(i, 1), Base_Frame_Selection_Constraint(i, 2), Base_Frame_Selection_Constraint(i, 3))
                Next
                product_add.Update()
            End If
            'H_Base_Frame_Create.DataGridView1.Columns(0).Width = 0
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Frame_Data_Open() 함수 오류!!", "", 1, 1)
        End Try
        H_Base_Frame_Create.DataGridView1.Refresh()
    End Sub


    Public Function Base_Frame_Comparison_Get_Excel()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 5)

        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            '--------------------------------------------------------------------------------------------------- 기종 구분...
            If Right(A_Main_Form.Lbl_ModelNumber.Text, 1) = "1" Then
                '--------------------------------------------------------------------------------------------------- Base Frame선택초기화 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, "SM21", 5, "A", "B", "C", "D", "E", "", "", "", "", "", _
                                              Base_Frame_Comparison_Count, Base_Frame_Comparison, "0")
            Else
                '--------------------------------------------------------------------------------------------------- Base Frame선택초기화 Constraint
                Call Get_Excel_Value_Function(sExcelFilePath, "SM", 7, "A", "C", "G", "H", "I", "J", "K", "", "", "", _
                                              Base_Frame_Comparison_Count, Base_Frame_Comparison, "0")
            End If
        Else : Return False
        End If

        Return False
    End Function


    Public Function Base_Frame_weld_mach_design_table_control()
        Try
            '--------------------------------------------------------------------------------------------------- 기종에 따른 Design Table Value 가져오기
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Base_Frame_folder_name & "\DT_BASE_FRAME.xls")

            Dim base_frame_DT_count = 1
            Dim Check = True
            Do   'Outer loop
                '--------------------------------------------------------------------------------------------------- Publish_Org_Name 유무 판별
                If Exc.Worksheets.Item(1).Range("A" & base_frame_DT_count + 1).Value = "" Then        '조건이 True이면
                    Check = False                                                                             '플래그 값을 False로 설정합니다.
                    Exit Do                                                                                       '내부 루프를 종료합니다.
                End If
                '--------------------------------------------------------------------------------------------------- Design Table Model Number Value
                Base_Frame_DT_Model_No(base_frame_DT_count) = Exc.Worksheets.Item(1).Range("A" & base_frame_DT_count + 1).Value
                '--------------------------------------------------------------------------------------------------- Base_Frame_DT_Motor_MAX
                Base_Frame_DT_Motor_MAX(base_frame_DT_count) = Exc.Worksheets.Item(1).Range("B" & base_frame_DT_count + 1).Value
                '--------------------------------------------------------------------------------------------------- Base_Frame_DT_W_BASE
                Base_Frame_DT_W_BASE(base_frame_DT_count) = Exc.Worksheets.Item(1).Range("C" & base_frame_DT_count + 1).Value
                '--------------------------------------------------------------------------------------------------- Base_Frame_DT_L_BASE
                Base_Frame_DT_L_BASE(base_frame_DT_count) = Exc.Worksheets.Item(1).Range("D" & base_frame_DT_count + 1).Value
                base_frame_DT_count = base_frame_DT_count + 1
            Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.
            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()

            '--------------------------------------------------------------------------------------------------- MSFlexGrid1에 Value 넣기
            HA_Base_Frame_Template_Create.DataGridView1.ColumnCount = 5
            GA_Enclosure_Template_Create.DataGridView1.Columns(0).HeaderText = ""
            GA_Enclosure_Template_Create.DataGridView1.Columns(1).HeaderText = "NO"
            GA_Enclosure_Template_Create.DataGridView1.Columns(2).HeaderText = "MODEL"
            GA_Enclosure_Template_Create.DataGridView1.Columns(3).HeaderText = "표준폭"
            GA_Enclosure_Template_Create.DataGridView1.Columns(4).HeaderText = "표준길이"
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            '--------------------------------------------------------------------------------------------------- MSFlexGrid1 Width
            HA_Base_Frame_Template_Create.DataGridView1.Columns(0).Width = 23
            HA_Base_Frame_Template_Create.DataGridView1.Columns(1).Width = 40
            HA_Base_Frame_Template_Create.DataGridView1.Columns(2).Width = 100
            HA_Base_Frame_Template_Create.DataGridView1.Columns(3).Width = 90
            HA_Base_Frame_Template_Create.DataGridView1.Columns(4).Width = 90
            '--------------------------------------------------------------------------------------------------- 가운데 정렬
            GA_Enclosure_Template_Create.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            GA_Enclosure_Template_Create.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Dim base_frame_DT_no = 1
            Dim first_base_frame_value = Nothing

            For i = 1 To base_frame_DT_count - 1
                If assy_value_Name = Left(Base_Frame_DT_Model_No(i), 4) Then
                    If Base_Frame_DT_Motor_MAX(i) > get_parameter_first_value(2) Then
                        HA_Base_Frame_Template_Create.DataGridView1.RowCount = HA_Base_Frame_Template_Create.DataGridView1.RowCount + 1
                        HA_Base_Frame_Template_Create.DataGridView1.Rows(base_frame_DT_no).Cells(1).Value() = base_frame_DT_no
                        HA_Base_Frame_Template_Create.DataGridView1.Rows(base_frame_DT_no).Cells(2).Value() = Base_Frame_DT_Model_No(i)
                        HA_Base_Frame_Template_Create.DataGridView1.Rows(base_frame_DT_no).Cells(3).Value() = Base_Frame_DT_W_BASE(i)
                        HA_Base_Frame_Template_Create.DataGridView1.Rows(base_frame_DT_no).Cells(4).Value() = Base_Frame_DT_L_BASE(i)

                        If base_frame_DT_no = 1 Then
                            first_base_frame_value = i
                            HA_Base_Frame_Template_Create.DataGridView1.Rows(base_frame_DT_no).Cells(0).Value() = ">"
                        End If
                        base_frame_DT_no = base_frame_DT_no + 1
                    End If
                End If
            Next

            '--------------------------------------------------------------------------------------------------- 선택되는 Data가 없을때....
            If base_frame_DT_no = 1 Then
                For i = 1 To base_frame_DT_count - 1
                    '---------------------------------------------------------------------------------------------------
                    If assy_value_Name = Left(Base_Frame_DT_Model_No(i), 4) Then
                        first_base_frame_value = i
                    End If
                Next
                MsgBox("일치하는 표준형Data가 없습니다. 기종의 제일 큰 표준형이 선택됩니다.", vbInformation, "ISPark_Automation")
            End If

            '--------------------------------------------------------------------------------------------------- Part List 1번에 Data 이름 넣기
            'Exc.Visible = True
            Exc.Workbooks.Open(Base_Frame_folder_name & "\" & STW_Template_Infomation(7, 7))
            '--------------------------------------------------------------------------------------------------- Data 검토
            Exc.Worksheets.Item(1).Range("Z1").Value = Left(HA_Base_Frame_Template_Create.DataGridView1.Rows(1).Cells(1).Value(), 4)
            Exc.ActiveWorkbook.Save()
            Call KillProcess("EXCEL")
            selection.clear
            '--------------------------------------------------------------------------------------------------- Design Table Search
            selection.search("Name='DT_BASE_CONFIG',all")

            '--------------------------------------------------------------------------------------------------- SMx100 일때 MSFlexGrid1 Value
            selection.Item(1).Value.Value = first_base_frame_value
            ActiveDocument.Product.Update()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Frame_weld_mach_design_table_control() 함수 오류!!", "", 1, 1)
        End Try

        Return True
    End Function


    Public Sub Base_Frame_Option_Setting_2()
        Try
            selection.clear
            '--------------------------------------------------------------------------------------------------- 기종 정보 변경
            selection.search("Name='Machine_Type',all")
            selection.Item(1).Value.Value = STW_Drawing_Infomation(1, 1)

            If Not Right(A_Main_Form.Lbl_ModelNumber.Text, 1) = "1" Then
                selection.clear
                '--------------------------------------------------------------------------------------------------- Template Update
                selection.search("Name='TR_PLN_DIST',all")
                selection.Item(1).Value.Value = 400
                CATIA.ActiveDocument.Product.Update()
                selection.Clear()
                selection.search("Name='TR_PLN_DIST',all")
                selection.Item(1).Value.Value = 0
                CATIA.ActiveDocument.Product.Update()
                '--------------------------------------------------------------------------------------------------- base_frame_Parameter 가져오기
                selection.clear
                selection.search("Name='KIM1',all")
                KIM1 = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIM2',all")
                KIM2 = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIM3',all")
                KIM3 = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIM4',all")
                KIM4 = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIM5',all")
                KIM5 = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIMa',all")
                KIMa = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIMb',all")
                KIMb = selection.Item(1).Value
                selection.clear
                selection.search("Name='KIMc',all")
                KIMc = selection.Item(1).Value
                selection.clear
                '--------------------------------------------------------------------------------------------------- LEN_BASE_L_Expand 가져오기
                selection.search("Name='LEN_BASE_L_Expand',all")
                LEN_BASE_L_Expand = selection.Item(1).Value
                HA_Base_Frame_Template_Create.LEN_BASE_L_Expand_Text.Text = LEN_BASE_L_Expand.Value
                selection.clear
                '--------------------------------------------------------------------------------------------------- LEN_BASE_L_NOW 가져오기
                selection.search("Name='LEN_BASE_L_NOW',all")
                LEN_BASE_L_NOW = selection.Item(1).Value
                HA_Base_Frame_Template_Create.NOW_BASE_L.Text = Math.Round(LEN_BASE_L_NOW.Value, 2)
            End If

            CATIA.ActiveDocument.Product.Update()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Frame_weld_mach_design_table_control2() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Public Sub Base_Frame_Drawing_Text_Setting()
        Call CATIA_BASIC03()
        Drawingselection.clear
        Drawingselection.search("Name=DWG_Name*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = STW_Drawing_Infomation(1, 1)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIM1_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIM1.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIM2_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIM2.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIM3_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIM3.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIM4_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIM4.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIM5_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIM5.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIMa_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIMa.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIMb_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIMb.Value, 2)
        Next
        Drawingselection.clear
        Drawingselection.search("Name=KIMc_Text*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Math.Round(KIMc.Value, 2)
        Next
    End Sub

End Module
