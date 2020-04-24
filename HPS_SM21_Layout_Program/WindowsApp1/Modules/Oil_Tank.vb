Module Oil_Tank
    '--------------------------------------------------------------------------------------------------- 오일탱크 선택 초기화 Constraint
    Public Oil_Tank_Initial_Constraint, Oil_Tank_Initial_Constraint_Count
    '--------------------------------------------------------------------------------------------------- 오일탱크 선택 초기화 Parameter
    Public Oil_Tank_Initial_Parameter, Oil_Tank_Initial_Parameter_Count
    '--------------------------------------------------------------------------------------------------- 오일탱크 선택 Constraint
    Public Oil_Tank_Selection_Constraint, Oil_Tank_Selection_Constraint_Count
    '--------------------------------------------------------------------------------------------------- 오일탱크 선택 Parameter
    Public Oil_Tank_Selection_Parameter, Oil_Tank_Selection_Parameter_Count
    '--------------------------------------------------------------------------------------------------- 오일탱크 Replace Parameter
    Public Oil_Tank_Replace_Data, Oil_Tank_Replace_Data_Count
    '--------------------------------------------------------------------------------------------------- 오일탱크 Tempate Folder 저장 이름
    Public Oil_Tank_folder_name
    '--------------------------------------------------------------------------------------------------- 오일탱크 DT
    Public oil_tank_DT_model_no(60), oil_tank_DT_partno(60), oil_tank_DT_Width_length(60), oil_tank_DT_Total_length(60), Oil_Tank_VVIP_Value(60), oil_tank_DT_EPN05_D10(60), oil_tank_DT_EPN05_D11(60)
    '--------------------------------------------------------------------------------------------------- Part List 변경 이름 변수
    Public PartList_No
    '--------------------------------------------------------------------------------------------------- Oil Tank Template Path
    Public Copy_Oil_Tank_Path
    '--------------------------------------------------------------------------------------------------- 변경되는 Replace Data 값 저장
    Public Oil_Tank_Replace_Data_Value



    Public Sub Oil_Tank_Excel_Open()
        '--------------------------------------------------------------------------------------------------- 도면 관련 정보 가져오기
        sExcelFilePath = Application.StartupPath & "\STW_Drawing_Infomation.xlsx"
        If System.IO.File.Exists(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", _
                                          STW_Drawing_Infomation_Count, STW_Drawing_Infomation, "5")
        Else
            Call SHOW_MESSAGE_BOX("엑셀 파일이 없어 종료합니다." & vbCrLf & "-> " & sExcelFilePath, "", 1, 1)
            Exit Sub
        End If
        Template_Data_Loading_Form.ProgressBar1_Change(40)

        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        '--------------------------------------------------------------------------------------------------- 오일탱크 선택초기화 Constraint
        If System.IO.File.Exists(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(7, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Oil_Tank_Initial_Constraint_Count, Oil_Tank_Initial_Constraint, "0")
            Template_Data_Loading_Form.ProgressBar1_Change(50)

            '--------------------------------------------------------------------------------------------------- 오일탱크 선택초기화 Parameter
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(7, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Oil_Tank_Initial_Parameter_Count, Oil_Tank_Initial_Parameter, "4")
            Template_Data_Loading_Form.ProgressBar1_Change(60)

            '--------------------------------------------------------------------------------------------------- 오일탱크 선택  Constraint
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(8, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Oil_Tank_Selection_Constraint_Count, Oil_Tank_Selection_Constraint, "0")
            Template_Data_Loading_Form.ProgressBar1_Change(70)

            '--------------------------------------------------------------------------------------------------- 오일탱크 선택 Parameter
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(8, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Oil_Tank_Selection_Parameter_Count, Oil_Tank_Selection_Parameter, "4")
            Template_Data_Loading_Form.ProgressBar1_Change(80)

            '--------------------------------------------------------------------------------------------------- 오일탱크 선택 Parameter
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(9, 1), 2, "B", "C", "", "", "", "", "", "", "", "", _
                                          Oil_Tank_Replace_Data_Count, Oil_Tank_Replace_Data, "0")
        Else
            Call SHOW_MESSAGE_BOX("엑셀 파일이 없어 종료합니다." & vbCrLf & "-> " & sExcelFilePath, "", 1, 1)
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        sExcelFilePath = sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & "\DT_Design_Rule_Version.xlsx"
        If System.IO.File.Exists(sExcelFilePath) Then
            ReDim Data_Version_Infomation(0 To 1)
            Call Get_Excel_Value_Function(sExcelFilePath, "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                          Data_Version_Infomation_Count, Data_Version_Infomation, "0")
        Else
            Call SHOW_MESSAGE_BOX("엑셀 파일이 없어 종료합니다." & vbCrLf & "-> " & sExcelFilePath, "", 1, 1)
            Exit Sub
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(90)
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Oil_tank_weld_mach_design_table_control  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub oil_tank_weld_mach_design_table_control()
        Try
            '--------------------------------------------------------------------------------------------------- Measure Parameter 보기 pyo
            DA_Oil_Tank_Template_Create.Oil_Tank_Measure_X.Text = get_parameter_first_value(get_parameter_first_count - 7) & " mm"
            DA_Oil_Tank_Template_Create.Oil_Tank_Measure_Y.Text = get_parameter_first_value(get_parameter_first_count - 6) & " mm"

            '--------------------------------------------------------------------------------------------------- 기종에 따른 Design Table Value 가져오기
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Dim oil_tank_DT_count = 1
            Dim Check
            If Mid(assy_value_Name, 4, 1) = "1" Then
                Exc.Workbooks.Open(Oil_Tank_folder_name & "\DT_SM2100_Main_Body.xlsx")
                Check = True
                Do   'Outer loop
                    '--------------------------------------------------------------------------------------------------- Publish_Org_Name 유무 판별
                    If Exc.Worksheets.Item(1).Range("A" & oil_tank_DT_count + 1).Value = "" Then        '조건이 True이면
                        Check = False                                                                             '플래그 값을 False로 설정합니다.
                        Exit Do                                                                                       '내부 루프를 종료합니다.
                    End If
                    '--------------------------------------------------------------------------------------------------- Design Table Model Number Value
                    oil_tank_DT_model_no(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("A" & oil_tank_DT_count + 1).Value
                    '--------------------------------------------------------------------------------------------------- Design Table Part Number Value
                    If DA_Oil_Tank_Template_Create.Option_KOR.Checked = True Then
                        oil_tank_DT_partno(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("L" & oil_tank_DT_count + 1).Value
                    Else
                        oil_tank_DT_partno(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("M" & oil_tank_DT_count + 1).Value
                    End If
                    '--------------------------------------------------------------------------------------------------- Design Table Oil tank 폭 Value
                    oil_tank_DT_Width_length(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("F" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------Design Table 최대수용길이Value
                    oil_tank_DT_Total_length(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("G" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------oil_tank_DT_EPN05_D10
                    oil_tank_DT_EPN05_D10(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("B" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------oil_tank_DT_EPN05_D11
                    oil_tank_DT_EPN05_D11(oil_tank_DT_count) = Exc.Worksheets.Item(1).Range("C" & oil_tank_DT_count + 1).Value
                    oil_tank_DT_count = oil_tank_DT_count + 1
                Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.
            ElseIf Left(assy_value_Name, 2) = "SM" Then
                Exc.Workbooks.Open(Oil_Tank_folder_name & "\SM-Oil_Tank_Weld_Mach_DesignTable.xls")
                Check = True
                Do   'Outer loop
                    '--------------------------------------------------------------------------------------------------- Design Table Model Number Value
                    oil_tank_DT_model_no(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("A" & oil_tank_DT_count + 1).Value
                    '--------------------------------------------------------------------------------------------------- Design Table Part Number Value
                    If DA_Oil_Tank_Template_Create.Option_KOR.Checked = True Then
                        oil_tank_DT_partno(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("C" & oil_tank_DT_count + 1).Value
                    Else
                        oil_tank_DT_partno(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("D" & oil_tank_DT_count + 1).Value
                    End If
                    '--------------------------------------------------------------------------------------------------- Design Table Oil tank 폭 Value
                    oil_tank_DT_Width_length(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("H" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------Design Table 최대수용길이Value
                    oil_tank_DT_Total_length(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("I" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------oil_tank_DT_EPN05_D10
                    oil_tank_DT_EPN05_D10(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("G" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------oil_tank_DT_EPN05_D11
                    oil_tank_DT_EPN05_D11(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("E" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------VVIP Optioin
                    Oil_Tank_VVIP_Value(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("AB" & oil_tank_DT_count + 1).Value
                    '--------------------------------------------------------------------------------------------------- Publish_Org_Name 유무 판별
                    If Exc.Worksheets("Sheet1").Range("C" & oil_tank_DT_count + 1).Value = "" Then        '조건이 True이면
                        Check = False                                                                             '플래그 값을 False로 설정합니다.
                        Exit Do                                                                                       '내부 루프를 종료합니다.
                    End If
                    oil_tank_DT_count = oil_tank_DT_count + 1
                Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.
            ElseIf Left(assy_value_Name, 2) = "TM" Then
                Exc.Workbooks.Open(Oil_Tank_folder_name & "\TM-Oil_Tank_Weld_Mach_DesignTable.xls")
                Check = True
                Do   'Outer loop
                    '--------------------------------------------------------------------------------------------------- Design Table Model Number Value
                    oil_tank_DT_model_no(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("A" & oil_tank_DT_count + 1).Value
                    '--------------------------------------------------------------------------------------------------- Design Table Part Number Value
                    oil_tank_DT_partno(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("C" & oil_tank_DT_count + 1).Value
                    '--------------------------------------------------------------------------------------------------- Design Table Oil tank 폭 Value
                    oil_tank_DT_Width_length(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("E" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------Design Table 최대수용길이Value
                    oil_tank_DT_Total_length(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("I" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------oil_tank_DT_EPN05_D10
                    oil_tank_DT_EPN05_D10(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("G" & oil_tank_DT_count + 1).Value
                    '---------------------------------------------------------------------------------------------------oil_tank_DT_EPN05_D11
                    oil_tank_DT_EPN05_D11(oil_tank_DT_count) = Exc.Worksheets("Sheet1").Range("H" & oil_tank_DT_count + 1).Value
                    '--------------------------------------------------------------------------------------------------- Publish_Org_Name 유무 판별
                    If Exc.Worksheets("Sheet1").Range("C" & oil_tank_DT_count + 1).Value = "" Then        '조건이 True이면
                        Check = False                                                                             '플래그 값을 False로 설정합니다.
                        Exit Do                                                                                       '내부 루프를 종료합니다.
                    End If
                    oil_tank_DT_count = oil_tank_DT_count + 1
                Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.
            End If

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(False)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL") 

            DA_Oil_Tank_Template_Create.DataGridView1.RowHeadersVisible = False
            DA_Oil_Tank_Template_Create.DataGridView1.RowCount = 0
            '--------------------------------------------------------------------------------------------------- MSFlexGrid1에 Value 넣기
            DA_Oil_Tank_Template_Create.DataGridView1.ColumnCount = 5
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(0).HeaderText = ""
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(1).HeaderText = "No"
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(2).HeaderText = "Part NO."
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(3).HeaderText = "폭"
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(4).HeaderText = "길이"
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            '--------------------------------------------------------------------------------------------------- MSFlexGrid1 Width
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(0).Width = 23
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(1).Width = 40
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(2).Width = 220
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(3).Width = 80
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(4).Width = 80
            '--------------------------------------------------------------------------------------------------- 가운데 정렬
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DA_Oil_Tank_Template_Create.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Dim oil_tank_DT_no = 1
            Dim first_oil_tank_value As Integer = 1
            For i = 1 To oil_tank_DT_count - 1
                '--------------------------------------------------------------------------------------------------- SMx100 일때 MSFlexGrid1 Value
                If Mid(assy_value_Name, 4, 1) = "1" Then
                    If assy_value_Name = Left(oil_tank_DT_model_no(i), 4) Then
                        DA_Oil_Tank_Template_Create.DataGridView1.RowCount = DA_Oil_Tank_Template_Create.DataGridView1.RowCount + 1
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(1).Value = oil_tank_DT_no
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(2).Value = oil_tank_DT_partno(i)
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(3).Value = oil_tank_DT_Width_length(i)
                        DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(4).Value = oil_tank_DT_Total_length(i)

                        If oil_tank_DT_no = 1 Then
                            first_oil_tank_value = i
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(0).Value = ">"
                        End If
                        oil_tank_DT_no = oil_tank_DT_no + 1
                    End If
                    '--------------------------------------------------------------------------------------------------- SM 일때 MSFlexGrid1 Value
                ElseIf Left(assy_value_Name, 2) = "SM" Then
                    If assy_value_Name = Left(oil_tank_DT_model_no(i), 3) Then
                        '--------------------------------------------------------------------------------------------------- EPN05_D10 , EPN05_D11 값보다 작을때만...
                        If get_parameter_first_value(get_parameter_first_count - 2) < oil_tank_DT_EPN05_D10(i) And get_parameter_first_value(get_parameter_first_count - 1) <= oil_tank_DT_EPN05_D11(i) Then
                            DA_Oil_Tank_Template_Create.DataGridView1.RowCount = DA_Oil_Tank_Template_Create.DataGridView1.RowCount + 1
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(1).Value = oil_tank_DT_no
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(2).Value = oil_tank_DT_partno(i)
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(3).Value = oil_tank_DT_Width_length(i)
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(4).Value = oil_tank_DT_Total_length(i)

                            If oil_tank_DT_no = 1 Then
                                first_oil_tank_value = i
                                DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(0).Value = ">"
                            End If
                            oil_tank_DT_no = oil_tank_DT_no + 1
                        End If
                    End If
                    '--------------------------------------------------------------------------------------------------- TM 일때 MSFlexGrid1 Value
                ElseIf Left(assy_value_Name, 2) = "TM" Then
                    If assy_value_Name = oil_tank_DT_model_no(i) Then
                        '--------------------------------------------------------------------------------------------------- EPN05_D10 , EPN05_D11 값보다 작을때만...
                        If get_parameter_first_value(get_parameter_first_count - 2) < oil_tank_DT_EPN05_D10(i) And _
                           get_parameter_first_value(get_parameter_first_count - 1) <= oil_tank_DT_EPN05_D11(i) Then
                            DA_Oil_Tank_Template_Create.DataGridView1.RowCount = DA_Oil_Tank_Template_Create.DataGridView1.RowCount + 1
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(1).Value = oil_tank_DT_no
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(2).Value = oil_tank_DT_partno(i)
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(3).Value = oil_tank_DT_Width_length(i)
                            DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(4).Value = oil_tank_DT_Total_length(i)

                            If oil_tank_DT_no = 1 Then
                                first_oil_tank_value = i
                                DA_Oil_Tank_Template_Create.DataGridView1.Rows(oil_tank_DT_no - 1).Cells(0).Value = ">"
                            End If
                            oil_tank_DT_no = oil_tank_DT_no + 1
                        End If
                    End If
                End If
            Next

            '--------------------------------------------------------------------------------------------------- Design Table 변경시 기종별 찾기
            Dim Type_count_value = 0
            If assy_value_Name = "SM3" Or assy_value_Name = "SM4" Or assy_value_Name = "SM6" Or assy_value_Name = "TMY" Or assy_value_Name = "TMZ" Then
                For i = 1 To oil_tank_DT_count
                    If assy_value_Name = "SM3" Or assy_value_Name = "SM4" Or assy_value_Name = "SM6" Then
                        If assy_value_Name = Left(oil_tank_DT_model_no(i), 3) Then
                            Type_count_value = i
                            Exit For
                        End If
                    Else
                        If assy_value_Name = oil_tank_DT_model_no(i) Then
                            Type_count_value = i
                            Exit For
                        End If
                    End If
                Next
            ElseIf assy_value_Name = "SM5" Or assy_value_Name = "TMX" Then
                Type_count_value = 0
            End If

            '--------------------------------------------------------------------------------------------------- 선택되는 Data가 없을때....
            If oil_tank_DT_no = 1 Then
                oil_tank_DT_no = 1
                For i = 1 To oil_tank_DT_count - 1
                    '--------------------------------------------------------------------------------------------------- SM 일때 표준없는경우
                    If Left(assy_value_Name, 2) = "SM" Then
                        If assy_value_Name = Left(oil_tank_DT_model_no(i), 3) Then
                            first_oil_tank_value = i
                        End If
                        '--------------------------------------------------------------------------------------------------- TM 일때 표준없는경우
                    ElseIf Left(assy_value_Name, 2) = "TM" Then
                        If assy_value_Name = oil_tank_DT_model_no(i) Then
                            first_oil_tank_value = i
                        End If
                    End If
                Next
                Call SHOW_MESSAGE_BOX("일치하는 표준형Data가 없습니다. 기종의 제일 큰 표준형이 선택됩니다.", "", 1, 1)
            End If

            '--------------------------------------------------------------------------------------------------- Part List 1번에 Data 이름 넣기
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Oil_Tank_folder_name & "\" & STW_Template_Infomation(3, 7))

            '--------------------------------------------------------------------------------------------------- Data 검토
            PartList_No = CInt(Exc.Worksheets.Item(1).Range("S1").Value)
            Exc.Worksheets.Item(1).Range("B" & PartList_No).Value = DA_Oil_Tank_Template_Create.DataGridView1.Rows(0).Cells(2).Value

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(True)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")
            selection.clear
            '--------------------------------------------------------------------------------------------------- Design Table Search
            selection.search("Name='Oil_Tank_Weld_DesignTable',all")

            '--------------------------------------------------------------------------------------------------- SMx100 일때 MSFlexGrid1 Value
            If Mid(assy_value_Name, 4, 1) = "1" Then
                selection.Item(1).Value.Value = first_oil_tank_value
            Else
                selection.Item(1).Value.configuration = first_oil_tank_value
            End If
            ActiveDocument.Product.Update()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("oil_tank_weld_mach_design_table_control() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Oil Tank Data Replace                               &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Oil_Tank_Data_Replace()
        On Error Resume Next        ' 템플릿 경로에 파트가 없는 경우, 그냥 넘기기 위해 사용..
        '-------------------------------------------------------------------------------------------------------- PartList Excel Open
        fso = CreateObject("Scripting.FileSystemObject")

        Exc = CreateObject("excel.application")
        'Exc.Visible = True
        Exc.Workbooks.Open(Copy_Oil_Tank_Path & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

        Dim Replace_Data_Name_Len
        Dim Replace_Data
        Dim data_replace

        Dim sw = 1
        ReDim Oil_Tank_Replace_Data_Value(0 To Oil_Tank_Replace_Data_Count - 1)
        For i = 1 To Input_Output_Excel_Value_Count - 1
            Dim sw2 = 0
            If Input_Output_Value(4, i) = "I" Then
                '-------------------------------------------------------------------------------------------------------- Excel 읽어서 Parameter에 입력.
                Call Input_Output_Excel_to_LayoutParameter(1, Input_Output_Value, i, 2, 1, O_BOM_Part_Value, 3, 1)
                '-------------------------------------------------------------------------------------------------------- Data Replace
                For j = 2 To O_BOM_Part_Value_Count - 1
                    If Input_Output_Value(2, i) = O_BOM_Part_Value(j, 3) Then
                        '--------------------------------------------------------------------------------------------------- 폴더안에 Data 가져오기
                        A_Main_Form.File_numbering.Items.Add(sDirectory_Path_Text & "\" & assy_value_Name & "\" & Input_Output_Value(2, i))

                        '--------------------------------------------------------------------------------------------------- 폴더안에 Data중 "_" 확인
                        For m = 0 To A_Main_Form.File_numbering.Items.Count - 1
                            If CStr(O_BOM_Part_Value(j, 1) & ".CATPart") = CStr(A_Main_Form.File_numbering.Items.Item(m)) Then
                                Replace_Data_Name_Len = Len(O_BOM_Part_Value(j, 1))
                                Oil_Tank_Replace_Data_Value(i) = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Input_Output_Value(2, i) & "\" & CStr(A_Main_Form.File_numbering.Items.Item(m))
                                '--------------------------------------------------------------------------------------------------- File 복사
                                fso.CopyFile(Oil_Tank_Replace_Data_Value(i), Copy_Oil_Tank_Path & "\")
                                Threading.Thread.Sleep(1000)
                                '--------------------------------------------------------------------------------------------------- Part Replace Part
                                Replace_Data = CATIA.ActiveDocument.Product.products.Item(Input_Output_Value(2, i))
                                data_replace = CATIA.ActiveDocument.Product.products.ReplaceComponent(Replace_Data, Oil_Tank_Replace_Data_Value(i), False)
                                Threading.Thread.Sleep(1000)
                                sw = sw + 1
                                sw2 = 1
                                Exit For
                            Else
                                Replace_Data_Name_Len = Len(O_BOM_Part_Value(j, 1))
                                If CStr(O_BOM_Part_Value(j, 1)) = Left(A_Main_Form.File_numbering.Items(m), Replace_Data_Name_Len) Then
                                    If Mid(A_Main_Form.File_numbering.Items.Item(m), Replace_Data_Name_Len + 1, 1) = "_" Then
                                        Oil_Tank_Replace_Data_Value(sw) = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Input_Output_Value(2, i) & "\" & CStr(A_Main_Form.File_numbering.Items.Item(m))

                                        '--------------------------------------------------------------------------------------------------- File 복사
                                        fso.CopyFile(Oil_Tank_Replace_Data_Value(sw), Copy_Oil_Tank_Path & "\")
                                        Threading.Thread.Sleep(1000)

                                        '--------------------------------------------------------------------------------------------------- Part Replace Part
                                        Replace_Data = CATIA.ActiveDocument.Product.products.Item(Input_Output_Value(2, i))
                                        data_replace = CATIA.ActiveDocument.Product.products.ReplaceComponent(Replace_Data, Copy_Oil_Tank_Path & "\" & A_Main_Form.File_numbering.Items.Item(m), False)
                                        Threading.Thread.Sleep(1000)
                                        sw = sw + 1
                                        sw2 = 1
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If sw2 = 1 Then Exit For
                Next j
            End If
        Next i

        Exc.Application.DisplayAlerts = False
        Exc.ActiveWorkbook.Close(True)
        Exc.Application.DisplayAlerts = True
        Exc.Quit()

        Call KillProcess("EXCEL") 
        ActiveDocument.Product.Update()
        '--------------------------------------------------------------------------------------------------- Oil Tank Constraint Delete & Create
        For i = 1 To Oil_Tank_Selection_Constraint_Count - 1
            Call Constraint_Delete(Oil_Tank_Initial_Constraint(i, 1), Oil_Tank_Initial_Constraint(i, 2), Oil_Tank_Initial_Constraint(i, 3))
            Call Constraint_Delete(Oil_Tank_Selection_Constraint(i, 1), Oil_Tank_Selection_Constraint(i, 2), Oil_Tank_Selection_Constraint(i, 3))
        Next
        On Error GoTo 0
    End Sub

End Module
