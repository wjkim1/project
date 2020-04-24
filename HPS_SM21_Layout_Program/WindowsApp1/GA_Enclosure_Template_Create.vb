Public Class GA_Enclosure_Template_Create
    Public oEPN18_D02, oEPN0118_D01


    Private Sub GA_Enclosure_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용
            Me.Top = 150
            Me.Left = 50
            Me.Width = 379
            Me.Height = 742
            Me.TopMost = True

            Call CATIA_BASIC02()

            '--------------------------------------------------------------------------------------------------- 
            Enclosure_Drawing_Create.Enabled = False
            Enclosure_Result_Part_Create.Enabled = False

            '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
            Call Data_O_BOM_Excel_Search()

            '--------------------------------------------------------------------------------------------------- Enclosure_Excel_Open
            If Enclosure_Excel_Open() = False Then
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- Enclosure_Template_Form_Inicial_Setting
            UI_STAGE_COMBO.Items.Add("2")
            UI_STAGE_COMBO.Items.Add("3")
            UI_STAGE_COMBO.Items.Add("4")
            '--------------------------------------------------------------------------------------------------  VVIP
            Enclosure_Parameter_Control.Enabled = False
            CONTROL_PANEL_Text.Enabled = False
            '--------------------------------------------------------------------------------------------------- 모터 상판 형상 변경
            Enclosure_Standard_Option.Enabled = False
            Enclosure_Remove_Option.Enabled = False
            Enclosure_Extrusion_Option.Enabled = False
            Enclosure_Motor_Top_Panel_Update.Enabled = False

            '--------------------------------------------------------------------------------------------------- 도면 관련 정보 가져오기
            sExcelFilePath = Application.StartupPath & "\STW_Drawing_Infomation.xlsx"
            If System.IO.File.Exists(sExcelFilePath) Then
                Call Get_Excel_Value_Function(sExcelFilePath, assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", _
                                              STW_Drawing_Infomation_Count, STW_Drawing_Infomation, "5")
            Else
                Call SHOW_MESSAGE_BOX("엑셀 파일이 없어 종료합니다." & vbCrLf & "-> " & sExcelFilePath, "", 1, 1)
                Exit Sub
            End If

            Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)

            '--------------------------------------------------------------------------------------------------- Input_Output Value  가져오기
            Call Input_Output_Excel_Value(STW_Template_Infomation(Form_Click_Index_Value, 1))

            '--------------------------------------------------------------------------------------------------- VVIP 활성화
            A_Main_Form.Lbl_ModelNumber.Text = "SM21"
            If Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
                Height = 800
                Frame4.Visible = False
                Frm_Enclosure.Visible = True
                Frm_Enclosure.Left = 12
                Enclosure_Parameter_Control.Top = 600
                Enclosure_Clash_Check.Top = 600
                Enclosure_Drawing_Create.Top = 640
                Enclosure_Result_Part_Create.Top = 680
                Message_Label.Top = 717
                ProgressBar1.Top = 739

                UI_PANEL_OUTSIDE.Enabled = False
                Label7.Enabled = False
                CONTROL_PANEL_Text.Enabled = False
            Else
                Height = 740
                Frame4.Visible = True
                Frm_Enclosure.Visible = False
                Enclosure_Parameter_Control.Top = 541
                Enclosure_Clash_Check.Top = 541
                Enclosure_Drawing_Create.Top = 581
                Enclosure_Result_Part_Create.Top = 621
                Message_Label.Top = 658
                ProgressBar1.Top = 680
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("GA_Enclosure_Template_Create_Load() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Enclosure_Template_Create_Click(sender As Object, e As EventArgs) Handles Enclosure_Template_Create.Click
        '--------------------------------------------------------------------------------------------------- Enclosure Part NO 가 없을때
        If Template_Part_Number_Text.Text = "" Then
            Call SHOW_MESSAGE_BOX("Enclosure Part No.가 없습니다.", "", 1, 1)
            Exit Sub
        End If

        Try
            '--------------------------------------------------------------------------------------------------- Template 실행여부 확인
            If SHOW_MESSAGE_BOX("Template01 실행 하겠습니까?", "", 2, 1) = vbYes Then
                ProgressBar1.Value = 0
                Message_Label.Text = "Template01 생성중..."
                '--------------------------------------------------------------------------------------------------- Template Data Folder Name
                ls_ss = Split(STW_Template_Infomation(Form_Click_Index_Value, 4), "\")
                Template_Data_Folder_Name = ls_ss(UBound(ls_ss))

                Call Folder_Delete(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value)

                fso.CopyFolder(sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4), sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\")

                Threading.Thread.Sleep(300)
                Dim Enclosure_folder = fso.GetFolder(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name)

                Enclosure_folder.Name = Template_Data_Folder_Name & "_" & date_value
                Enclosure_folder_name = Enclosure_folder.Path

                '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
                Call wirte_transfer_file(Enclosure_folder_name)
                ProgressBar1.Value = 10

                '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
                Call Clash_Check_Folder_Create()
                ProgressBar1.Value = 20

                '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
                Call Data_Constraint_Delete(Enclosure_Initial_Constraint, Enclosure_Initial_Constraint_Count, Result_Template_Item_Value)
                ProgressBar1.Value = 30

                '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
                Call Get_Parameter_Data_Comparison(Enclosure_Initial_Parameter, Enclosure_Initial_Parameter_Count, Me)

                '--------------------------------------------------------------------------------------------------- Enclosure Template Data Open
                Call CATIA_BASIC02()

                '--------------------------------------------------------------------------------------------------- Template Data Open
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Enclosure_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5)
                products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                Threading.Thread.Sleep(500)

                Dim Template_Skeleton_Part = Documents.Item("SKELETON.CATPart").Part

                '--------------------------------------------------------------------------------------------------- Input 변경
                For i = 1 To Input_Output_Excel_Value_Count - 1
                    If Input_Output_Value(4, i) = "D" Then
                        Call Constraint_Delete(Input_Output_Value(1, i), Input_Output_Value(2, i), Input_Output_Value(3, i))
                    ElseIf Input_Output_Value(4, i) = "A" Then
                        Call Input_Output_Axis_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                    ElseIf Input_Output_Value(4, i) = "B" Then
                        Call Input_Output_Element_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                    ElseIf Input_Output_Value(4, i) = "C" Then
                        Call Input_Output_Initial_Parameter(Input_Output_Value(1, i), Input_Output_Value(2, i))
                        '--------------------------------------------------------------------------------------------------- Oil_Tank Template TYPE
                    ElseIf Input_Output_Value(4, i) = "L" Then
                        Call Input_Output_UI_to_LayoutParameter_L(Me, Input_Output_Value(1, i), Input_Output_Value(2, i))
                    ElseIf Input_Output_Value(4, i) = "G" Then
                        Call Input_Output_Insert_G_Type()
                    ElseIf Input_Output_Value(4, i) = "H" Then
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        Exc = CreateObject("excel.application")
                        'Exc.Visible = True
                        Exc.Workbooks.Open(Enclosure_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

                        Call Input_Output_Excel_to_Excel(Exc, 1, Input_Output_Value(1, i), Input_Output_Value(2, i), O_BOM_Part_Value, 3, 1)

                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Close(True)
                        Exc.Application.DisplayAlerts = True
                        Exc.Quit()

                        Call KillProcess("EXCEL") 
                    ElseIf Input_Output_Value(4, i) = "K" Then
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        Exc = CreateObject("excel.application")
                        'Exc.Visible = True
                        Exc.Workbooks.Open(Enclosure_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

                        Call Input_Output_LayoutParameter_to_Excel(Exc, STW_Drawing_Infomation(1, 4), Input_Output_Value(1, i), Input_Output_Value(2, i))

                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Close(True)
                        Exc.Application.DisplayAlerts = True
                        Exc.Quit()

                        Call KillProcess("EXCEL") 
                    End If
                Next

                ProgressBar1.Value = 50
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value       'hwheo 180718 수정

                '--------------------------------------------------------------------------------------------------- Enclosure Template Open
                Dim Enclosure_open = CATIA.Documents.Open(Enclosure_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5))

                Call CATIA_BASIC02()
                CATIA.ActiveDocument.Product.Update()

                '---------------------------------------------------------------------------------------------------
                If Mid(assy_value_Name, 4, 1) = "1" Then
                    Call Put_Result_Data()
                End If

                '--------------------------------------------------------------------------------------------------- Enclosure Option Setting
                Call Enclosure_Option_Setting()
                ProgressBar1.Value = 60

                '--------------------------------------------------------------------------------------------------- Template Data Get Parameter
                Call Template_Data_Get_Parameter()
                ProgressBar1.Value = 70

                '--------------------------------------------------------------------------------------------------- VVIP Option Check
                If VVIP_Check.Checked = True Then
                    '--------------------------------------------------------------------------------------------------- Input 변경
                    For i = 1 To Input_Output_Excel_Value_Count - 1
                        If Input_Output_Value(4, i) = "M" Then
                            Call Input_Output_UI_to_LayoutParameter_M(Me, Input_Output_Value(1, i), Input_Output_Value(2, i))
                        End If
                    Next i
                    selection.clear
                    selection.search("Name='CONTROL_PANEL_외장_폭',all")
                    If Not selection.Count = 0 Then
                        CONTROL_PANEL_Text.Text = selection.Item(1).Value.Value
                        Enclosure_Parameter_Update.Enabled = True
                    End If

                    If UI_CONTROL_PANEL.Checked = True Then
                        CONTROL_PANEL_Text.Enabled = True
                    End If
                End If
                ProgressBar1.Value = 90

                '--------------------------------------------------------------------------------------------------- Product File Save
                CATIA.ActiveDocument.Product.Update()
                CATIA.ActiveDocument.Save()

                '--------------------------------------------------------------------------------------------------- 모터 상판 형상 Option 가져오기
                Call Enclosure_Motor_Top_Panel_Function()
            Else
                Exit Sub
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Enclosure_Template_Create_Click() 함수 오류!!", "", 1, 1)
        End Try

        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Product.Update()

        '--------------------------------------------------------------------------------------------------- Enclosure 간섭체크
        If Auto_Clash_Check.Checked = True Then
            Call Enclosure_Clash_Check_Click(sender, Nothing)
        End If

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template01 생성완료"
    End Sub


    Private Sub Enclosure_Clash_Check_Click(sender As Object, e As EventArgs) Handles Enclosure_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Enclosure 간섭 체크중.."
        Call Template_Clash_Check_Function()
        Enclosure_Drawing_Create.Enabled = True
        Enclosure_Result_Part_Create.Enabled = True
        Message_Label.Text = "Enclosure 간섭체크 완료"
    End Sub


    Private Sub Enclosure_Drawing_Create_Click(sender As Object, e As EventArgs) Handles Enclosure_Drawing_Create.Click

        ProgressBar1.Value = 0
        Message_Label.Text = "Template02 생성중..."

        ProgressBar1.Value = 10

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
        '--------------------------------------------------------------------------------------------------- COMP 저장
        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Save()

        '--------------------------------------------------------------------------------------------------- Title Block 생성
        Dim Enclosure_drawing_open = CATIA.Documents.NewFrom(Enclosure_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 6))
        ProgressBar1.Value = 40

        Call CATIA_BASIC03()

        '--------------------------------------------------------------------------------------------------- 도면 정보에서 Item 가져오기
        Dim Drawing_Item_Folder_Value = Enclosure_folder_name
        Dim Drawing_Item_Count
        For Drawing_Item_Count = 1 To 20
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = STW_Drawing_Infomation(0, Drawing_Item_Count) Then
                Exit For
            End If
        Next

        '--------------------------------------------------------------------------------------------------- 도면 Part List 생성하기
        For i = 1 To STW_Drawing_Infomation_Count
            If Not STW_Drawing_Infomation(i, STW_Drawing_Infomation_Count) = "" Then
                DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(i, Drawing_Item_Count), i)
            End If
        Next

        ProgressBar1.Value = 60

        On Error Resume Next
        Drawingselection.clear
        '--------------------------------------------------------------------------------------------------- 도면 DWG_No 변경
        Drawingselection.search("Name=Model*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = STW_Drawing_Infomation(1, 1)
        Next
        Drawingselection.clear
        '--------------------------------------------------------------------------------------------------- DWG_No
        Drawingselection.search("Name=DWG_No*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Template_Part_Number_Text.Text
        Next
        On Error GoTo 0

        ProgressBar1.Value = 80
        '--------------------------------------------------------------------------------------------------- 표제란 입력
        Call drawing_title_box_write()
        DrawingViews.Item("Main View").Activate()
        '--------------------------------------------------------------------------------------------------- 파일 Update
        CATIA.ActiveDocument.Update()
        '--------------------------------------------------------------------------------------------------- 도면 파일 저장
        CATIA.ActiveDocument.SaveAs(Enclosure_folder_name & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATDrawing")

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template02 생성완료"
    End Sub


    Private Sub Enclosure_Result_Part_Create_Click(sender As Object, e As EventArgs) Handles Enclosure_Result_Part_Create.Click
        Try
            ProgressBar1.Value = 0
            Message_Label.Text = "Template03 Part 변환..."

            ProgressBar1.Value = 5
            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
            '--------------------------------------------------------------------------------------------------- Template 파일 저장 및 Generate CATPart from Product 명령어 실행
            CATIA.ActiveDocument.Save()
            selection.Clear()
            '--------------------------------------------------------------------------------------------------- Product -> Part 변환
            Call Generate_CATPart_from_Product()
            '--------------------------------------------------------------------------------------------------- Geometrycal 이름에 "\" 기호 삭제
            Call Delete_Symbol()

            Message_Label.Text = "Template03 Part 확정중..."

            '--------------------------------------------------------------------------------------------------- PartNumber 변경
            CATIA.ActiveDocument.Product.PartNumber = Template_Part_Number_Text.Text & "_" & date_value
            publications1 = CATIA.ActiveDocument.Product.publications
            For i = 1 To Input_Output_Excel_Value_Count - 1
                '--------------------------------------------------------------------------------------------------- Publication 생성
                If Input_Output_Value(4, i) = "E" Then
                    Call publication_axis_create(Input_Output_Value(1, i))
                    '--------------------------------------------------------------------------------------------------- Parameter 생성
                ElseIf Input_Output_Value(4, i) = "F" Then
                    Call part_parameter_create(Input_Output_Value(1, i), Input_Output_Parameter_Value(i), Input_Output_Value(2, i))
                End If
            Next
            '--------------------------------------------------------------------------------------------------- 동일한 이름의 Data가 있는지 확인
            Call Save_File_Check(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            '--------------------------------------------------------------------------------------------------- Part File Save
            CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            '--------------------------------------------------------------------------------------------------- Part Data Close
            CATIA.ActiveDocument.Close()

            '--------------------------------------------------------------------------------------------------- 도면 파일 Close
            Call CATIA_BASIC02()

            Call Active_Window_Change("CATDrawing", 1, 10)
            If Strings.Right(CATIA.ActiveDocument.Name, 10) = "CATDrawing" Then
                CATIA.ActiveDocument.Close()
            End If

            '--------------------------------------------------------------------------------------------------- Template 파일 Close
            Call CATIA_BASIC02()

            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
            If CATIA.ActiveDocument.Name = STW_Template_Infomation(Form_Click_Index_Value, 5) Then
                CATIA.ActiveDocument.Close()
            End If

            '--------------------------------------------------------------------------------------------------- Enclosure 초기위치 삭제
            Call Constraint_Delete("EPN01_P16", "TR_EPN01_P16", 1)
            '--------------------------------------------------------------------------------------------------- Control Panel Constraint 삭제
            Call Constraint_Delete("EPN11_P03", "EPN10_P01", 1)
            ProgressBar1.Value = 30

            '--------------------------------------------------------------------------------------------------- COMP 에서 Template Data 제거
            Call CATIA_BASIC02()
            selection.Clear()
            selection.Add(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count))
            selection.Delete()
            Threading.Thread.Sleep(500)

            '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
            Call Data_Constraint_Delete(Enclosure_Initial_Constraint, Enclosure_Initial_Constraint_Count, Result_Template_Item_Value)
            '--------------------------------------------------------------------------------------------------- COMP Data에 실적 DATA 삽입
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart"
            CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            Threading.Thread.Sleep(500)

            '--------------------------------------------------------------------------------------------------- Enclosure EPN 변경
            If Strings.Left(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name, 1) <> "E" Then
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
            End If

            '--------------------------------------------------------------------------------------------------- Enclosure Data 저장
            Dim Enclosure_Part = Documents.Item(Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            Enclosure_Part.Save()
            ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- Enclosure Constraint Delete & Create
            For i = 1 To Enclosure_Selection_Constraint_Count - 1
                Call Constraint_Delete(Enclosure_Selection_Constraint(i, 1), Enclosure_Selection_Constraint(i, 2), Enclosure_Selection_Constraint(i, 3))
            Next
            ProgressBar1.Value = 80

            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Call O_BOM_Excel_Writing(Result_Template_Item_Value, Template_Part_Number_Text.Text & "_" & date_value)
            Call Btn_Refresh_Click()

            '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
            fso = CreateObject("Scripting.FileSystemObject")
            Dim file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            file_Attributes_value.Attributes = 33

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Message_Label.Text = "Template04 생성완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Enclosure_Result_Part_Create_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(379, 668, Me)
    End Sub


    Private Sub Enclosure_Motor_Top_Panel_Update_Click(sender As Object, e As EventArgs) Handles Enclosure_Motor_Top_Panel_Update.Click
        Template_Data_Path_text.text = "Update 중..."

        '--------------------------------------------------------------------------------------------------- 흡음제 제거 선택
        If Enclosure_Remove_Option.Checked = True Then
            TOP_MOTOR_PNL_OPTION.Value = "SPLIT"
        ElseIf Enclosure_Extrusion_Option.Checked = True Then
            TOP_MOTOR_PNL_OPTION.Value = "EXTRUDE"
        End If

        CATIA.ActiveDocument.Product.Update()

        Template_Data_Path_text.text = "Update 완료"
    End Sub


    Private Sub LEN_ENC_L_Expand_Update_Click(sender As Object, e As EventArgs) Handles LEN_ENC_L_Expand_Update.Click
        Me.Message_Label.Text = "Enclosure 길이조정 중..."

        If Not LEN_ENC_L_Expand_Text.Text = "" Then
            LEN_ENC_L_Expand.Value = LEN_ENC_L_Expand_Text.Text
            ActiveDocument.Product.Update()
            Threading.Thread.Sleep(300)
        End If
        selection.clear
        selection.search("Name='DIST_ENCLOSURE_L',all")
        If Not selection.Count = 0 Then
            DIST_ENCLOSURE_L = selection.Item(1).Value
            NOW_ENC_L.Text = DIST_ENCLOSURE_L.Value
            selection.Clear()
            ActiveDocument.Product.Update()
        End If

        Me.Message_Label.Text = "Enclosure 길이조정 완료"
    End Sub


    Private Sub Enclosure_Parameter_Update_Click(sender As Object, e As EventArgs) Handles Enclosure_Parameter_Update.Click
        VVIP_Message.Text = "Update 진행중"

        Call VVIP_Option(VVIP_Check, UI_CANOPY, "CANOPY적용")
        Call VVIP_Option(VVIP_Check, UI_DOOR_LOCK, "VVIP_DOOR_LOCK")
        Call VVIP_Option(VVIP_Check, UI_PANEL_OUTSIDE, "CONTROL_PANEL_외장")
        selection.clear
        '--------------------------------------------------------------------------------------------------- CONTROL_PANEL_외장_폭 변경
        selection.search("Name='CONTROL_PANEL_외장_폭',all")
        If Not selection.Count = 0 Then
            Dim CONTROL_PANEL = selection.Item(1).Value
            CONTROL_PANEL.Value = CONTROL_PANEL_Text.Text

            '--------------------------------------------------------------------------------------------------- Product File Save
            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Save()
        End If

        VVIP_Message.Text = "Update 완료"
    End Sub


    Private Sub Enclosure_Parameter_Control_Click(sender As Object, e As EventArgs) Handles Enclosure_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- 화면 전환
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        Call Get_Template_Parameter_Excel_Value(STW_COMP_Infomation(7, 6))
        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()
        Call Template_Parameter_Modify(STW_COMP_Infomation(7, 6))
        Call Folder_File_Count()
    End Sub

End Class