Public Class HA_Base_Frame_Template_Create


    Private Sub HA_Base_Frame_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- Form Position
        Me.Top = 150
        Me.Left = 50
        Me.Width = 400
        Me.Height = 645
        Me.TopMost = True

        Base_Frame_Drawing_Create.Enabled = False
        Base_Frame_Result_Part_Create.Enabled = False
        VVIP_Frame.Visible = False
        '--------------------------------------------------------------------------------------------------- 기종에 따라 폼 디자인 변경
        Frame5.Enabled = False
        Frm_BaseFrame_2.Visible = False
        Option_Motor_Default.Enabled = False
        Option_Under_8.Enabled = False
        Option_Over_8.Enabled = False

        If A_Main_Form.Lbl_ModelNumber.Text = "SM61" Then
            Frame5.Enabled = True
            Option_Motor_Default.Enabled = True
            Option_Under_8.Enabled = True
            Option_Over_8.Enabled = True
            Frm_BaseFrame_2.Visible = True
        ElseIf Strings.Right(A_Main_Form.Lbl_ModelNumber.Text, 1) = "1" Then
            Frm_BaseFrame_1.Visible = False
            Frm_BaseFrame_2.Visible = True
            Frm_BaseFrame_2.Left = 12
        Else
            Frm_BaseFrame_1.Visible = True
            Frm_BaseFrame_2.Visible = False
        End If

        Call CATIA_BASIC02()

        Template_Data_Loading_Form.ProgressBar1_Change(20)

        '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
        Call Data_O_BOM_Excel_Search()

        '--------------------------------------------------------------------------------------------------- Base_Frame_Excel_Open
        Call Base_Frame_Excel_Open()

        '--------------------------------------------------------------------------------------------------- 도면 관련 정보 가져오기
        Call Get_Excel_Value_Function(Application.StartupPath & "\STW_Drawing_Infomation.xlsx", assy_value_Name, 9, _
                                      "A", "B", "C", "D", "E", "F", "G", "H", "I", "", STW_Drawing_Infomation_Count, STW_Drawing_Infomation, "5")

        Template_Data_Loading_Form.ProgressBar1_Change(80)

        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)

        '--------------------------------------------------------------------------------------------------- 활성화
        If Strings.Left(A_Main_Form.Lbl_ModelNumber.Text, 2) = "SM" And Not Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
            Auto_Trap_Option1.Enabled = True
            Auto_Trap_Option2.Enabled = True
            VVIP_Check.Enabled = True
        End If
    End Sub


    Private Sub Base_Frame_Template_Create_Click(sender As Object, e As EventArgs) Handles Base_Frame_Template_Create.Click, Base_Frame_Template_Create2.Click
        Try
            '--------------------------------------------------------------------------------------------------- Base_Frame Part NO 가 없을때
            If Template_Part_Number_Text.Text = "" Then
                Call SHOW_MESSAGE_BOX("Base_Frame Part No.가 없습니다.", "", 1, 1)
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- Motor 무게
            If Frame5.Enabled = True Then
                If Option_Motor_Default.Text = True Then
                    MsgBox("Motor 무게를 선택하세요.", vbInformation, "ISPark_Automation")
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------------------------------------- Enclosure 유무
            If Option_Default.Checked = True Then
                MsgBox("Enclosure 유무를 선택하세요.", vbInformation, "ISPark_Automation")
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- Template 실행여부 확인
            If SHOW_MESSAGE_BOX("Template01 실행 하겠습니까?", "", 2, 1) = vbYes Then
                ProgressBar1.Value = 0
                Message_Label.Text = "Template01 생성중..."

                '--------------------------------------------------------------------------------------------------- Template Data Folder Name
                ls_ss = Split(STW_Template_Infomation(Form_Click_Index_Value, 4), "\")
                Template_Data_Folder_Name = ls_ss(UBound(ls_ss))

                Call Folder_Delete(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value)
                fso.CopyFolder(sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4), sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\")
                Threading.Thread.Sleep(500)

                Dim Base_Frame_folder = fso.GetFolder(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name)
                Base_Frame_folder.Name = Template_Data_Folder_Name & "_" & date_value
                Base_Frame_folder_name = Base_Frame_folder.Path

                '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
                Call wirte_transfer_file(Base_Frame_folder_name)
                ProgressBar1.Value = 10

                '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
                Call Clash_Check_Folder_Create()
                ProgressBar1.Value = 20

                '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
                Call Data_Constraint_Delete(Base_Frame_Initial_Constraint, Base_Frame_Initial_Constraint_Count, Result_Template_Item_Value)
                ProgressBar1.Value = 30

                '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
                Call Get_Parameter_Data_Comparison(Base_Frame_Initial_Parameter, Base_Frame_Initial_Parameter_Count, Me)

                '--------------------------------------------------------------------------------------------------- Base_Frame Template Data Open
                Call CATIA_BASIC02()
                '--------------------------------------------------------------------------------------------------- Template Data Open
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Base_Frame_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5)
                products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

                Dim Template_Skeleton_Part = Documents.Item("#Skeleton_Base_Frame.CATPart").Part
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
                        If Exc.Name = "" Then
                            Exc = CreateObject("excel.application")
                            Exc.Workbooks.Open(Base_Frame_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))
                            'Exc.Visible = True
                        End If

                        Call Input_Output_Excel_to_Excel(Exc, 1, Input_Output_Value(1, i), Input_Output_Value(2, i), O_BOM_Part_Value, 3, 1)

                        If Not Input_Output_Value(4, i + 1) = "H" Then
                            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
                            Exc.ActiveWorkbook.Save()
                            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
                            Exc.Quit()
                        End If
                    ElseIf Input_Output_Value(4, i) = "K" Then
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        If Exc.Name = "" Then
                            Exc = CreateObject("excel.application")
                            Exc.Workbooks.Open(Base_Frame_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))
                            'Exc.Visible = True
                        End If

                        Call Input_Output_LayoutParameter_to_Excel(Exc, "PartList", Input_Output_Value(1, i), Input_Output_Value(2, i))

                        If Not Input_Output_Value(4, i + 1) = "K" Then
                            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
                            Exc.ActiveWorkbook.Save()
                            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
                            Exc.Quit()
                        End If
                    End If
                Next
                CATIA.ActiveDocument.Product.Update()
                ProgressBar1.Value = 50

                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value

                '--------------------------------------------------------------------------------------------------- Base_Frame Template Open
                Dim Base_Frame_open = CATIA.Documents.Open(Base_Frame_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5))

                Call CATIA_BASIC02()
                CATIA.ActiveDocument.Product.Update()

                '---------------------------------------------------------------------------------------------------
                If Mid(assy_value_Name, 4, 1) = "1" Then
                    Call Base_Frame_weld_mach_design_table_control()
                End If

                '--------------------------------------------------------------------------------------------------- Base_Frame Option Setting
                Call Base_Frame_Option_Setting_2()
                ProgressBar1.Value = 70

                '--------------------------------------------------------------------------------------------------- VVIP Option Check
                If VVIP_Check.Checked = True And UI_CONTROL_PNL.Checked = True Then
                    '--------------------------------------------------------------------------------------------------- Input 변경
                    For i = 1 To Input_Output_Excel_Value_Count - 1
                        If Input_Output_Value(4, i) = "M" Then
                            Call Input_Output_UI_to_LayoutParameter_M(Me, Input_Output_Value(1, i), Input_Output_Value(2, i))
                        End If
                    Next i
                    Base_Frame_Parameter_Update.Enabled = True
                End If
                ProgressBar1.Value = 90

                '--------------------------------------------------------------------------------------------------- Product File Save
                CATIA.ActiveDocument.Product.Update()
                CATIA.ActiveDocument.Save()
            Else
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- Base_Frame 간섭체크
            If Auto_Clash_Check.Checked = True Then
                Call Base_Frame_Clash_Check_Click(sender, Nothing)
            End If

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            '--------------------------------------------------------------------------------------------------- 업데이트 한번 더
            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Save()
            ProgressBar1.Value = 100

            '--------------------------------------------------------------------------------------------------- Base_Frame 길이 조정
            If UI_ENCLOSURE_TYPE_2.Checked = True Then        'UI_ENCLOSURE_TYPE_2 : No : 수정 가능
                LEN_BASE_L_Expand_Text.Enabled = True
            ElseIf UI_ENCLOSURE_TYPE.Checked = True Or Option_Default.Checked = True Then   'UI_ENCLOSURE_TYPE : Yes / Option_Default : Default : 수정 불가능
                LEN_BASE_L_Expand_Text.Enabled = True
            End If

            '--------------------------------------------------------------------------------------------------- 레이아웃 Product, 베이스프레임 Product 선언
            Dim Product_Main = Nothing
            Dim Product_BaseFrame = Nothing
            '--------------------------------------------------------------------------------------------------- 레이아웃 Product 경로
            Dim Product_Main_Path = Nothing

            '--------------------------------------------------------------------------------------------------- 레이아웃 selection, 베이스프레임 selection
            Dim EPN1101_EPN11_D02
            Dim Product_Main_Sel, Product_BaseFrame_Sel
            If UI_ENCLOSURE_TYPE.Checked = True Then
                Dim O_BOM_File_Name = Nothing
                If O_BOM_File_Name = "" Then
                    Product_Main_Path = Replace(Strings.Right(now_o_bom_excel_path, InStrRev(now_o_bom_excel_path, "\") + 1), ".xlsx", "")
                End If
                Product_Main = CATIA.Documents.Item(Product_Main_Path & ".CATProduct")
                Product_Main.Activate()
                Product_Main_Sel = Product_Main.selection
                Product_Main_Sel.Clear()
                Product_Main_Sel.Clear()
                Product_Main_Sel.search("CATProductSearch.Part.Name=EPN1101,all")

                Product_BaseFrame = CATIA.Documents.Item(STW_Template_Infomation(Form_Click_Index_Value, 5))
                If Product_Main_Sel.Count <> 0 Then
                    Product_Main_Sel.search("(Name=EPN11_D02),sel")
                    If Product_Main_Sel.Count = 0 Then
                        GoTo End_If
                    End If
                    EPN1101_EPN11_D02 = Product_Main_Sel.Item(1).Value.Value
                    Product_BaseFrame = CATIA.Documents.Item(STW_Template_Infomation(Form_Click_Index_Value, 5))
                    Product_BaseFrame_Sel = Product_BaseFrame.selection
                    Product_BaseFrame_Sel.Clear()
                    Product_Main_Sel.search("(Name=EPN18_MANUAL_CHANGE),all")
                    If Product_Main_Sel.Count = 0 Then
                        GoTo End_If
                    End If
                    Product_Main_Sel.Item(1).Value.Value = True
                    Product_BaseFrame_Sel.Clear()
                    Product_Main_Sel.search("(Name=L_BASE),all")
                    If Product_Main_Sel.Count = 0 Then
                        GoTo End_If
                    End If
                    Product_Main_Sel.Item(1).Value.Value = EPN1101_EPN11_D02
                End If
End_If:
            End If

            Product_BaseFrame.Activate()
            Product_BaseFrame.Product.Update()
            Product_BaseFrame.Save()

            Message_Label.Text = "Template01 생성완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)

        End Try
    End Sub


    Private Sub Base_Frame_Clash_Check_Click(sender As Object, e As EventArgs) Handles Base_Frame_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Base_Frame 간섭 체크중.."

        Call Template_Clash_Check_Function()

        Base_Frame_Drawing_Create.Enabled = True
        Base_Frame_Result_Part_Create.Enabled = True

        Message_Label.Text = "Base_Frame 간섭체크 완료"
    End Sub


    Private Sub Base_Frame_Drawing_Create_Click(sender As Object, e As EventArgs) Handles Base_Frame_Drawing_Create.Click

        ProgressBar1.Value = 0
        Message_Label.Text = "Template02 생성중..."

        ProgressBar1.Value = 10

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- COMP 저장
        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Save()

        '--------------------------------------------------------------------------------------------------- Title Block 생성
        Dim Base_Frame_drawing_open = CATIA.Documents.NewFrom(Base_Frame_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 6))
        ProgressBar1.Value = 40

        Call CATIA_BASIC03()
        '--------------------------------------------------------------------------------------------------- 도면 정보에서 Item 가져오기(140905 name명 오기)
        Dim Drawing_Item_Folder_Value = Base_Frame_folder_name
        Dim Drawing_Item_Count = Nothing
        For Drawing_Item_Count = 1 To 20
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = STW_Drawing_Infomation(0, Drawing_Item_Count) Then
                Exit For
            End If
        Next

        '--------------------------------------------------------------------------------------------------- 도면 Part List 생성하기
        For i = 1 To STW_Drawing_Infomation_Count
            If Not STW_Drawing_Infomation(i, Drawing_Item_Count) = "" Then
                DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(i, Drawing_Item_Count), i)
            End If
        Next
        ProgressBar1.Value = 60

        '--------------------------------------------------------------------------------------------------- 도면 Text 변경
        Call Base_Frame_Drawing_Text_Setting()

        On Error Resume Next
        '--------------------------------------------------------------------------------------------------- 도면 DWG_No 변경
        Drawingselection.clear
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
        CATIA.ActiveDocument.SaveAs(Base_Frame_folder_name & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATDrawing")

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template02 생성완료"
    End Sub


    Private Sub Base_Frame_Result_Part_Create_Click(sender As Object, e As EventArgs) Handles Base_Frame_Result_Part_Create.Click
        Try
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
            ProgressBar1.Value = 15

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
            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
            If CATIA.ActiveDocument.Name = STW_Template_Infomation(Form_Click_Index_Value, 5) Then
                CATIA.ActiveDocument.Close()
            End If
            ProgressBar1.Value = 30

            '--------------------------------------------------------------------------------------------------- Base Frame 초기위치 Constraint 삭제
            '--------------------------------------------------------------------------------------------------- COMP 에서 Template Data 제거
            Call CATIA_BASIC02()
            selection.Clear()
            selection.Add(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count))
            selection.Delete()
            Threading.Thread.Sleep(300)

            '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
            Call Data_Constraint_Delete(Base_Frame_Initial_Constraint, Base_Frame_Initial_Constraint_Count, Result_Template_Item_Value)
            '--------------------------------------------------------------------------------------------------- COMP Data에 실적 DATA 삽입
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart"
            CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            Threading.Thread.Sleep(300)

            '--------------------------------------------------------------------------------------------------- Base_Frame EPN 변경
            If Strings.Left(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name, 1) <> "E" Then
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
            End If

            '--------------------------------------------------------------------------------------------------- Base_Frame Data 저장
            Dim Base_Frame_Part = Documents.Item(Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            Base_Frame_Part.Save()
            ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- Base_Frame Constraint Delete & Create
            For i = 1 To Base_Frame_Selection_Constraint_Count - 1
                Call Constraint_Delete(Base_Frame_Selection_Constraint(i, 1), Base_Frame_Selection_Constraint(i, 2), Base_Frame_Selection_Constraint(i, 3))
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
            Message_Label.Text = "Template03 Part 확정 완료."
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Frame_Result_Part_Create_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Base_Frame_Parameter_Control_Click(sender As Object, e As EventArgs) Handles Base_Frame_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- 화면 전환
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        Call Get_Template_Parameter_Excel_Value(STW_COMP_Infomation(8, 6))

        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()
        Call Template_Parameter_Modify(STW_COMP_Infomation(8, 6))
        Call Folder_File_Count()
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(400, 610, Me)
    End Sub


    Private Sub LEN_BASE_L_Expand_Update_Click(sender As Object, e As EventArgs) Handles LEN_BASE_L_Expand_Update.Click
        LEN_BASE_L_Expand.Value = LEN_BASE_L_Expand_Text.Text
        CATIA.ActiveDocument.Product.Update()
        selection.clear
        selection.search("Name='LEN_BASE_L_NOW',all")
        If Not selection.Count = 0 Then
            LEN_BASE_L_NOW = selection.Item(1).Value
            NOW_BASE_L.Text = Math.Round(LEN_BASE_L_NOW.Value, 2)
        End If
    End Sub


    Private Sub Base_Frame_Parameter_Update_Click(sender As Object, e As EventArgs) Handles Base_Frame_Parameter_Update.Click
        VVIP_Message.text = "Update 진행중"

        Call VVIP_Option(VVIP_Check, UI_CONTROL_PNL, "CONTROL_PANEL_외장")
        '--------------------------------------------------------------------------------------------------- Product File Save
        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Save()

        VVIP_Message.text = "Update 완료"
    End Sub

End Class