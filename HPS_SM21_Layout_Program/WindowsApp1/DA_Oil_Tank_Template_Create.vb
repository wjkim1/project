Public Class DA_Oil_Tank_Template_Create


    Private Sub DA_Oil_Tank_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
            Me.Top = 150
            Me.Left = 50
            Me.Width = 439
            Me.Height = 655
            Me.TopMost = True

            Call CATIA_BASIC02()

            Template_Data_Loading_Form.ProgressBar1_Change(30)

            '--------------------------------------------------------------------------------------------------- Oil_Tank_Template_Form_Inicial_Setting
            'VVIP_Frame.Visible = False                      ' VVIP 관련 프레임 Hide
            Oil_Tank_Parameter_Update.Enabled = False       ' Update Botton Enable
            Oil_Tank_Drawing_Create.Enabled = False
            Oil_Tank_Result_Part_Create.Enabled = False

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

            '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
            Call Data_O_BOM_Excel_Search()

            '--------------------------------------------------------------------------------------------------- Oil_Tank_Excel_Open
            Call Oil_Tank_Excel_Open()

            Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("DA_Oil_Tank_Template_Create_Load() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Oil_Tank_Template_Create_Click(sender As Object, e As EventArgs) Handles Oil_Tank_Template_Create.Click
        '--------------------------------------------------------------------------------------------------- Oil Tank Part NO 가 없을때
        If Template_Part_Number_Text.Text = "" Then
            Call SHOW_MESSAGE_BOX("Oil Tank Part No.가 없습니다.", "", 1, 1)
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- Oil Tank Type 이 없을때
        If Option_Default.Checked = True And Strings.Left(sLbl_ModelNumber, 2) = "SM" Then
            Call SHOW_MESSAGE_BOX("Oil Tank Type을 선택하세요.", "", 1, 1)
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

                Call Folder_Delete(sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value)

                fso.CopyFolder(sTemplate_Data_Path_text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4), sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\")

                Threading.Thread.Sleep(300)
                Dim Oil_Tank_folder = fso.getfolder(sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name)

                Oil_Tank_folder.Name = Template_Data_Folder_Name & "_" & date_value
                Oil_Tank_folder_name = Oil_Tank_folder.path

                '--------------------------------------------------------------------------------------------------- Oil_Tank Path ( Replace 에 사용 )
                Copy_Oil_Tank_Path = sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value

                '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
                Call wirte_transfer_file(Oil_Tank_folder_name)
                ProgressBar1.Value = 10

                '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
                Call Clash_Check_Folder_Create()
                ProgressBar1.Value = 20

                '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
                Call Data_Constraint_Delete(Oil_Tank_Initial_Constraint, Oil_Tank_Initial_Constraint_Count, Result_Template_Item_Value)
                ProgressBar1.Value = 30

                '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
                Call Get_Parameter_Data_Comparison(Oil_Tank_Initial_Parameter, Oil_Tank_Initial_Parameter_Count, Me)

                '--------------------------------------------------------------------------------------------------- Oil_Tank Template Data Open
                Call CATIA_BASIC02()

                '--------------------------------------------------------------------------------------------------- Template Data Open
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Oil_Tank_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5)
                products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                Threading.Thread.Sleep(500)

                Dim Template_Skeleton_Part
                If Mid(sLbl_ModelNumber, 4, 1) = "1" Then
                    Template_Skeleton_Part = Documents.Item("SKELETON_OIL_TANK.CATPart").Part
                Else
                    Template_Skeleton_Part = Documents.Item("Input_Oil_Tank.CATPart").Part
                End If

                '--------------------------------------------------------------------------------------------------- Input 변경
                For i = 1 To Input_Output_Excel_Value_Count - 1
                    If Input_Output_Value(4, i) = "D" Then
                        Call Constraint_Delete(Input_Output_Value(1, i), Input_Output_Value(2, i), Input_Output_Value(3, i))
                    ElseIf Input_Output_Value(4, i) = "A" Then
                        Call Input_Output_Axis_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                    ElseIf Input_Output_Value(4, i) = "B" Then
                        Call Input_Output_Element_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                        Template_Skeleton_Part.Update   '200312
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
                        Exc.Workbooks.Open(Oil_Tank_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

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
                        Exc.Workbooks.Open(Oil_Tank_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

                        Call Input_Output_LayoutParameter_to_Excel(Exc, STW_Drawing_Infomation(1, 4), Input_Output_Value(1, i), Input_Output_Value(2, i))

                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Close(True)
                        Exc.Application.DisplayAlerts = True
                        Exc.Quit()

                        Call KillProcess("EXCEL") 
                    End If
                Next

                ProgressBar1.Value = 50
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value

                '--------------------------------------------------------------------------------------------------- Oil_Tank Template Open
                Dim Oil_Tank_open = CATIA.Documents.Open(Oil_Tank_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5))

                '--------------------------------------------------------------------------------------------------- oil_tank_weld_mach_design_table_control
                Call oil_tank_weld_mach_design_table_control()

                '--------------------------------------------------------------------------------------------------- 기종 정보 입력
                selection.Clear
                selection.search("Name='MACHINE_TYPE',all")
                If Not selection.Count = 0 Then
                    selection.Item(1).Value.Value = STW_Drawing_Infomation(1, 1)
                    ProgressBar1.Value = 60
                End If

                '--------------------------------------------------------------------------------------------------- Template Data Get Parameter
                Call Template_Data_Get_Parameter()
                ProgressBar1.Value = 70

                '--------------------------------------------------------------------------------------------------- Oil Tank Template Data Replace
                Call Oil_Tank_Data_Replace()
                ProgressBar1.Value = 90

                '--------------------------------------------------------------------------------------------------- Product File Save
                CATIA.ActiveDocument.Product.Update()
                CATIA.ActiveDocument.Save()
            Else
                Exit Sub
            End If

            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Product.Update()

            '--------------------------------------------------------------------------------------------------- Oil Tank 간섭체크
            If Auto_Clash_Check.Checked = True Then
                Call Oil_Tank_Clash_Check_Click()
            End If
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Message_Label.Text = "Template01 생성완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("DA_Oil_Tank_Template_Create_Load 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Oil_Tank_Clash_Check_Click() Handles Oil_Tank_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Oil Tank 간섭 체크중.."
        Call Template_Clash_Check_Function()
        Oil_Tank_Drawing_Create.Enabled = True
        Oil_Tank_Result_Part_Create.Enabled = True
        Message_Label.Text = "Oil Tank 간섭체크 완료"
    End Sub


    Private Sub Oil_Tank_Parameter_Control_Click(sender As Object, e As EventArgs) Handles Oil_Tank_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- 화면 전환
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        Call Get_Template_Parameter_Excel_Value(STW_COMP_Infomation(4, 6))
        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()
        Call Template_Parameter_Modify(STW_COMP_Infomation(4, 6))
        Call Folder_File_Count()
    End Sub


    Private Sub Oil_Tank_Drawing_Create_Click(sender As Object, e As EventArgs) Handles Oil_Tank_Drawing_Create.Click
        ProgressBar1.Value = 0
        Message_Label.Text = "Template02 생성중..."

        ProgressBar1.Value = 10

        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- COMP 저장
        CATIA.ActiveDocument.Save()
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 파일 Open
        Dim Oil_Tank_drawing_open = CATIA.Documents.NewFrom(Oil_Tank_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 6))
        ProgressBar1.Value = 40

        Call CATIA_BASIC03()
        DrawingSheets.Item(1).Activate()
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 정보에서 Item 가져오기
        Dim Drawing_Item_Folder_Value = Oil_Tank_folder_name
        Dim Drawing_Item_Count As Integer
        For Drawing_Item_Count = 1 To 20
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = STW_Drawing_Infomation(0, Drawing_Item_Count) Then
                Exit For
            End If
        Next
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 Part List 생성하기
        For i = 1 To STW_Drawing_Infomation_Count - 1
            If Not STW_Drawing_Infomation(i, Drawing_Item_Count) = "" Then
                DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(i, Drawing_Item_Count), i)
            End If
        Next
        RuntimeCheck()
        Call CATIA_BASIC03()
        RuntimeCheck()
        On Error Resume Next
        '--------------------------------------------------------------------------------------------------- 도면 DWG_No 변경
        Drawingselection.clear
        Drawingselection.search("Name=Model*,all")
        If Not Drawingselection.Count = 0 Then
            For i = 1 To Drawingselection.Count
                Drawingselection.Item(i).Value.Text = STW_Drawing_Infomation(1, 1)
            Next
        End If
        '--------------------------------------------------------------------------------------------------- DWG_No
        Drawingselection.clear
        Drawingselection.search("Name=DWG_No*,all")
        If Not Drawingselection.Count = 0 Then
            For i = 1 To Drawingselection.Count
                Drawingselection.Item(i).Value.Text = Template_Part_Number_Text.Text
            Next
        End If
        On Error GoTo 0

        ProgressBar1.Value = 80
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 표제란 입력
        Call drawing_title_box_write()
        DrawingViews.Item("Main View").Activate()

        '-------RuntimeCheck-------------------------------------------------------------------------------------------- Layout Parameter -> Drawing
        For i = 1 To Input_Output_Excel_Value_Count - 1
            If Input_Output_Value(4, i) = "J" Then
                Call Input_Output_LayoutParameter_to_Drawing(Input_Output_Value(1, i), Input_Output_Value(2, i))
            End If
            RuntimeCheck()
        Next i
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 파일 Update
        CATIA.ActiveDocument.Update()
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 파일 저장
        CATIA.ActiveDocument.SaveAs(Oil_Tank_folder_name & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATDrawing")
        RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template02 생성완료"
    End Sub


    Private Sub Oil_Tank_Result_Part_Create_Click(sender As Object, e As EventArgs) Handles Oil_Tank_Result_Part_Create.Click
        Try
            ProgressBar1.Value = 0
            Message_Label.Text = "Template03 Part 변환..."

            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

            '--------------------------------------------------------------------------------------------------- Replace Data Hide
            selection.Clear()
            Dim visPropertySet1 = selection.VisProperties

            For i = 1 To Input_Output_Excel_Value_Count - 1
                If Input_Output_Value(4, i) = "I" Then
                    For j = 1 To CATIA.ActiveDocument.Product.products.Count
                        '--------------------------------------------------------------------------------------------------- Part 이름이 동일할때...
                        If Input_Output_Value(2, i) = CATIA.ActiveDocument.Product.products.Item(j).Name Then
                            '--------------------------------------------------------------------------------------------------- Part Selection
                            selection.Add(CATIA.ActiveDocument.Product.products.Item(j))
                            Exit For
                        End If
                    Next j
                End If
            Next i

            '--------------------------------------------------------------------------------------------------- Part Hide
            visPropertySet1 = visPropertySet1.Parent
            visPropertySet1.SetShow(1)
            selection.Clear()

            '--------------------------------------------------------------------------------------------------- Template 파일 저장 및 Generate CATPart from Product 명령어 실행
            CATIA.ActiveDocument.Save()

            '--------------------------------------------------------------------------------------------------- Product -> Part 변환
            Call Generate_CATPart_from_Product()

            '--------------------------------------------------------------------------------------------------- Geometrycal 이름에 "\" 기호 삭제
            Call Delete_Symbol()

            Message_Label.Text = "Template03 Part 확정중..."
            ProgressBar1.Value = 10

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
            Call Save_File_Check(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & _
                                 Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")

            '--------------------------------------------------------------------------------------------------- Part File Save
            CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & _
                                        Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")

            '--------------------------------------------------------------------------------------------------- Part Data Close
            CATIA.ActiveDocument.Close()

            '--------------------------------------------------------------------------------------------------- 도면 파일 Close
            Call CATIA_Basic02()
            Call Active_Window_Change("CATDrawing", 1, 10)
            If Strings.Right(CATIA.ActiveDocument.Name, 10) = "CATDrawing" Then
                CATIA.ActiveDocument.Close()
            End If
            '--------------------------------------------------------------------------------------------------- Template 파일 Close
            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
            '--------------------------------------------------------------------------------------------------- Replace Data Hide
            selection.Clear()
            visPropertySet1 = selection.VisProperties

            For i = 1 To Input_Output_Excel_Value_Count - 1
                If Input_Output_Value(4, i) = "I" Then
                    For j = 1 To CATIA.ActiveDocument.Product.products.Count
                        '--------------------------------------------------------------------------------------------------- Part 이름이 동일할때...
                        If Input_Output_Value(2, i) = CATIA.ActiveDocument.Product.products.Item(j).Name Then
                            '--------------------------------------------------------------------------------------------------- Part Selection
                            selection.Add(CATIA.ActiveDocument.Product.products.Item(j))
                            Exit For
                        End If
                    Next j
                End If
            Next i
            '--------------------------------------------------------------------------------------------------- Part Hide
            visPropertySet1 = visPropertySet1.Parent
            visPropertySet1.SetShow(0)
            selection.Clear()

            If CATIA.ActiveDocument.Name = STW_Template_Infomation(Form_Click_Index_Value, 5) Then
                CATIA.ActiveDocument.Save()
                CATIA.ActiveDocument.Close()
            End If
            ProgressBar1.Value = 30
            '--------------------------------------------------------------------------------------------------- COMP 에서 Template Data 제거
            Call CATIA_Basic02()
            selection.Clear()
            selection.Add(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count))
            selection.Delete()
            Threading.Thread.Sleep(500)
            '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
            Call Data_Constraint_Delete(Oil_Tank_Initial_Constraint, Oil_Tank_Initial_Constraint_Count, Result_Template_Item_Value)
            '--------------------------------------------------------------------------------------------------- COMP Data에 실적 DATA 삽입
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart"
            Threading.Thread.Sleep(500)
            CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            '--------------------------------------------------------------------------------------------------- Oil_Tank EPN 변경
            If Strings.Left(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name, 1) <> "E" Then
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
            End If

            '--------------------------------------------------------------------------------------------------- Oil Tank Data 저장
            Dim Oil_Tank_Part = Documents.Item(Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            Oil_Tank_Part.Save()
            ProgressBar1.Value = 60
            '--------------------------------------------------------------------------------------------------- Oil Tank Constraint Delete & Create
            For i = 1 To Oil_Tank_Selection_Constraint_Count - 1
                Call Constraint_Delete(Oil_Tank_Selection_Constraint(i, 1), Oil_Tank_Selection_Constraint(i, 2), Oil_Tank_Selection_Constraint(i, 3))
            Next
            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Call O_BOM_Excel_Writing(Result_Template_Item_Value, Template_Part_Number_Text.Text & "_" & date_value)
            Call Btn_Refresh_Click()
            '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
            Threading.Thread.Sleep(500)

            fso = CreateObject("Scripting.FileSystemObject")
            Dim file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            file_Attributes_value.Attributes = 33
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Message_Label.Text = "Template03 생성완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            ProgressBar1.Value = 100
            Message_Label.Text = "Template03 생성 오류!!"
            Call SHOW_MESSAGE_BOX("Template03 생성 함수 오류!!", "", 1, 1)
        End Try
    End Sub

End Class