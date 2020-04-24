Public Class BA_Coupling_Template_Create


    Private Sub BA_Coupling_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- Form Position
        Me.Top = 150
        Me.Left = 50
        Me.Width = 544
        Me.Height = 340
        Me.TopMost = True

        Call CATIA_Basic02()

        '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
        Exc = CreateObject("excel.application")
        Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")

        Coupling_Constraint_Excel_Path_Text.Text = Exc.Worksheets("Sheet1").Range("A" & 11).Value
        Coupling_BOM_File_Name.Text = Exc.Worksheets("Sheet1").Range("A" & 12).Value

        Exc.Application.DisplayAlerts = False
        Exc.ActiveWorkbook.Close(False)
        Exc.Application.DisplayAlerts = True
        Exc.Quit()

        Call KillProcess("EXCEL") 
    End Sub


    Private Sub Coupling_Path_Save_Click(sender As Object, e As EventArgs) Handles Coupling_Path_Save.Click
        '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
        Exc = CreateObject("excel.application")
        Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")

        Exc.Worksheets("Sheet1").Range("A" & 11).Value = Coupling_Constraint_Excel_Path_Text.Text

        Exc.Application.DisplayAlerts = False
        Exc.ActiveWorkbook.Close(True)
        Exc.Application.DisplayAlerts = True
        Exc.Quit()

        Call KillProcess("EXCEL") 
    End Sub


    Private Sub Coupling_Parameter_Control_Click(sender As Object, e As EventArgs) Handles Coupling_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        If Get_Template_Parameter_Excel_Value("EPN03") = False Then
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()

        Call Template_Parameter_Modify("EPN03")
        Call Folder_File_Count()
    End Sub


    Private Sub Coupling_BOM_File_Click(sender As Object, e As EventArgs) Handles Coupling_BOM_File.Click
        '지워도 될지 판단해야됨
        Dim template_value = Nothing

        Dim data_replace

        ProgressBar1.Value = 0
        '--------------------------------------------------------------------------------------------------- 다중 작업 일때...
        If Job_Repeat_Check.Checked = True Then
            On Error Resume Next
            Dim oShell As Object, oFolder As Object
            '--------------------------------------------------------------------------------------------------- 폴더 선택하기
            oShell = CreateObject("shell.application")
            oFolder = oShell.browseforfolder(0, "Coupling 다중 작업 - 폴더 찾기", &H200&)
            Coupling_Repeat_Folder = oFolder.self.path & "\"
            Dim FName = Dir(Coupling_Repeat_Folder, vbNormal)  '목록을 생성해서 파일명을 가져옵니다.
            On Error GoTo 0

            If FName = "" Then
                Call SHOW_MESSAGE_BOX("선택한 폴더에 COUPLING BOM File이 없습니다.", "", 1, 1)
                Coupling_Message_Label.Text = "Coupling 생성 종료."
                Exit Sub
            Else
                '--------------------------------------------------------------------------------------------------- BOM_Excel_Open
                Dim sw = False          'Loof 비교 값.
                On Error Resume Next
                Do                  '폴더안에 Excel 파일 수 만큼 Loof
                    Coupling_Message_Label.Text = "Coupling 생성중..."
                    ProgressBar1.Value = 10
                    If Strings.Right(FName, 4) = ".xls" Or Strings.Right(FName, 4) = "xlsx" Then
                        sfile = Coupling_Repeat_Folder & FName
                        Exc = CreateObject("excel.application")
                        'Exc.Visible = True
                        Exc.Workbooks.Open(sfile)            'CommonDialog입력 받은 엑셀파일
                        Exc.Worksheets("sheet1").Select()

                        '--------------------------------------------------------------------------------------------------- 파일 경로 가져오기
                        Coupling_BOM_File_Name.Text = FName
                        '--------------------------------------------------------------------------------------------------- Bom Part가 없을때 까지 루프
                        Dim coupling_value_count = 1
                        Dim Check = True
                        Do   'Outer loop
                            '--------------------------------------------------------------------------------------------------- Part 이름 가져오기
                            coupling_create_value_name(0, coupling_value_count) = Exc.Worksheets("Sheet1").Range("A" & coupling_value_count + 1).Value
                            coupling_create_value_name(1, coupling_value_count) = Exc.Worksheets("Sheet1").Range("B" & coupling_value_count + 1).Value
                            coupling_value_count = coupling_value_count + 1
                            '--------------------------------------------------------------------------------------------------- Coupling 유무 판별
                            If Exc.Worksheets("Sheet1").Range("B" & coupling_value_count + 1).Value = "" Then        '조건이 True이면
                                Check = False                                                                             '플래그 값을 False로 설정합니다.
                                Exit Do                                                                                       '내부 루프를 종료합니다.
                            End If
                        Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.
                        '--------------------------------------------------------------------------------------------------- Excel File 닫기
                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Close(False)
                        Exc.Application.DisplayAlerts = True
                        Exc.Quit()
                        '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기
                        Call Get_Excel_Value_Function(Application.StartupPath & "\STW_COMP_Infomation.xlsx", coupling_create_value_name(1, 2), 6, "A", "B", "C", "D", "E", "F", "", "", "", "", STW_COMP_Infomation_Count, STW_COMP_Infomation, "0")
                        Call Data_O_BOM_Excel_Search()
                        Call Input_Output_Excel_Value("COUPLING")
                        '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
                        Exc = CreateObject("excel.application")
                        'Exc.Visible = True
                        Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")
                        Exc.Worksheets("Sheet1").Range("A" & 12).Value = ""
                        Exc.Worksheets("Sheet1").Range("A" & 12).Value = Coupling_BOM_File_Name.Text

                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Close(True)
                        Exc.Application.DisplayAlerts = True
                        Exc.Quit()

                        Call KillProcess("EXCEL") 
                        ProgressBar1.Value = 20
                        '---------------------------------------------------------------------------------------------------
                        Call Get_Excel_Value_Function(Application.StartupPath & "\STW_Template_Infomation.xlsx", coupling_create_value_name(1, 2), 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", STW_Template_Infomation_Count, STW_Template_Infomation, "0")
                        For j = 1 To UBound(STW_Template_Infomation)
                            If STW_Template_Infomation(j, 1) = "COUPLING" Then
                                Exit For
                            End If
                        Next j
                        '--------------------------------------------------------------------------------------------------- Data Open
                        Dim coupling_open = CATIA.Documents.Open(oFolder.parentfolder.self.path & "\" & coupling_create_value_name(1, 3) & "\TEMPLATE\" & coupling_create_value_name(1, 3) & "_" & STW_Template_Infomation(1, 5))
                        'coupling_open = CATIA.Documents.Open(oFolder.parentfolder.self.path & "\" & coupling_create_value_name(1, 3) & "\TEMPLATE\" & coupling_create_value_name(1, 3) & "_" & STW_Template_Infomation(i, 5))
                        Threading.Thread.Sleep(500)

                        ProgressBar1.Value = 30
                        If Not Err.Number = 0 Then
                            Call SHOW_MESSAGE_BOX("다중 작업1 - 폴더 경로가 잘 못 되었습니다. 다시 시작하여 주세요.", "", 1, 1)
                            Coupling_Message_Label.Text = "Coupling 생성 종료."
                            Exit Sub
                        End If
                        '--------------------------------------------------------------------------------------------------- Data Replace
                        Dim replace_product = coupling_open.Product.products.Item(1)
                        '--------------------------------------------------------------------------------------------------- SMx100 기종 일때...
                        If Mid(coupling_create_value_name(1, 2), 4, 1) = 1 Then
                            data_replace = coupling_open.Product.products.ReplaceComponent(replace_product, oFolder.parentfolder.self.path & "\" & coupling_create_value_name(1, 3) & "\COMMON\" & coupling_create_value_name(1, 4) & ".CATPart", False)
                            '--------------------------------------------------------------------------------------------------- SMx100 기종이 아닐때...
                        Else
                            data_replace = coupling_open.Product.products.ReplaceComponent(replace_product, oFolder.parentfolder.self.path & "\" & coupling_create_value_name(1, 3) & "\COMMON\" & coupling_create_value_name(1, 4) & ".CATProduct", False)
                        End If
                        If Not Err.Number = 0 Then
                            Call SHOW_MESSAGE_BOX("다중 작업2 - 폴더 경로가 잘 못 되었습니다. 다시 시작하여 주세요.", "", 1, 1)
                            Coupling_Message_Label.Text = "Coupling 생성 종료."
                            Exit Sub
                        End If
                        CATIA.ActiveDocument.Product.Update()
                        ProgressBar1.Value = 40

                        '--------------------------------------------------------------------------------------------------- Keyway 입력
                        Call CATIA_Basic02()
                        Call Input_Output_Insert_G_Type()
                        ProgressBar1.Value = 70

                        '--------------------------------------------------------------------------------------------------- coupling_get_parameter 가져오기
                        Call coupling_get_parameter()

                        '--------------------------------------------------------------------------------------------------- Zoom All
                        If template_value = "Fail" Then
                            '--------------------------------------------------------------------------------------------------- Part Data Close
                            CATIA.ActiveDocument.Close()
                            ProgressBar1.Value = 0
                            Coupling_Message_Label.Text = "Coupling BOM을 다시 선택하세요"
                            Exit Sub
                        End If

                        CATIA.ActiveWindow.ActiveViewer.Reframe()

                        ProgressBar1.Value = 100
                        Coupling_Message_Label.Text = "Coupling 생성완료."

                        Call Create_Publication_Parameter_Click()
                    End If
                    On Error GoTo 0

                    FName = Dir()    ' 다음 항목을 읽어들입니다.
                    If FName = "" Then
                        sw = True
                    End If
                Loop Until sw = True
            End If
            '--------------------------------------------------------------------------------------------------- 다중 작업 아닐때...
        Else
            Coupling_Message_Label.Text = "Coupling 생성중..."
            ProgressBar1.Value = 10
            '--------------------------------------------------------------------------------------------------- BOM_Excel_Open
            Call Coupling_BOM_Excel_Open()
            '--------------------------------------------------------------------------------------------------- BOM List Cancel시 에러 처리
            If Coupling_BOM = "" Then
                Call SHOW_MESSAGE_BOX("COUPLING BOM File을 선택하세요", "", 1, 1)
                Coupling_Message_Label.Text = "Coupling 생성 종료."
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기
            Call Get_Excel_Value_Function(Application.StartupPath & "\STW_COMP_Infomation.xlsx", coupling_create_value_name(1, 2), 6, "A", "B", "C", "D", "E", "F", "", "", "", "", STW_COMP_Infomation_Count, STW_COMP_Infomation, "0")
            Call Input_Output_Excel_Value("COUPLING")
            '--------------------------------------------------------------------------------------------------- BOM List Cancel시 에러 처리
            If Folder_Name = "" Then
                Call SHOW_MESSAGE_BOX("COUPLING BOM File을 선택하세요", "", 1, 1)
                now_o_bom_value = 0
                Coupling_Message_Label.Text = " "
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")
            Exc.Worksheets("Sheet1").Range("A" & 12).Value = ""
            Exc.Worksheets("Sheet1").Range("A" & 12).Value = Coupling_BOM_File_Name.Text

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(True)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL") 
            ProgressBar1.Value = 20
            '---------------------------------------------------------------------------------------------------
            Call Get_Excel_Value_Function(Application.StartupPath & "\STW_Template_Infomation.xlsx", coupling_create_value_name(1, 2), 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", STW_Template_Infomation_Count, STW_Template_Infomation, "0")
            For i2 = 1 To UBound(STW_Template_Infomation)
                If STW_Template_Infomation(i2, 1) = "COUPLING" Then
                    Exit For
                End If
            Next i2
            '--------------------------------------------------------------------------------------------------- Data Open
            On Error Resume Next

            Dim coupling_open = CATIA.Documents.Open(sTemplate_Data_Path_text & "\" & STW_Template_Infomation(1, 4) & "\" & _
                                                 coupling_create_value_name(1, 3) & "\TEMPLATE\" & coupling_create_value_name(1, 3) & "_" & _
                                                 STW_Template_Infomation(1, 5))
            'coupling_open = CATIA.Documents.Open(sTemplate_Data_Path_text & "\" & STW_Template_Infomation(i, 4) & "\" & coupling_create_value_name(1, 3) & "\TEMPLATE\" & coupling_create_value_name(1, 3) & "_" & STW_Template_Infomation(i, 5))

            If Err.Number <> 0 And STW_Template_Infomation(1, 5) = Nothing Then
                Call SHOW_MESSAGE_BOX("BOM 정보를 확인해주세요.", "", 1, 1)
                ProgressBar1.Value = 0
                Coupling_Message_Label.Text = ""
                Exit Sub
            End If

            If Err.Number <> 0 Then
                Call SHOW_MESSAGE_BOX("파일 경로를 확인해주세요.", "", 1, 1)
                ProgressBar1.Value = 0
                Coupling_Message_Label.Text = ""
                Exit Sub
            End If

            On Error GoTo 0
            ProgressBar1.Value = 30
            '--------------------------------------------------------------------------------------------------- Data Replace
            '--------------------------------------------------------------------------------------------------- SMx100 기종일때...
            Dim replace_product
            If Mid(coupling_create_value_name(1, 2), 4, 1) = 1 Then
                replace_product = coupling_open.Product.products.Item(1)
                data_replace = coupling_open.Product.products.ReplaceComponent(replace_product, sTemplate_Data_Path_text & "\" & _
                               STW_Template_Infomation(1, 4) & "\" & coupling_create_value_name(1, 3) & "\COMMON\" & _
                               coupling_create_value_name(1, 4) & ".CATPart", False)
            Else
                replace_product = coupling_open.Product.products.Item(1)
                data_replace = coupling_open.Product.products.ReplaceComponent(replace_product, sTemplate_Data_Path_text & "\" & _
                               STW_Template_Infomation(1, 4) & "\" & coupling_create_value_name(1, 3) & "\COMMON\" & _
                               coupling_create_value_name(1, 4) & ".CATProduct", False)
            End If
            On Error Resume Next
            CATIA.ActiveDocument.Product.Update()
            On Error GoTo 0
            ProgressBar1.Value = 40
            '--------------------------------------------------------------------------------------------------- Keyway 입력
            Call CATIA_Basic02()
            Call Input_Output_Insert_G_Type()
            ProgressBar1.Value = 70
            '--------------------------------------------------------------------------------------------------- coupling_get_parameter 가져오기
            Call coupling_get_parameter()
            '--------------------------------------------------------------------------------------------------- Zoom All
            If template_value = "Fail" Then
                '--------------------------------------------------------------------------------------------------- Part Data Close
                CATIA.ActiveDocument.Close()
                ProgressBar1.Value = 0
                Coupling_Message_Label.Text = "Coupling BOM을 다시 선택하세요."
                Exit Sub
            End If
            CATIA.ActiveDocument.Product.Update()
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Coupling_Message_Label.Text = "Coupling 생성완료."
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' 파트 확정 -> 실적파트로 변경                    &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Private Sub Create_Publication_Parameter_Click()
        Try
            date_value = Format(Now, "yymmddhhmm")

            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call Active_Window_Change("Product", 1, 7)

            '--------------------------------------------------------------------------------------------------- Product -> Part 변환
            Call Generate_CATPart_from_Product()

            '--------------------------------------------------------------------------------------------------- Geometrycal 이름에 "\" 기호 삭제
            Call Delete_Symbol()

            Coupling_Message_Label.Text = "Publication, Parameter 생성중..."
            ProgressBar1.Value = 0
            publications1 = product_add.publications
            '--------------------------------------------------------------------------------------------------- SMx100 기종이 아닐때, 불필요한 파라메터 제거
            If Not Mid(coupling_create_value_name(1, 2), 4, 1) = 1 And coupling_create_value_name(1, 3) = "DISK" Then
                selection.clear
                selection.search("Name='EPN03_P02',all")             'EPN 찾기.
                Dim coupling_axis_data = selection.Item(1).Value
                coupling_axis_data.Count()
                selection.Clear()
                selection.Add(coupling_axis_data)
                selection.Delete()
            End If

            ProgressBar1.Value = 30
            For i = 1 To Input_Output_Excel_Value_Count - 1
                '--------------------------------------------------------------------------------------------------- EPN04_P01 Publication 생성
                If Input_Output_Value(4, i) = "E" Then
                    Call publication_axis_create(Input_Output_Value(1, i))
                    '--------------------------------------------------------------------------------------------------- EPN04_D01 parameter 생성
                ElseIf Input_Output_Value(4, i) = "F" Then
                    Call part_parameter_create(Input_Output_Value(1, i), Input_Output_Parameter_Value(i), Input_Output_Value(2, i))
                End If
            Next

            product_add.PartNumber = coupling_create_value_name(1, 1)
            product_add.Name = "EPN03"
            CATIA.ActiveDocument.Part.Update()

            '--------------------------------------------------------------------------------------------------- File Save
            If Job_Repeat_Check.Checked = False Then
                CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & coupling_create_value_name(1, 2) & "\EPN03\" & _
                                            coupling_create_value_name(1, 1) & "_" & date_value & ".CATPart")
            Else
                CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & coupling_create_value_name(1, 2) & "\EPN03\" & _
                                            coupling_create_value_name(1, 1) & ".CATPart")
            End If
            Threading.Thread.Sleep(500)

            '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
            fso = CreateObject("Scripting.FileSystemObject")

            Dim file_Attributes_value
            If Job_Repeat_Check.Checked = False Then
                file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & coupling_create_value_name(1, 2) & "\EPN03\" & coupling_create_value_name(1, 1) & "_" & date_value & ".CATPart")
            Else
                file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & coupling_create_value_name(1, 2) & "\EPN03\" & coupling_create_value_name(1, 1) & ".CATPart")
            End If
            file_Attributes_value.Attributes = 33

            '--------------------------------------------------------------------------------------------------- motor_parameter 생성
            ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- Template Data Close
            Call Active_Window_Change("Product", 1, 7)
            If Strings.Right(CATIA.ActiveDocument.Name, 7) = "Product" Then
                CATIA.ActiveDocument.Close()
            End If

            '--------------------------------------------------------------------------------------------------- Drawing Data Close
            Call Active_Window_Change("Drawing", 1, 7)

            '--------------------------------------------------------------------------------------------------- ActiveDocument가 Drawing 일 경우만 창 종료
            If Strings.Right(CATIA.ActiveDocument.Name, 7) = "Drawing" Then
                CATIA.ActiveDocument.Close()
            End If

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Coupling_Message_Label.Text = "Publication, Parameter 생성완료."
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Create_Publication_Parameter_Click 함수 오류!!", "", 1, 1)
        End Try
    End Sub

End Class