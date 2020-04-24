Module Coupling
    '--------------------------------------------------------------------------------------------------- 커플링 선택 초기화 Constraint
    Public Coupling_Initial_Constraint, Coupling_Initial_Constraint_Count
    '--------------------------------------------------------------------------------------------------- 커플링 선택 초기화 Parameter
    Public Coupling_Initial_Parameter, Coupling_Initial_Parameter_Count
    '--------------------------------------------------------------------------------------------------- 커플링 선택 Constraint
    Public Coupling_Selection_Constraint, Coupling_Selection_Constraint_Count
    '--------------------------------------------------------------------------------------------------- 커플링 선택 Parameter
    Public Coupling_Selection_Parameter, Coupling_Selection_Parameter_Count
    Public Coupling_BOM
    '--------------------------------------------------------------------------------------------------- 커플링 Template Excel Name
    Public coupling_create_value_name(1, 10)


    Public Function Coupling_Excel_Open()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        '--------------------------------------------------------------------------------------------------- 커플링선택초기화 Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(3, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Coupling_Initial_Constraint_Count, Coupling_Initial_Constraint, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(30)

        '--------------------------------------------------------------------------------------------------- 커플링선택초기화 Parameter
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(3, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Coupling_Initial_Parameter_Count, Coupling_Initial_Parameter, "4")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(40)

        '--------------------------------------------------------------------------------------------------- 커플링선택  Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(4, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Coupling_Selection_Constraint_Count, Coupling_Selection_Constraint, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(50)

        '--------------------------------------------------------------------------------------------------- 커플링선택 Parameter
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(4, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Coupling_Selection_Parameter_Count, Coupling_Selection_Parameter, "4")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(60)

        '--------------------------------------------------------------------------------------------------- Template 관련 정보 가져오기
        sExcelFilePath = Application.StartupPath & "\STW_Template_Infomation.xlsx"
        If System.IO.File.Exists(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, sLbl_ModelNumber, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", _
                                          STW_Template_Infomation_Count, STW_Template_Infomation, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(70)

        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        ReDim Data_Version_Infomation(0 To 1)

        sExcelFilePath = sTemplate_Data_Path_text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & "\DT_Design_Rule_Version.xlsx"
        If System.IO.File.Exists(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                          Data_Version_Infomation_Count, Data_Version_Infomation, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(80)

        Return True
    End Function


    '---------------------------------------------------------------------------------------------------
    ' Dummy_Data_open     &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Function Dummy_Data_open()
        Call CATIA_BASIC02()

        Try
            '--------------------------------------------------------------------------------------------------- EPN0501 Search
            Dim Tank_Dummy_Count = 1
            For i = 1 To product_add.products.Count
                If Right(product_add.products.Item(i).Name, 7) = "EPN0501" Then
                    Tank_Dummy_Count = 0
                    Exit For
                End If
            Next

            Dim arrayOfVariantOfBSTR1(0)
            If Tank_Dummy_Count = 1 Then
                '--------------------------------------------------------------------------------------------------- Tank Dummy Loading
                If EXISTS_FILE_CHECK(arrayOfVariantOfBSTR1(0)) Then
                    arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & assy_value_Name & "\EPN0501\" & assy_value_Name & "-EPN0501-DMY.CATPart"
                Else : Return False
                End If

                CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                Threading.Thread.Sleep(500)
            End If

            '--------------------------------------------------------------------------------------------------- EPN1101 Search
            Dim Enclosure_Dummy_Count = 1
            For i = 1 To product_add.products.Count
                If Right(product_add.products.Item(i).Name, 7) = "EPN1101" Then
                    Enclosure_Dummy_Count = 0
                    Exit For
                End If
            Next

            If Enclosure_Dummy_Count = 1 Then
                '--------------------------------------------------------------------------------------------------- Enclosure Dummy Loading
                arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & assy_value_Name & "\EPN1101\" & assy_value_Name & "-EPN1101-DMY.CATPart"
                CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                Threading.Thread.Sleep(500)
            End If
            Return True
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Dummy_Data_open() 함수 오류!!", "", 1, 1)
            Return False
        End Try
    End Function


    Public Sub coupling_get_parameter()
        Call CATIA_BASIC02()

        Try
            If CATIA.ActiveDocument.Product.products.Item(1).PartNumber = "Common_Assy" Then
                MsgBox("기종 정보가 일치하지 않습니다.", vbOKCancel Or vbInformation, "ISPark Automation")
                Exit Sub
            End If

            ReDim Input_Output_Parameter_Value(0 To 100)
            For i = 1 To Input_Output_Excel_Value_Count - 1
                If Input_Output_Value(4, i) = "F" Then
                    selection.clear
                    selection.search("Name='" & Input_Output_Value(1, i) & "',all")
                    Input_Output_Parameter_Value(i) = selection.Item(1).Value
                End If
            Next
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("coupling_get_parameter() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Coupling_BOM_Excel_Open                 &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Coupling_BOM_Excel_Open()
        '--------------------------------------------------------------------------------------------------- File Open
        With BA_Coupling_Template_Create.OpenFileDialog1
            .Title = "Open File"
            .Filter = "Excel File(*.xls, *.xlsx)|*.xls;*.xlsx"          '엑셀파일만
            .FileName = ""
            .ShowReadOnly = True            '읽기 전용
            .Multiselect = False
            .ShowDialog()
            'BA_Coupling_Template_Create.CommonDialog1.InitDir = BA_Coupling_Template_Create.Coupling_Constraint_Excel_Path_Text.Text
            'BA_Coupling_Template_Create.CommonDialog1.ShowOpen()

            If .FileName = "" Then
                Coupling_BOM = ""
                Exit Sub
            End If

            sfile = .FileName             '파일 오픈하고 파일선택
            If .FileName = "" Then
                Coupling_BOM = ""
                Exit Sub
            Else
                Coupling_BOM = "1"
            End If
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(sfile)            'CommonDialog입력 받은 엑셀파일
            Exc.Worksheets("sheet1").Select()
        End With
        '--------------------------------------------------------------------------------------------------- 파일 경로 가져오기
        Folder_Name = Split(BA_Coupling_Template_Create.OpenFileDialog1.FileName, vbNullChar)(0)
        ls_ss = Split(Folder_Name, "\")
        Origin_File_Name = ls_ss(UBound(ls_ss))
        BA_Coupling_Template_Create.Coupling_BOM_File_Name.Text = Origin_File_Name
        Dim coupling_name = Split(Origin_File_Name, ".")
        Dim coupling_name_org = coupling_name(0)
        Dim Folder_Chr = "\"
        For i = 0 To UBound(ls_ss) - 1
            If i = 0 Then BA_Coupling_Template_Create.Coupling_Constraint_Excel_Path_Text.Text = ""
            BA_Coupling_Template_Create.Coupling_Constraint_Excel_Path_Text.Text = BA_Coupling_Template_Create.Coupling_Constraint_Excel_Path_Text.Text + ls_ss(i) & Folder_Chr
        Next i
        '--------------------------------------------------------------------------------------------------- Bom Part가 없을때 까지 루프
        Dim coupling_value_count = 1
        'part_name_count = 1
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

        Call KillProcess("EXCEL")
    End Sub

End Module