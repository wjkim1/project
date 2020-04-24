Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Scripting
Imports INFITF
Imports MECMOD
Imports ProductStructureTypeLib
Imports main = main_form

Module O_BOM_Layout_Construction
    '--------------------------------------------------------------------------------------------------- 메인 폼의 값
    Public sOBomPath As String
    '--------------------------------------------------------------------------------------------------- VB 파일 시스템 관련
    Public fso As FileSystemObject
    '--------------------------------------------------------------------------------------------------- 진행중 Product 일때 사용
    Public Get_Layout_List_Number
    '--------------------------------------------------------------------------------------------------- BOM 폴더 생성 관련
    Public date_value, Folder_Value As String, o_bom_excel_file_name As String
    Public now_o_bom_value, now_o_bom_excel_path            'now_o_bom_value = 0 O-BOM 가져오기
    Public Folder_Name As String, Origin_File_Name As String, Test_One_Origin_File_Name As String, One_Origin_File_Name As String
    Public ls_ss
    '--------------------------------------------------------------------------------------------------- O-BOM 관련 Excel 파일 정보 가져오기
    ' 선택한 O_BOM(BOM_LIST)  정보
    Public O_BOM_Part_Value, O_BOM_Part_Value_Count
    '--------------------------------------------------------------------------------------------------- Constraint 가져오기
    'O_BOM_Constraint_Value : Constraint(1.표준) 정보
    Public O_BOM_Constraint_Value, O_BOM_Constraint_Value_Count
    '--------------------------------------------------------------------------------------------------- LAYOUT 대상 부품 구분 파트 가져오기
    Public O_BOM_Layout_Part_Value, O_BOM_Layout_Part_Value_Count
    '--------------------------------------------------------------------------------------------------- O-BOM Standard Excel Name
    Public O_BOM_Standard_Excel_Name
    '--------------------------------------------------------------------------------------------------- O-BOM Standard Value 가져오기
    Public O_BOM_Standard_Part_Value, O_BOM_Standard_Part_Value_Count
    '--------------------------------------------------------------------------------------------------- O-BOM 기준 Data Loading
    Public pub_number, get_part_publication(1000), axis_number, get_part_axisSystem(1000)
    Public get_part_publication2017 As New Dictionary(Of String, Publication)
    Public Tmp_Duplication_Check_Dic As New Dictionary(Of String, String)

    Public std_p_number_count As Integer
    '--------------------------------------------------------------------------------------------------- O-BOM 기준 Data Loading
    Public layout_EPN_ref, layout_part_name_ref
    '--------------------------------------------------------------------------------------------------- O-BOM 구속조건 생성
    Public Publish_Org_Name(1000) As String, Publish_Sub_Name(1000) As String, constraint_part01(1000), constraint_part02(1000), constraint_count
    '--------------------------------------------------------------------------------------------------- O-BOM  Publication 관련 선언
    Public constraints1, name_ref, product_fix, Product1, publications1, Pub1, Pub_Search, Pub_Search02, Pub_reference1, Pub_reference2, coincidence_constraint(1000), Pub1_Search, Pub2_Search
    Public Pub1_DisplayName As String, Pub2_DisplayName As String
    Public add_axis_Product                                                                                                     'Product Axis와 동일축을 사용하기 위해 생성된 Part
    '--------------------------------------------------------------------------------------------------- Motor 정보
    Public Motor_EPN02_D01, Motor_EPN02_D11

    Public Main_Form
    Sub New()
    End Sub
    '---------------------------------------------------------------------------------------------------
    ' Bom_File_Open_Function   &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& ' 사용자가 직접 파일을 선택하여 Open
    '---------------------------------------------------------------------------------------------------
    Public Sub Bom_File_Open_Function()
        '--------------------------------------------------------------------------------------------------- 파일 경로 가져오기
        Dim WkBook, BOM_File_Path, O_BOM_DB_Path, Path_Size_1, Path_Size_2, Temp_Size
        Dim ls_ss_folder, assy_value_Name_ref

        A_Main_Form.OpenFileDialog1.Title = "Open File"
        A_Main_Form.OpenFileDialog1.Filter = "Excel File(*.xls, *.xlsx)|*.xls;*.xlsx"          '엑셀파일만
        A_Main_Form.OpenFileDialog1.FileName = ""
        A_Main_Form.OpenFileDialog1.ReadOnlyChecked = True
        A_Main_Form.OpenFileDialog1.ShowReadOnly = True '= cdlOFNNoReadOnlyReturn            '읽기 전용
        A_Main_Form.OpenFileDialog1.Multiselect = False

        'If(now_o_bom_value = 1, CStr(sOBomPath), CStr(sO_BOM_Path_Text
        If now_o_bom_value = 1 Then
            A_Main_Form.OpenFileDialog1.InitialDirectory = CStr(sOBomPath)
        Else
            A_Main_Form.OpenFileDialog1.InitialDirectory = CStr(sO_BOM_Path_Text)
        End If


        A_Main_Form.OpenFileDialog1.ShowDialog(A_Main_Form)
        'OpenFileDialog1 -> OpenFile() 읽기 전용으로 오픈
        '--------------------------------------------------------------------------------------------------- O-BOM 파일 없을 때
        If A_Main_Form.OpenFileDialog1.FileName = "" Then
            assy_value_Name = ""
            Error_Count = 1
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- O-BOM 경로 확인하기(O-BOM File 가져오기만 실행)
        If First_O_BOM = 1 Then
            BOM_File_Path = Split(A_Main_Form.OpenFileDialog1.FileName, "\")
            O_BOM_DB_Path = Split(sO_BOM_Path_text, "\")

            Path_Size_1 = UBound(BOM_File_Path)
            Path_Size_2 = UBound(O_BOM_DB_Path)

            If Path_Size_1 <= Path_Size_2 Then
                Temp_Size = Path_Size_1
            Else
                Temp_Size = Path_Size_2
            End If

            For i = 0 To Temp_Size
                If Not BOM_File_Path(i) = O_BOM_DB_Path(i) Then
                    Call SHOW_MESSAGE_BOX("O-BOM은 """ & sO_BOM_Path_text & """ 경로에서 가져와야 합니다.", "", 1, 1)
                    Diff_Path = 1
                    Exit Sub
                End If
            Next i
        End If

        sfile = A_Main_Form.OpenFileDialog1.FileName             '파일 오픈하고 파일선택
        Exc = CreateObject("excel.application")
        'Exc.Visible = True
        WkBook = Exc.Workbooks.Open(sfile)    'ReadOnly:=False
        Exc.Worksheets("BOM_LIST").Select()
        '--------------------------------------------------------------------------------------------------- 파일 경로 가져오기(PYO)
        Folder_Name = Split(A_Main_Form.OpenFileDialog1.FileName, vbNullChar)(0)
        ls_ss = Split(Folder_Name, "\")
        Origin_File_Name = ls_ss(UBound(ls_ss))
        ls_ss_folder = Split(Folder_Name, "\" & Origin_File_Name)
        A_Main_Form.re_o_bom_path.Text = ls_ss_folder(0)
        A_Main_Form.O_BOM_File_Name.Text = Origin_File_Name
        assy_value_Name_ref = Exc.Worksheets("BOM_LIST").Range("A" & 3).Value

        Dim O_BOM_FileNameSM21Check
        O_BOM_FileNameSM21Check = Split(Origin_File_Name, "_")
        If O_BOM_FileNameSM21Check(1) <> "SM21" Then
            Call SHOW_MESSAGE_BOX("O_BOM의 기종정보를 확인하세요.", "", 1, 1)
            Error_Count = 1
            Exit Sub
        End If
        If Left(assy_value_Name_ref, 4) <> "SM21" Then
            Call SHOW_MESSAGE_BOX("O_BOM의 기종정보를 확인하세요.", "", 1, 1)
            now_o_bom_value = 0
            Error_Count = 1
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------- TM
        If Left(assy_value_Name_ref, 2) = "TM" Then
            If Mid(assy_value_Name_ref, 3, 3) = "150" Or Mid(assy_value_Name_ref, 3, 3) = "125" Then
                assy_value_Name = "TMX"
            ElseIf Mid(assy_value_Name_ref, 3, 3) = "900" Or Mid(assy_value_Name_ref, 3, 3) = "800" Or Mid(assy_value_Name_ref, 3, 3) = "700" Then
                assy_value_Name = "TMY"
            ElseIf Mid(assy_value_Name_ref, 3, 3) = "400" Or Mid(assy_value_Name_ref, 3, 3) = "500" Or Mid(assy_value_Name_ref, 3, 3) = "600" Then
                assy_value_Name = "TMZ"
            Else
                Call SHOW_MESSAGE_BOX("O_BOM의 기종정보를 확인하세요.", "", 1, 1)
                now_o_bom_value = 0
                Error_Count = 1
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------------------- Lbl_ModelNumber에 기종정보 표시.
            A_Main_Form.Lbl_ModelNumber.Text = assy_value_Name
            '--------------------------------------------------------------------------------------------------- SME
        ElseIf Left(assy_value_Name_ref, 3) = "SME" Then
            assy_value_Name = Left(assy_value_Name_ref, 4)
            A_Main_Form.Lbl_ModelNumber.Text = assy_value_Name
            '--------------------------------------------------------------------------------------------------- SMx100
        ElseIf Mid(assy_value_Name_ref, 4, 1) = "1" Then
            assy_value_Name = Left(assy_value_Name_ref, 4)
            A_Main_Form.Lbl_ModelNumber.Text = assy_value_Name
            '--------------------------------------------------------------------------------------------------- SM
        Else
            assy_value_Name = Left(assy_value_Name_ref, 3)
            A_Main_Form.Lbl_ModelNumber.Text = assy_value_Name
        End If

        If Left(Origin_File_Name, 1) = "C" Then
            Sales_No = Left(Origin_File_Name, 10)
        Else
            Call SHOW_MESSAGE_BOX("O_BOM File Name의 Sales NO.를 확인하세요.", "", 1, 1)
            now_o_bom_value = 0
            Error_Count = 1
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- Excel File 닫기
        Exc.Quit()
        'WkBook.Close()
        Call KillProcess("EXCEL")
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' New_Folder_Create   &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub New_Folder_Create()
        Try
            Dim O_Bom_Path, o_bom_file_value, o_bom_file_value_Name, Exc_File_Path, Exc_File_Name
            Dim Exc_Extension, O_Bom_Name
            Dim WkBook = Nothing
            '--------------------------------------------------------------------------------------------------- 새로운 폴더 생성
            o_bom_excel_file_name = Sales_No & "_" & assy_value_Name

            fso = New FileSystemObject
            If Not (fso.FolderExists(A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\")) Then
                fso.CreateFolder(A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name)
                fso.CreateFolder(A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\" & "CLASH_CHECK")
                Threading.Thread.Sleep(1000)
            End If

            date_value = Format(Now, "yymmddhhmm")

            If now_o_bom_value = 1 Then
                O_Bom_Path = A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name

                Debug.Print("aa" & A_Main_Form.O_BOM_File_Name.Text)
                o_bom_file_value = fso.GetFile(O_Bom_Path & "\" & A_Main_Form.O_BOM_File_Name.Text)
                '--------------------------------------------------------------------------------------------------- O-BOM 파일 복사 및 이름 변경
                If Right(o_bom_file_value.Name, 1) = "x" Then

                    '--------------------------------------------------------------------------------------------------- 180313 C의 Temp로 복사
                    fso.CopyFile(A_Main_Form.re_o_bom_path.Text & "\" & A_Main_Form.O_BOM_File_Name.Text, "C:\Temp\" & A_Main_Form.O_BOM_File_Name.Text)

                    Threading.Thread.Sleep(1500)
                    '--------------------------------------------------------------------------------------------------- 180313 C의 Temp로 복사
                    o_bom_file_value = fso.GetFile("C:\Temp\" & A_Main_Form.O_BOM_File_Name.Text)

                    o_bom_file_value.Name = Sales_No & "_" & assy_value_Name & "_" & date_value & ".xlsx"
                    '--------------------------------------------------------------------------------------------------- 180313 C의 Temp로 복사
                    fso.CopyFile("C:\Temp\" & o_bom_file_value.Name, A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\")
                    Threading.Thread.Sleep(1500)
                    '--------------------------------------------------------------------------------------------------- C:\ 경로에 복사한 엑셀 삭제
                    o_bom_file_value_Name = o_bom_file_value.Name
                    o_bom_file_value.Delete()

                    '--------------------------------------------------------------------------------------------------- DB 경로에 복사한 엑셀
                    o_bom_file_value = fso.GetFile(A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\" & o_bom_file_value_Name)

                    A_Main_Form.O_BOM_File_Name.Text = o_bom_file_value.Name
                Else
                    Exc_File_Path = A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name
                    Exc_File_Name = Exc_File_Path & "\" & Sales_No & "_" & assy_value_Name & "_" & date_value & ".xlsx"

                    Exc = CreateObject("excel.application")
                    'Exc.Visible = true
                    Exc.Application.DisplayAlerts = False
                    WkBook = Exc.Workbooks.Open(A_Main_Form.re_o_bom_path.Text & "\" & A_Main_Form.O_BOM_File_Name.Text)
                    Threading.Thread.Sleep(2000)
                    Call Exc.ActiveWorkbook.SaveAs(Exc_File_Name, 51)
                    Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()

                    A_Main_Form.O_BOM_File_Name.Text = Sales_No & "_" & assy_value_Name & "_" & date_value & ".xlsx"
                End If
            Else
                O_Bom_Path = A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name

                If Right(A_Main_Form.O_BOM_File_Name.Text, 1) = "x" Then
                    Exc_Extension = ".xlsx"
                    O_Bom_Name = Sales_No & "_" & assy_value_Name & "_" & date_value & Exc_Extension

                    '--------------------------------------------------------------------------------------------------- O-BOM 파일 복사 및 이름 변경
                    fso.CopyFile(sO_BOM_Path_Text & "\" & A_Main_Form.O_BOM_File_Name.Text, O_Bom_Path & "\" & O_Bom_Name)
                Else
                    Exc_Extension = ".xls"

                    '--------------------------------------------------------------------------------------------------- O-BOM 파일 복사 및 이름 변경
                    fso.CopyFile(A_Main_Form.O_BOM_Path_Text.Text & "\" & A_Main_Form.O_BOM_File_Name.Text, O_Bom_Path & "\" & date_value & Exc_Extension)
                    Threading.Thread.Sleep(1500)

                    O_Bom_Name = Sales_No & "_" & assy_value_Name & "_" & date_value & ".xlsx"

                    o_bom_file_value = fso.GetFile(O_Bom_Path & "\" & date_value & Exc_Extension)

                    Exc = CreateObject("excel.application")
                    'Exc.Visible = true
                    Exc.Application.DisplayAlerts = False
                    WkBook = Exc.Workbooks.Open(o_bom_file_value)
                    Threading.Thread.Sleep(1500)
                    Call Exc.ActiveWorkbook.SaveAs(O_Bom_Path & "\" & O_Bom_Name, 51)
                    Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
                    Threading.Thread.Sleep(1500)

                    o_bom_file_value.Delete()
                End If

                A_Main_Form.O_BOM_File_Name.Text = O_Bom_Name
            End If
            Try
                Exc.Quit()
                If Not (WkBook Is Nothing) Then
                    WkBook.Close()
                End If
                Exc = Nothing
                o_bom_file_value = Nothing
                '--------------------------------------------------------------------------------------------------- 실행된 Excel이 없어서 오류나는 경우 무시 
            Catch ex As NullReferenceException
            Catch ex As System.Runtime.InteropServices.COMException
            Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            End Try
            Call KillProcess("EXCEL")
            A_Main_Form.re_o_bom_path.Text = A_Main_Form.Project_DB_Path_Text.Text & "\" & Sales_No & "_" & assy_value_Name
            now_o_bom_excel_path = A_Main_Form.re_o_bom_path.Text & "\" & A_Main_Form.O_BOM_File_Name.Text

            '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
            Exc = CreateObject("excel.application")
            'Exc.Visible = true
            WkBook = Exc.Workbooks.Open(System.Windows.Forms.Application.StartupPath & "\data_path.xlsx")
            Exc.Worksheets("Sheet1").Range("C" & 3).Value = A_Main_Form.re_o_bom_path.Text
            Exc.Worksheets("Sheet1").Range("A" & 5).Value = A_Main_Form.O_BOM_File_Name.Text
            Exc.Worksheets("Sheet1").Range("A" & 10).Value = A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\" & A_Main_Form.O_BOM_File_Name.Text
            Try
                Exc.Application.DisplayAlerts = False
                Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
                Exc.ActiveWorkbook.Save()
                Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
                '
                Exc.Quit()
                If Not (WkBook Is Nothing) Then
                    WkBook.Close()
                End If
                '--------------------------------------------------------------------------------------------------- 실행된 Excel이 없어서 오류나는 경우 무시 
            Catch ex As NullReferenceException
            Catch ex As System.Runtime.InteropServices.COMException
            Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            End Try
            Call KillProcess("EXCEL")
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("New_Folder_Create 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Data Loading    &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Data_Loading_Click()
        Try
            Dim i, j, k
            Dim arrayOfVariantOfBSTR1(0)
            ReDim layout_EPN_ref(0 To O_BOM_Part_Value_Count)
            ReDim layout_part_name_ref(0 To O_BOM_Part_Value_Count)
            Dim ProductsTmp As ProductStructureTypeLib.Products = Nothing
            pub_number = 1

            A_Main_Form.Message_Label.Text = "Data Loading중 입니다"
            '--------------------------------------------------------------------------------------------------- MSFlexGrid1에 Value 넣기(대기)
            A_Main_Form.DataGridView1.ColumnHeadersVisible = True
            A_Main_Form.DataGridView1.Columns(0).HeaderText = "No"
            A_Main_Form.DataGridView1.Columns(1).HeaderText = "Location 정보"
            A_Main_Form.DataGridView1.Columns(2).HeaderText = "Missing"
            A_Main_Form.DataGridView1.Columns(3).HeaderText = "O-BOM PartNumber"
            A_Main_Form.DataGridView1.Columns(4).HeaderText = "O-BOM Part Name"
            '--------------------------------------------------------------------------------------------------- MSFlexGrid1 Width
            A_Main_Form.DataGridView1.Columns(0).Width = 30
            A_Main_Form.DataGridView1.Columns(1).Width = 120
            A_Main_Form.DataGridView1.Columns(2).Width = 60
            A_Main_Form.DataGridView1.Columns(3).Width = 160
            A_Main_Form.DataGridView1.Columns(4).Width = 400

            '--------------------------------------------------------------------------------------------------- STD BOM List를 기준으로 Part Number List 생성
            std_p_number_count = 1
            '--------------------------------------------------------------------------------------------------- O_BOM_Layout_Part Count(PYO)
            Try
                For i = 2 To O_BOM_Part_Value_Count - 1     'O-BOM 파일
                    For j = 1 To O_BOM_Layout_Part_Value_Count
                        If O_BOM_Part_Value(i, 3).ToString() = O_BOM_Layout_Part_Value(j, 1).ToString() Then
                            A_Main_Form.DataGridView1.Rows.Add()
                            A_Main_Form.DataGridView1.Item(0, std_p_number_count - 1).Value = CStr(std_p_number_count)
                            A_Main_Form.DataGridView1.Item(1, std_p_number_count - 1).Value = CStr(O_BOM_Part_Value(i, 3))
                            A_Main_Form.DataGridView1.Item(3, std_p_number_count - 1).Value = CStr(O_BOM_Part_Value(i, 1))
                            A_Main_Form.DataGridView1.Item(4, std_p_number_count - 1).Value = CStr(O_BOM_Part_Value(i, 2))
                            layout_EPN_ref(std_p_number_count) = O_BOM_Part_Value(i, 3)
                            std_p_number_count = std_p_number_count + 1
                            Exit For
                        End If
                    Next
                Next
            Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                Call SHOW_MESSAGE_BOX("연결 정보가 잘 못 설정되어 종료합니다.", "", 1, 1)
                Exit Sub
            End Try

            ' 진행중 Product가 아닐때
            If Get_Layout_List_Number = 0 Then
                '--------------------------------------------------------------------------------------------------- 새로운 Product 추가
                product_add = Documents.Add("Product")
                products = product_add.Product.products
                ProductsTmp = products
            End If

            Dim FName(20)
            Dim part_value_count(200) As Integer
            Dim part_loading_count, numbering_data_path
            pub_number = 1
            axis_number = 1
            part_loading_count = 0
            FileListCounter = 0
            Dim partCount
            fso = CreateObject("Scripting.FileSystemObject")
            Dim folder
            '--------------------------------------------------------------------------------------------------- BOM List를 기준으로 Part Name List 생성
            For i = 1 To std_p_number_count - 1
                For j = 2 To O_BOM_Part_Value_Count - 1
                    '--------------------------------------------------------------------------------------------------- O-Bom EPN과 STD Bom EPN 비교
                    If CStr(layout_EPN_ref(i)) = CStr(O_BOM_Part_Value(j, 3)) Then
                        '--------------------------------------------------------------------------------------------------- Data에서 _파일 찾기
                        For k = 1 To 255
                            If Mid(O_BOM_Part_Value(j, 1), k, 1) = "" Then
                                Exit For
                            End If
                        Next

                        '--------------------------------------------------------------------------------------------------- Folder File Count
                        numbering_data_path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & layout_EPN_ref(i)
                        A_Main_Form.File_numbering.Path = "C:\"
                        '---------------------------------------------------------------------------------------------------
                        ' 생 성 일  : 20.04.16
                        ' 생 성 자  : 김원주
                        ' 수 정 일  : -
                        ' 수 정 자  : -
                        ' 수정사유  : 파일 존재여부 확인 후 없으면  Error Message 출력
                        ' Parameter : 
                        '---------------------------------------------------------------------------------------------------
                        If Dir(numbering_data_path, vbDirectory) = "" Then
                            MsgBox(layout_EPN_ref(i) & "폴더가 존재하지 않음", vbSystemModal)
                            Continue For

                        End If
                        '---------------------------------------------------------------------------------------------------
                        A_Main_Form.File_numbering.Path = numbering_data_path
                        If IsNumeric(O_BOM_Part_Value(j, 4)) Then
                            partCount = O_BOM_Part_Value(j, 4)
                        Else
                            partCount = 1
                        End If
                        If A_Main_Form.File_numbering.Path = numbering_data_path Then
                            '---------------------------------------------------------------------------------------------------
                            ' 생 성 일  : 20.02.19
                            ' 생 성 자  : 허혜원
                            ' 수 정 일  : -
                            ' 수 정 자  : -
                            ' 수정사유  : BOM_LIST 수량 상관없이 1개만 가져오도록 수정 -> For pcount = 1 To 1 'partCount 로 변경
                            ' Parameter : 
                            '---------------------------------------------------------------------------------------------------
                            For m = 0 To A_Main_Form.File_numbering.Items.Count - 1
                                If CStr(O_BOM_Part_Value(j, 1) & ".CATPart") = CStr(A_Main_Form.File_numbering.Items(m)) Then
                                    For pcount = 1 To 1 'partCount
                                        A_Main_Form.data_open_path_list.Items.Add(sDirectory_Path_Text & "\" & assy_value_Name & "\" & layout_EPN_ref(i) & "\" & O_BOM_Part_Value(j, 1) & ".CATPart")
                                        A_Main_Form.data_open_list.Items.Add(O_BOM_Part_Value(j, 3))
                                        fso = CreateObject("Scripting.FileSystemObject")
                                        folder = fso.GetFolder(sDirectory_Path_Text & "\" & assy_value_Name & "\" & layout_EPN_ref(i))
                                        For Each EachFile In folder.Files
                                            Dim fileName As String = EachFile.Name
                                            If fileName.StartsWith(O_BOM_Part_Value(j, 1) & "_") And fileName.EndsWith(".CATPart") Then
                                                Debug.Print("A")
                                            End If
                                        Next

                                    Next pcount

                                    Exit For
                                ElseIf CStr(O_BOM_Part_Value(j, 1)) = Mid(A_Main_Form.File_numbering.Items(m), 1, k - 1) Then
                                    If Mid(A_Main_Form.File_numbering.Items(m), k, 1) = "_" Then
                                        For pcount = 1 To 1 'partCount
                                            Debug.Print(sDirectory_Path_Text & "\" & assy_value_Name & "\" & layout_EPN_ref(i) & "\" & A_Main_Form.File_numbering.Items(m))
                                            A_Main_Form.data_open_path_list.Items.Add(sDirectory_Path_Text & "\" & assy_value_Name & "\" & layout_EPN_ref(i) & "\" & A_Main_Form.File_numbering.Items(m))
                                            A_Main_Form.data_open_list.Items.Add(O_BOM_Part_Value(j, 3))
                                        Next pcount
                                        '---------------------------------------------------------------------------------------------------
                                        ' 생 성 일  : 20.02.20
                                        ' 생 성 자  : 허혜원
                                        ' 수 정 일  : -
                                        ' 수 정 자  : -
                                        ' 수정사유  :    1. 기존
                                        '               2. Exit For 주석처리 하면 폴더 내 파일명이 앞 O_BOM_Part_Value(j, 1) & "_" 까지 일치하는 모든 Part 가져옴
                                        ' Parameter : 
                                        '---------------------------------------------------------------------------------------------------
                                        'Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Next

            A_Main_Form.File_numbering.Refresh()
            If Get_Layout_List_Number = 0 Then
                '--------------------------------------------------------------------------------------------------- 새로운 Part 추가
                For i = 0 To A_Main_Form.data_open_path_list.Items.Count - 1
                    If i Mod 10 = 0 Then
                        '--------------------------------------------------------------------------------------------------- Zoom All
                        CATIA.ActiveWindow.ActiveViewer.Reframe()
                    End If
                    arrayOfVariantOfBSTR1(0) = A_Main_Form.data_open_path_list.Items(i)

                    ProductsTmp.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

                    '--------------------------------------------------------------------------------------------------- Part 수량 체크
                    part_value_count(i) = products.Count
                    If i = 0 Then
                        If part_value_count(0) = 1 Then
                            Try
                                products.Item(1).Name = A_Main_Form.data_open_list.Items(0)
                            Catch ex As Exception

                            End Try
                        End If
                    Else
                        If Not part_value_count(i) = part_value_count(i - 1) Then
                            Debug.Print(A_Main_Form.data_open_list.Items(i))
                            Try
                                products.Item(products.Count).Name = A_Main_Form.data_open_list.Items(i)
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                Next
                '--------------------------------------------------------------------------------------------------- Publication 가져오기
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : 
                ' 수 정 일  : 19.12.19
                ' 수 정 자  : 
                ' 수정사유  : EPN 중복항목 체크 추가
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                get_part_publication2017.Clear()
                For i = 1 To products.Count
                    publications1 = products.Item(i).publications
                    For k = 1 To publications1.Count
                        If Not get_part_publication2017.ContainsKey(publications1.Item(k).Name.ToString()) Then
                            get_part_publication2017.Add(publications1.Item(k).Name.ToString(), publications1.Item(k))
                            get_part_publication(pub_number) = publications1.Item(k)
                            pub_number = pub_number + 1
                        Else
                            If Not Tmp_Duplication_Check_Dic.ContainsKey(publications1.Item(k).Name.ToString()) Then
                                Tmp_Duplication_Check_Dic.Add(publications1.Item(k).Name.ToString(), publications1.Item(k).Name.ToString())
                            Else
                                Debug.Print(publications1.Item(k).Name.ToString())
                                'MsgBox(publications1.Item(k).Name.ToString() & " 이 중복됩니다.")
                            End If
                        End If
                    Next
                Next
            End If

            '--------------------------------------------------------------------------------------------------- Zoom All
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            A_Main_Form.Message_Label.Text = "Data Loading이 완료되었습니다"
            A_Main_Form.ProgressBar1.Value = 50
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Data_Loading_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Comparison Data &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Comparison_Data()
        Try
            '--------------------------------------------------------------------------------------------------- 미확정 표시 List 생성(PYO)
            '--------------------------------------------------------------------------------------------------- Data Missing List 생성
            Dim layout_EPN_ref_count = 1
            Dim jNum As Integer
            Dim products_name As String = ""
            For i = 1 To std_p_number_count - 1
                For j = 1 To O_BOM_Part_Value_Count
                    If layout_EPN_ref(i) = O_BOM_Part_Value(j, 3) Then
                        A_Main_Form.DataGridView1.Refresh()

                        For k = 1 To CATIA.ActiveDocument.Product.products.Count
                            '--------------------------------------------------------------------------------------------------- Excel 이름 가져오기
                            For m = 1 To 255
                                If Mid(O_BOM_Part_Value(j, 2), m, 1) = "" Then
                                    Exit For
                                End If
                            Next
                            '--------------------------------------------------------------------------------------------------- Excel 이름과 자리수 비교
                            products_name = CATIA.ActiveDocument.Product.products.Item(k).Name
                            If O_BOM_Part_Value(j, 3) = products_name Then
                                A_Main_Form.DataGridView1.Item(2, i - 1).Value = ""
                                Exit For
                            End If
                        Next

                        jNum = j
                        Exit For
                    End If
                Next

                If Not O_BOM_Part_Value(jNum, 3) = products_name Then
                    A_Main_Form.DataGridView1.Item(2, i - 1).Value = "V"
                End If
            Next

            A_Main_Form.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            A_Main_Form.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            A_Main_Form.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            A_Main_Form.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            A_Main_Form.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            '--------------------------------------------------------------------------------------------------- STD BOM List를 기준으로 Part Name List 생성
            A_Main_Form.ProgressBar1.Value = 60
            A_Main_Form.DataGridView1.Refresh()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Comparison_Data 함수 오류!!", "", 1, 1)
        End Try
    End Sub


#Region " --- Axis_Setting() : Axis 생성 및 퍼플리케이션 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.03
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Axis_Setting()
        ' Constraint 재구성 여러번 실행시 axis_selection 이름이 중복이라 에러떠서 try Catch 안씀 / lsm / 190719
        On Error Resume Next
        Call CATIA_BASIC02()

        Dim oProducts As Products
        oProducts = ActiveDocument.Product.Products()
        '--------------------------------------------------------------------------------------------------- 기존에 남아있는 불필요한 axis_selection part 삭제
        selection = CATIA.ActiveDocument.selection
        selection.clear()
        selection.Search("(Name=axis_selection* & CATPrtSearch.PartFeature),all")
        If selection.Count > 0 Then
            selection.Delete()
        End If
        oProducts.AddNewComponent("Part", "axis_selection")
        Dim partDocument1 = Documents.Item(Documents.Count)
        Dim add_part = partDocument1.Part
        '--------------------------------------------------------------------------------------------------- Axis 생성
        Dim axisSystems1 As AxisSystems = add_part.AxisSystems
        Dim axisSystem1 As AxisSystem = axisSystems1.Add()
        axisSystem1.Name = "axissystem_selection"
        axisSystem1.IsCurrent = True

        Dim arrayOfVariantOfDouble1(2)
        arrayOfVariantOfDouble1(0) = 0.0#
        arrayOfVariantOfDouble1(1) = 0.0#
        arrayOfVariantOfDouble1(2) = 0.0#

        axisSystem1.OriginType = 1
        axisSystem1.PutOrigin(arrayOfVariantOfDouble1)

        arrayOfVariantOfDouble1(0) = 1.0#
        arrayOfVariantOfDouble1(1) = 0.0#
        arrayOfVariantOfDouble1(2) = 0.0#

        axisSystem1.XAxisType = 1
        axisSystem1.PutXAxis(arrayOfVariantOfDouble1)

        arrayOfVariantOfDouble1(0) = 0.0#
        arrayOfVariantOfDouble1(1) = 1.0#
        arrayOfVariantOfDouble1(2) = 0.0#

        axisSystem1.YAxisType = 1
        axisSystem1.PutYAxis(arrayOfVariantOfDouble1)

        arrayOfVariantOfDouble1(0) = 0.0#
        arrayOfVariantOfDouble1(1) = 0.0#
        arrayOfVariantOfDouble1(2) = 1.0#

        axisSystem1.ZAxisType = 1
        axisSystem1.PutZAxis(arrayOfVariantOfDouble1)

        add_part.UpdateObject(axisSystem1)
        add_part.Update()
        Threading.Thread.Sleep(300)

        '--------------------------------------------------------------------------------------------------- Axis Publication 생성
        add_axis_Product = partDocument1.GetItem("axis_selection")

        Dim reference1 = add_axis_Product.CreateReferenceFromName("axis_selection/!axissystem_selection")
        publications1 = add_axis_Product.publications

        Dim publication2 = publications1.Add("axissystem_selection")
        publications1.SetDirect("axissystem_selection", reference1)

        '--------------------------------------------------------------------------------------------------- Product Fix 구속조건
        Dim productDocument1 = CATIA.ActiveDocument
        Product1 = productDocument1.Product
        constraints1 = Product1.Connections("CATIAConstraints")
        name_ref = Product1.CreateReferenceFromName(Product1.Name & "/" & Product1.products.Item(Product1.products.Count).Name & "/!" & Product1.Name & "/" & Product1.products.Item(Product1.products.Count).Name & "/")
        product_fix = constraints1.AddMonoEltCst(0, name_ref)   'FIX

        '--------------------------------------------------------------------------------------------------- Axis 생성
        Call CATIA_BASIC02()
        selection.clear
        '--------------------------------------------------------------------------------------------------- 하위 Item 구속조건 적용
        '--------------------------------------------------------------------------------------------------- 1번째 Publication 찾기
        selection.search("Name='EPN010101_P01',all")
        Pub_Search = selection.Item(2).Value
        Pub1_DisplayName = Pub_Search.Valuation.DisplayName
        selection.clear
        '--------------------------------------------------------------------------------------------------- 2번째 Publication 이름 정하기
        selection.search("Name=axissystem_selection*,all")
        Pub_Search = selection.Item(2).Value
        Pub2_DisplayName = Pub_Search.Valuation.DisplayName

        '--------------------------------------------------------------------------------------------------- 1번째 Publication 이름 정하기
        Pub_reference1 = Product1.CreateReferenceFromName(Pub1_DisplayName)
        Pub_reference2 = Product1.CreateReferenceFromName(Pub2_DisplayName)
        coincidence_constraint(0) = constraints1.AddBiEltCst(2, Pub_reference1, Pub_reference2)
        Threading.Thread.Sleep(500)

        If Publish_Sub_Name(0) = "" Then
            'coincidence_constraint(0).Name = Publish_Sub_Name(0)
        Else
            coincidence_constraint(0).Name = Publish_Sub_Name(0)
        End If
        Product1.Update()
        Threading.Thread.Sleep(300)

        '--------------------------------------------------------------------------------------------------- 불필요한 Constraint 제거
        Call CATIA_BASIC02()

        selection.Clear()
        selection.Add(Product1.products.Item(Product1.products.Count))  'axis_selection
        selection.Add(constraints1.Item(1))                             'Constraint : Fix
        selection.Add(constraints1.Item(2))                             'Constraint : EPN010101
        selection.Delete()
        Threading.Thread.Sleep(300)

        '--------------------------------------------------------------------------------------------------- Bull Gear Fix
        name_ref = Product1.CreateReferenceFromName(Product1.Name & "/EPN010101/!" & Product1.Name & "/EPN010101/")
        product_fix = constraints1.AddMonoEltCst(0, name_ref)

        A_Main_Form.ProgressBar1.Value = 65
        On Error GoTo 0
    End Sub
#End Region


    '---------------------------------------------------------------------------------------------------
    '구속조건 생성    &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Constraint_Setting()
        A_Main_Form.Message_Label.Text = " Data Position Setting하고 있습니다"
        CATIA_BASIC02()
        fso = CreateObject("Scripting.FileSystemObject")
        '--------------------------------------------------------------------------------------------------- Product Fix 구속조건 적용
        constraints1 = product_add.Connections("CATIAConstraints")
        '--------------------------------------------------------------------------------------------------- Constraint 오류 출력용 file
        Dim file As System.IO.StreamWriter
        Try
            '--------------------------------------------------------------------------------------------------- 하위 Item 구속조건 적용
            Dim Stop_Constraint = 0
            For i = 1 To O_BOM_Constraint_Value_Count - 1
                Dim sw = 0
                '--------------------------------------------------------------------------------------------------- CON_TMP_01 이면...
                If O_BOM_Constraint_Value(i, 2) = "CON_TMP_01" Then
                    For k = 1 To Stop_Constraint
                        If O_BOM_Constraint_Value(i, 1) = O_BOM_Constraint_Value(k, 1) Then
                            constraints1 = product_add.Connections("CATIAConstraints")
                            For zb = 1 To constraints1.Count
                                If constraints1.Item(zb).Name = O_BOM_Constraint_Value(k, 2) Then
                                    sw = 1
                                    Exit For
                                End If
                            Next zb
                        End If
                        If sw = 1 Then
                            Exit For
                        End If
                    Next k
                    If Not sw = 1 Then
                        For k = 1 To Stop_Constraint
                            If O_BOM_Constraint_Value(i, 1) = O_BOM_Constraint_Value(k, 2) Then
                                constraints1 = product_add.Connections("CATIAConstraints")
                                For zb = 1 To constraints1.Count
                                    If constraints1.Item(zb).Name = O_BOM_Constraint_Value(k, 2) Then
                                        sw = 1
                                        Exit For
                                    End If
                                Next zb
                            End If
                            If sw = 1 Then
                                Exit For
                            End If
                        Next k
                    End If
                Else
                    Stop_Constraint = i
                End If
                '--------------------------------------------------------------------------------------------------- 이미 Constraint가 적용 되었으면 넘김.
                If Not sw = 1 Then
                    '--------------------------------------------------------------------------------------------------- 1번째 Publication 찾기
                    '/// Constraint1 또는 Constraint2 값이 현재 Product 내의 publication List에 없을 경우 Continue For
                    'If Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 1)) Or Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 2)) Then
                    If Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 1)) Or Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 2)) Then
                        If (fso.FileExists("C:\SM21_Exe_File\Publications.txt") = True) Then
                            file = My.Computer.FileSystem.OpenTextFileWriter("C:\SM21_Exe_File\Publications.txt", True)

                            If Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 1)) Then
                                file.WriteLine(Format(Now, "yymmddhhmm") & " publication 1 없음 : " & O_BOM_Constraint_Value(i, 1))
                            End If

                            If Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 2)) Then
                                file.WriteLine(Format(Now, "yymmddhhmm") & " publication 2 없음 : " & O_BOM_Constraint_Value(i, 2))
                            End If
                            file.Close()
                        End If
                        Continue For
                    End If

                    '--------------------------------------------------------------------------------------------------- 1번째 Publication 찾기
                    Call GetPub_DisplayName(Pub1_DisplayName, O_BOM_Constraint_Value(i, 1))
                    If (Pub1_DisplayName = "Empty") Then
                        Continue For
                    End If

                    Call GetPub_reference(Pub_reference1, Pub1_DisplayName, O_BOM_Constraint_Value(i, 1))
                    '--------------------------------------------------------------------------------------------------- 2번째 Publication 찾기
                    Call GetPub_DisplayName(Pub2_DisplayName, O_BOM_Constraint_Value(i, 2))
                    If (Pub1_DisplayName = "Empty" Or Pub2_DisplayName = "Empty") Then
                        Continue For
                    End If
                    Call GetPub_reference(Pub_reference2, Pub2_DisplayName, O_BOM_Constraint_Value(i, 2))

                    '---------------------------------------------------------------------------------------------------
                    ' 생 성 일  : 19.08.29
                    ' 생 성 자  : 허혜원
                    ' 수 정 일  : -
                    ' 수 정 자  : -
                    ' 수정사유  : Constraint 생성안될 경우 무시하고 Err.txt에 출력하기
                    ' Parameter : 
                    '---------------------------------------------------------------------------------------------------
                    'fso = CreateObject("Scripting.FileSystemObject")
                    'If (fso.FileExists(System.Windows.Forms.Application.StartupPath & "\Err.txt") = False) Then
                    'fso.CreateTextFile(System.Windows.Forms.Application.StartupPath & "\Err.txt")
                    'End If
                    '--------------------------------------------------------------------------------------------------- 구속 생성
                    Try
                        coincidence_constraint(i) = constraints1.AddBiEltCst(2, Pub_reference1, Pub_reference2)
                        coincidence_constraint(i).Name = O_BOM_Constraint_Value(i, 2)
                    Catch exPrint As Exception
                        Debug.Print(exPrint.StackTrace)
                        If (fso.FileExists(System.Windows.Forms.Application.StartupPath & "\Err.txt") = True) Then
                            file = My.Computer.FileSystem.OpenTextFileWriter(System.Windows.Forms.Application.StartupPath & "\Err.txt", True)
                            file.WriteLine("Constraint Err : " & Format(Now, "mmddhhmm") & "  " & Pub1_DisplayName & " ■ " & Pub2_DisplayName)
                            file.Close()
                        End If
                    End Try
                    'VB6.0->2017 업데이트 + R26부터 발생하는 Over-Constrain 알림창 제거.
                    Call Constraint_Update_ExceptionHandling(product_add, constraints1)

                    selection.Clear()
                    For j = 1 To constraints1.Count
                        '--------------------------------------------------------------------------------------------------- CON_TMP이면 Constraint 삭제.
                        If Left(constraints1.Item(j).Name, 7) = "CON_TMP" Then
                            Threading.Thread.Sleep(200)
                            selection.Add(constraints1.Item(j))
                        End If
                    Next j
                    If selection.Count <> 0 Then
                        selection.Delete()
                        Threading.Thread.Sleep(200)
                    End If
                End If
                Debug.Print(i & "/ " & O_BOM_Constraint_Value_Count - 1)
            Next i
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()
            product_add.Update()
            Threading.Thread.Sleep(200)
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Constraint_Setting 함수 오류!!", "", 1, 1)
        End Try

        '--------------------------------------------------------------------------------------------------- MOTOR 정보가 있는지 검색...
        Dim Motor_Keyway = 0
        selection.clear
        selection.search("Name=EPN02_D01,all")
        Motor_Keyway = selection.Count
        selection.clear
        selection.search("Name=EPN02_D01,all")
        Motor_Keyway = Motor_Keyway + selection.Count

        If Motor_Keyway = 0 Then
            Call SHOW_MESSAGE_BOX("모터 또는 모터에 Keyway 정보가 없습니다.", "", 1, 1)
        End If

        A_Main_Form.Message_Label.Text = "Data Position Setting이 완료 되었습니다"
        '--------------------------------------------------------------------------------------------------- Axis Enable
        product_add.Update()
        A_Main_Form.Radio1.Enabled = True
        A_Main_Form.Radio2.Enabled = True
        A_Main_Form.ProgressBar1.Value = 85
    End Sub


    Public Function GetPub_DisplayName(ByRef Pub_DisplayName, Constraint_Name)
        Pub_DisplayName = vbEmpty.ToString()
        Try
            If get_part_publication2017.ContainsKey(Constraint_Name) Then
                Pub_DisplayName = get_part_publication2017.Item(Constraint_Name).Valuation.DisplayName
                GetPub_DisplayName = True
            Else
                GetPub_DisplayName = False
            End If
        Catch ex As Exception
        End try
    End Function


#Region " --- Constraint_Update_ExceptionHandling() : 오류인 Constraint -> Deactivate "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.07.02
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : product_add : 최상위 프로덕트
    '             constraints1 : 최상위 프로덕트의 모든 Constraints
    '---------------------------------------------------------------------------------------------------
    Public Sub Constraint_Update_ExceptionHandling(product_add, constraints1)
        If A_Main_Form.Check1.Checked = True Then
            Try
                product_add.Update()
                Threading.Thread.Sleep(200)
            Catch ex As System.Runtime.InteropServices.COMException
                '--------------------------------------------------------------------------------------------------- constraint update 오류가 발생하면..
                For j = 1 To constraints1.Count
                    '--------------------------------------------------------------------------------------------------- 연결오류이면 Constraint Deactivate
                    If Not (constraints1.Item(j).Status = 0) Then
                        constraints1.Item(j).Deactivate()
                    End If
                Next j
            End Try
        End If

        '--------------------------------------------------------------------------------------------------- R26부터 발생하는 Over-Constrain 알림창 제거.
        '--------------------------------------------------------------------------------------------------- Update Diagnosis Information 창 닫기
        Dim Window_HWND As IntPtr = 0
        Window_HWND = FindWindow(vbNullString, "Update Diagnosis Information: ")
        Threading.Thread.Sleep(200)

        If Window_HWND = 0 Then
            Window_HWND = FindWindow(vbNullString, "Update Diagnosis: Product1")
            Threading.Thread.Sleep(200)

            If Window_HWND = 0 Then
                Dim tmp_str = Split(CATIA.ActiveDocument.Name, ".")
                Dim ActiveDocument_Name = tmp_str(0)

                Window_HWND = FindWindow(vbNullString, "Update Diagnosis: " & ActiveDocument_Name)
                Threading.Thread.Sleep(200)
            End If
        End If

        If Not Window_HWND = 0 Then
            Dim Button_HWND = FindWindowEx(Window_HWND, IntPtr.Zero, "Button", Nothing) 'IntPtr.Zero
            Threading.Thread.Sleep(200)

            SendMessage(Button_HWND, BM_CLICK, 0, 0)
            Threading.Thread.Sleep(200)
        End If
    End Sub
#End Region


    '---------------------------------------------------------------------------------------------------
    ' Error_Constraint_Delete &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Error_Constraint_Delete()
        Try
            Call CATIA_BASIC02()

            If Get_Layout_List_Number = 1 Then  '진행중 Product 일 때
                constraints1 = CATIA.ActiveDocument.Product.Connections("CATIAConstraints")
                selection.Clear()
                '--------------------------------------------------------------------------------------------------- CON_TMP 삭제
                For i = 1 To constraints1.Count
                    If Left(constraints1.Item(i).Name, 7) = "CON_TMP" Then
                        selection.Add(constraints1.Item(i))
                    End If
                Next

                If selection.Count > 0 Then
                    selection.Delete()
                End If
            End If

            A_Main_Form.ProgressBar1.Value = 90
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Error_Constraint_Delete() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Comp_Product_Name_Change  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Function Comp_Product_Name_Change()
        Dim i, excel_name_value
        Comp_Product_Name_Change = True
        '--------------------------------------------------------------------------------------------------- Product name 변경
        Try
            For i = 1 To 255
                If Mid(A_Main_Form.O_BOM_File_Name.Text, i, 1) = "." Then
                    Exit For
                End If
            Next
            excel_name_value = Mid(A_Main_Form.O_BOM_File_Name.Text, 1, i - 1)
            CATIA.ActiveDocument.Product.PartNumber = Sales_No & "_" & assy_value_Name & "_" & date_value
            '--------------------------------------------------------------------------------------------------- 생성된 Product 저장
            CATIA.ActiveDocument.Product.Update()

            COMP_Product_File_Name = Sales_No & "_" & assy_value_Name & "_" & date_value

            Dim sSavePath As String
            sSavePath = A_Main_Form.Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\" & Sales_No & "_" & assy_value_Name & "_" & date_value & ".CATProduct"

            CATIA.DisplayFileAlerts = False
            CATIA.ActiveDocument.SaveAs(sSavePath)
            CATIA.DisplayFileAlerts = True
        Catch ex As System.ArgumentException
            Call SHOW_MESSAGE_BOX("CATProduct 파일 저장시 오류가 발생하였습니다.", "", 1, 1)
            Comp_Product_Name_Change = False
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("CATProduct 파일 저장시 오류가 발생하였습니다.", "", 1, 1)
            Comp_Product_Name_Change = False
        End Try

        If now_o_bom_value = 1 Then
            '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
            Call KillProcess("EXCEL")

            Exc = CreateObject("excel.application")
            Exc.Workbooks.Open(System.Windows.Forms.Application.StartupPath & "\data_path.xlsx")
            Exc.Worksheets("Sheet1").Range("C" & 3).Value = A_Main_Form.re_o_bom_path.Text
            Exc.Worksheets("Sheet1").Range("A" & 5).Value = A_Main_Form.O_BOM_File_Name.Text
            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
            Exc.ActiveWorkbook.Save()
            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()

            Call KillProcess("EXCEL")
            A_Main_Form.Message_Label.Text = "진행중 O-BOM Data 구성이 완료 되었습니다"
        Else
            A_Main_Form.Message_Label.Text = "O-BOM Data 구성이 완료 되었습니다"
        End If
        A_Main_Form.ProgressBar1.Value = 100

        'Return 0
    End Function


    '---------------------------------------------------------------------------------------------------
    ' Get_Layout_List_File_Open_Function  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Get_Layout_List_File_Open_Function()
        Try
            Get_Layout_List_Number = 1
            '--------------------------------------------------------------------------------------------------- File Open
            A_Main_Form.OpenFileDialog1.Title = "Open File"
            A_Main_Form.OpenFileDialog1.Filter = "CATProduct File(*.CATProduct)|*.CATProduct"          'CATProduct 파일만
            A_Main_Form.OpenFileDialog1.FileName = ""
            A_Main_Form.OpenFileDialog1.ReadOnlyChecked = True
            A_Main_Form.OpenFileDialog1.ShowReadOnly = True '= cdlOFNNoReadOnlyReturn            '읽기 전용
            A_Main_Form.OpenFileDialog1.Multiselect = False
            A_Main_Form.OpenFileDialog1.InitialDirectory = A_Main_Form.re_o_bom_path.Text
            A_Main_Form.OpenFileDialog1.ShowDialog(A_Main_Form)

            If A_Main_Form.OpenFileDialog1.FileName = "" Then
                assy_value_Name = ""
                Get_Layout_List_Number = 0
                Exit Sub
            End If

            Dim CATIA_Product_Open_Path = A_Main_Form.OpenFileDialog1.FileName             '파일 오픈하고 파일선택
            'Exc.Visible = True
            A_Main_Form.Message_Label.Text = "진행중 Product 구성 진행중..."
            '--------------------------------------------------------------------------------------------------- 파일 경로 가져오기
            ls_ss = Split(CATIA_Product_Open_Path, "\")
            COMP_Product_File_Name = ls_ss(UBound(ls_ss))
            '--------------------------------------------------------------------------------------------------- now_o_bom_excel_path
            Dim date_value_Len = Len(COMP_Product_File_Name)
            date_value = Mid(COMP_Product_File_Name, date_value_Len - 20, 10)

            fso = CreateObject("Scripting.FileSystemObject")
            Dim len_now_o_bom_excel_path = Len(A_Main_Form.OpenFileDialog1.FileName)
            Dim Exe_Extension = ".xlsx"

            'Dim fso2 As FileSystemObject
            '--------------------------------------------------------------------------------------------------- Folder Copy 중복 방지
            Dim now_o_bom_excel_path_value
            Try
                Exe_Extension = ".xls"
                If (fso.FileExists(Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & ".xls") = True) Then
                    now_o_bom_excel_path_value = fso.GetFile(Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & ".xls")
                ElseIf (fso.FileExists(Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & ".xlsx") = True) Then
                    now_o_bom_excel_path_value = fso.GetFile(Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & ".xlsx")
                    Exe_Extension = ".xlsx"
                Else
                    MessageBox.Show("진행중 Product에 상응하는 O-BOM 이 없습니다." & vbCrLf & Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & ".xlsx 파일을 확인해주세요.")
                    Return
                End If
            Catch ex As Exception ':  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                Try
                    Exe_Extension = ".xlsx"
                    now_o_bom_excel_path_value = fso.GetFile(Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & ".xlsx")
                Catch exx As Exception
                End Try
            End Try

            Try
                '--------------------------------------------------------------------------------------------------- Excel 파일 경로 가져오기
                ls_ss = Split(Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11), "\")

                Dim Get_Layout_Bom_Excel_File_Name = ls_ss(UBound(ls_ss))
                A_Main_Form.O_BOM_File_Name.Text = Get_Layout_Bom_Excel_File_Name & Exe_Extension
                '--------------------------------------------------------------------------------------------------- Sales_No 적용
                If Left(A_Main_Form.O_BOM_File_Name.Text, 1) = "C" Then
                    Sales_No = Left(A_Main_Form.O_BOM_File_Name.Text, 10)
                Else
                    Call SHOW_MESSAGE_BOX("O_BOM File Name의 Sales NO.를 확인하세요.", "", 1, 1)
                    now_o_bom_value = 0
                    Error_Count = 1
                    Exit Sub
                End If
                '--------------------------------------------------------------------------------------------------- 진행중 Product_open
                Dim O_Bom_Product_Open = CATIA.Documents.Open(CATIA_Product_Open_Path)
                now_o_bom_excel_path = Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - 11) & Exe_Extension
            Catch ex As System.IO.FileNotFoundException
                Call SHOW_MESSAGE_BOX("진행중 Product 파일과 일치하는 O-BOM 파일이 없습니다" & vbCrLf & "파일형식 : 세일즈오더_기종_년월일시분.xls" & vbCrLf & "          예) 3211000_SM5_1201011111.xls", "", 1, 1)
                Error_Count = 1
                Exit Sub
            Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                Call SHOW_MESSAGE_BOX("Get_Layout_List_File_Open_Function 함수 오류!!", "", 1, 1)
                Error_Count = 1
            End Try

            '--------------------------------------------------------------------------------------------------- 기종 정보 가져오기
            Dim assy_value_Name_Split = Split(A_Main_Form.O_BOM_File_Name.Text, "_")
            assy_value_Name = assy_value_Name_Split(1)
            '--------------------------------------------------------------------------------------------------- Lbl_ModelNumber에 기종정보 표시.
            A_Main_Form.Lbl_ModelNumber.Text = assy_value_Name
            '--------------------------------------------------------------------------------------------------- O_BOM_File_Name
            Dim assy_value_Name_Len = Len(assy_value_Name)
            Dim Sales_No_Name_Len = Len(Sales_No)
            A_Main_Form.re_o_bom_path.Text = Left(CATIA_Product_Open_Path, len_now_o_bom_excel_path - (24 + Sales_No_Name_Len + assy_value_Name_Len))
            '--------------------------------------------------------------------------------------------------- now_o_bom_excel_path 가져오기 (Template 작업시 사용)
            Exc = CreateObject("excel.application")
            Exc.Workbooks.Open(System.Windows.Forms.Application.StartupPath & "\data_path.xlsx")
            Exc.Worksheets("Sheet1").Range("A" & 10).Value = now_o_bom_excel_path
            Exc.Worksheets("Sheet1").Range("A" & 5).Value = A_Main_Form.O_BOM_File_Name.Text
            Exc.Worksheets("Sheet1").Range("C" & 3).Value = A_Main_Form.re_o_bom_path.Text
            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
            Exc.ActiveWorkbook.Save()
            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
            Call KillProcess("EXCEL")
            '--------------------------------------------------------------------------------------------------- ListBox 초기화
            A_Main_Form.data_open_path_list.Items.Clear()
            A_Main_Form.data_open_list.Items.Clear()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Get_Layout_List_File_Open_Function() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Public Function Over_Constraint_List_Function()
        Call CATIA_BASIC02()

        A_Main_Form.ProgressBar1.Value = 20
        '--------------------------------------------------------------------------------------------------- 새로운 Excel Open
        Exc = CreateObject("excel.application")
        Exc.Visible = True
        Exc.Workbooks.Add()
        '--------------------------------------------------------------------------------------------------- 새로운 Excel에 기본값 입력
        Exc.Worksheets("sheet1").Cells(1, 1).Value = "No"
        Exc.Worksheets("sheet1").Cells(1, 2).Value = "축번호_1"
        Exc.Worksheets("sheet1").Cells(1, 3).Value = "축번호_2"
        Exc.Worksheets("sheet1").Cells(1, 4).Value = "연결부품명_1"
        Exc.Worksheets("sheet1").Cells(1, 5).Value = "연결부품명_2"
        Exc.Worksheets("sheet1").Cells(1, 6).Value = "부품간거리"

        Dim constraint1 = Nothing
        Try
            constraint1 = CATIA.ActiveDocument.Product.Connections("CATIAConstraints")

            CATIA.ActiveDocument.Product.Update()
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call KillProcess("EXCEL")
            Call SHOW_MESSAGE_BOX("Over_Constraint_List_Function 함수 오류!!", "", 1, 1)
            Return False
        End Try

        Dim over_constraint_number As Integer = 1
        '--------------------------------------------------------------------------------------------------- constraint 수량 만큼 Loof
        For i = 1 To constraint1.Count
            '--------------------------------------------------------------------------------------------------- Over constraint이 있을때 Excel에 기입
            If Not constraint1.Item(i).Status = 0 Then
                For j = 1 To O_BOM_Constraint_Value_Count - 1
                    If constraint1.Item(i).Name = O_BOM_Constraint_Value(j, 2) Then
                        Exc.Worksheets("sheet1").Cells(over_constraint_number + 1, 1).Value = over_constraint_number
                        Exc.Worksheets("sheet1").Cells(over_constraint_number + 1, 2).Value = O_BOM_Constraint_Value(j, 1)
                        Exc.Worksheets("sheet1").Cells(over_constraint_number + 1, 3).Value = O_BOM_Constraint_Value(j, 2)
                        Exc.Worksheets("sheet1").Cells(over_constraint_number + 1, 4).Value = O_BOM_Constraint_Value(j, 3)
                        Exc.Worksheets("sheet1").Cells(over_constraint_number + 1, 5).Value = O_BOM_Constraint_Value(j, 4)

                        Call axis_distance(O_BOM_Constraint_Value(j, 1), O_BOM_Constraint_Value(j, 2))

                        Exc.Worksheets("sheet1").Cells(over_constraint_number + 1, 6).Value = axis_axis_measure
                        over_constraint_number = over_constraint_number + 1
                    End If
                Next
            End If
        Next

        A_Main_Form.ProgressBar1.Value = 80
        '--------------------------------------------------------------------------------------------------- Excel 정리
        Exc.Columns("A:F").Select()
        Exc.Columns("A:F").EntireColumn.AutoFit()
        A_Main_Form.ProgressBar1.Value = 90

        Return True
    End Function


    '---------------------------------------------------------------------------------------------------
    ' Clash_Check_Function  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Clash_Check_Function()
        A_Main_Form.ProgressBar1.Value = 0

        Call CATIA_BASIC02()
        If ActiveDocument Is Nothing Then
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- Clash Data 없을때..
        If CATIA.Documents.Count = 0 Then
            Call SHOW_MESSAGE_BOX("간섭 CHECK Data 가 없습니다.", "", 1, 1)
            A_Main_Form.ProgressBar1.Value = 100
            Exit Sub
        End If
        A_Main_Form.ProgressBar1.Value = 20

        '--------------------------------------------------------------------------------------------------- Clash 추가
        Dim cClashes = CATIA.ActiveDocument.Product.GetTechnologicalObject("Clashes")
        Dim oClash = cClashes.AddFromSel

        '--------------------------------------------------------------------------------------------------- Clash Type 추가
        oClash.ComputationType = 1 'catClashComputationTypeInsideOne
        oClash.InterferenceType = 0 'catClashInterferenceTypeContact
        '--------------------------------------------------------------------------------------------------- Clash 계산
        oClash.Compute()
        A_Main_Form.ProgressBar1.Value = 50
        '--------------------------------------------------------------------------------------------------- Clash Report
        If Not oClash.Conflicts.Count = 0 Then
            '--------------------------------------------------------------------------------------------------- Excel File Open
            Exc = CreateObject("excel.application")
            Exc.Visible = True
            Exc.Workbooks.Add()
            '--------------------------------------------------------------------------------------------------- Clash Excel 표기
            Exc.Worksheets("sheet1").Cells(1, 1).Value = "No"
            Exc.Worksheets("sheet1").Cells(1, 2).Value = "Product01"
            Exc.Worksheets("sheet1").Cells(1, 3).Value = "Product02"
            Exc.Worksheets("sheet1").Cells(1, 4).Value = "Type"
            Exc.Worksheets("sheet1").Cells(1, 5).Value = "Value"
            For i = 1 To oClash.Conflicts.Count
                '--------------------------------------------------------------------------------------------------- Clash Excel 표기
                Dim FirstProduct_PartNumber As String
                Dim SecondProduct_PartNumber As String
                FirstProduct_PartNumber = oClash.Conflicts.Item(i).FirstProduct.PartNumber
                SecondProduct_PartNumber = oClash.Conflicts.Item(i).SecondProduct.PartNumber
                Exc.Worksheets("sheet1").Cells(i + 1, 1).Value = i
                Exc.Worksheets("sheet1").Cells(i + 1, 2).Value = FirstProduct_PartNumber
                Exc.Worksheets("sheet1").Cells(i + 1, 3).Value = SecondProduct_PartNumber
                '--------------------------------------------------------------------------------------------------- Clash Type 표기
                If oClash.Conflicts.Item(i).Type = 0 Then
                    Exc.Worksheets("sheet1").Cells(i + 1, 4).Value = "Clash"
                ElseIf oClash.Conflicts.Item(i).Type = 1 Then
                    Exc.Worksheets("sheet1").Cells(i + 1, 4).Value = "Contact"
                End If
                Exc.Worksheets("sheet1").Cells(i + 1, 5).Value = oClash.Conflicts.Item(i).Value
            Next
            '--------------------------------------------------------------------------------------------------- Excel 정리
            Exc.Columns("A:E").Select()
            Exc.Columns("A:E").EntireColumn.AutoFit()
        End If
        A_Main_Form.ProgressBar1.Value = 90
    End Sub

End Module
