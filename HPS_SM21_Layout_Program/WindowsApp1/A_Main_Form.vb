Imports Scripting
Imports Microsoft.WindowsAPICodePack
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class A_Main_Form
    Public A_Main_Form As A_Main_Form

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ProgramStartDate = Date.Now     ' 프로그램 실행 이후 엑셀 종료를 위해..
        TabPage1.TabIndex = 0
        '--------------------------------------------------------------------------------------------------- Form 최상위
        DataGridView1.ReadOnly = False
        '--------------------------------------------------------------------------------------------------- A_Main_Form 화면 위치 조정
        'lWidth = Screen.Width - 13000
        'lHeight = 1080
        Me.Height = 750
        Me.Width = 805
        '--------------------------------------------------------------------------------------------------- CATIA_Basic
        Call CATIA_BASIC()
        '--------------------------------------------------------------------------------------------------- 프로그램 실행전 Open Excel 수량 Check
        Call Excle_Process_Count("EXCEL.EXE")
        '--------------------------------------------------------------------------------------------------- 마지막 구동 프로그램 Path 가져오기
        Call Find_data_Path()
        '--------------------------------------------------------------------------------------------------- coupling 유무 판단
        coupling_number = 0
        Message_Label.Text = "O-BOM File 가져오기 버튼을 클릭후 O-BOM Excel File 파일을 선택하세요."
        '--------------------------------------------------------------------------------------------------- 진행중 Product 일때 사용
        Get_Layout_List_Number = 0
        '--------------------------------------------------------------------------------------------------- Main Form 위치 초기화
        Call Form_Position_Initial(Me)
        '--------------------------------------------------------------------------------------------------- 기종정보 초기화
        assy_value_Name = "Not Value"
        '--------------------------------------------------------------------------------------------------- Manual 정보 가져오기
        Call Get_Excel_Value_Function(Application.StartupPath & "\STW_COMP_Infomation.xlsx", "Manual", 5, "A", "B", "C", "D", "E", "", "", "", "", "", Manual_Infomation_Count, Manual_Infomation, "0")
        '--------------------------------------------------------------------------------------------------- 메인 폼의 경로들 변수화
        sDirectory_Path_Text = Directory_Path_Text.Text
        sConstraint_Excel_Path_Text = Constraint_Excel_Path_Text.Text
        sO_BOM_Path_Text = O_BOM_Path_Text.Text
        sProject_DB_Path_Text = Project_DB_Path_Text.Text
        sTemplate_Data_Path_Text = Template_Data_Path_Text.Text

        fso = New FileSystemObject
        If (fso.FileExists(System.Windows.Forms.Application.StartupPath & "\Err.txt") = True) Then
            Dim file As System.IO.StreamWriter
            System.IO.File.WriteAllText(System.Windows.Forms.Application.StartupPath & "\Err.txt", "")
        End If

        If fso.FolderExists("C:\SM21_Exe_File") = False Then
            fso.CreateFolder("C:\SM21_Exe_File")
        End If
        If (fso.FileExists("C:\SM21_Exe_File\Publications.txt") = False) Then
            System.IO.File.Create("C:\SM21_Exe_File\Publications.txt").Dispose()
        End If
        If fso.FileExists("C:\SM21_Exe_File\Publications.txt") = True Then
            Dim file As System.IO.StreamWriter
            System.IO.File.WriteAllText("C:\SM21_Exe_File\Publications.txt", "")
        End If
    End Sub


    Private Sub GetFolderPath(Path_Text As TextBox)
        Me.TopMost = False
        Dim GetDirectory_value = GetDirectory("폴더를 선택하여 주세요.", Path_Text.Text)
        Me.TopMost = True
        If GetDirectory_value = "" Then Exit Sub
        Path_Text.Text = GetDirectory_value
    End Sub


    Private Sub DB_Path_Save_Click(sender As Object, e As EventArgs) Handles DB_Path_Save.Click
        Me.ProgressBar1.Value = 10
        Exc = CreateObject("excel.Application")
        Me.ProgressBar1.Value = 20
        '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
        Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")
        Me.ProgressBar1.Value = 30
        Exc.Worksheets("Sheet1").Range("A1").Value = Me.Directory_Path_Text.Text
        Exc.Worksheets("Sheet1").Range("A2").Value = Me.Constraint_Excel_Path_Text.Text
        Exc.Worksheets("Sheet1").Range("A3").Value = Me.O_BOM_Path_Text.Text
        Exc.Worksheets("Sheet1").Range("A4").Value = Me.Project_DB_Path_Text.Text
        Exc.Worksheets("Sheet1").Range("A16").Value = Me.Template_Data_Path_Text.Text
        Me.ProgressBar1.Value = 50

        Exc.Application.DisplayAlerts = False
        Exc.ActiveWorkbook.Save()
        Exc.ActiveWorkbook.Close(True)
        Exc.Application.DisplayAlerts = True
        Exc.Quit()

        Me.ProgressBar1.Value = 80

        Call KillProcess("EXCEL")

        Me.ProgressBar1.Value = 100
        Call SHOW_MESSAGE_BOX("경로 저장 완료", "", 1, 1)

        Me.Message_Label.Text = " 경로 저장 완료"
    End Sub


    Private Sub Directory_Path_Cmd_Click(sender As Object, e As EventArgs) Handles Directory_Path_Cmd.Click
        Call GetFolderPath(Directory_Path_Text)
    End Sub


    Private Sub O_BOM_Path_Cmd_Click(sender As Object, e As EventArgs) Handles O_BOM_Path_Cmd.Click
        Call GetFolderPath(O_BOM_Path_Text)
    End Sub


    Private Sub Template_Data_Path_Cmd_Click(sender As Object, e As EventArgs) Handles Template_Data_Path_Cmd.Click
        Call GetFolderPath(Template_Data_Path_Text)
    End Sub


    Private Sub Constraint_Excel_Path_Cmd_Click(sender As Object, e As EventArgs) Handles Constraint_Excel_Path_Cmd.Click
        Call GetFolderPath(Constraint_Excel_Path_Text)
    End Sub


    Private Sub Project_DB_Path_Cmd_Click(sender As Object, e As EventArgs) Handles Project_DB_Path_Cmd.Click
        Call GetFolderPath(Project_DB_Path_Text)
    End Sub


    Private Sub re_o_bom_path_TextChanged(sender As Object, e As EventArgs) Handles re_o_bom_path.TextChanged
        sOBomPath = re_o_bom_path.Text
    End Sub


    Private Sub O_BOM_File_Name_TextChanged(sender As Object, e As EventArgs) Handles O_BOM_File_Name.TextChanged
        sO_BOM_File_Name = O_BOM_File_Name.Text
    End Sub


    Public Sub Get_Create_Template_Button()
        For i = 1 To Result_Part_Create.Length - 1
            Result_Part_Create(i) = MainTab_GroupBox.Controls("Result_Part_Create" & i)
        Next

        For i = 1 To Template_Name_Label.Length - 1
            Template_Name_Label(i) = MainTab_GroupBox.Controls("Template_Name_Label" & i)
        Next

        For i = 1 To Template_Create.Length - 1
            Template_Create(i) = MainTab_GroupBox.Controls("Template_Create" & i)
        Next
    End Sub


    Private Sub O_BOM_File_Click(sender As Object, e As EventArgs) Handles O_BOM_File.Click
        Call Call_O_BOM_File_Click()
        Call KillProcess("EXCEL")

        sLbl_ModelNumber = Lbl_ModelNumber.Text
    End Sub


    Public Sub Call_O_BOM_File_Click()
        Call CATIA_BASIC02()
        '--------------------------------------------------------------------------------------------------- Now BOM_Excel_Open시 변수 사용
        now_o_bom_value = 0
        Error_Count = 0
        '--------------------------------------------------------------------------------------------------- ListBox 초기화
        Me.data_open_path_list.Items.Clear()
        Me.data_open_list.Items.Clear()
        '--------------------------------------------------------------------------------------------------- BOM_Excel_Open
        First_O_BOM = 1
        Diff_Path = 0
        Call Bom_File_Open_Function()
        '--------------------------------------------------------------------------------------------------- BOM List Cancel시 에러 처리
        If assy_value_Name = "" Then
            Call SHOW_MESSAGE_BOX("O_BOM File을 선택하세요.", "", 1, 1)
            now_o_bom_value = 0
            Exit Sub
        End If
        If Error_Count = 1 Then
            Exit Sub
        End If
        If Diff_Path = 1 Then
            Exit Sub
        End If

        If Get_Excel_Name_Check(Sales_No & "_" & assy_value_Name) = False Then
            Call SHOW_MESSAGE_BOX(Sales_No & "_" & assy_value_Name & "에 한글이 포함되어 O_BOM File 불러오기가 종료됩니다.", "", 1, 1)
            Exit Sub
        End If

        Me.DataGridView1.Rows.Clear()
        Me.DataGridView1.RowHeadersVisible = False
        Me.ProgressBar1.Value = 0

        '--------------------------------------------------------------------------------------------------- COMP 정보 가져오기
        If Get_Excel_Value_Function(Application.StartupPath & "\STW_COMP_Infomation.xlsx", assy_value_Name, 6, "A", "B", "C", "D", "E", "F", "", "", "", "",
                                    STW_COMP_Infomation_Count, STW_COMP_Infomation, "0") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Application.StartupPath & "\STW_COMP_Infomation.xlsx", 1, 1)
        End If
        '--------------------------------------------------------------------------------------------------- Template 관련 정보 가져오기
        If Get_Excel_Value_Function(Application.StartupPath & "\STW_Template_Infomation.xlsx", assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "",
                                    STW_Template_Infomation_Count, STW_Template_Infomation, "0") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Application.StartupPath & "\STW_Template_Infomation.xlsx", 1, 1)
        End If
        '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기
        ' 생 성 일  : -
        ' 생 성 자  : -
        ' 수 정 일  : 1. 20.02.13
        ' 수 정 자  : 허혜원
        ' 수정사유  : 1. BOM_LIST 시트 J열(개수) 가져오기 추가
        ' Parameter : -
        '---------------------------------------------------------------------------------------------------
        If Get_Excel_Value_Function(sfile, "BOM_LIST", 4, "B", "H", "K", "J", "", "", "", "", "", "", O_BOM_Part_Value_Count, O_BOM_Part_Value, "1") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", sfile, 1, 1)
        End If
        Me.Message_Label.Text = "BOM Excel 파일 Loading 완료."

        '--------------------------------------------------------------------------------------------------- BOM_Excel 파일 경로 입력 하기
        Call O_BOM_Excel_Path()
        Me.ProgressBar1.Value = 10
        '--------------------------------------------------------------------------------------------------- Constraint 정보 가져오기
        If Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), STW_COMP_Infomation(2, 1), 4, "A", "B", "C", "D", "", "", "", "", "", "",
                                    O_BOM_Constraint_Value_Count, O_BOM_Constraint_Value, "2") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), 1, 1)
        End If
        Me.Message_Label.Text = "Constraint Excel 파일 Loading 완료."
        Me.ProgressBar1.Value = 20
        '--------------------------------------------------------------------------------------------------- LAYOUT 대상 부품 구분 정보 가져오기
        If Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 3), STW_COMP_Infomation(2, 3), 3, "B", "C", "D", "", "", "", "", "", "", "",
                                    O_BOM_Layout_Part_Value_Count, O_BOM_Layout_Part_Value, "3") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 3), 1, 1)
        End If
        Me.Message_Label.Text = "LAYOUT 대상 부품 구분 Excel 파일 Loading 완료."
        '--------------------------------------------------------------------------------------------------- O_BOM Standard 정보 가져오기(PYO std필요없음)        
        '--------------------------------------------------------------------------------------------------- New_Folder_Create
        Call New_Folder_Create()
        '--------------------------------------------------------------------------------------------------- Data Loading
        Tmp_Duplication_Check_Dic.Clear()
        Call Data_Loading_Click()
        '--------------------------------------------------------------------------------------------------- EPN No , BOM List 비교 구문
        Call Comparison_Data()
        ''''--------------------------------------------------------------------------------------------------- Axis Setting
        Call Axis_Setting()
        ''''--------------------------------------------------------------------------------------------------- 구속조건 생성
        Call Constraint_Setting()
        ''''--------------------------------------------------------------------------------------------------- 중복 구속조건 제거
        Call Error_Constraint_Delete()
        ''''--------------------------------------------------------------------------------------------------- Product name 변경 및 저장
        If Comp_Product_Name_Change() = False Then
            Exit Sub
        End If
        ''''--------------------------------------------------------------------------------------------------- Result_Data_Replace 버튼 활성화
        Result_Data_Replace.Enabled = True
        Call KillProcess("EXCEL")
        Call ShowDuplication_Check_Dic()
    End Sub

    Public Function ShowDuplication_Check_Dic()
        Dim DupEPNS = ""
        If Tmp_Duplication_Check_Dic.Count <> 0 Then
            For Each EPNName In Tmp_Duplication_Check_Dic
                DupEPNS = DupEPNS & EPNName.Value & vbCrLf
            Next
            MsgBox(DupEPNS & "축 이름이 중복됩니다. 확인 후 수정해주세요.", vbSystemModal)
        End If
        Tmp_Duplication_Check_Dic.Clear()
    End Function

    Private Sub BtEx_Click(sender As Object, e As EventArgs) Handles BtEx.Click
        If SHOW_MESSAGE_BOX("프로그램을 종료 하시겠습니까?", "", 2, 1) = vbYes Then
            Call KillProcess("EXCEL")
            Me.Close()
        Else
            Me.TopMost = True
        End If
    End Sub


    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        '--------------------------------------------------------------------------------------------------- Missing 'V' 셀이 아닐 경우 종료
        If Not DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value.ToString() = "V" Then
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- 
        Dim EPN_Name = Me.DataGridView1.Item(1, Me.DataGridView1.CurrentRow.Index).Value

        '--------------------------------------------------------------------------------------------------- 3D 형상위치\기종폴더\Missing EPN 폴더가 없을 경우 생성
        Call ShowOpenFileDialog("Open Missing File", "Part File(*.CATPart)|*.CATPart", sDirectory_Path_Text & "\" & Lbl_ModelNumber.Text & "\" & EPN_Name)
        If Me.OpenFileDialog1.FileName = "" Then
            DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value = "V"
            Exit Sub
        End If

        Dim CATIA_Product_Open_Path = Me.OpenFileDialog1.FileName

        Me.Message_Label.Text = " Data Loading중 입니다."
        Me.ProgressBar1.Value = 10

        Try
            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call Active_Window_Change(COMP_Product_File_Name, 0, 0)

            Dim products1 = ActiveDocument.Product.Products

            '--------------------------------------------------------------------------------------------------- InstanceName 중복 오류 방지를 위해 Product상에 EPN 이미 존재하는지 확인
            For i = 1 To products1.Count
                If products1.Item(i).Name = EPN_Name Then
                    Call SHOW_MESSAGE_BOX("EPN (" & EPN_Name & ")이 기존 Product에 이미 존재합니다." & vbCrLf & "기존 데이터 제거 후 작업해주세요.",
                                          "EPN 중복", 1, 1)
                    Me.ProgressBar1.Value = 0
                    DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value = "V"
                    Exit Sub
                End If
            Next

            '--------------------------------------------------------------------------------------------------- Instance Name 변경
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = CATIA_Product_Open_Path
            products1.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            Dim Part = products1.Item(products1.Count)
            Dim PartNumber = Part.PartNumber

            Dim fileName
            'Me.OpenFileDialog1.FileName
            Dim fileNames
            fileNames = Split(CATIA_Product_Open_Path, "\")
            fileName = Replace(fileNames(UBound(fileNames)), ".CATPart", "")
            Part.Name = EPN_Name
            Part.PartNumber = fileName

            Me.ProgressBar1.Value = 20
            '--------------------------------------------------------------------------------------------------- 엑셀경로 가져오기
            Dim Excel_Path
            If o_bom_excel_file_name <> "" Then
                Excel_Path = Project_DB_Path_Text.Text & "\" & o_bom_excel_file_name & "\" & O_BOM_File_Name.Text
            Else
                Excel_Path = now_o_bom_excel_path
            End If
            '엑셀의 워크북이 이미 실행되어 있는 경우 예외처리 하기

            ' "BOM 엑셀 파일 변경(Missing 되어 신규로 추가하는 EPN의 PartNumber) 후 저장 종료"
            Exc = CreateObject("excel.application")
            Dim WkBook = Exc.Workbooks.Open(Excel_Path)
            'Exc.Visible = True

            Dim FoundCell = WkBook.Sheets("BOM_LIST").Range("K:K").Find(What:=EPN_Name, LookAt:=1)
            If Not (FoundCell Is Nothing) Then




                If InStr(PartNumber, "_") = 0 Then
                    WkBook.Sheets("BOM_LIST").Cells(FoundCell.Row, 2).Value = PartNumber

                    Me.DataGridView1.Item(3, Me.DataGridView1.CurrentRow.Index).Value = PartNumber
                Else
                    WkBook.Sheets("BOM_LIST").Cells(FoundCell.Row, 2).Value = Strings.Left(PartNumber, InStr(PartNumber, "_") - 1)
                    Me.DataGridView1.Item(3, Me.DataGridView1.CurrentRow.Index).Value = PartNumber
                End If
                Exc.Application.DisplayAlerts = False
                Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
                Exc.ActiveWorkbook.Save()
                Exc.ActiveWorkbook.Close(True)
                Call KillProcess("EXCEL")
                '--------------------------------------------------------------------------------------------------- Missing "V" 체크 없애기
                Me.DataGridView1.Item(2, Me.DataGridView1.CurrentRow.Index).Value = ""

                ActiveDocument.Save()
            Else
                Call SHOW_MESSAGE_BOX(WkBook.Name & " 파일에 " & EPN_Name & " 이 존재하지 않습니다." & vbCrLf & "파일을 다시 확인해주세요.", "", 1, 1)
                Me.ProgressBar1.Value = 0
                DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value = "V"
                Exit Sub
            End If

            ' "Dictionary PublicationDic에 CATIA ActiveDocument에 존재하는 기존 Publication, 현재 추가할 Publication 이름 추가"
            '/// 특정 Part의 Publication을 Dictionary에 추가
            Dim PublicationDic = Nothing
            Call GetOnePartPublication(Part, PublicationDic, True)

            Me.ProgressBar1.Value = 30
            product_add = CATIA.ActiveDocument.Product
            products = product_add.products
            selection = CATIA.ActiveDocument.Selection
            constraints1 = product_add.Connections("CATIAConstraints")
            '/// Product상의 모든 Part의 Publication을 배열과 Dictionary에 추가
            'Call GetProductAllPartPublication(False)
            Call GetProductAllPartPublication(True)
            Me.ProgressBar1.Value = 60

            For i = 1 To O_BOM_Constraint_Value_Count - 1
                '--------------------------------------------------------------------------------------------------- CON_TMP_01 이면...
                ' "EPN_Name Constraint가 Publication이 있는지 체크하고 없을 경우 Continue"
                If (Not PublicationDic.Exists(O_BOM_Constraint_Value(i, 1))) And (Not (PublicationDic.Exists(O_BOM_Constraint_Value(i, 2)))) Then
                    Continue For
                Else
                    '/// Debug.Print("Missing Constraint : " & O_BOM_Constraint_Value(i, 1) & " / " & O_BOM_Constraint_Value(i, 2))
                End If
                '/// Constraint1 또는 Constraint2 값이 현재 Product 내의 publication List에 없을 경우 Continue For
                If Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 1)) Or Not get_part_publication2017.ContainsKey(O_BOM_Constraint_Value(i, 2)) Then
                    Continue For
                End If
                '/// Debug.Print("publication 다 존재함 : " & O_BOM_Constraint_Value(i, 1) & ", " & O_BOM_Constraint_Value(i, 2))

                '--------------------------------------------------------------------------------------------------- 1번째 Publication 찾기
                Call GetPub_DisplayName(Pub1_DisplayName, O_BOM_Constraint_Value(i, 1))
                Call GetPub_reference(Pub_reference1, Pub1_DisplayName, O_BOM_Constraint_Value(i, 1))

                '--------------------------------------------------------------------------------------------------- 2번째 Publication 찾기
                Call GetPub_DisplayName(Pub2_DisplayName, O_BOM_Constraint_Value(i, 2))
                Call GetPub_reference(Pub_reference2, Pub2_DisplayName, O_BOM_Constraint_Value(i, 2))

                '--------------------------------------------------------------------------------------------------- 구속 생성
                coincidence_constraint(i) = constraints1.AddBiEltCst(2, Pub_reference1, Pub_reference2)
                coincidence_constraint(i).Name = O_BOM_Constraint_Value(i, 2)
                '--------------------------------------------------------------------------------------------------- 단계별 보기 Check
                'VB6.0->2017 업데이트 + R26부터 발생하는 Over-Constrain 알림창 제거.                
                Call Constraint_Update_ExceptionHandling(product_add, constraints1)
                selection.Clear()
                For j = 1 To constraints1.Count
                    '--------------------------------------------------------------------------------------------------- CON_TMP이면 Constraint 삭제.
                    If Strings.Left(constraints1.Item(j).Name, 7) = "CON_TMP" Then
                        Threading.Thread.Sleep(200)
                        selection.Add(constraints1.Item(j))
                    End If
                Next j
                If selection.Count <> 0 Then
                    selection.Delete()
                End If
            Next i
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX(ex.ToString & " / " & ex.Source & " / " & ex.Message, "", 1, 1)
        End Try
        Me.ProgressBar1.Value = 80

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        CATIA.DisplayFileAlerts = False
        product_add.Update()
        CATIA.DisplayFileAlerts = True
        Me.ProgressBar1.Value = 100

        Call SHOW_MESSAGE_BOX("Missing 파트 불러오기 완료", "Missing 파트", 1, 1)
        Me.Message_Label.Text = " Missing 파트 불러오기 완료"
    End Sub


    Private Sub Get_Now_O_BOM_File_Click(sender As Object, e As EventArgs) Handles Get_Now_O_BOM_File.Click
        Call CATIA_BASIC02()
        now_o_bom_value = 1
        First_O_BOM = 0
        Diff_Path = 0
        '--------------------------------------------------------------------------------------------------- Now BOM_Excel_Open시 변수 사용        
        Error_Count = 0
        Call Bom_File_Open_Function()
        If Error_Count = 1 Then
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------- BOM List Cancel시 에러 처리
        If assy_value_Name = "" Then
            Call SHOW_MESSAGE_BOX("O_BOM File을 선택하세요.", "", 1, 1)
            now_o_bom_value = 0
            Exit Sub
        End If
        Me.DataGridView1.Rows.Clear()
        Me.ProgressBar1.Value = 0
        '--------------------------------------------------------------------------------------------------- ListBox 초기화
        Me.data_open_path_list.Items.Clear()
        Me.data_open_list.Items.Clear()
        '--------------------------------------------------------------------------------------------------- BOM_Excel_Open        
        '--------------------------------------------------------------------------------------------------- COMP 정보 가져오기
        If Get_Excel_Value_Function(Application.StartupPath & "\STW_COMP_Infomation.xlsx", assy_value_Name, 6, "A", "B", "C", "D", "E", "F", "", "", "", "", STW_COMP_Infomation_Count, STW_COMP_Infomation, "0") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File을 선택하세요.", Application.StartupPath & "\STW_COMP_Infomation.xlsx", 1, 1)
        End If
        '--------------------------------------------------------------------------------------------------- Template 관련 정보 가져오기
        If Get_Excel_Value_Function(Application.StartupPath & "\STW_Template_Infomation.xlsx", assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", STW_Template_Infomation_Count, STW_Template_Infomation, "0") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Application.StartupPath & "\STW_Template_Infomation.xlsx", 1, 1)
        End If
        '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기
        If Get_Excel_Value_Function(sfile, "BOM_LIST", 4, "B", "H", "K", "J", "", "", "", "", "", "", O_BOM_Part_Value_Count, O_BOM_Part_Value, "1") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", sfile, 1, 1)
        End If
        Me.Message_Label.Text = "BOM Excel 파일 Loading 완료."
        '--------------------------------------------------------------------------------------------------- BOM_Excel 파일 경로 입력 하기
        Call O_BOM_Excel_Path()
        Me.ProgressBar1.Value = 10
        '--------------------------------------------------------------------------------------------------- Constraint 정보 가져오기
        If Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), STW_COMP_Infomation(2, 1), 4, "A", "B", "C", "D", "", "", "", "", "", "", O_BOM_Constraint_Value_Count, O_BOM_Constraint_Value, "2") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), 1, 1)
        End If
        Me.Message_Label.Text = "Constraint Excel 파일 Loading 완료."
        Me.ProgressBar1.Value = 20
        '--------------------------------------------------------------------------------------------------- LAYOUT 대상 부품 구분 정보 가져오기
        If Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 3), STW_COMP_Infomation(2, 3), 3, "B", "C", "D", "", "", "", "", "", "", "", O_BOM_Layout_Part_Value_Count, O_BOM_Layout_Part_Value, "3") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 3), 1, 1)
        End If
        Me.Message_Label.Text = "LAYOUT 대상 부품 구분 Excel 파일 Loading 완료."
        '--------------------------------------------------------------------------------------------------- O_BOM Standard 정보 가져오기(PYO std필요없음)        
        '--------------------------------------------------------------------------------------------------- New_Folder_Create
        Call New_Folder_Create()
        '--------------------------------------------------------------------------------------------------- Data Loading
        Tmp_Duplication_Check_Dic.Clear()
        Call Data_Loading_Click()
        '--------------------------------------------------------------------------------------------------- EPN No , BOM List 비교 구문
        Call Comparison_Data()
        ''''--------------------------------------------------------------------------------------------------- Axis Setting
        Call Axis_Setting()
        ''''--------------------------------------------------------------------------------------------------- 구속조건 생성
        Call Constraint_Setting()
        ''''--------------------------------------------------------------------------------------------------- 중복 구속조건 제거
        Call Error_Constraint_Delete()
        ''''--------------------------------------------------------------------------------------------------- Product name 변경 및 저장
        Call Comp_Product_Name_Change()
        ''''--------------------------------------------------------------------------------------------------- Result_Data_Replace 버튼 활성화
        Result_Data_Replace.Enabled = True
        Call KillProcess("EXCEL")
        sLbl_ModelNumber = Lbl_ModelNumber.Text
        Call ShowDuplication_Check_Dic()
    End Sub


    Private Sub Get_Layout_List_Click(sender As Object, e As EventArgs) Handles Get_Layout_List.Click
        Call CATIA_BASIC02()      '170321 R26
        '--------------------------------------------------------------------------------------------------- 진행중 Product File Open
        Call Get_Layout_List_File_Open_Function()
        If Error_Count = 1 Then
            Exit Sub
        End If
        If Me.OpenFileDialog1.FileName = "" Then
            assy_value_Name = ""
            Get_Layout_List_Number = 0
            Call SHOW_MESSAGE_BOX("진행중 Product를 선택하세요.", "", 1, 1)
            Exit Sub
        End If
        Me.DataGridView1.Rows.Clear()
        Me.ProgressBar1.Value = 0
        '--------------------------------------------------------------------------------------------------- Now BOM_Excel_Open시 변수 사용
        Error_Count = 0
        Me.ProgressBar1.Value = 10
        '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기
        If Get_Excel_Value_Function(now_o_bom_excel_path, "BOM_LIST", 4, "B", "H", "K", "J", "", "", "", "", "", "", O_BOM_Part_Value_Count, O_BOM_Part_Value, "1") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", now_o_bom_excel_path, 1, 1)
        End If
        Me.Message_Label.Text = "BOM Excel 파일 Loading 완료."
        '--------------------------------------------------------------------------------------------------- COMP 정보 가져오기
        If Get_Excel_Value_Function(Application.StartupPath & "\STW_COMP_Infomation.xlsx", assy_value_Name, 6, "A", "B", "C", "D", "E", "F", "", "", "", "", STW_COMP_Infomation_Count, STW_COMP_Infomation, "0") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Application.StartupPath & "\STW_COMP_Infomation.xlsx", 1, 1)
        End If
        '--------------------------------------------------------------------------------------------------- Template 관련 정보 가져오기
        If Get_Excel_Value_Function(Application.StartupPath & "\STW_Template_Infomation.xlsx", assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", STW_Template_Infomation_Count, STW_Template_Infomation, "0") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Application.StartupPath & "\STW_Template_Infomation.xlsx", 1, 1)
        End If
        '--------------------------------------------------------------------------------------------------- BOM_Excel 파일 경로 입력 하기
        Call O_BOM_Excel_Path()
        Me.ProgressBar1.Value = 20
        '--------------------------------------------------------------------------------------------------- Constraint 정보 가져오기
        If Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), STW_COMP_Infomation(2, 1), 4, "A", "B", "C", "D", "", "", "", "", "", "", O_BOM_Constraint_Value_Count, O_BOM_Constraint_Value, "2") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), 1, 1)
        End If
        Me.Message_Label.Text = "Constraint Excel 파일 Loading 완료."
        Me.ProgressBar1.Value = 30
        '--------------------------------------------------------------------------------------------------- LAYOUT 대상 부품 구분 정보 가져오기
        If Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 3), STW_COMP_Infomation(2, 3), 3, "B", "C", "D", "", "", "", "", "", "", "", O_BOM_Layout_Part_Value_Count, O_BOM_Layout_Part_Value, "3") = False Then
            Call SHOW_MESSAGE_BOX("O_BOM File 불러오기가 종료됩니다.", Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 3), 1, 1)
        End If
        Me.Message_Label.Text = "LAYOUT 대상 부품 구분 Excel 파일 Loading 완료."
        '--------------------------------------------------------------------------------------------------- Data Loading
        Tmp_Duplication_Check_Dic.Clear()
        Call Data_Loading_Click()
        '--------------------------------------------------------------------------------------------------- EPN No , BOM List 비교 구문
        Call Comparison_Data()
        ''''--------------------------------------------------------------------------------------------------- 중복 구속조건 제거
        Call Error_Constraint_Delete()
        Me.Radio1.Enabled = True
        Me.Radio2.Enabled = True
        '--------------------------------------------------------------------------------------------------- Result_Data_Replace 버튼 활성화
        Result_Data_Replace.Enabled = True
        Me.Message_Label.Text = "진행중 Product 구성 완료."
        Me.ProgressBar1.Value = 100
        Call KillProcess("EXCEL")
        sLbl_ModelNumber = Lbl_ModelNumber.Text
        Call ShowDuplication_Check_Dic()
    End Sub


    Private Sub Result_Data_Replace_Click(sender As Object, e As EventArgs) Handles Result_Data_Replace.Click
        AB_Part_Replace.ShowDialog()
    End Sub


    Private Sub Over_Constraint_List_Click(sender As Object, e As EventArgs) Handles Over_Constraint_List.Click
        Me.ProgressBar1.Value = 0
        Message_Label.Text = "Over Constraint List 생성 중..."

        If Over_Constraint_List_Function() = False Then
            Me.ProgressBar1.Value = 0
            Message_Label.Text = "Over Constraint List 생성 작업 취소..."
            Exit Sub
        End If

        'Exc.Visible = True

        Me.ProgressBar1.Value = 100
        Message_Label.Text = "Over Constraint List 생성 완료."
    End Sub


    Private Sub Element_ALL_Hide_Click(sender As Object, e As EventArgs) Handles Element_ALL_Hide.Click
        Try
            Call CATIA_BASIC02()
            selection.Clear()
            '--------------------------------------------------------------------------------------------------- Axis Search
            selection.search("(((((((((((((((((((((((((((((((((((FreeStyle.'Axis System' + 'Part Design'.'Axis System') + 'Generative Shape Design'.'Axis System') + 'Functional Molded Part'.'Axis System') + FreeStyle.Curve) + '2D Layout for 3D Design'.Curve) + Sketcher.Curve) + Drafting.Curve) + 'Part Design'.Curve) + 'Generative Shape Design'.Curve) + 'Functional Molded Part'.Curve) + FreeStyle.Line) + '2D Layout for 3D Design'.Line) + Sketcher.Line) + Drafting.Line) + 'Part Design'.Line) + 'Generative Shape Design'.Line) + 'Functional Molded Part'.Line) + FreeStyle.Point) + '2D Layout for 3D Design'.Point) + Sketcher.Point) + Drafting.Point) + 'Part Design'.Point) + 'Generative Shape Design'.Point) + 'Functional Molded Part'.Point) + 'Part Design'.Sketch) + 'Generative Shape Design'.Sketch) + 'Functional Molded Part'.Sketch) + FreeStyle.Plane) + 'Part Design'.Plane) + 'Generative Shape Design'.Plane) + 'Functional Molded Part'.Plane) + FreeStyle.Surface) + 'Part Design'.Surface) + 'Generative Shape Design'.Surface) + 'Functional Molded Part'.Surface),scr")

            '--------------------------------------------------------------------------------------------------- Axis Hide
            selection.VisProperties.Parent.SetShow(1)
            selection.Clear()

            Message_Label.Text = "Element ALL Hide 완료."
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Message_Label.Text = "Element ALL Hide 작업 오류."
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Constraint 재구성  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Private Sub retry_constraint_Click(sender As Object, e As EventArgs) Handles retry_constraint.Click
        Try
            Me.ProgressBar1.Value = 0
            '--------------------------------------------------------------------------------------------------- Data 없을때..
            If CATIA.Documents.Count = 0 Then
                Call SHOW_MESSAGE_BOX("Model정보가 없습니다.", "", 1, 1)
                Exit Sub
            End If

            Message_Label.Text = " Data Position Setting하고 있습니다..."
            '--------------------------------------------------------------------------------------------------- 기종 정보 가져오기
            Dim assy_value_Name_Split
            assy_value_Name_Split = Split(Me.O_BOM_File_Name.Text, "_")
            assy_value_Name = assy_value_Name_Split(1)

            '--------------------------------------------------------------------------------------------------- constraint delete
            selection = CATIA.ActiveDocument.selection
            selection.Clear()

            Dim constraint1 = CATIA.ActiveDocument.Product.Connections("CATIAConstraints")
            If constraint1.Count > 0 Then
                For i = 1 To constraint1.Count
                    selection.Add(constraint1.Item(i))
                    Threading.Thread.Sleep(200)
                Next

                If Not selection.Count = 0 Then
                    selection.Delete()
                    Threading.Thread.Sleep(300)
                End If
            End If
            Me.ProgressBar1.Value = 10

            '--------------------------------------------------------------------------------------------------- Constraint_Excel_Open
            Call Get_Excel_Value_Function(Me.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 1), STW_COMP_Infomation(2, 1), 4,
                                          "A", "B", "C", "D", "", "", "", "", "", "", O_BOM_Constraint_Value_Count, O_BOM_Constraint_Value, "2")

            '---------------------------------------------------------------------------------------------------
            '   기    능 : publication 저장된 배열 get_part_publication 초기화한 후 재입력
            '---------------------------------------------------------------------------------------------------
            get_part_publication2017.Clear()
            pub_number = 1
            For i = 1 To 1000
                get_part_publication(i) = Nothing
            Next
            '---------------------------------------------------------------------------------------------------
            '   기    능 : products 가져오기 추가  19.08.29
            '---------------------------------------------------------------------------------------------------
            Call CATIA_BASIC02()
            For i = 1 To products.Count
                publications1 = products.Item(i).publications

                For k = 1 To publications1.Count
                    get_part_publication(pub_number) = publications1.Item(k)

                    If Not get_part_publication2017.ContainsKey(publications1.Item(k).Name.ToString()) Then
                        get_part_publication2017.Add(publications1.Item(k).Name.ToString(), publications1.Item(k))
                    End If

                    pub_number = pub_number + 1
                Next
            Next

            Call Axis_Setting()
            Call Constraint_Setting()
            '---------------------------------------------------------------------------------------------------
            ' 생 성 일  : -
            ' 생 성 자  : -
            ' 수 정 일  : 1. 19.08.22
            '             2. 19.08.28
            ' 수 정 자  : 허혜원
            ' 수정사유  : 1. Constraint 재구성 실행 후 Constraint가 생성되지 않았을 경우 한번 더 실행.(반복문으로 실행시 무한루프 돌 가능성이 있어 1회만 실행)
            '             2. 5번 반복으로 변경
            ' Parameter : -
            '---------------------------------------------------------------------------------------------------
            Dim Check
            Check = 0
            Do While (constraints1.Count <= 1)
                Call Constraint_Setting()
                If Check >= 5 Then
                    Exit Do
                End If
                Check = Check + 1
            Loop
            'If constraints1.Count <= 1 Then
            '    Call Constraint_Setting()
            'End If

            Me.Message_Label.Text = "Data Position Setting 완료."
            Me.ProgressBar1.Value = 100
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Constraint 재구성() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Clash_Check_Click(sender As Object, e As EventArgs) Handles Clash_Check.Click
        Me.ProgressBar1.Value = 0
        Me.Message_Label.Text = "간섭 체크 중..."

        Call Clash_Check_Function()

        Me.ProgressBar1.Value = 100
        Me.Message_Label.Text = "간섭 체크 완료."
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' 실적 선택     &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Private Sub Result_Part_Create1_Click(sender As Object, e As EventArgs) Handles Result_Part_Create1.Click, Result_Part_Create2.Click, Result_Part_Create3.Click, Result_Part_Create4.Click, Result_Part_Create5.Click, Result_Part_Create6.Click, Result_Part_Create7.Click, Result_Part_Create8.Click, Result_Part_Create9.Click
        Form_Click_Index_Value = CInt(Strings.Right(sender.Name, 1))

        Template_Data_Loading_Form.Show()
        Template_Data_Loading_Form.Label1.Text = "실적 가져오기 Form Loading 중... 잠시만 기다려 주세요..."

        Call Form_Loading_Function(STW_Template_Infomation(Form_Click_Index_Value, 8), "Result_Data")

        Template_Data_Loading_Form.Close()  '종료
    End Sub


    Private Sub Template_Create1_Click(sender As Object, e As EventArgs) Handles Template_Create1.Click, Template_Create2.Click, Template_Create3.Click, Template_Create4.Click, Template_Create5.Click, Template_Create6.Click, Template_Create7.Click, Template_Create8.Click, Template_Create9.Click
        Form_Click_Index_Value = CInt(Strings.Right(sender.Name, 1))

        '--------------------------------------------------------------------------------------------------- COUPLING COVER 일때...
        Dim Coupling_Count_Value = Nothing
        If STW_Template_Infomation(Form_Click_Index_Value, 1) = "COUPLING COVER" Then
            Call CATIA_BASIC02()
            For i = 1 To ActiveDocument.Product.products.Count
                If ActiveDocument.Product.products.Item(i).Name = "EPN03" Then
                    Coupling_Count_Value = 1
                    Exit For
                End If
            Next

            If Coupling_Count_Value = 0 Then
                If SHOW_MESSAGE_BOX("Coupling이 없습니다. 계속 진행할까요?", "", 2, 1) = vbNo Then
                    Exit Sub
                End If
            End If
        End If

        Template_Data_Loading_Form.Show()

        Call Form_Loading_Function(STW_Template_Infomation(Form_Click_Index_Value, 9), STW_Template_Infomation(Form_Click_Index_Value, 1))

        Template_Data_Loading_Form.Close()
    End Sub


    Private Sub MainTab_Click(sender As Object, e As EventArgs) Handles MainTab.Click
        If MainTab.SelectedTab.TabIndex = 1 And Click_Count = 0 Then
            Call Get_Create_Template_Button()
            Call Template_Data_Form_Inicial_Setting()
        ElseIf MainTab.SelectedTab.TabIndex = 2 Then
            Me.project_data_path.Text = Me.re_o_bom_path.Text
            Drawing_RepresentationMode_Value = 2
        End If
    End Sub


    Private Sub Project_Data_Path_Open_Click(sender As Object, e As EventArgs) Handles Project_Data_Path_Open.Click
        '--------------------------------------------------------------------------------------------------- 폴더 선택
        Me.TopMost = False
        Dim GetDirectory_value = GetDirectory("폴더를 선택하여 주세요.", project_data_path.Text)
        Me.TopMost = True

        project_data_path.Text = GetDirectory_value

        '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
        Exc = CreateObject("excel.application")
        Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")
        Exc.Worksheets("Sheet1").Range("A" & 9).Value = project_data_path.Text
        Exc.ActiveWorkbook.Save()
        Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()

        Call KillProcess("EXCEL")
    End Sub


    Private Sub Comp_Layout_Save_Click(sender As Object, e As EventArgs) Handles Comp_Layout_Save.Click
        Try
            Call CATIA_BASIC02()

            '--------------------------------------------------------------------------------------------------- LAYOUT_Folder 초기화
            Dim LAYOUT_Folder = Nothing
            '--------------------------------------------------------------------------------------------------- LAYOUT Folder 가져오기
            fso = CreateObject("Scripting.FileSystemObject")
            '--------------------------------------------------------------------------------------------------- 현재 시간
            date_value = Format(Now, "yymmddhhmm")
            '--------------------------------------------------------------------------------------------------- LAYOUT Folder가 없을때...
            If System.IO.Directory.Exists(project_data_path.Text & "\LAYOUT") = False Then
                fso.CreateFolder(project_data_path.Text & "\LAYOUT")
                LAYOUT_Folder = project_data_path.Text & "\LAYOUT\"
            Else
                '--------------------------------------------------------------------------------------------------- LAYOUT Folder가 있을때...
                LAYOUT_Folder = fso.GetFolder(project_data_path.Text & "\LAYOUT")
                LAYOUT_Folder.Name = LAYOUT_Folder.Name & "_" & date_value

                fso.CreateFolder(project_data_path.Text & "\LAYOUT")
            End If

            '--------------------------------------------------------------------------------------------------- 최상위 해제해야 Save Management 이 바로 보임
            CATIA.StartCommand("Save Management")
            Me.TopMost = False
            Base_Drawing.Enabled = True
            Ref_Drawing.Enabled = True
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Comp_Layout_Save_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Base_Drawing_Click(sender As Object, e As EventArgs) Handles Base_Drawing.Click
        Try
            If SHOW_MESSAGE_BOX("실행 하시겠습니까?", "", 2, 1) = vbYes Then
                Call CATIA_BASIC02()
                Me.ProgressBar1.Value = 0

                '--------------------------------------------------------------------------------------------------- LAYOUT_Folder 초기화
                Dim LAYOUT_Folder = Nothing
                '--------------------------------------------------------------------------------------------------- LAYOUT Folder 가져오기
                fso = CreateObject("Scripting.FileSystemObject")
                LAYOUT_Folder = fso.GetFolder(project_data_path.Text & "\LAYOUT")
                '--------------------------------------------------------------------------------------------------- 현재 시간
                date_value = Format(Now, "yymmddhhmm")
                '--------------------------------------------------------------------------------------------------- LAYOUT Folder가 없을때...
                If LAYOUT_Folder Is Nothing Then
                    fso.CreateFolder(project_data_path.Text & "\LAYOUT")
                End If

                Message_Label.Text = "도면 생성 중..."

                Call Base_Drawing_Create()

                Message_Label.Text = "도면 생성이 완료 되었습니다"
                Base_DWG_Create.Enabled = True
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Drawing_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Base_DWG_Create_Click(sender As Object, e As EventArgs) Handles Base_DWG_Create.Click
        ProgressBar1.Value = 0
        Message_Label.Text = "DWG 변환중..."
        ProgressBar1.Value = 50

        CATIA.ActiveDocument.ExportData(project_data_path.Text & "\LAYOUT\" & Sales_No & "_" & assy_value_Name & "_" & date_value & ".dwg", "dwg")

        ProgressBar1.Value = 100
        Message_Label.Text = "DWG 변환이 완료 되었습니다"
    End Sub


    Private Sub Ref_Drawing_Click(sender As Object, e As EventArgs) Handles Ref_Drawing.Click
        Try
            If SHOW_MESSAGE_BOX("실행 하시겠습니까?", "", 2, 1) = vbYes Then
                Call CATIA_BASIC02()
                Me.ProgressBar1.Value = 0

                '--------------------------------------------------------------------------------------------------- LAYOUT_Folder 초기화
                Dim LAYOUT_Folder = Nothing
                '--------------------------------------------------------------------------------------------------- LAYOUT Folder 가져오기
                fso = CreateObject("Scripting.FileSystemObject")
                LAYOUT_Folder = fso.GetFolder(project_data_path.Text & "\LAYOUT")
                '--------------------------------------------------------------------------------------------------- 현재 시간
                date_value = Format(Now, "yymmddhhmm")
                '--------------------------------------------------------------------------------------------------- LAYOUT Folder가 없을때...
                If LAYOUT_Folder Is Nothing Then
                    fso.CreateFolder(project_data_path.Text & "\LAYOUT")
                End If

                Message_Label.Text = "도면 생성 중..."
                Call Ref_Drawing_Create()

                Message_Label.Text = "도면 생성이 완료 되었습니다."
                Ref_DWG_Create.Enabled = True
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Ref_Drawing_Click() 함수 오류!!", "", 1, 1)
            Message_Label.Text = "도면 생성이 종료되었습니다."
        End Try
    End Sub


    Private Sub Ref_DWG_Create_Click(sender As Object, e As EventArgs) Handles Ref_DWG_Create.Click
        ProgressBar1.Value = 0
        Message_Label.Text = "DWG 변환중..."
        ProgressBar1.Value = 50

        If Enclosure_Value = "Without Enclosure" Then
            CATIA.ActiveDocument.ExportData(project_data_path.Text & "\LAYOUT\" & Sales_No & "_" & assy_value_Name & "_" & date_value & "_REF_W.dwg", "dwg")
        Else
            CATIA.ActiveDocument.ExportData(project_data_path.Text & "\LAYOUT\" & Sales_No & "_" & assy_value_Name & "_" & date_value & "_REF_H.dwg", "dwg")
        End If

        ProgressBar1.Value = 100
        Message_Label.Text = "DWG 변환이 완료 되었습니다"
    End Sub

    Private Sub A_Main_Form_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Call KillProcess("EXCEL")
    End Sub

    Private Sub Oli_Piping_Re_Constraint_Click(sender As Object, e As EventArgs) Handles Oli_Piping_Re_Constraint.Click

    End Sub

    Private Sub Title_Block_Create_Click(sender As Object, e As EventArgs) Handles Title_Block_Create.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ZZZB_Parameter_Create.Show()
    End Sub

    Private Sub Motor_Template_Create_Click(sender As Object, e As EventArgs) Handles Motor_Template_Create.Click

    End Sub
End Class


Public Module STW_New_Platform_Common
    '--------------------------------------------------------------------------------------------------- Excel Open 경로
    Public sfile
    Public lWidth, lHeight


    '--- 경로 도우미
    Public Function GetDirectory(Optional strMsg As String = "", Optional initial_Path As String = "") As String
        GetDirectory = ""
        Dim WinDialg As Dialogs.CommonOpenFileDialog
        WinDialg = New Dialogs.CommonOpenFileDialog
        WinDialg.IsFolderPicker = True
        WinDialg.Multiselect = False
        If strMsg = "" Then
            WinDialg.Title = "폴더를 선택하세요"
        Else
            WinDialg.Title = ""
        End If


        If initial_Path <> "" Then
            WinDialg.InitialDirectory = CStr(initial_Path)
        End If


        If WinDialg.ShowDialog() = DialogResult.OK Then
            GetDirectory = WinDialg.FileName
        End If
    End Function


    '---------------------------------------------------------------------------------------------------
    ' Form Position  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Form_Position_Initial(Form_Name)
        Try
            Form_Name.Left = lWidth
            Form_Name.Top = lHeight
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Get_Excel_Value_Function  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Function Get_Excel_Value_Function(Excel_Path As Object, sheet_name As Object, Excel_Column_Count As Object, Excel_Column_1 As Object, Excel_Column_2 As Object, Excel_Column_3 As Object,
                                             Excel_Column_4 As Object, Excel_Column_5 As Object, Excel_Column_6 As Object, Excel_Column_7 As Object, Excel_Column_8 As Object, Excel_Column_9 As Object,
                                             Excel_Column_10 As Object, ByRef Excel_Low_Count As Object, ByRef Excel_Value As Object, Excel_Option As Object)

        Get_Excel_Value_Function = True

        Try
            Dim Column_Value(10) As String, Check, layout_EPN_count
            '--------------------------------------------------------------------------------------------------- Excle Value Count
            Excel_Low_Count = 1
            '--------------------------------------------------------------------------------------------------- Excle Excel Open
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Excel_Path)

            Check = True
            '--------------------------------------------------------------------------------------------------- Excel_Option = 4 ---> 실적 선택 및 Template 관련 Excel 가져올때...(실적 검색대상 Counting)
            If Excel_Option = 4 Then
                Do
                    If CStr(Exc.Worksheets(sheet_name).Range("I" & Excel_Low_Count + 1).Value) = "" Then
                        Check = False
                        Exit Do
                    Else
                        Excel_Low_Count = Excel_Low_Count + 1
                    End If
                Loop Until Check = False
                '--------------------------------------------------------------------------------------------------- Excel_Option = 5 ---> Template 도면  정보 Excel 가져올때...(도면파트리스트의 Sheet count와 기종정보)
            ElseIf Excel_Option = 5 Then
                Do
                    If CStr(Exc.Worksheets(sheet_name).Range("A" & Excel_Low_Count + 1).Value) = "" Then
                        Check = False
                        Exit Do
                    Else
                        Excel_Low_Count = Excel_Low_Count + 1
                    End If
                Loop Until Check = False
                '--------------------------------------------------------------------------------------------------- Excel_Option = 6 ---> Rotor Assy Constraint 가져올때
            ElseIf Excel_Option = 6 Then
                Do
                    If CStr(Exc.Worksheets(sheet_name).Range("A" & Excel_Low_Count + 1).Value) = "" Then
                        Check = False
                        Exit Do
                    Else
                        Excel_Low_Count = Excel_Low_Count + 1
                    End If
                Loop Until Check = False
                '    '--------------------------------------------------------------------------------------------------- Excel_Option = 0 ---> DT_Design_Rule_Version 정보가져오기(template폴더안 Designrule 버전가져오기)
            Else
                Do
                    If CStr(Exc.Worksheets(sheet_name).Range(Excel_Column_1 & Excel_Low_Count + 1).Value) = "" Then
                        Check = False
                        Exit Do
                    Else
                        Excel_Low_Count = Excel_Low_Count + 1
                    End If
                Loop Until Check = False
            End If

            '--------------------------------------------------------------------------------------------------- Excel Value 변수 크기 정의
            ReDim Excel_Value(0 To Excel_Low_Count - 1, 0 To Excel_Column_Count)
            '--------------------------------------------------------------------------------------------------- Column 만큼 For 문 적용
            Column_Value(1) = Excel_Column_1
            Column_Value(2) = Excel_Column_2
            Column_Value(3) = Excel_Column_3
            Column_Value(4) = Excel_Column_4
            Column_Value(5) = Excel_Column_5
            Column_Value(6) = Excel_Column_6
            Column_Value(7) = Excel_Column_7
            Column_Value(8) = Excel_Column_8
            Column_Value(9) = Excel_Column_9
            Column_Value(10) = Excel_Column_10
            layout_EPN_count = 1

            '--------------------------------------------------------------------------------------------------- Excel 파일 정보 가져오기
            For i = 1 To Excel_Column_Count
                For j = 0 To Excel_Low_Count - 1
                    '--------------------------------------------------------------------------------------------------- Excel_Option = 3 ---> BOM Excel 일때...
                    If Excel_Option = 3 Then    'LAYOUT대상부품구분
                        If Exc.Worksheets(sheet_name).Range("D" & j + 1).Value.ToString() = "Y" Then
                            '--------------------------------------------------------------------------------------------------- Publish_Org_Name 이름 가져오기
                            Excel_Value(layout_EPN_count, i) = Exc.Worksheets(sheet_name).Range(Column_Value(i) & j + 1).Value
                            layout_EPN_count = layout_EPN_count + 1
                        End If
                    Else
                        Excel_Value(j, i) = Exc.Worksheets(sheet_name).Range(Column_Value(i) & j + 1).Value
                        '--------------------------------------------------------------------------------------------------- Nothing 값 비교위해 문자열로 변경 20.01.03
                        If IsNothing(Excel_Value(j, i)) = True Then
                            Excel_Value(j, i) = ""
                        End If
                    End If

                    '--------------------------------------------------------------------------------------------------- Excel_Option = 1 ---> BOM Excel 일때...
                    If Excel_Option = 1 Then
                        '--------------------------------------------------------------------------------------------------- Coupling 유무 판별
                        If Excel_Value(j, i).ToString() = "EPN03" Then
                            coupling_number = 1
                        End If

                        '--------------------------------------------------------------------------------------------------- Oil_Piping_Part Name 정의 ( Oil Piping Layout에 사용 )
                        If Excel_Value(j, i).ToString() = "EPN0601" Then
                            '사용확인 20.03.19 허혜원
                            Oil_Piping_Part_Value = Exc.Worksheets(sheet_name).Range("B" & j).Value
                        End If

                        '--------------------------------------------------------------------------------------------------- Oil Return Piping 도면 생성시 Oil_Piping_BOM_PartNumber가져오기
                        If i = 1 Then
                            If Not (String.IsNullOrEmpty(Exc.Worksheets(sheet_name).Range("I" & j + 1).Value)) Then
                                If Exc.Worksheets(sheet_name).Range("I" & j + 1).Value.ToString() = "EPN060115" Then
                                    Oil_Piping_BOM_PartNumber = Exc.Worksheets(sheet_name).Range("B" & j + 1).Value
                                End If
                            End If
                        End If
                    End If
                Next
                '--------------------------------------------------------------------------------------------------- 마지막 열이 아닐때 변수 초기화
                If Not i = Excel_Column_Count Then
                    layout_EPN_count = 1
                End If
            Next

            '--------------------------------------------------------------------------------------------------- Excel_Option = 2 ---> Constraint Excel 일때...(constraint엑셀에 기종정보기록함,현재는 안쓰는걸로 판단됨)
            If Excel_Option = 2 Then
                Exc.Worksheets(STW_COMP_Infomation(2, 1)).Select()
                'Exc.Worksheets(STW_COMP_Infomation(2, 1)).Range("Z" & 2).Value = assy_value_Name

                '사용확인 20.03.19 허혜원
                Excel_Value(0, 0) = Exc.Worksheets(STW_COMP_Infomation(2, 1)).Range("Z" & 1).Value
                Excel_Value(1, 0) = Exc.Worksheets(STW_COMP_Infomation(2, 1)).Range("Z" & 2).Value
            End If

            '--------------------------------------------------------------------------------------------------- Excel_Option = 3 ---> BOM Excel 일때...
            If Excel_Option = 3 Then
                Excel_Low_Count = layout_EPN_count - 1
            End If

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(False)
            Exc.Quit()

            Call KillProcess("EXCEL")
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Get_Excel_Value_Function() 함수 오류!!", "", 1, 1)
            Call KillProcess("EXCEL")
        End Try
    End Function


    '---------------------------------------------------------------------------------------------------
    '  Btn_Refresh_Click &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Function Btn_Refresh_Click()
        Try
            A_Main_Form.DataGridView1.Rows.Clear()
            A_Main_Form.ProgressBar1.Value = 0
            std_p_number_count = 1

            '--------------------------------------------------------------------------------------------------- 커플링선택초기화 Constraint
            sExcelFilePath = now_o_bom_excel_path
            If EXISTS_FILE_CHECK(sExcelFilePath) Then
                Call Get_Excel_Value_Function(sExcelFilePath, "BOM_LIST", 4, "B", "H", "K", "J", "", "", "", "", "", "",
                                              O_BOM_Part_Value_Count, O_BOM_Part_Value, "1")
            Else : Return False
            End If

            '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기
            Dim iJvalue As Integer = 0
            For i = 2 To O_BOM_Part_Value_Count - 1
                For j = 1 To O_BOM_Layout_Part_Value_Count
                    If String.IsNullOrEmpty(O_BOM_Part_Value(i, 3)) Or String.IsNullOrEmpty(O_BOM_Layout_Part_Value(j, 1)) Then
                        Exit For
                    End If
                    If O_BOM_Part_Value(i, 3).ToString() = O_BOM_Layout_Part_Value(j, 1).ToString() Then
                        A_Main_Form.DataGridView1.Rows.Add()
                        A_Main_Form.DataGridView1.Item(0, std_p_number_count - 1).Value = std_p_number_count
                        A_Main_Form.DataGridView1.Item(1, std_p_number_count - 1).Value = O_BOM_Part_Value(i, 3)
                        A_Main_Form.DataGridView1.Item(3, std_p_number_count - 1).Value = O_BOM_Part_Value(i, 1)
                        A_Main_Form.DataGridView1.Item(4, std_p_number_count - 1).Value = O_BOM_Part_Value(i, 2)
                        layout_EPN_ref(std_p_number_count) = O_BOM_Part_Value(i, 3)
                        'layout_part_name_ref(std_p_number_count) = O_BOM_Part_Value(i, 3)
                        std_p_number_count = std_p_number_count + 1

                        iJvalue = j
                        Exit For
                    End If
                Next
            Next

            A_Main_Form.DataGridView1.Refresh()
            Dim products_name = Nothing
            For i = 1 To std_p_number_count - 1
                For j = 1 To O_BOM_Part_Value_Count
                    If layout_EPN_ref(i) = O_BOM_Part_Value(j, 3) Then
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

                        iJvalue = j
                        Exit For
                    End If
                Next

                If Not O_BOM_Part_Value(iJvalue, 3) = products_name Then
                    A_Main_Form.DataGridView1.Item(2, i - 1).Value = "V"
                End If
            Next

            Return True
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Btn_Refresh_Click() 함수 오류!!", "", 1, 1)
            Return False
        End Try
    End Function

End Module