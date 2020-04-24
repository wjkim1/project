
Public Class ZA_Clash_Check_Create
    Dim oldGridX As Integer, oldGridY As Integer


    Private Sub ZA_Clash_Check_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 897
        Me.Height = 405
        Me.TopMost = True

        Call CATIA_BASIC02()

        '--------------------------------------------------------------------------------------------------- 폼 X 버튼 비활성화
        'Call REMOVED_CLOSE_BUTTON(Me.Handle)
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '--------------------------------------------------------------------------------------------------- 체크 버튼, Distance 값 확인
        For i = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Rows(i).Cells(0).Value = False Then
                Call SHOW_MESSAGE_BOX("간섭체크 항목을 Check 해주세요.", "", 1, 1)
                Exit Sub
            End If

            If InStr(UCase(DataGridView1.Rows(i).Cells(4).Value), UCase("Distance fail")) <> 0 Then
                If DataGridView1.Rows(i).Cells(4).Value = "" Then
                    Call SHOW_MESSAGE_BOX("Distance fail 항목의 간섭검증 값을 입력해주세요.", "", 1, 1)
                    Exit Sub
                End If
            End If

            If InStr(UCase(DataGridView1.Rows(i).Cells(4).Value), UCase("CLASH")) <> 0 Then
                If (InStr(UCase(DataGridView1.Rows(i).Cells(4).Value), UCase("Non CLASH")) <> 0) Then
                    GoTo Next_i
                End If

                If InStr(UCase(DataGridView1.Rows(i).Cells(4).Value), UCase("CLASH : 0")) <> 0 Then
                    GoTo Next_i
                End If

                If DataGridView1.Rows(i).Cells(5).Value = "" Then
                    Call SHOW_MESSAGE_BOX("CLASH 항목의 간섭검증 값을 입력해주세요.", "", 1, 1)
                    Exit Sub
                End If
            End If

Next_i:
        Next i

        Call Part_Component_Clash_Excel_Export_Click()
        Me.Close()
    End Sub


#Region " --- Part_Component_Clash_Excel_Export_Click() : 간섭 체크 폼의 내용을 엑셀로 저장 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.02
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Private Sub Part_Component_Clash_Excel_Export_Click()
        Try
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Add()
            Exc.Worksheets(1).Select()

            '--------------------------------------------------------------------------------------------------- 헤더 정보 입력
            For i = 1 To 10
                Exc.Worksheets(1).Cells(1, i).Value = DataGridView1.Columns(i - 1).HeaderText
            Next

            For i = 0 To Clash_Check_Excel_Count - 2
                For j = 1 To 10
                    Exc.Worksheets(1).Cells(i + 2, j).Value = DataGridView1.Rows(i).Cells(j).Value
                Next
            Next
            Exc.Worksheets(1).Columns("A:J").Select()
            Exc.Worksheets(1).Columns("A:J").EntireColumn.AutoFit()

            Dim assy_value_Name_Split = Split(sO_BOM_File_Name, "_")
            assy_value_Name = assy_value_Name_Split(1)

            Exc.ActiveWorkbook.SaveAs(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\CLASH_CHECK\" & EPN_NO_Label.Text & "_Crash_" & Format(Now, "yyMMddhhmm") & ".xlsx")

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(True)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL") 
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Part_Component_Clash_Excel_Export_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- DataGridView1_CellContentClick() : 그리드 클릭시.. "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.02
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If

            If DataGridView1.Columns(e.ColumnIndex).Name.ToString = "btnCheck" Then
                DataGridView1.Rows(e.RowIndex).Cells(0).Value = True
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("DataGridView1_CellContentClick() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region



End Class