'모든 프로젝트 필수 설정★Visual Studio 편집하며 계속하기 
'1. 메뉴>디버그>옵션>디버깅>일반>편집하며 계속하기 사용 및 아래 항목 모두 체크
'2. 프로젝트 속성>컴파일>대상 CPU>x86
'3. 참조 속성>Interop 형식 포함>False
'4. 솔루션 구성 Debug 모드
'Dim CATIA As INFITF.Application
'CATIA = GetObject("", "CATIA.Application") ' "", 쓰기

'CATIA 프로젝트 설정
'1. C:\Program Files\Dassault Systemes\B26\win_b64\code\bin\InfTypeLib.tlb 참조 추가
'2. Imports INFITF or Not

''VB6.0->2017 변경사항
'0. 'VB6.0->2017 로 검색
'1. Type -> Structure
'2. 메서드 괄호
'3. hWnd -> Handle
'4. Set 없어짐
'5. App -> Application
'	App.EXEName -> Application.ProductName
'    Application.StartPath = 실행 파일 이름 제외 경로
'	Application.ExecutablePath = 전체 경로
'6.	Imports Microsoft.VisualBasic.FileIO    '[FileSystem] FileCopy, Rename 등
'	Imports System.IO                       '[DirectoryInfo] GetFiles, FullName, Extension 등
'Imports Scripting                       '[FileSystemObject]Microsoft Scripting Runtime 참조 필요. FolderExists, FileExists 등

Imports System.Runtime.InteropServices

'// 추가 한 것들
Public Module ISParkvb

    '---------- ---------- ---------- ---------- ---------- ----------
    ' 기 능 : 파일 선택 창 띄우기
    ' 입력값 : Title : 창 Title, Filter : 선택 파일 확장자 지정, InitialDirectory : 초기 경로
    ' 반환값 : 선택 파일 경로 반환
    ' 생성일 : 19.03.00
    ' 생성자 : 허혜원
    ' 수정일 : ----
    ' 수정자 : ----
    '---------- ---------- ---------- ---------- ---------- ----------
    Public Function ShowOpenFileDialog(Title As String, Filter As String, InitialDirectory As String)
        Dim Split_Path, StrCheckAndCreatePath

        A_Main_Form.OpenFileDialog1.Title = Title
        A_Main_Form.OpenFileDialog1.Filter = Filter        'CATPart파일만
        A_Main_Form.OpenFileDialog1.FileName = ""
        A_Main_Form.OpenFileDialog1.ReadOnlyChecked = True
        A_Main_Form.OpenFileDialog1.ShowReadOnly = True      '= cdlOFNNoReadOnlyReturn            '읽기 전용
        A_Main_Form.OpenFileDialog1.Multiselect = False

        Split_Path = Split(InitialDirectory, "\")
        StrCheckAndCreatePath = ""
        For Each parentPath In Split_Path
            StrCheckAndCreatePath = StrCheckAndCreatePath & parentPath '& "\"
            If Dir(StrCheckAndCreatePath, vbDirectory) = "" Then
                MkDir(StrCheckAndCreatePath)
            End If
            StrCheckAndCreatePath = StrCheckAndCreatePath & "\"
        Next
        A_Main_Form.OpenFileDialog1.InitialDirectory = InitialDirectory 'A_Main_Form.Directory_Path_Text.Text & "\" & A_Main_Form.Lbl_ModelNumber.Text & "\" & EPN_Name
        A_Main_Form.OpenFileDialog1.ShowDialog(A_Main_Form)
        ShowOpenFileDialog = A_Main_Form.OpenFileDialog1.FileName
    End Function


    '---------- ---------- ---------- ---------- ---------- ----------
    ' 기 능 : BOM 문자열 한글 포함 확인하기
    ' 입력값_Name : BOM 파일 명
    ' 반환값 : 한글미포함 True, 한글포함이면 False
    ' 생성일 : 19.03.00
    ' 생성자 : 허혜원
    ' 수정일 : ----
    ' 수정자 : ----
    '---------- ---------- ---------- ---------- ---------- ----------
    Public Function Get_Excel_Name_Check(Excel_Name)
        'Strings.GetChar(Excel_Name, 1)
        Get_Excel_Name_Check = True
        For Each vGetChar In Excel_Name
            '한글은 Globalizationi.UnicodeGategory.OtherLetter에 속함
            If Char.GetUnicodeCategory(vGetChar) = System.Globalization.UnicodeCategory.OtherLetter Then
                'Debug.Print(vGetChar & " 은 한글!!!")
                Get_Excel_Name_Check = False
                Exit Function
            End If
        Next
    End Function

    '---------- ---------- ---------- ---------- ---------- ----------
    ' 기 능 : 특정 Part의 publication을 Dictionary에 추가
    ' 입력값 : Part : Part, PublicationDic : Dictionary, Initial : True시 Dictionary 기존 데이터 제거 후 생성
    ' 반환값 : 
    ' 생성일 : 19.03.00
    ' 생성자 : 허혜원
    ' 수정일 : ----
    ' 수정자 : ----
    '---------- ---------- ---------- ---------- ---------- ----------
    Public Sub GetOnePartPublication(Part As Object, ByRef PublicationDic As Object, Initial As Boolean)
        If TypeName(PublicationDic).ToString() = "Nothing" Then
            PublicationDic = CreateObject("Scripting.Dictionary")
        End If

        If Initial = True Then
            PublicationDic.RemoveAll()
        End If

        For i = 1 To Part.publications.Count
            If PublicationDic.Exists(Part.publications.Item(i).Name.ToString()) = False Then
                PublicationDic.Add(Part.publications.Item(i).Name.ToString(), Part.publications.Item(i).Name.ToString())
            End If
        Next i
    End Sub


    '---------- ---------- ---------- ---------- ---------- ----------
    ' 기 능 : Product상의 모든 Part의 Publication을 배열과 Dictionary에 추가
    ' 입력값 : Initial : True시 Dictionary 기존 데이터 제거 후 생성
    ' 반환값 : 
    ' 생성일 : 19.03.00
    ' 생성자 : 허혜원
    ' 수정일 : ----
    ' 수정자 : ----
    '---------- ---------- ---------- ---------- ---------- ----------
    Public Sub GetProductAllPartPublication(Initial As Boolean)
        pub_number = 1
        If Initial = True Then
            '/// publication 저장된 Dictionary get_part_publication2017 초기화한 후 재입력
            get_part_publication2017.Clear()
        End If
        For i = 1 To products.Count
            publications1 = products.Item(i).publications
            For k = 1 To publications1.Count
                If Not get_part_publication2017.ContainsKey(publications1.Item(k).Name.ToString()) Then
                    get_part_publication2017.Add(publications1.Item(k).Name.ToString(), publications1.Item(k))
                End If
                pub_number = pub_number + 1
            Next
        Next
    End Sub


    '---------- ---------- ---------- ---------- ---------- ----------
    ' 기 능 : Publication 변수에 입력하기
    ' 입력값 : Pub_reference : Publiincation 입력 변수, Pub_DisplayName : Publication 경로, O_BOM_Constraint_Value_One : Publication Name
    ' 반환값 : 
    ' 생성일 : 19.03.00
    ' 생성자 : 허혜원
    ' 수정일 : ----
    ' 수정자 : ----
    '---------- ---------- ---------- ---------- ---------- ----------
    Public Function GetPub_reference(ByRef Pub_reference As Object, Pub_DisplayName As String, O_BOM_Constraint_Value_One As String)
        Dim Split_DisplayName, Geometry_Name, High_Geometry_Name, parentName, Parent
        Pub_reference = vbEmpty.ToString()
        'Pub_DisplayName = "A"
        Try
            If Pub_DisplayName <> "" Then
                Pub_reference = product_add.CreateReferenceFromName(Pub_DisplayName)
                '--------------------------------------------------------------------------------------------------- Publication 못 찾을 경우 1
                If Pub_reference.ToString() = vbEmpty.ToString() Then   'VB6.0->2017
                    Split_DisplayName = Split(Pub_DisplayName, "!")
                    Geometry_Name = Split(Split_DisplayName(UBound(Split_DisplayName)), "/")

                    selection.Clear()
                    selection.search("Name='" & Geometry_Name(0) & "',all")
                    High_Geometry_Name = selection.Item(1).Value.Parent.Name
                    Pub_reference = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & High_Geometry_Name & "/" & Split_DisplayName(1))
                End If
                '--------------------------------------------------------------------------------------------------- Publication 못 찾을 경우 2
                If Pub_reference.ToString() = vbEmpty.ToString() Or Pub_reference.Name = "" Then
                    selection.Clear()
                    selection.search("Name='" & O_BOM_Constraint_Value_One & "',all")
                    Split_DisplayName = Split(Pub_DisplayName, "!")
                    Parent = selection.Item(1).Value.Parent.Parent
                    parentName = ""

                    Parent = Parent.Parent
                    parentName = Parent.Name & "/" & parentName
                    Pub_reference = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & parentName & Split_DisplayName(1))

                    If Pub_reference.ToString() = vbEmpty.ToString() Or Pub_reference.Name = "" Then
                        GetPub_reference = False
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.Print("A")
        End Try

        If Pub_reference.ToString = vbEmpty.ToString() Then
            Try
                Split_DisplayName = Split(Pub_DisplayName, "!")
                Dim tmpName = Strings.Left(Split_DisplayName(0), Split_DisplayName(0).length - 1)
                Geometry_Name = Split(tmpName, "/")
                selection.search("Name='" & Geometry_Name(UBound(Geometry_Name)) & "',all")
                If selection.count = 0 Then
                    Exit Function
                End If

                Geometry_Name = Split(Split_DisplayName(1), "/")
                tmpName = Geometry_Name(UBound(Geometry_Name) - 1)
                Geometry_Name = tmpName

                selection.search("Name='" & Geometry_Name & "',sel")
                High_Geometry_Name = "" 'Geometry_Name
                Dim j
                Dim tmp
                tmp = selection.Item(1).Value
                If selection.count = 0 Then
                    Exit Function
                End If
                For j = 1 To 5
                    tmp = tmp.Parent
                    If TypeName(tmp) <> "Part" And TypeName(tmp) <> "PartDocument" Then
                        High_Geometry_Name = tmp.Name & "/" & High_Geometry_Name
                    Else
                        Exit For
                    End If
                Next
                'High_Geometry_Name = selection.Item(1).Value.Parent.Name
                Pub_reference = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & High_Geometry_Name & Split_DisplayName(1))
            Catch ex As Exception
                Debug.Print("A")
            End Try
        End If


        Return 0
    End Function




End Module
