# VBA-Number_Puzzle
![Number_Puzzle](https://user-images.githubusercontent.com/66747535/100056576-39237180-2e69-11eb-929f-1c9b68246ea9.gif)


엑셀에서 VBA 매크로를 통해 실행할 수 있는 숫자 퍼즐 게임이다.

## 구현한 핵심 알고리즘
1. 버튼 입력에 따른 퍼즐 변경
2. 퍼즐이 완성된 상태인지 점검

## 적용법
1. VBA 편집창에 들어간다.
2. 모듈이 아니라 적용할 시트의 코드 창에 아래의 코드를 모두 넣는다.
3. 매크로 직접 실행으로 Format 실행

## 코드
<details>
    <summary>코드보기</summary>

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells(1, 8) = "=rows(" + Selection.Address + ")"
    m = Cells(1, 8) '행 크기
    Cells(1, 8) = "=columns(" + Selection.Address + ")"
    n = Cells(1, 8) '열 크기
    
    If m > 1 Or n > 1 Then '다중선택방지
    ElseIf Selection.Address = Cells(1, 7).Address Then
        Start '게임시작
    ElseIf Selection.Address = Cells(3, 7).Address Then
        Mov (1) '상
    ElseIf Selection.Address = Cells(5, 7).Address Then
        Mov (2) ' 하
    ElseIf Selection.Address = Cells(4, 6).Address Then
        Mov (3) '좌
    ElseIf Selection.Address = Cells(4, 8).Address Then
        Mov (4) '우
    End If
    
    '클리어 테스트
    Clear
    
    Cells(4, 7).Select
End Sub


Function Mov(tp)
    Selection.Font.Color = RGB(100, 255, 100)
    
    Application.ScreenUpdating = False
    
    m = Cells(2, 8)
    n = Cells(3, 8)
    
    Select Case tp
    Case 1 '상
        If m <> 5 Then
            With Cells(m + 1, n)
                .Copy Cells(m, n)
                .Value = ""
                .Interior.Color = RGB(0, 0, 0)
                Cells(2, 8) = .Row
                Cells(3, 8) = .Column
            End With
        End If
    Case 2 '하
        If m <> 1 Then
            With Cells(m - 1, n)
                .Copy Cells(m, n)
                .Value = ""
                .Interior.Color = RGB(0, 0, 0)
                 Cells(2, 8) = .Row
                 Cells(3, 8) = .Column
            End With
        End If
    Case 3 '좌
        If n <> 5 Then
            With Cells(m, n + 1)
                .Copy Cells(m, n)
                .Value = ""
                .Interior.Color = RGB(0, 0, 0)
                Cells(2, 8) = .Row
                Cells(3, 8) = .Column
            End With
        End If
    Case 4 '우
        If n <> 1 Then
            With Cells(m, n - 1)
                .Copy Cells(m, n)
                .Value = ""
                .Interior.Color = RGB(0, 0, 0)
                Cells(2, 8) = .Row
                Cells(3, 8) = .Column
            End With
        End If
    End Select
    
    Application.ScreenUpdating = True
    
    Selection.Font.Color = RGB(255, 255, 255)
End Function

Function Clear()
    k = 0
    For Each ce In Range(Cells(1, 1), Cells(5, 5))
        k = k + 1
        If k = 25 Then '클리어
            Cells(5, 5) = 25
            Cells(5, 5).Interior.Color = RGB(100, 100, 255)
            Cells(2, 7) = "Clear!"
        ElseIf ce.Value <> k Then '클리어 실패
            Exit Function
        End If
    Next
End Function

Sub Format() '포맷
    Application.ScreenUpdating = False
    
    Range("A1：XFD1048576").EntireRow.Clear
    Range("A1：XFD1048576").EntireColumn.Clear
    Range("I7：XFD1048576").EntireRow.Hidden = True
    Range("I7：XFD1048576").EntireColumn.Hidden = True
    
    With Range(Cells(1, 1), Cells(5, 5))
        .ColumnWidth = 9
        .RowHeight = 60
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 153)
        .Font.Size = 40
    End With
    With Range(Cells(1, 6), Cells(5, 8))
        .ColumnWidth = 9
        .RowHeight = 60
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(O, O, O)
        .Font.Color = RGB(255, 255, 255)
    End With
    With Range(Cells(6, 1), Cells(6, 8))
        .MergeCells = True
        .RowHeight = 25
        .Interior.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Value = "원쪽 위에서부터 1~24가 위치하게 만들면 된다."
    End With
    With Cells(1, 7)
        .Value = "시작"
        .Interior.Color = RGB(100, 100, 255)
        .Font.Size = 15
        .Font.Bold = True
    End With
    With Cells(2, 7)
        .Font.Color = RGB(100, 100, 255)
        .Font.Size = 15
        .Font.Bold = True
    End With
    Range("H1.H2,H3").Font.Color = RGB(0, 0, 0)
    Range("G3,G5,F4,H4").Font.Size = 30
    
    Cells(4, 7) = "방향키로" & vbCrLf & "이동"
    Cells(3, 7) = "▲"
    Cells(4, 6) = "◀"
    Cells(4, 8) = "▶"
    Cells(5, 7) = "▼"
    
    Start
End Sub
Function Start()
    Application.ScreenUpdating = False

    Range(Cells(1, 1), Cells(5, 5)).Interior.Color = RGB(100, 100, 255)
    Cells(2, 7) = ""

    For Each ce In Range(Cells(1, 1), Cells(5, 5))
        i = i + 1
        ce.Value = i
    Next

    For i = 1 To 20 '섞어주기
        Cells(1, 8) = "=RANDBETWEEN(1,5)"
        j = Cells(1, 8)
        Cells(1, 8) = "=RANDBETWEEN(1,5)"
        k = Cells(1, 8)
        Cells(1, 8) = "=RANDBETWEEN(1,5)"
        m = Cells(1, 8)
        Cells(1, 8) = "=RANDBETWEEN(1,5)"
        n = Cells(1, 8)
        tmp = Cells(j, k)
        Cells(j, k) = Cells(m, n)
        Cells(m, n) = tmp
    Next

    For Each ce In Range(Cells(1, 1), Cells(5, 5))
        If ce.Value = 25 Then
            ce.Value = ""
            ce.Interior.Color = RGB(0, 0, 0)

            Cells(2, 8) = ce.Row '빈칸' 위치 저장
            Cells(3, 8) = ce.Column

            Application.ScreenUpdating = True
            Exit Function
        End If
    Next
End Function

Private Sub Worksheet_Activate()
    Cells(4, 7).Select
End Sub
```

</details>
