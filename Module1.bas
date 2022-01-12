Attribute VB_Name = "Module1"
Dim ip, op, project As Workbook
Dim ttxt, toan, van, lsu, anh, nv, chitieu, diemchuan, dstrg1, dstrg2, dstrg3, dstr4, dstrg5, proj As Worksheet

Dim sbd(10000), ten(10000) As String
Dim n, nv1(10000), nv2(10000), nv3(10000), ctieu(10) As Integer
Dim dtong(10000) As Double
Dim dtoan(10000), dvan(10000), danh(10000), dlichsu(10000) As Double

'n1t1 là mang luu vi tri cua nhung nguoi chon nv1 tai trg 1(tuong tu voi n1t2,n1t3,n1t4,n1t5)
'sn1t1 là so nguoi chon nv1 tai trg 1(tuong tu voi sn1t2, sn1t3, sn1t4, sn1t5)
'dcn1t1 là diem chuan nv1 cua trg 1( tuong tu voi dcn1t2, dcn1t3, dcn1t4, dcn1t5)
Dim n1t1(10000), n1t2(10000), n1t3(10000), n1t4(10000), n1t5(10000) As Integer
Dim sn1t1, sn1t2, sn1t3, sn1t4, sn1t5 As Integer
Dim dcn1t1, dcn1t2, dcn1t3, dcn1t4, dcn1t5 As Double

'tuong tu nhu tren nhung da loc nhung hoc sinh trung tuyen nv1
Dim sn2t1, sn2t2, sn2t3, sn2t4, sn2t5 As Integer
Dim n2t1(10000), n2t2(10000), n2t3(10000), n2t4(10000), n2t5(10000) As Integer
Dim dcn2t1, dcn2t2, dcn2t3, dcn2t4, dcn2t5 As Double

'tuong tu nhu tren nhung da loc nhung hoc sinh trung tuyen nv1,nv2
Dim sn3t1, sn3t2, sn3t3, sn3t4, sn3t5 As Integer
Dim n3t1(10000), n3t2(10000), n3t3(10000), n3t4(10000), n3t5(10000) As Integer
Dim dcn3t1, dcn3t2, dcn3t3, dcn3t4, dcn3t5 As Double

'sot1 là so hoc sinh trung tuyen vao truong 1(tuong tu voi sot2, sot3, sot5, sot4)
'trg1() là mang luu lai vi tri cua cac hoc sinh da trung tuyen vao trg 1(tuong tu voi trg2,trg3,trg4,trg5)
Dim sot1, sot2, sot3, sot5, sot4 As Integer
Dim trg1(10000), trg2(10000), trg3(10000), trg4(10000), trg5(10000) As Integer

Sub Openfile()
Workbooks.Open ("D:\PROJECT\input.xlsx")

Set ip = Workbooks("input.xlsx")
Set ttxt = ip.Worksheets("thong_tin_xet_tuyen")
Set toan = ip.Worksheets("diem_toan")
Set van = ip.Worksheets("diem_van")
Set lsu = ip.Worksheets("diem_lich_su")
Set anh = ip.Worksheets("diem_ngoai_ngu")
Set nv = ip.Worksheets("nguyen_vong")
Set chitieu = ip.Worksheets("chi_tieu")

'Tim so hoc sinh co trong danh sach luu vao bien n
n = 0
i = 1
While (ttxt.Cells(i, 1) <> "")
    n = n + 1
    i = i + 1
Wend

End Sub

'Tao workbook
Sub taoFile()
Dim i As Double
Application.DisplayAlerts = False
Set output = Workbooks.Add

output.SaveAs "D:\PROJECT" & "\" & "output" & ".xlsx"
i = 2
output.Sheets.Add(before:=Sheets(i - 1)).Name = "diem_chuan"
i = i + 1
output.Sheets.Add(before:=Sheets(i - 1)).Name = "danh_sach_1"
i = i + 1
output.Sheets.Add(before:=Sheets(i - 1)).Name = "danh_sach_2"
i = i + 1
output.Sheets.Add(before:=Sheets(i - 1)).Name = "danh_sach_3"
i = i + 1
output.Sheets.Add(before:=Sheets(i - 1)).Name = "danh_sach_4"
i = i + 1
output.Sheets.Add(before:=Sheets(i - 1)).Name = "danh_sach_5"
i = i + 1

output.Sheets(7).Delete


Workbooks.Open ("D:\PROJECT\output.xlsx")

Set op = Workbooks("output.xlsx")

Set diemchuan = op.Worksheets("diem_chuan")
Set dstrg1 = op.Worksheets("danh_sach_1")
Set dstrg2 = op.Worksheets("danh_sach_2")
Set dstrg3 = op.Worksheets("danh_sach_3")
Set dstrg4 = op.Worksheets("danh_sach_4")
Set dstrg5 = op.Worksheets("danh_sach_5")


For i = 1 To 5
    diemchuan.Cells(i, 1) = i
Next

'In ra diem chuan cua tung truong cho nao diem chuan khong co thi khong ghi
If dcn1t1 = 60 Then
    diemchuan.Cells(1, 2) = ""
Else: diemchuan.Cells(1, 2) = dcn1t1
End If

If dcn1t2 = 60 Then
    diemchuan.Cells(2, 2) = ""
Else: diemchuan.Cells(2, 2) = dcn1t2
End If

If dcn1t3 = 60 Then
    diemchuan.Cells(3, 2) = ""
Else: diemchuan.Cells(3, 2) = dcn1t3
End If

If dcn1t4 = 60 Then
    diemchuan.Cells(4, 2) = ""
Else: diemchuan.Cells(4, 2) = dcn1t4
End If

If dcn1t5 = 60 Then
    diemchuan.Cells(5, 2) = ""
Else: diemchuan.Cells(5, 2) = dcn1t5
End If


If dcn2t1 = 60 Then
    diemchuan.Cells(1, 3) = ""
Else: diemchuan.Cells(1, 3) = dcn2t1
End If

If dcn2t2 = 60 Then
    diemchuan.Cells(2, 3) = ""
Else: diemchuan.Cells(2, 3) = dcn2t2
End If

If dcn2t3 = 60 Then
    diemchuan.Cells(3, 3) = ""
Else: diemchuan.Cells(3, 3) = dcn2t3
End If

If dcn2t4 = 60 Then
    diemchuan.Cells(4, 3) = ""
Else: diemchuan.Cells(4, 3) = dcn2t4
End If

If dcn2t5 = 60 Then
    diemchuan.Cells(5, 3) = ""
Else: diemchuan.Cells(5, 3) = dcn2t5
End If


If dcn3t1 = 60 Then
    diemchuan.Cells(1, 4) = ""
Else: diemchuan.Cells(1, 4) = dcn3t1
End If

If dcn3t2 = 60 Then
    diemchuan.Cells(2, 4) = ""
Else: diemchuan.Cells(2, 4) = dcn3t2
End If

If dcn3t3 = 60 Then
    diemchuan.Cells(3, 4) = ""
Else: diemchuan.Cells(3, 4) = dcn3t3
End If

If dcn3t4 = 60 Then
    diemchuan.Cells(4, 4) = ""
Else: diemchuan.Cells(4, 4) = dcn3t4
End If

If dcn3t5 = 60 Then
    diemchuan.Cells(5, 4) = ""
Else: diemchuan.Cells(5, 4) = dcn3t5
End If

For i1 = 1 To sot1
    dstrg1.Cells(i1, 1) = ten(trg1(i1))
Next


For i2 = 1 To sot2
    dstrg2.Cells(i2, 1) = ten(trg2(i2))
Next

For i3 = 1 To sot3
    dstrg3.Cells(i3, 1) = ten(trg3(i3))
Next

dstrg4.Cells(2, 2) = hovatentv
For i4 = 1 To sot4
    dstrg4.Cells(i4, 1) = ten(trg4(i4))
Next

For i5 = 1 To sot5
    dstrg5.Cells(i5, 1) = ten(trg5(i5))
Next

End Sub
Sub ganvaomang()
'gan Sbd và tên vào mang sbd va ten
For i = 1 To n
    sbd(i) = ttxt.Cells(i, 1)
    ten(i) = ttxt.Cells(i, 2)
Next

' tim cac diem và các nguyen vong dung vi tri thu n trong mang luu diem va nguyen vong
For i = 1 To n
     ' tim dung SBD o trong sheet diem anh roi gan diem anh
    For j = 1 To n
        If anh.Cells(j, 1) = sbd(i) Then
        danh(i) = anh.Cells(j, 2)
        End If
    Next
    For j = 1 To n
        ' neu diem tieng anh nho hon 2 thi gan diem toan bang -100
        ' con neu khong thi tim dung SBD trong sheet toan roi gan diem toan
        If danh(i) < 2 Then
            dtoan(i) = -100
        ElseIf toan.Cells(j, 1) = sbd(i) Then
        dtoan(i) = toan.Cells(j, 2)
        Exit For
        End If
    Next
    For j = 1 To n
        If van.Cells(j, 1) = sbd(i) Then
        dvan(i) = van.Cells(j, 2)
        Exit For
        End If
    Next
    For j = 1 To n
        If lsu.Cells(j, 1) = sbd(i) Then
        dlichsu(i) = lsu.Cells(j, 2)
        Exit For
        End If
    Next
 
    For j = 1 To n
        If nv.Cells(j, 1) = sbd(i) Then
        nv1(i) = nv.Cells(j, 2)
        Exit For
        End If
    Next
    For j = 1 To n
        If nv.Cells(j, 1) = sbd(i) Then
        nv2(i) = nv.Cells(j, 3)
        Exit For
        End If
    Next
    For j = 1 To n
        If nv.Cells(j, 1) = sbd(i) Then
        nv3(i) = nv.Cells(j, 4)
        Exit For
        End If
    Next
    
     ' cach tinh diem tong de gan vao mang diem tong
    dtong(i) = dtoan(i) * 2 + dvan(i) * 2 + dlichsu(i)
    
    ' neu diem tong nho hon 0 thi khong xet den nhung thi sinh nay
    If dtong(i) < 0 Then
        nv1(i) = 0
        nv2(i) = 0
        nv3(i) = 0
    End If
Next

' gan chi tieu 5 truong vao mang
For i = 1 To 5
    ctieu(i) = chitieu.Cells(i, 1)
Next
End Sub
Sub gannv1()
sn1t1 = 0
sn1t2 = 0
sn1t3 = 0
sn1t4 = 0
sn1t5 = 0
' Gan vi tri cua cac hoc sinh dang ky nguyen vong 1 tai cac trg vao cac mang co ten n1t1,n1t2,..,n1t5 va dem so hoc sinh dang ky do
For i = 1 To n
     Select Case nv1(i)
        Case 1
        sn1t1 = sn1t1 + 1
        n1t1(sn1t1) = i
        Case 2
        sn1t2 = sn1t2 + 1
        n1t2(sn1t2) = i
        Case 3
        sn1t3 = sn1t3 + 1
        n1t3(sn1t3) = i
        Case 4
        sn1t4 = sn1t4 + 1
        n1t4(sn1t4) = i
        Case 5
        sn1t5 = sn1t5 + 1
        n1t5(sn1t5) = i
    End Select
Next
End Sub
Sub gannv2()
sn2t1 = 0
sn2t2 = 0
sn2t3 = 0
sn2t4 = 0
sn2t5 = 0
' Gan vi tri cua cac hoc sinh dang ky nguyen vong 2 sau khi loai bo hoc sinh trung tuyen nv1 tai cac trg vao cac mang co ten n2t1,n2t2,..,n2t5 va dem so hoc sinh dang ky do
For i = 1 To n
     Select Case nv2(i)
        Case 1
        sn2t1 = sn2t1 + 1
        n2t1(sn2t1) = i
        Case 2
        sn2t2 = sn2t2 + 1
        n2t2(sn2t2) = i
        Case 3
        sn2t3 = sn2t3 + 1
        n2t3(sn2t3) = i
        Case 4
        sn2t4 = sn2t4 + 1
        n2t4(sn2t4) = i
        Case 5
        sn2t5 = sn2t5 + 1
        n2t5(sn2t5) = i
    End Select
Next
End Sub
Sub gannv3()
sn3t1 = 0
sn3t2 = 0
sn3t3 = 0
sn3t4 = 0
sn3t5 = 0
For i = 1 To n
     Select Case nv3(i)
        Case 1
        sn3t1 = sn3t1 + 1
        n3t1(sn3t1) = i
        Case 2
        sn3t2 = sn3t2 + 1
        n3t2(sn3t2) = i
        Case 3
        sn3t3 = sn3t3 + 1
        n3t3(sn3t3) = i
        Case 4
        sn3t4 = sn3t4 + 1
        n3t4(sn3t4) = i
        Case 5
        sn3t5 = sn3t5 + 1
        n3t5(sn3t5) = i
    End Select
Next
End Sub
Sub locnv1trg1()
'''neu chi tieu lon hon hoac bang so hoc sinh o nguyen vong 1 cua truong 1 thi lay het so hoc sinh do
If (ctieu(1) >= sn1t1) Then
    'hoc sinh duoc chon roi thi gan nv2,nv3 bang 0 de khong xet nua
    For i = 1 To sn1t1
        nv2(n1t1(i)) = 0
        nv3(n1t1(i)) = 0
        'tim diem chuan bang cach tim diem tong nho nhat
        If (dcn1t1 > dtong(n1t1(i))) Then
            dcn1t1 = dtong(n1t1(i))
        End If
        sot1 = sot1 + 1
        trg1(sot1) = n1t1(i)
    Next
    ctieu(1) = ctieu(1) - sn1t1
'neu chi tieu nho hon so hoc sinh o nv1 trg1 thi loc ra dung so chi tieu do
Else
    'lay so hoc sinh o dau mang n1t1 bang so chi tieu cho vao mang laysn1t1
    Dim laysn1t1(10000) As Integer
    For i = 1 To ctieu(1)
        laysn1t1(i) = (n1t1(i))
    Next
    'chay tu vi tri ctieu(1) +1 den cuoi mang n1t1 cu thay the hoc sinh co diem cao hon vao vi tri cua hoc sinh co diem thap nhat trong mang laysn1t1
    For i = ctieu(1) + 1 To sn1t1
        Dim luu As Integer
        luu = 1
        'tim hoc sinh co diem thap nhat trong mang laysn1t1
        For j = 2 To ctieu(1)
            '
            If (dtong(laysn1t1(j)) < dtong(laysn1t1(luu))) Or ((dtong(laysn1t1(j)) = dtong(laysn1t1(luu))) And danh(laysn1t1(j)) < danh(laysn1t1(luu))) Then
                luu = j
            End If
        Next
        'thay the neu hoc sinh ben ngoai co diem cao hon hoc sinh co diem thap nhat trong mang laysn1t1 vao dung vi tri do
        If (dtong(n1t1(i)) > dtong(laysn1t1(luu))) Or ((dtong(n1t1(i)) = dtong(laysn1t1(luu))) And danh(n1t1(i)) > danh(laysn1t1(luu))) Then
            laysn1t1(luu) = n1t1(i)
        End If
    Next
    'tim hoc sinh co diem thap nhat de thanh diemchuan va gan nv2 nv3 cua cac hoc sinh da duoc chon bang 0
    For i = 1 To ctieu(1)
        nv2(laysn1t1(i)) = 0
        nv3(laysn1t1(i)) = 0
        If (dcn1t1 > dtong(laysn1t1(i))) Then
            dcn1t1 = dtong(laysn1t1(i))
        End If
        'tinh so hoc sinh trung tuyen vao truong
        sot1 = sot1 + 1
        'luu lai vi tri cua hoc sinh da trung tuyen vao truong
        trg1(sot1) = laysn1t1(i)
    Next
    ctieu(1) = 0
End If
End Sub
Sub locnv1trg2()
If (ctieu(2) >= sn1t2) Then
    For i = 1 To sn1t2
        nv2(n1t2(i)) = 0
        nv3(n1t2(i)) = 0
        If (dcn1t2 > dtong(n1t2(i))) Then
            dcn1t2 = dtong(n1t2(i))
        End If
        sot2 = sot2 + 1
        trg2(sot2) = n1t2(i)
    Next
    ctieu(2) = ctieu(2) - sn1t2
Else
    Dim laysn1t2(10000) As Integer
    For i = 1 To ctieu(2)
        laysn1t2(i) = (n1t2(i))
    Next
    For i = ctieu(2) + 1 To sn1t2
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(2)
            If (dtong(laysn1t2(j)) < dtong(laysn1t2(luu))) Or ((dtong(laysn1t2(j)) = dtong(laysn1t2(luu))) And danh(laysn1t2(j)) < danh(laysn1t2(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n1t2(i)) > dtong(laysn1t2(luu))) Or ((dtong(n1t2(i)) = dtong(laysn1t2(luu))) And danh(n1t2(i)) > danh(laysn1t2(luu))) Then
            laysn1t2(luu) = n1t2(i)
        End If
    Next
    For i = 1 To ctieu(2)
        nv2(laysn1t2(i)) = 0
        nv3(laysn1t2(i)) = 0
        If (dcn1t2 > dtong(laysn1t2(i))) Then
            dcn1t2 = dtong(laysn1t2(i))
        End If
        sot2 = sot2 + 1
        trg2(sot2) = laysn1t2(i)
    Next
    ctieu(2) = 0
End If
End Sub
Sub locnv1trg3()
If (ctieu(3) >= sn1t3) Then
    For i = 1 To sn1t3
        nv2(n1t3(i)) = 0
        nv3(n1t3(i)) = 0
        If (dcn1t3 > dtong(n1t3(i))) Then
            dcn1t3 = dtong(n1t3(i))
        End If
        sot3 = sot3 + 1
        trg3(sot3) = n1t3(i)
    Next
    ctieu(3) = ctieu(3) - sn1t3
Else
    Dim laysn1t3(10000) As Integer
    For i = 1 To ctieu(3)
        laysn1t3(i) = (n1t3(i))
    Next
    For i = ctieu(3) + 1 To sn1t3
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(3)
            If (dtong(laysn1t3(j)) < dtong(laysn1t3(luu))) Or ((dtong(laysn1t3(j)) = dtong(laysn1t3(luu))) And danh(laysn1t3(j)) < danh(laysn1t3(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n1t3(i)) > dtong(laysn1t3(luu))) Or ((dtong(n1t3(i)) = dtong(laysn1t3(luu))) And danh(n1t3(i)) > danh(laysn1t3(luu))) Then
            laysn1t3(luu) = n1t3(i)
        End If
    Next
    For i = 1 To ctieu(3)
        nv2(laysn1t3(i)) = 0
        nv3(laysn1t3(i)) = 0
        If (dcn1t3 > dtong(laysn1t3(i))) Then
            dcn1t3 = dtong(laysn1t3(i))
        End If
        sot3 = sot3 + 1
        trg3(sot3) = laysn1t3(i)
    Next
    ctieu(3) = 0
End If
End Sub
Sub locnv1trg4()
If (ctieu(4) >= sn1t4) Then
    For i = 1 To sn1t4
        nv2(n1t4(i)) = 0
        nv3(n1t4(i)) = 0
        If (dcn1t4 > dtong(n1t4(i))) Then
            dcn1t4 = dtong(n1t4(i))
        End If
        sot4 = sot4 + 1
        trg4(sot4) = n1t4(i)
    Next
    ctieu(4) = ctieu(4) - sn1t4
Else
    Dim laysn1t4(10000) As Integer
    For i = 1 To ctieu(4)
        laysn1t4(i) = (n1t4(i))
    Next
    For i = ctieu(4) + 1 To sn1t4
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(4)
            If (dtong(laysn1t4(j)) < dtong(laysn1t4(luu))) Or ((dtong(laysn1t4(j)) = dtong(laysn1t4(luu))) And danh(laysn1t4(j)) < danh(laysn1t4(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n1t4(i)) > dtong(laysn1t4(luu))) Or ((dtong(n1t4(i)) = dtong(laysn1t4(luu))) And danh(n1t4(i)) > danh(laysn1t4(luu))) Then
            laysn1t4(luu) = n1t4(i)
        End If
    Next
    For i = 1 To ctieu(4)
        nv2(laysn1t4(i)) = 0
        nv3(laysn1t4(i)) = 0
        If (dcn1t4 > dtong(laysn1t4(i))) Then
            dcn1t4 = dtong(laysn1t4(i))
        End If
        sot4 = sot4 + 1
        trg4(sot4) = laysn1t4(i)
    Next
    ctieu(4) = 0
End If
End Sub
Sub locnv1trg5()
If (ctieu(5) >= sn1t5) Then
    For i = 1 To sn1t5
        nv2(n1t5(i)) = 0
        nv3(n1t5(i)) = 0
        If (dcn1t5 > dtong(n1t5(i))) Then
            dcn1t5 = dtong(n1t5(i))
        End If
        sot5 = sot5 + 1
        trg5(sot5) = n1t5(i)
    Next
    ctieu(5) = ctieu(5) - sn1t5
Else
    Dim laysn1t5(10000) As Integer
    For i = 1 To ctieu(5)
        laysn1t5(i) = (n1t5(i))
    Next
    For i = ctieu(5) + 1 To sn1t5
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(5)
            If (dtong(laysn1t5(j)) < dtong(laysn1t5(luu))) Or ((dtong(laysn1t5(j)) = dtong(laysn1t5(luu))) And danh(laysn1t5(j)) < danh(laysn1t5(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n1t5(i)) > dtong(laysn1t5(luu))) Or ((dtong(n1t5(i)) = dtong(laysn1t5(luu))) And danh(n1t5(i)) > danh(laysn1t5(luu))) Then
            laysn1t5(luu) = n1t5(i)
        End If
    Next
    For i = 1 To ctieu(5)
        nv2(laysn1t5(i)) = 0
        nv3(laysn1t5(i)) = 0
        If (dcn1t5 > dtong(laysn1t5(i))) Then
            dcn1t5 = dtong(laysn1t5(i))
        End If
        sot5 = sot5 + 1
        trg5(sot5) = laysn1t5(i)
    Next
    ctieu(5) = 0
End If
End Sub
Sub locnv2trg1()
'Neu ctieu còn moi loc tiep
If ctieu(1) > 0 Then
    If (ctieu(1) >= sn2t1) Then
    For i = 1 To sn2t1
        nv3(n2t1(i)) = 0
        If (dcn2t1 > dtong(n2t1(i))) Then
            dcn2t1 = dtong(n2t1(i))
        End If
        sot1 = sot1 + 1
        trg1(sot1) = n2t1(i)
    Next
    ctieu(1) = ctieu(1) - sn2t1
    Else
    Dim laysn2t1(10000) As Integer
    For i = 1 To ctieu(1)
        laysn2t1(i) = (n2t1(i))
    Next
    For i = ctieu(1) + 1 To sn2t1
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(1)
            If (dtong(laysn2t1(j)) < dtong(laysn2t1(luu))) Or ((dtong(laysn2t1(j)) = dtong(laysn2t1(luu))) And danh(laysn2t1(j)) < danh(laysn2t1(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n2t1(i)) > dtong(laysn2t1(luu))) Or ((dtong(n2t1(i)) = dtong(laysn2t1(luu))) And danh(n2t1(i)) > danh(laysn2t1(luu))) Then
            laysn2t1(luu) = n2t1(i)
        End If
    Next
    For i = 1 To ctieu(1)
        nv3(laysn2t1(i)) = 0
        If (dcn2t1 > dtong(laysn2t1(i))) Then
            dcn2t1 = dtong(laysn2t1(i))
        End If
        sot1 = sot1 + 1
        trg1(sot1) = laysn2t1(i)
    Next
    ctieu(1) = 0
    End If
End If
End Sub
Sub locnv2trg2()
If ctieu(2) > 0 Then
    If (ctieu(2) >= sn2t2) Then
    For i = 1 To sn2t2
        nv3(n2t2(i)) = 0
        If (dcn2t2 > dtong(n2t2(i))) Then
            dcn2t2 = dtong(n2t2(i))
        End If
        sot2 = sot2 + 1
        trg2(sot2) = n2t2(i)
    Next
    ctieu(2) = ctieu(2) - sn2t2
    Else
    Dim laysn2t2(10000) As Integer
    For i = 1 To ctieu(2)
        laysn2t2(i) = (n2t2(i))
    Next
    For i = ctieu(2) + 1 To sn2t2
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(2)
            If (dtong(laysn2t2(j)) < dtong(laysn2t2(luu))) Or ((dtong(laysn2t2(j)) = dtong(laysn2t2(luu))) And danh(laysn2t2(j)) < danh(laysn2t2(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n2t2(i)) > dtong(laysn2t2(luu))) Or ((dtong(n2t2(i)) = dtong(laysn2t2(luu))) And danh(n2t2(i)) > danh(laysn2t2(luu))) Then
            laysn2t2(luu) = n2t2(i)
        End If
    Next
    For i = 1 To ctieu(2)
        nv3(laysn2t2(i)) = 0
        If (dcn2t2 > dtong(laysn2t2(i))) Then
            dcn2t2 = dtong(laysn2t2(i))
        End If
        sot2 = sot2 + 1
        trg2(sot2) = laysn2t2(i)
    Next
    ctieu(2) = 0
    End If
End If
End Sub
Sub locnv2trg3()
If ctieu(3) > 0 Then
    If (ctieu(3) >= sn2t3) Then
    For i = 1 To sn2t3
        nv3(n2t3(i)) = 0
        If (dcn2t3 > dtong(n2t3(i))) Then
            dcn2t3 = dtong(n2t3(i))
        End If
        sot3 = sot3 + 1
        trg3(sot3) = n2t3(i)
    Next
    ctieu(3) = ctieu(3) - sn2t3
    Else
    Dim laysn2t3(10000) As Integer
    For i = 1 To ctieu(3)
        laysn2t3(i) = (n2t3(i))
    Next
    For i = ctieu(3) + 1 To sn2t3
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(3)
            If (dtong(laysn2t3(j)) < dtong(laysn2t3(luu))) Or ((dtong(laysn2t3(j)) = dtong(laysn2t3(luu))) And danh(laysn2t3(j)) < danh(laysn2t3(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n2t3(i)) > dtong(laysn2t3(luu))) Or ((dtong(n2t3(i)) = dtong(laysn2t3(luu))) And danh(n2t3(i)) > danh(laysn2t3(luu))) Then
            laysn2t3(luu) = n2t3(i)
        End If
    Next
    For i = 1 To ctieu(3)
        nv3(laysn2t3(i)) = 0
        If (dcn2t3 > dtong(laysn2t3(i))) Then
            dcn2t3 = dtong(laysn2t3(i))
        End If
        sot3 = sot3 + 1
        trg3(sot3) = laysn2t3(i)
    Next
    ctieu(3) = 0
    End If
End If
End Sub
Sub locnv2trg4()
If ctieu(4) > 0 Then
    If (ctieu(4) >= sn2t4) Then
    For i = 1 To sn2t4
        nv3(n2t4(i)) = 0
        If (dcn2t4 > dtong(n2t4(i))) Then
            dcn2t4 = dtong(n2t4(i))
        End If
        sot4 = sot4 + 1
        trg4(sot4) = n2t4(i)
    Next
    ctieu(4) = ctieu(4) - sn2t4
    Else
    Dim laysn2t4(10000) As Integer
    For i = 1 To ctieu(4)
        laysn2t4(i) = (n2t4(i))
    Next
    For i = ctieu(4) + 1 To sn2t4
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(4)
            If (dtong(laysn2t4(j)) < dtong(laysn2t4(luu))) Or ((dtong(laysn2t4(j)) = dtong(laysn2t4(luu))) And danh(laysn2t4(j)) < danh(laysn2t4(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n2t4(i)) > dtong(laysn2t4(luu))) Or ((dtong(n2t4(i)) = dtong(laysn2t4(luu))) And danh(n2t4(i)) > danh(laysn2t4(luu))) Then
            laysn2t4(luu) = n2t4(i)
        End If
    Next
    For i = 1 To ctieu(4)
        nv3(laysn2t4(i)) = 0
        If (dcn2t4 > dtong(laysn2t4(i))) Then
            dcn2t4 = dtong(laysn2t4(i))
        End If
        sot4 = sot4 + 1
        trg4(sot4) = laysn2t4(i)
    Next
    ctieu(4) = 0
    End If
End If
End Sub
Sub locnv2trg5()
If ctieu(5) > 0 Then
    If (ctieu(5) >= sn2t5) Then
    For i = 1 To sn2t5
        nv3(n2t5(i)) = 0
        If (dcn2t5 > dtong(n2t5(i))) Then
            dcn2t5 = dtong(n2t5(i))
        End If
        sot5 = sot5 + 1
        trg5(sot5) = n2t5(i)
    Next
    ctieu(5) = ctieu(5) - sn2t5
    Else
    Dim laysn2t5(10000) As Integer
    For i = 1 To ctieu(5)
        laysn2t5(i) = (n2t5(i))
    Next
    For i = ctieu(5) + 1 To sn2t5
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(5)
            If (dtong(laysn2t5(j)) < dtong(laysn2t5(luu))) Or ((dtong(laysn2t5(j)) = dtong(laysn2t5(luu))) And danh(laysn2t5(j)) < danh(laysn2t5(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n2t5(i)) > dtong(laysn2t5(luu))) Or ((dtong(n2t5(i)) = dtong(laysn2t5(luu))) And danh(n2t5(i)) > danh(laysn2t5(luu))) Then
            laysn2t5(luu) = n2t5(i)
        End If
    Next
    For i = 1 To ctieu(5)
        nv3(laysn2t5(i)) = 0
        If (dcn2t5 > dtong(laysn2t5(i))) Then
            dcn2t5 = dtong(laysn2t5(i))
        End If
        sot5 = sot5 + 1
        trg5(sot5) = laysn2t5(i)
    Next
    ctieu(5) = 0
    End If
End If
End Sub
Sub locnv3trg1()
If ctieu(1) > 0 Then
    If (ctieu(1) >= sn3t1) Then
    For i = 1 To sn3t1
        If (dcn3t1 > dtong(n3t1(i))) Then
            dcn3t1 = dtong(n3t1(i))
        End If
        sot1 = sot1 + 1
        trg1(sot1) = n3t1(i)
    Next
    ctieu(1) = ctieu(1) - sn3t1
    Else
    Dim laysn3t1(10000) As Integer
    For i = 1 To ctieu(1)
        laysn3t1(i) = (n3t1(i))
    Next
    For i = ctieu(1) + 1 To sn3t1
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(1)
            If (dtong(laysn3t1(j)) < dtong(laysn3t1(luu))) Or ((dtong(laysn3t1(j)) = dtong(laysn3t1(luu))) And danh(laysn3t1(j)) < danh(laysn3t1(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n3t1(i)) > dtong(laysn3t1(luu))) Or ((dtong(n3t1(i)) = dtong(laysn3t1(luu))) And danh(n3t1(i)) > danh(laysn3t1(luu))) Then
            laysn3t1(luu) = n3t1(i)
        End If
    Next
    For i = 1 To ctieu(1)
        If (dcn3t1 > dtong(laysn3t1(i))) Then
            dcn3t1 = dtong(laysn3t1(i))
        End If
        sot1 = sot1 + 1
        trg1(sot1) = laysn3t1(i)
    Next
    ctieu(1) = 0
    End If
End If
End Sub
Sub locnv3trg2()
If ctieu(2) > 0 Then
    If (ctieu(2) >= sn3t2) Then
    For i = 1 To sn3t2
        If (dcn3t2 > dtong(n3t2(i))) Then
            dcn3t2 = dtong(n3t2(i))
        End If
        sot2 = sot2 + 1
        trg2(sot2) = n3t2(i)
    Next
    ctieu(2) = ctieu(2) - sn3t2
    Else
    Dim laysn3t2(10000) As Integer
    For i = 1 To ctieu(2)
        laysn3t2(i) = (n3t2(i))
    Next
    For i = ctieu(2) + 1 To sn3t2
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(2)
            If (dtong(laysn3t2(j)) < dtong(laysn3t2(luu))) Or ((dtong(laysn3t2(j)) = dtong(laysn3t2(luu))) And danh(laysn3t2(j)) < danh(laysn3t2(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n3t2(i)) > dtong(laysn3t2(luu))) Or ((dtong(n3t2(i)) = dtong(laysn3t2(luu))) And danh(n3t2(i)) > danh(laysn3t2(luu))) Then
            laysn3t2(luu) = n3t2(i)
        End If
    Next
    For i = 1 To ctieu(2)
        If (dcn3t2 > dtong(laysn3t2(i))) Then
            dcn3t2 = dtong(laysn3t2(i))
        End If
        sot2 = sot2 + 1
        trg2(sot2) = laysn3t2(i)
    Next
    ctieu(2) = 0
    End If
End If
End Sub
Sub locnv3trg3()
If ctieu(3) > 0 Then
    If (ctieu(3) >= sn3t3) Then
    For i = 1 To sn3t3
        If (dcn3t3 > dtong(n3t3(i))) Then
            dcn3t3 = dtong(n3t3(i))
        End If
        sot3 = sot3 + 1
        trg3(sot3) = n3t3(i)
    Next
    ctieu(3) = ctieu(3) - sn3t3
    Else
    Dim laysn3t3(10000) As Integer
    For i = 1 To ctieu(3)
        laysn3t3(i) = (n3t3(i))
    Next
    For i = ctieu(3) + 1 To sn3t3
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(3)
            If (dtong(laysn3t3(j)) < dtong(laysn3t3(luu))) Or ((dtong(laysn3t3(j)) = dtong(laysn3t3(luu))) And danh(laysn3t3(j)) < danh(laysn3t3(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n3t3(i)) > dtong(laysn3t3(luu))) Or ((dtong(n3t3(i)) = dtong(laysn3t3(luu))) And danh(n3t3(i)) > danh(laysn3t3(luu))) Then
            laysn3t3(luu) = n3t3(i)
        End If
    Next
    For i = 1 To ctieu(3)
        If (dcn3t3 > dtong(laysn3t3(i))) Then
            dcn3t3 = dtong(laysn3t3(i))
        End If
        sot3 = sot3 + 1
        trg3(sot3) = laysn3t3(i)
    Next
    ctieu(3) = 0
    End If
End If
End Sub
Sub locnv3trg4()
If ctieu(4) > 0 Then
    If (ctieu(4) >= sn3t4) Then
    For i = 1 To sn3t4
        If (dcn3t4 > dtong(n3t4(i))) Then
            dcn3t4 = dtong(n3t4(i))
        End If
        sot4 = sot4 + 1
        trg4(sot4) = n3t4(i)
    Next
    ctieu(4) = ctieu(4) - sn3t4
    Else
    Dim laysn3t4(10000) As Integer
    For i = 1 To ctieu(4)
        laysn3t4(i) = (n3t4(i))
    Next
    For i = ctieu(4) + 1 To sn3t4
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(4)
            If (dtong(laysn3t4(j)) < dtong(laysn3t4(luu))) Or ((dtong(laysn3t4(j)) = dtong(laysn3t4(luu))) And danh(laysn3t4(j)) < danh(laysn3t4(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n3t4(i)) > dtong(laysn3t4(luu))) Or ((dtong(n3t4(i)) = dtong(laysn3t4(luu))) And danh(n3t4(i)) > danh(laysn3t4(luu))) Then
            laysn3t4(luu) = n3t4(i)
        End If
    Next
    For i = 1 To ctieu(4)
        If (dcn3t4 > dtong(laysn3t4(i))) Then
            dcn3t4 = dtong(laysn3t4(i))
        End If
        sot4 = sot4 + 1
        trg4(sot4) = laysn3t4(i)
    Next
    ctieu(4) = 0
    End If
End If
End Sub
Sub locnv3trg5()
If ctieu(5) > 0 Then
    If (ctieu(5) >= sn3t5) Then
    For i = 1 To sn3t5
        If (dcn3t5 > dtong(n3t5(i))) Then
            dcn3t5 = dtong(n3t5(i))
        End If
        sot5 = sot5 + 1
        trg5(sot5) = n3t5(i)
    Next
    ctieu(5) = ctieu(5) - sn3t5
    Else
    Dim laysn3t5(10000) As Integer
    For i = 1 To ctieu(5)
        laysn3t5(i) = (n3t5(i))
    Next
    For i = ctieu(5) + 1 To sn3t5
        Dim luu As Integer
        luu = 1
        For j = 2 To ctieu(5)
            If (dtong(laysn3t5(j)) < dtong(laysn3t5(luu))) Or ((dtong(laysn3t5(j)) = dtong(laysn3t5(luu))) And danh(laysn3t5(j)) < danh(laysn3t5(luu))) Then
                luu = j
            End If
        Next
        If (dtong(n3t5(i)) > dtong(laysn3t5(luu))) Or ((dtong(n3t5(i)) = dtong(laysn3t5(luu))) And danh(n3t5(i)) > danh(laysn3t5(luu))) Then
            laysn3t5(luu) = n3t5(i)
        End If
    Next
    For i = 1 To ctieu(5)
        If (dcn3t5 > dtong(laysn3t5(i))) Then
            dcn3t5 = dtong(laysn3t5(i))
        End If
        sot5 = sot5 + 1
        trg5(sot5) = laysn3t5(i)
    Next
    ctieu(5) = 0
    End If
End If
End Sub
Sub locdiem()

'Gan diem chuan cua cac truong o cac nguyen vong bang 60
dcn1t1 = 60
dcn1t2 = 60
dcn1t3 = 60
dcn1t4 = 60
dcn1t5 = 60

dcn2t1 = 60
dcn2t2 = 60
dcn2t3 = 60
dcn2t4 = 60
dcn2t5 = 60

dcn3t1 = 60
dcn3t2 = 60
dcn3t3 = 60
dcn3t4 = 60
dcn3t5 = 60

'Gan so nguoi duoc chon o moi truong ban dau bang 0
sot1 = 0
sot2 = 0
sot3 = 0
sot4 = 0
sot5 = 0

'Gan vi tri cua nhung nguoi chon nguyen vong 1 o tung truong lan luot vào các mang n1t1 ->n1t5
Call gannv1

'Lan luot loc nguoi trung tuyen nguyen vong 1 o tung truong nhung nguoi nao duoc chon thi nv2 nv3 cho bang 0 de khong xet tiep
Call locnv1trg1
Call locnv1trg2
Call locnv1trg3
Call locnv1trg4
Call locnv1trg5

'Gan vi tri cua nhung nguoi chon nguyen vong 2 con lai o tung truong lan luot vào các mang n2t1 ->n2t5
Call gannv2

'Lan luot loc nguoi trung tuyen nguyen vong 2 o tung truong nhung nguoi nao duoc chon thi nv3 cho bang 0 de khong xet tiep
Call locnv2trg1
Call locnv2trg2
Call locnv2trg3
Call locnv2trg4
Call locnv2trg5

'Gan vi tri cua nhung nguoi chon nguyen vong 3 con lai o tung truong lan luot vào các mang n3t1 ->n3t5
Call gannv3

'Lan luot loc nguoi trung tuyen nguyen vong 3 o tung truong
Call locnv3trg1
Call locnv3trg2
Call locnv3trg3
Call locnv3trg4
Call locnv3trg5

End Sub
Sub Sapxepvitri()

Dim maxx As Double
Dim trunggian As Integer

'Sap xep vi tri hoc sinh cho truong 1 tu cao xuong thap
For i = 1 To sot1
    maxx = i
    'Tim hoc sinh cao nhat tu vi tri thu i den cuoi day
    For j = i To sot1
        If (dtong(trg1(j)) > dtong(trg1(maxx))) Or (dtong(trg1(j)) = dtong(trg1(maxx)) And danh(trg1(j)) > danh(trg1(maxx))) Then
            maxx = j
        End If
    Next
    'So nao lon nhat tu i ve cuoi day ma khac vi tri i thi doi cho cho vi tri i
    If maxx <> i Then
        trunggian = trg1(i)
        trg1(i) = trg1(maxx)
        trg1(maxx) = trunggian
    End If
Next

'tuong tu voi trg 2 den trg 5
For i = 1 To sot2
    maxx = i
    For j = i To sot2
        If (dtong(trg2(j)) > dtong(trg2(maxx))) Or (dtong(trg2(j)) = dtong(trg2(maxx)) And danh(trg2(j)) > danh(trg2(maxx))) Then
            maxx = j
        End If
    Next
    If maxx <> i Then
        trunggian = trg2(i)
        trg2(i) = trg2(maxx)
        trg2(maxx) = trunggian
    End If
Next

For i = 1 To sot3
    maxx = i
    For j = i To sot3
        If (dtong(trg3(j)) > dtong(trg3(maxx))) Or (dtong(trg3(j)) = dtong(trg3(maxx)) And danh(trg3(j)) > danh(trg3(maxx))) Then
            maxx = j
        End If
    Next
    If maxx <> i Then
        trunggian = trg3(i)
        trg3(i) = trg3(maxx)
        trg3(maxx) = trunggian
    End If
Next

For i = 1 To sot4
    maxx = i
    For j = i To sot4
        If (dtong(trg4(j)) > dtong(trg4(maxx))) Or (dtong(trg4(j)) = dtong(trg4(maxx)) And danh(trg4(j)) > danh(trg4(maxx))) Then
            maxx = j
        End If
    Next
    If maxx <> i Then
        trunggian = trg4(i)
        trg4(i) = trg4(maxx)
        trg4(maxx) = trunggian
    End If
Next

For i = 1 To sot5
    maxx = i
    For j = i To sot5
        If (dtong(trg5(j)) > dtong(trg5(maxx))) Or (dtong(trg5(j)) = dtong(trg5(maxx)) And danh(trg5(j)) > danh(trg5(maxx))) Then
            maxx = j
        End If
    Next
    If maxx <> i Then
        trunggian = trg5(i)
        trg5(i) = trg5(maxx)
        trg5(maxx) = trunggian
    End If
Next
End Sub
Sub Closefile()

Workbooks("input.xlsx").Close SaveChanges:=True
Workbooks("output.xlsx").Close SaveChanges:=True

End Sub

Sub main()
' Buoc 1: Mo file
Call Openfile

'Buoc2: Tim kiem va gan sbd, tên, các diem, nv1,nv2,nv3, chitieu vào mang de luu tru
Call ganvaomang

'Buoc3: Loc hoc sinh trung tuyen các trg lan luot vao mang trg1(), trg2(),...,trg(5) và luu lai diem chuan tai cac nguyen vong cua cac truong
Call locdiem

'Buoc4: Sap xep lai vi tri hoc sinh trong tung truong o cac mang trg1(), trg2(),...,trg(5)sao cho thoa man tu cao xuong thap
Call Sapxepvitri

'Buoc5: Tao file output de in ra diem chuan va danh sach trung tuyen cua tung truong
Call taoFile

'Buoc 6: Dong file input, output
Call Closefile

End Sub
