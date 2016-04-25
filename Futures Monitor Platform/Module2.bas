Attribute VB_Name = "Module2"
Sub renew_globe()
    
    Dim i As Integer
    
    With Workbooks("Monitor Platform 201503 for APP.xlsm").Worksheets("犀利環球")
        
        '黃金
        .Cells(3, 5) = "=[犀利環球DDE.xls]DDE!$M$25"
        .Cells(3, 6) = "=[犀利環球DDE.xls]DDE!$N$25"
        .Cells(3, 7) = "=[犀利環球DDE.xls]DDE!$O$25"
        .Cells(3, 8) = "=[犀利環球DDE.xls]DDE!$G$25"
        .Cells(3, 9) = "=[犀利環球DDE.xls]DDE!$H$25"
        .Cells(3, 11) = "=[犀利環球DDE.xls]DDE!$J$25"
        .Cells(3, 12) = "=[犀利環球DDE.xls]DDE!$K$25"
        
        '澳幣 英鎊 加幣 歐元
        For i = 4 To 7
            .Cells(i, 5) = "=[犀利環球DDE.xls]DDE!$M$" & i + 25 & " *10000"
            .Cells(i, 6) = "=[犀利環球DDE.xls]DDE!$N$" & i + 25 & " *10000"
            .Cells(i, 7) = "=[犀利環球DDE.xls]DDE!$O$" & i + 25 & " *10000"
            .Cells(i, 8) = "=[犀利環球DDE.xls]DDE!$G$" & i + 25 & " *10000"
            .Cells(i, 9) = "=[犀利環球DDE.xls]DDE!$H$" & i + 25 & " *10000"
            .Cells(i, 11) = "=[犀利環球DDE.xls]DDE!$J$" & i + 25
            .Cells(i, 12) = "=[犀利環球DDE.xls]DDE!$K$" & i + 25
        Next i
        
        '日元
        .Cells(8, 5) = "=[犀利環球DDE.xls]DDE!$M$33 *1000000"
        .Cells(8, 6) = "=[犀利環球DDE.xls]DDE!$N$33 *1000000"
        .Cells(8, 7) = "=[犀利環球DDE.xls]DDE!$O$33 *1000000"
        .Cells(8, 8) = "=[犀利環球DDE.xls]DDE!$G$33 *1000000"
        .Cells(8, 9) = "=[犀利環球DDE.xls]DDE!$H$33 *1000000"
        .Cells(8, 11) = "=[犀利環球DDE.xls]DDE!$J$33"
        .Cells(8, 12) = "=[犀利環球DDE.xls]DDE!$K$33"
        
        '瑞郎
        .Cells(9, 5) = "=[犀利環球DDE.xls]DDE!$M$34 *10000"
        .Cells(9, 6) = "=[犀利環球DDE.xls]DDE!$N$34 *10000"
        .Cells(9, 7) = "=[犀利環球DDE.xls]DDE!$O$34 *10000"
        .Cells(9, 8) = "=[犀利環球DDE.xls]DDE!$G$34 *10000"
        .Cells(9, 9) = "=[犀利環球DDE.xls]DDE!$H$34 *10000"
        .Cells(9, 11) = "=[犀利環球DDE.xls]DDE!$J$34"
        .Cells(9, 12) = "=[犀利環球DDE.xls]DDE!$K$34"
        
        '輕原油
        .Cells(10, 5) = "=[犀利環球DDE.xls]DDE!$M$38"
        .Cells(10, 6) = "=[犀利環球DDE.xls]DDE!$N$38"
        .Cells(10, 7) = "=[犀利環球DDE.xls]DDE!$O$38"
        .Cells(10, 8) = "=[犀利環球DDE.xls]DDE!$G$38"
        .Cells(10, 9) = "=[犀利環球DDE.xls]DDE!$H$38"
        .Cells(10, 11) = "=[犀利環球DDE.xls]DDE!$J$38"
        .Cells(10, 12) = "=[犀利環球DDE.xls]DDE!$K$38"
        
        '美元指數
        .Cells(11, 5) = "=[犀利環球DDE.xls]DDE!$M$71"
        .Cells(11, 6) = "=[犀利環球DDE.xls]DDE!$N$71"
        .Cells(11, 7) = "=[犀利環球DDE.xls]DDE!$O$71"
        .Cells(11, 8) = "=[犀利環球DDE.xls]DDE!$G$71"
        .Cells(11, 9) = "=[犀利環球DDE.xls]DDE!$H$71"
        .Cells(11, 11) = "=[犀利環球DDE.xls]DDE!$J$71"
        .Cells(11, 12) = "=[犀利環球DDE.xls]DDE!$K$71"

    End With

End Sub
