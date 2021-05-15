Module Module1

    Sub Main()
        SetSignature()

    End Sub

    Sub SetSignature()
        Dim objWB, objExcel, message

        objExcel = CreateObject("Excel.Application")
        objWB = objExcel.WorkBooks.Open("C:\Temp\Test1.xlsx")
        If objWB.Equals(vbNull) Then
            MsgBox("objWB is null: ")
        Else
            objExcel.ActiveWorkbook.Signatures.AddNonVisibleSignature("4157e14126de03994c7e41c6d36d8ea7")
            objWB.close
            objExcel.Quit
        End If

    End Sub


End Module
