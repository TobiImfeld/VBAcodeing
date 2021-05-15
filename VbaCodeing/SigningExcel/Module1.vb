Module Module1

    Sub Main()
        SetSignature()

    End Sub

    Sub SetSignature()
        Dim objWB, objExcel, activeWB

        objExcel = CreateObject("Excel.Application")
        objWB = objExcel.Application.Workbooks.Open("C:\Temp\Test1.xlsx")
        If objWB.Equals(vbNull) Then
            MsgBox("Could not obpen Excel file")
        Else
            'objExcel.ActiveWorkbook.Signatures.AddNonVisibleSignature("4157e14126de03994c7e41c6d36d8ea7")
            'objExcel.ActiveWorkbook.Signatures.AddSignatureLine 'Works for manually set signature.
            activeWB = objExcel.ActiveWorkbook

            MsgBox("Excel has signature: " & activeWB.Signatures.AddNonVisibleSignature("4157e14126de03994c7e41c6d36d8ea7"))

            objWB.close
            objExcel.Quit
        End If

    End Sub

End Module
