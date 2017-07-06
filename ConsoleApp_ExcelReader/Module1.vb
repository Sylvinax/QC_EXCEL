Imports Microsoft.Office.Interop.Excel



Module Module1

	Sub Main()
        ' Dim test As New ExcelClass
        '"C:\Users\Administrator\Desktop\TestCase.xlsx"

        'MsgBox(test.OpenExcel("C:\Users\Administrator\Desktop\OneKeySnippet.xlsx").Range("A1").Text)
        'test.OpenExcel("C:\Users\Administrator\Desktop\OneKeySnippet.xlsx").Quit()





        'Dim xx = New ExcelClass

        'Dim y As Workbooks
        'Dim j As Sheets
        'Dim z As Workbook

        'y = xx.OpenExcel("C:\Users\Administrator\Desktop", True)


        'For Each z In y
        '    MsgBox(z.Sheets(1).UsedRange.Rows.count)
        '    MsgBox(z.Sheets(1).UsedRange.Columns.count)


        'Next
        'y.Application.Quit()



        Dim xx = New ExcelClass

        Dim y As Dictionary(Of String, ArrayList)
        y = xx.GetTestStep("C:\Users\Administrator\Desktop\testdata")

        Dim z = y.Item("1").Item(0).Item("StepDescription")
        MsgBox(z)
        z = y.Item("2").Item(1).Item("StepDescription")
        MsgBox(z)
        z = y.Item("3").Item(2).Item("StepDescription")
        MsgBox(z)


        'Dim testSteps As New ArrayList
        'Dim teststep As New Dictionary(Of String, String)
        'teststep.Add("StepName", "1")
        'teststep.Add("StepDescription", "2")
        'teststep.Add("StepExpectedResult", "3")
        'testSteps.Add(teststep)
        ''teststep.Clear()

        'MsgBox(1)






        'objWorkSheet = Workbook.Sheets(1)

        ' MsgBox(objWorkSheet.Range("A1").Text)





        'objExcel.Quit()



        'Dim x As ArrayList
        ' x.Add({{1, 2, 3}, {4, 5, 6}})

        'Dim y = x.Item(0)

        ' MsgBox(y(0, 0))



    End Sub

End Module
