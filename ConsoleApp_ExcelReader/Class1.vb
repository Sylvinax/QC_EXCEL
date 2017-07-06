Imports Microsoft.Office.Interop.Excel

Public Class ExcelClass


    Function OpenExcel(path As String, visible As Boolean) As Workbooks

        Dim objExcel As New Application
        objExcel.Visible = visible
        Dim strFileList As String() = IO.Directory.GetFiles(path)
        Dim intFilecount As Integer = IO.Directory.GetFiles(path).Count
        For Each fileName In strFileList
            If InStr(fileName, ".xls") > 0 Then
                objExcel.Workbooks.Open(fileName)

            End If
        Next
        Return objExcel.Workbooks


    End Function


    'Get the TestStep from each Excel
    Function GetTestStep(path As String, Optional startRow As Integer = 2, Optional neededColumn As Integer = 3) As Dictionary(Of String, ArrayList)

        Dim objWorkbook As Workbook
        Dim objWorkbooks As Workbooks
        'Dim objSheets As Sheets
        Dim NeededDatas As New Dictionary(Of String, ArrayList)
        Dim testSteps As New ArrayList

        Dim intUsedRange As Integer


        objWorkbooks = OpenExcel(path, True)

        For Each objWorkbook In objWorkbooks
            intUsedRange = objWorkbook.Sheets(1).UsedRange.Rows.Count

            For i = startRow To intUsedRange

                Dim strStep As String = objWorkbook.Sheets(1).Cells(i, 1).Text
                Dim strDescription As String = objWorkbook.Sheets(1).Cells(i, 2).Text
                Dim strExpected As String = objWorkbook.Sheets(1).Cells(i, 3).Text
                Dim teststep As New Dictionary(Of String, String)
                teststep.Add("StepName", strStep)
                teststep.Add("StepDescription", strDescription)
                teststep.Add("StepExpectedResult", strExpected)
                testSteps.Add(teststep)

            Next


            NeededDatas.Add(IO.Path.GetFileNameWithoutExtension(objWorkbook.Name), testSteps)


        Next objWorkbook

        objWorkbooks.Application.Quit()

        Return NeededDatas



    End Function



    Function GetExcelFileName(path) As ArrayList
        Dim strFileList As String() = IO.Directory.GetFiles(path)
        Dim intFilecount As Integer = IO.Directory.GetFiles(path).Count
        Dim filenames As New ArrayList
        For Each fileName In strFileList
            If InStr(fileName, ".xls") > 0 Then
                filenames.Add(IO.Path.GetFileNameWithoutExtension(fileName))
            End If
        Next
        Return filenames
    End Function
End Class
