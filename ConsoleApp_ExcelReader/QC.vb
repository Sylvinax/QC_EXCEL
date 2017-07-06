Imports TDAPIOLELib

Public Class QC

    Function Get_Test() As Test
        Dim objTest As Test
        Dim objTestFactory As TestFactory
        Return objTest
    End Function

    Function Add_Test(path)









    End Function

    Sub WriteTestSteps(test As Test, source As String(,))
        Dim objDesignFactory As DesignStepFactory
        Dim objDesignStep As DesignStep
        objDesignFactory = test.DesignStepFactory
        objDesignStep = objDesignFactory.AddItem(vbNull)
        objDesignStep.StepName = "Step1"
        objDesignStep.StepDescription = "StepDescription"
        objDesignStep.StepExpectedResult = "StepExpectedResult"
        ' ArrayList SSS
    End Sub

End Class