Imports TDAPIOLELib

Public Class QC

    Function Get_Test() As Test
        Dim objTest As Test
        Dim objTestFactory As TestFactory
        Return objTest
    End Function

    Sub WriteTestSteps(test As Test, source As String(,))
        Dim objDesignFactory As DesignStepFactory
        Dim objDesignStep As DesignStep
        objDesignFactory = test.DesignStepFactory
        objDesignStep = objDesignFactory.AddItem("Step1")
        objDesignStep.StepDescription = "StepDescription"
        objDesignStep.StepExpectedResult = "StepExpectedResult"
        ' ArrayList SS
    End Sub

End Class