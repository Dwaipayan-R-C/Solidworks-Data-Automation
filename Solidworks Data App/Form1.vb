Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Declaration of all global variable 
        Dim swApp As SldWorks.SldWorks
        Dim errors As Long
        Dim warnings As Long
        Dim swModel As SldWorks.ModelDoc2
        Dim result As Object
        Dim moiArray As Array
        Dim listBoxVal As String

        'Connecting to the Solidworks Application 
        swApp = New SldWorks.SldWorks()

        'Storing the file in global swModel 
        swModel = swApp.OpenDoc6("E:\Tandemloop\Project\3d-local\Door.SLDPRT", SwConst.swDocumentTypes_e.swDocPART, SwConst.swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)

        If Me.ComboBox1.SelectedIndex = 0 Then
            result = COM(swModel)
            listBoxVal = "COM " + result.ToString

            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If
        If Me.ComboBox1.SelectedIndex = 1 Then

            'Calling the Density function and storing the double value to the variable as Result. 

            result = Density(swModel)
            listBoxVal = "Density for the model is " + result.ToString
            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If

        If Me.ComboBox1.SelectedIndex = 2 Then

            'Calling the Volume function and storing the double value to the variable as Result. 

            result = Volume(swModel)
            listBoxVal = "Volume for the model is " + result.ToString

            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If

        If Me.ComboBox1.SelectedIndex = 3 Then

            'Calling the Mass function and storing the double value to the variable as Result. 

            result = Mass(swModel)
            listBoxVal = "Mass for the model is " + result.ToString

            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If

        If Me.ComboBox1.SelectedIndex = 4 Then

            'Calling the surfaceArea function and storing the double value to the variable as Result. 

            result = surfaceArea(swModel)
            listBoxVal = "Surface Area for the model is " + result.ToString

            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If

        If Me.ComboBox1.SelectedIndex = 5 Then

            moiArray = MOI(swModel)
            ListBox1.Items.Add("Moment of Inertia of the Model is")
            listBoxVal = "Lxx =" + moiArray(6).ToString + " " + "m4"
            ListBox1.Items.Add(listBoxVal)
            listBoxVal = "Lyy =" + moiArray(7).ToString + " " + "m4"
            ListBox1.Items.Add(listBoxVal)
            listBoxVal = "Lzz =" + moiArray(8).ToString + " " + "m4"
            ListBox1.Items.Add(listBoxVal)
            listBoxVal = "Lxy =" + moiArray(9).ToString + " " + "m4"
            ListBox1.Items.Add(listBoxVal)
            listBoxVal = "Lxz =" + moiArray(10).ToString + " " + "m4"
            ListBox1.Items.Add(listBoxVal)
            listBoxVal = "Lyz =" + moiArray(11).ToString + " " + "m4"
            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If

        If Me.ComboBox1.SelectedIndex = 6 Then

            'Calling the path function and storing the double value to the variable as Result. 

            result = path(swModel)
            listBoxVal = "Path of the file is :" + result

            ListBox1.Items.Add(listBoxVal)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If

        If Me.ComboBox1.SelectedIndex = 7 Then
            listBoxVal = "FIle Details are "

            ListBox1.Items.Add(listBoxVal)
            fileDetails(swModel)
            ListBox1.Items.Add("")
            ListBox1.TopIndex = ListBox1.Items.Count - 1

        End If

        If Me.ComboBox1.SelectedIndex = 8 Then
            configuration(swModel)
        End If

        If Me.ComboBox1.SelectedIndex = 9 Then
            colorDetails(swModel)
        End If

        If Me.ComboBox1.SelectedIndex = 10 Then
            scaleInfo(swModel)
        End If

    End Sub

    Function COM(ByVal swModel As Object) As Object

        Dim swMassProp As SldWorks.MassProperty
        'Dim result As String
        'swModel.Visible = True
        swMassProp = swModel.Extension.CreateMassProperty
        Return (swMassProp.CenterOfMass(0), swMassProp.CenterOfMass(1), swMassProp.CenterOfMass(2))
        Console.WriteLine(swMassProp.CenterOfMass(0))
        Console.WriteLine(swMassProp.CenterOfMass(1))
        Console.WriteLine(swMassProp.CenterOfMass(2))
    End Function

    Function Density(ByVal swModel As Object) As Object

        Dim swModelExt As SldWorks.ModelDocExtension
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swComp As SldWorks.Component2


        Dim nStatus As Long
        Dim vMassProp As Object
        Dim i As Long
        Dim nbrSelections As Long

        swModelExt = swModel.Extension
        swSelMgr = swModel.SelectionManager

        nbrSelections = swSelMgr.GetSelectedObjectCount2(-1)

        nbrSelections = nbrSelections - 1
        vMassProp = swModelExt.GetMassProperties2(1, nStatus, True)
        Return vMassProp(5) / vMassProp(3)
        Console.WriteLine("Density = " & vMassProp(5) / vMassProp(3))


    End Function

    Function Volume(ByVal swModel As Object) As Object

        Dim swModelExt As SldWorks.ModelDocExtension
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swComp As SldWorks.Component2


        Dim nStatus As Long
        Dim vMassProp As Object
        Dim i As Long
        Dim nbrSelections As Long

        swModelExt = swModel.Extension
        swSelMgr = swModel.SelectionManager

        nbrSelections = swSelMgr.GetSelectedObjectCount2(-1)

        nbrSelections = nbrSelections - 1
        vMassProp = swModelExt.GetMassProperties2(1, nStatus, True)
        Return vMassProp(3)
        Console.WriteLine("Volume = " & vMassProp(3))


    End Function

    Function Mass(ByVal swModel As Object) As Object

        Dim swModelExt As SldWorks.ModelDocExtension
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swComp As SldWorks.Component2


        Dim nStatus As Long
        Dim vMassProp As Object
        Dim i As Long
        Dim nbrSelections As Long

        swModelExt = swModel.Extension
        swSelMgr = swModel.SelectionManager

        nbrSelections = swSelMgr.GetSelectedObjectCount2(-1)

        nbrSelections = nbrSelections - 1
        vMassProp = swModelExt.GetMassProperties2(1, nStatus, True)
        Return vMassProp(5)
        Console.WriteLine("Volume = " & vMassProp(5))


    End Function

    Function surfaceArea(ByVal swModel As Object) As Object

        Dim swModelExt As SldWorks.ModelDocExtension
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swComp As SldWorks.Component2


        Dim nStatus As Long
        Dim vMassProp As Object
        Dim i As Long
        Dim nbrSelections As Long

        swModelExt = swModel.Extension
        swSelMgr = swModel.SelectionManager

        nbrSelections = swSelMgr.GetSelectedObjectCount2(-1)

        nbrSelections = nbrSelections - 1
        vMassProp = swModelExt.GetMassProperties2(1, nStatus, True)
        Console.WriteLine("Surface Area = " & vMassProp(4))
        Return vMassProp(4)


    End Function

    Function MOI(ByVal swModel As Object) As Array

        Dim swModelExt As SldWorks.ModelDocExtension
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swComp As SldWorks.Component2
        Dim moiArray As Array


        Dim nStatus As Long
        Dim vMassProp As Object
        Dim i As Long
        Dim nbrSelections As Long

        swModelExt = swModel.Extension
        swSelMgr = swModel.SelectionManager

        nbrSelections = swSelMgr.GetSelectedObjectCount2(-1)

        nbrSelections = nbrSelections - 1
        vMassProp = swModelExt.GetMassProperties2(1, nStatus, True)

        'moiArray(0) = vMassProp(6)
        'moiArray(1) = "Lyy = " & vMassProp(7)
        'moiArray(2) = "Lzz = " & vMassProp(8)
        'moiArray(3) = "Lxy = " & vMassProp(9)
        'moiArray(4) = "Lzx = " & vMassProp(10)
        'moiArray(5) = "Lyz = " & vMassProp(11)
        'Return ("Lxx = " & vMassProp(6), ("Lyy = " & vMassProp(7)),
        '    ("Lzz = " & vMassProp(8)), ("Lxy = " & vMassProp(9)),
        '    ("Lzx = " & vMassProp(10)), ("Lyz = " & vMassProp(11)))
        Return vMassProp
        'Console.WriteLine()
        'Console.WriteLine()
        'Console.WriteLine()
        'Console.WriteLine()
        'Console.WriteLine()

        'Return (vMassProp(6), vMassProp(7), vMassProp(8), vMassProp(9), vMassProp(10), vMassProp(11))


    End Function

    Function path(ByVal swModel As Object) As Object
        Console.WriteLine(swModel.GetPathName)
        Return swModel.GetPathName

    End Function

    Function fileDetails(ByVal swModel As Object)
        Dim swSumInfoSavedBy As Integer = 5
        Dim swSumInfoCreateDate As Integer = 6
        Dim swSumInfoSaveDate As Integer = 7
        Dim fileDetailsVal As String
        fileDetailsVal = "Saved by =" + swModel.SummaryInfo(swSumInfoSavedBy).ToString

        ListBox1.Items.Add(fileDetailsVal)
        fileDetailsVal = "Date created =" + swModel.SummaryInfo(swSumInfoCreateDate).ToString

        ListBox1.Items.Add(fileDetailsVal)
        fileDetailsVal = "Date saved =" + swModel.SummaryInfo(swSumInfoSaveDate).ToString

        ListBox1.Items.Add(fileDetailsVal)
        ListBox1.Items.Add("")


    End Function

    Function configuration(ByVal swModel As Object)

        Dim configNames() As String
        Dim configName As String
        Dim swConfig As SldWorks.Configuration
        Dim i As Long


        'Get and print active model path and name        
        'Get and print name of active configuration
        swConfig = swModel.GetActiveConfiguration
        Console.WriteLine("Name of active configuration = " & swConfig.Name)


        'Get and print names of all configurations
        configNames = swModel.GetConfigurationNames
        For i = 0 To UBound(configNames)
            configName = configNames(i)
            swConfig = swModel.GetConfigurationByName(configName)
            Console.WriteLine("  Name of configuration(" & i & ") = " & configName)
            Console.WriteLine(" Use alternate name in BOM  = " & swConfig.UseAlternateNameInBOM)
            Console.WriteLine(" Alternate name  = " + swConfig.AlternateName)
            Console.WriteLine(" Comment = " + swConfig.Comment)
        Next i
    End Function

    Function colorDetails(ByVal swModel As Object)

        Dim swBody As SldWorks.Body2
        Dim vBodyArr As Object
        Dim vBody As Object
        Dim vMatProp As Object
        Dim i As Long
        For i = 0 To 5
            vBodyArr = swModel.GetBodies2(i, False)

            If Not IsNothing(vBodyArr) Then

                For Each vBody In vBodyArr
                    swBody = vBody
                    vMatProp = swBody.MaterialPropertyValues2

                    If Not IsNothing(vMatProp) Then
                        Console.WriteLine("abc")
                        Console.WriteLine("RGB = [" & vMatProp(0) * 255.0# & ", " & vMatProp(1) * 255.0# & ", " & vMatProp(2) * 255.0# & "]")
                    End If
                Next

            End If

        Next i
        MsgBox("Done")
    End Function

    Function scaleInfo(ByVal swModel As Object)
        Dim swModelDocExt As SldWorks.ModelDocExtension
        Dim swModView As SldWorks.ModelView
        Dim swModViews As Object
        Dim index As Long
        Dim Count As Long

        swModelDocExt = swModel.Extension
        ' Get model views
        swModViews = swModelDocExt.GetModelViews
        ' Get number of model views
        Count = swModelDocExt.GetModelViewCount
        Console.WriteLine("Number of model views: " & Count)
        ' Get the scale factor of each model view
        For index = LBound(swModViews) To UBound(swModViews)
            swModView = swModViews(index)
            Console.WriteLine("Scale of this model view is: " & swModView.Scale2)
        Next index
    End Function
End Class
