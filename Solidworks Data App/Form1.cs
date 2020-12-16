using System;
using System.Collections;
using System.Linq;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Solidworks_Data_App
{
    public partial class Form1
    {
        public Form1()
        {
            InitializeComponent();
            _Button1.Name = "Button1";
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            // Declaration of all global variable 
            SldWorks.SldWorks swApp;
            var errors = default(long);
            var warnings = default(long);
            SldWorks.ModelDoc2 swModel;
            object result;
            Array moiArray;
            string listBoxVal;

            // Connecting to the Solidworks Application 
            swApp = new SldWorks.SldWorks();

            // Storing the file in global swModel 
            SldWorks.ModelDoc2 localOpenDoc6() { int argErrors = Conversions.ToInteger(errors); int argWarnings = Conversions.ToInteger(warnings); var ret = swApp.OpenDoc6(@"E:\Tandemloop\Project\3d-local\Door.SLDPRT", (int)SwConst.swDocumentTypes_e.swDocPART, (int)SwConst.swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref argErrors, ref argWarnings); return ret; }

            swModel = localOpenDoc6();
            if (this.ComboBox1.SelectedIndex == 0)
            {
                result = COM(swModel);
                listBoxVal = "COM " + result.ToString();
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 1)
            {

                // Calling the Density function and storing the double value to the variable as Result. 

                result = Density(swModel);
                listBoxVal = "Density for the model is " + result.ToString();
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 2)
            {

                // Calling the Volume function and storing the double value to the variable as Result. 

                result = Volume(swModel);
                listBoxVal = "Volume for the model is " + result.ToString();
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 3)
            {

                // Calling the Mass function and storing the double value to the variable as Result. 

                result = Mass(swModel);
                listBoxVal = "Mass for the model is " + result.ToString();
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 4)
            {

                // Calling the surfaceArea function and storing the double value to the variable as Result. 

                result = surfaceArea(swModel);
                listBoxVal = "Surface Area for the model is " + result.ToString();
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 5)
            {
                moiArray = MOI(swModel);
                this.ListBox1.Items.Add("Moment of Inertia of the Model is");
                listBoxVal = "Lxx =" + moiArray(6).ToString() + " " + "m4";
                this.ListBox1.Items.Add(listBoxVal);
                listBoxVal = "Lyy =" + moiArray(7).ToString() + " " + "m4";
                this.ListBox1.Items.Add(listBoxVal);
                listBoxVal = "Lzz =" + moiArray(8).ToString() + " " + "m4";
                this.ListBox1.Items.Add(listBoxVal);
                listBoxVal = "Lxy =" + moiArray(9).ToString() + " " + "m4";
                this.ListBox1.Items.Add(listBoxVal);
                listBoxVal = "Lxz =" + moiArray(10).ToString() + " " + "m4";
                this.ListBox1.Items.Add(listBoxVal);
                listBoxVal = "Lyz =" + moiArray(11).ToString() + " " + "m4";
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 6)
            {

                // Calling the path function and storing the double value to the variable as Result. 

                result = path(swModel);
                listBoxVal = (string)Operators.AddObject("Path of the file is :", result);
                this.ListBox1.Items.Add(listBoxVal);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 7)
            {
                listBoxVal = "FIle Details are ";
                this.ListBox1.Items.Add(listBoxVal);
                fileDetails(swModel);
                this.ListBox1.Items.Add("");
                this.ListBox1.TopIndex = this.ListBox1.Items.Count - 1;
            }

            if (this.ComboBox1.SelectedIndex == 8)
            {
                configuration(swModel);
            }

            if (this.ComboBox1.SelectedIndex == 9)
            {
                colorDetails(swModel);
            }

            if (this.ComboBox1.SelectedIndex == 10)
            {
                scaleInfo(swModel);
            }
        }

        public object COM(object swModel)
        {
            SldWorks.MassProperty swMassProp;
            // Dim result As String
            // swModel.Visible = True
            swMassProp = (SldWorks.MassProperty)swModel.Extension.CreateMassProperty;
            return (swMassProp.CenterOfMass((object)0), swMassProp.CenterOfMass((object)1), swMassProp.CenterOfMass((object)2));
            Console.WriteLine(swMassProp.CenterOfMass((object)0));
            Console.WriteLine(swMassProp.CenterOfMass((object)1));
            Console.WriteLine(swMassProp.CenterOfMass((object)2));
        }

        public object Density(object swModel)
        {
            SldWorks.ModelDocExtension swModelExt;
            SldWorks.SelectionMgr swSelMgr;
            SldWorks.Component2 swComp;
            var nStatus = default(long);
            object vMassProp;
            long i;
            long nbrSelections;
            swModelExt = (SldWorks.ModelDocExtension)swModel.Extension;
            swSelMgr = (SldWorks.SelectionMgr)swModel.SelectionManager;
            nbrSelections = swSelMgr.GetSelectedObjectCount2(-1);
            nbrSelections = nbrSelections - 1;
            object localGetMassProperties2() { int argStatus = Conversions.ToInteger(nStatus); var ret = swModelExt.GetMassProperties2(1, out argStatus, true); return ret; }

            vMassProp = localGetMassProperties2();
            return Operators.DivideObject(vMassProp((object)5), vMassProp((object)3));
            Console.WriteLine(Operators.ConcatenateObject("Density = ", Operators.DivideObject(vMassProp((object)5), vMassProp((object)3))));
        }

        public object Volume(object swModel)
        {
            SldWorks.ModelDocExtension swModelExt;
            SldWorks.SelectionMgr swSelMgr;
            SldWorks.Component2 swComp;
            var nStatus = default(long);
            object vMassProp;
            long i;
            long nbrSelections;
            swModelExt = (SldWorks.ModelDocExtension)swModel.Extension;
            swSelMgr = (SldWorks.SelectionMgr)swModel.SelectionManager;
            nbrSelections = swSelMgr.GetSelectedObjectCount2(-1);
            nbrSelections = nbrSelections - 1;
            object localGetMassProperties2() { int argStatus = Conversions.ToInteger(nStatus); var ret = swModelExt.GetMassProperties2(1, out argStatus, true); return ret; }

            vMassProp = localGetMassProperties2();
            return vMassProp(3);
            Console.WriteLine(Operators.ConcatenateObject("Volume = ", vMassProp((object)3)));
        }

        public object Mass(object swModel)
        {
            SldWorks.ModelDocExtension swModelExt;
            SldWorks.SelectionMgr swSelMgr;
            SldWorks.Component2 swComp;
            var nStatus = default(long);
            object vMassProp;
            long i;
            long nbrSelections;
            swModelExt = (SldWorks.ModelDocExtension)swModel.Extension;
            swSelMgr = (SldWorks.SelectionMgr)swModel.SelectionManager;
            nbrSelections = swSelMgr.GetSelectedObjectCount2(-1);
            nbrSelections = nbrSelections - 1;
            object localGetMassProperties2() { int argStatus = Conversions.ToInteger(nStatus); var ret = swModelExt.GetMassProperties2(1, out argStatus, true); return ret; }

            vMassProp = localGetMassProperties2();
            return vMassProp(5);
            Console.WriteLine(Operators.ConcatenateObject("Volume = ", vMassProp((object)5)));
        }

        public object surfaceArea(object swModel)
        {
            SldWorks.ModelDocExtension swModelExt;
            SldWorks.SelectionMgr swSelMgr;
            SldWorks.Component2 swComp;
            var nStatus = default(long);
            object vMassProp;
            long i;
            long nbrSelections;
            swModelExt = (SldWorks.ModelDocExtension)swModel.Extension;
            swSelMgr = (SldWorks.SelectionMgr)swModel.SelectionManager;
            nbrSelections = swSelMgr.GetSelectedObjectCount2(-1);
            nbrSelections = nbrSelections - 1;
            object localGetMassProperties2() { int argStatus = Conversions.ToInteger(nStatus); var ret = swModelExt.GetMassProperties2(1, out argStatus, true); return ret; }

            vMassProp = localGetMassProperties2();
            Console.WriteLine(Operators.ConcatenateObject("Surface Area = ", vMassProp((object)4)));
            return vMassProp(4);
        }

        public Array MOI(object swModel)
        {
            SldWorks.ModelDocExtension swModelExt;
            SldWorks.SelectionMgr swSelMgr;
            SldWorks.Component2 swComp;
            Array moiArray;
            var nStatus = default(long);
            object vMassProp;
            long i;
            long nbrSelections;
            swModelExt = (SldWorks.ModelDocExtension)swModel.Extension;
            swSelMgr = (SldWorks.SelectionMgr)swModel.SelectionManager;
            nbrSelections = swSelMgr.GetSelectedObjectCount2(-1);
            nbrSelections = nbrSelections - 1;
            object localGetMassProperties2() { int argStatus = Conversions.ToInteger(nStatus); var ret = swModelExt.GetMassProperties2(1, out argStatus, true); return ret; }

            vMassProp = localGetMassProperties2();

            // moiArray(0) = vMassProp(6)
            // moiArray(1) = "Lyy = " & vMassProp(7)
            // moiArray(2) = "Lzz = " & vMassProp(8)
            // moiArray(3) = "Lxy = " & vMassProp(9)
            // moiArray(4) = "Lzx = " & vMassProp(10)
            // moiArray(5) = "Lyz = " & vMassProp(11)
            // Return ("Lxx = " & vMassProp(6), ("Lyy = " & vMassProp(7)),
            // ("Lzz = " & vMassProp(8)), ("Lxy = " & vMassProp(9)),
            // ("Lzx = " & vMassProp(10)), ("Lyz = " & vMassProp(11)))
            return (Array)vMassProp;
            // Console.WriteLine()
            // Console.WriteLine()
            // Console.WriteLine()
            // Console.WriteLine()
            // Console.WriteLine()

            // Return (vMassProp(6), vMassProp(7), vMassProp(8), vMassProp(9), vMassProp(10), vMassProp(11))


        }

        public object path(object swModel)
        {
            Console.WriteLine(swModel.GetPathName);
            return swModel.GetPathName;
        }

        public object fileDetails(object swModel)
        {
            int swSumInfoSavedBy = 5;
            int swSumInfoCreateDate = 6;
            int swSumInfoSaveDate = 7;
            string fileDetailsVal;
            fileDetailsVal = "Saved by =" + swModel.SummaryInfo(swSumInfoSavedBy).ToString();
            this.ListBox1.Items.Add(fileDetailsVal);
            fileDetailsVal = "Date created =" + swModel.SummaryInfo(swSumInfoCreateDate).ToString();
            this.ListBox1.Items.Add(fileDetailsVal);
            fileDetailsVal = "Date saved =" + swModel.SummaryInfo(swSumInfoSaveDate).ToString();
            this.ListBox1.Items.Add(fileDetailsVal);
            this.ListBox1.Items.Add("");
            return default;
        }

        public object configuration(object swModel)
        {
            string[] configNames;
            string configName;
            SldWorks.Configuration swConfig;
            long i;


            // Get and print active model path and name        
            // Get and print name of active configuration
            swConfig = (SldWorks.Configuration)swModel.GetActiveConfiguration;
            Console.WriteLine("Name of active configuration = " + swConfig.Name);


            // Get and print names of all configurations
            configNames = (string[])swModel.GetConfigurationNames;
            var loopTo = (long)Information.UBound(configNames);
            for (i = 0; i <= loopTo; i++)
            {
                configName = configNames[Conversions.ToInteger(i)];
                swConfig = (SldWorks.Configuration)swModel.GetConfigurationByName(configName);
                Console.WriteLine("  Name of configuration(" + i + ") = " + configName);
                Console.WriteLine(" Use alternate name in BOM  = " + swConfig.UseAlternateNameInBOM);
                Console.WriteLine(" Alternate name  = " + swConfig.AlternateName);
                Console.WriteLine(" Comment = " + swConfig.Comment);
            }

            return default;
        }

        public object colorDetails(object swModel)
        {
            SldWorks.Body2 swBody;
            object vBodyArr;
            object vMatProp;
            long i;
            for (i = 0; i <= 5; i++)
            {
                vBodyArr = swModel.GetBodies2(i, false);
                if (!Information.IsNothing(vBodyArr))
                {
                    foreach (var vBody in (IEnumerable)vBodyArr)
                    {
                        swBody = (SldWorks.Body2)vBody;
                        vMatProp = swBody.MaterialPropertyValues2;
                        if (!Information.IsNothing(vMatProp))
                        {
                            Console.WriteLine("abc");
                            Console.WriteLine(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("RGB = [", Operators.MultiplyObject(vMatProp((object)0), 255.0D)), ", "), Operators.MultiplyObject(vMatProp((object)1), 255.0D)), ", "), Operators.MultiplyObject(vMatProp((object)2), 255.0D)), "]"));
                        }
                    }
                }
            }

            Interaction.MsgBox("Done");
            return default;
        }

        public object scaleInfo(object swModel)
        {
            SldWorks.ModelDocExtension swModelDocExt;
            SldWorks.ModelView swModView;
            object swModViews;
            long index;
            long Count;
            swModelDocExt = (SldWorks.ModelDocExtension)swModel.Extension;
            // Get model views
            swModViews = swModelDocExt.GetModelViews();
            // Get number of model views
            Count = swModelDocExt.GetModelViewCount();
            Console.WriteLine("Number of model views: " + Count);
            // Get the scale factor of each model view
            var loopTo = (long)Information.UBound((Array)swModViews);
            for (index = Information.LBound((Array)swModViews); index <= loopTo; index++)
            {
                swModView = (SldWorks.ModelView)swModViews(index);
                Console.WriteLine("Scale of this model view is: " + swModView.Scale2);
            }

            return default;
        }
    }
}