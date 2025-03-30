using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhExcel.Components
{
    public class GH_Ex_Data_Set : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Data_Set class.
        /// </summary>
        public GH_Ex_Data_Set()
          : base("Set Excel Data", "XL Set",
              "Send data to Excel",
              "CORE", "Excel")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Worksheet", "Wks", "An Excel Worksheet Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddTextParameter("Data", "D", "A datatree of values. Each branch of the datatree will be used as a column and each item in the list will be populated to the rows in the column", GH_ParamAccess.tree);
            pManager[1].Optional = false;
            pManager.AddTextParameter("Starting Cell", "S", "The start Cell Address (ex. A1)", GH_ParamAccess.item);
            pManager[2].Optional = true;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated.", GH_ParamAccess.item, false);
            pManager[3].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Range", "Rng", "An Excel Range Object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool activate = false;
            if (DA.GetData(3, ref activate))
            {
                if (activate)
                {
                    IGH_Goo gooA = null;
                    if (!DA.GetData(0, ref gooA)) return;

                    if (!gooA.CastTo<ExWorksheet>(out ExWorksheet worksheet))
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Wbk input must be an Excel Worksheet Object");
                        return;
                    }

                    string address = "A1";
                    DA.GetData(2, ref address);

                    List<List<GH_String>> dataSet = new List<List<GH_String>>();
                    if (!DA.GetDataTree(1, out GH_Structure<GH_String> ghData)) return;

                    foreach (List<GH_String> data in ghData.Branches)
                    {
                        dataSet.Add(data);
                    }

                    ExRange range = worksheet.WriteData(dataSet, address);
                    DA.SetData(0, range);
                }
            }
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.Icons_Data_Set;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("2F5F04E9-B75C-4049-92AD-8B8AC81D92C3"); }
        }
    }
}