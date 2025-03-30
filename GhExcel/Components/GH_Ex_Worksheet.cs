using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhExcel.Components
{
    public class GH_Ex_Worksheet : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Worksheet class.
        /// </summary>
        public GH_Ex_Worksheet()
          : base("Excel Worksheet", "XL Wks",
              "Get a Worksheet by name or the active Worksheet from a Workbook",
              "CORE", "Excel")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Workbook", "Wkb", "An Excel Workbook Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddGenericParameter("Name", "*N", "OPTIONAL: The name of a new or existing worksheet in the current Workbook", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Worksheet", "Wks", "An Excel Worksheet Object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.CastTo<ExWorkbook>(out ExWorkbook workbook))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Wbk input must be an Excel Workbook Object");
                return;
            }

            string name = string.Empty;

            if (DA.GetData(1, ref name))
            {
                DA.SetData(0, workbook.GetWorksheetByName(name));
            }
            else
            {
                DA.SetData(0, workbook.GetActiveWorksheet());
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
                return Properties.Resources.Icons_Worksheet;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("44D0A7E9-5C0E-4A72-A474-B09DCD1CC7FD"); }
        }
    }
}