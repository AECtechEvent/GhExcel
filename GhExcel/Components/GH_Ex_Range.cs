using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhExcel.Components
{
    public class GH_Ex_Range : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Range class.
        /// </summary>
        public GH_Ex_Range()
          : base("Excel Range", "XL Rng",
              "Get a Range from a Worksheet",
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
            pManager.AddTextParameter("Starting Cell", "S", "The start Cell Address (ex. A1)", GH_ParamAccess.item);
            pManager[1].Optional = false;
            pManager.AddTextParameter("End Cell", "E", "The end Cell Address (ex. A1)", GH_ParamAccess.item);
            pManager[2].Optional = false;
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
            IGH_Goo gooA = null;
            if (!DA.GetData(0, ref gooA)) return;

            if (!gooA.CastTo<ExWorksheet>(out ExWorksheet worksheet))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Wbk input must be an Excel Worksheet Object");
                return;
            }

            string min = "A1";
            if (!DA.GetData(1, ref min)) return;
            
            string max = "B2";
            if (!DA.GetData(2, ref max)) return;

            ExRange range = worksheet.GetRange(min, max);

            DA.SetData(0, range);
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
                return Properties.Resources.Icons_Range;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("A16145B2-1651-46AB-A2CF-97D9AFF4BE77"); }
        }
    }
}