using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhExcel.Components
{
    public class GH_Ex_Workbook : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Workbook class.
        /// </summary>
        public GH_Ex_Workbook()
          : base("Excel Workbook", "XL Wbk",
              "Get a Workbook from a file path or the active Workbook",
              "CORE", "Excel")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Application", "App", "An Excel Application Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddGenericParameter("Filepath", "*F", "OPTIONAL: The full filepath to an Excel Workbook (File)", GH_ParamAccess.item);
            pManager[1].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Workbook", "Wkb", "An Excel Workbook Object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo gooA = null;
            if(!DA.GetData(0, ref gooA))return;

            if (!gooA.CastTo<ExApp>(out ExApp app))
            {
                this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "App input must be an Excel Application Object");
                return;
            }

            string filepath = string.Empty;

            if(DA.GetData(1,ref filepath))
            {
                if (!System.IO.File.Exists(filepath))
                {
                    this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The specified Filepath does not exist");
                    return;
                }
                DA.SetData(0, app.LoadWorkbook(filepath));
            }
            else
            {
                DA.SetData(0, app.GetActiveWorkbook());
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
                return Properties.Resources.Icons_Workbook;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("49168CC9-54AF-4845-ABBB-0F3C3BCD9098"); }
        }
    }
}