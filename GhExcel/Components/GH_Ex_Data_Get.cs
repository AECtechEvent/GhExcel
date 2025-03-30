using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhExcel.Components
{
    public class GH_Ex_Data_Get : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GH_Ex_Data_Get class.
        /// </summary>
        public GH_Ex_Data_Get()
          : base("Get Excel Data", "XL Get",
              "Read data from Excel",
              "CORE", "Excel")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Excel Range", "Rng", "An Excel Range Object", GH_ParamAccess.item);
            pManager[0].Optional = false;
            pManager.AddBooleanParameter("Activate", "_A", "If true, the component will be activated.", GH_ParamAccess.item, false);
            pManager[1].Optional = false;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Data", "D", "A datatree of values. Each branch of the datatree will be used as a column and each item in the list will be populated to the rows in the column", GH_ParamAccess.tree);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool activate = false;
            if (DA.GetData(1, ref activate))
            {
                if (activate)
                {
                    IGH_Goo gooA = null;
                    if (!DA.GetData(0, ref gooA)) return;

                    if (!gooA.CastTo<ExRange>(out ExRange range))
                    {
                        this.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Rng input must be an Excel Range Object");
                        return;
                    }


                    GH_Path path = new GH_Path();

                    if (this.Params.Input[0].VolatileData.PathCount > 1) path = this.Params.Input[0].VolatileData.get_Path(this.RunCount - 1);
                    path = path.AppendElement(this.RunCount - 1);

                    GH_Structure<GH_String> data = range.ReadData(path);

                    DA.SetDataTree(0, data);
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
                return Properties.Resources.Icons_Data_Get;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("AC6E6190-1CB6-4FCD-B83B-D3B729E2BEFD"); }
        }
    }
}