using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace GhExcel
{
    public class GhExcelInfo : GH_AssemblyInfo
    {
        public override string Name => "GhExcel";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "Grasshopper Plugin for Microsoft Excel interoperability";

        public override Guid Id => new Guid("50FB97AC-9EA9-4FD1-9905-5B263989B12D");

        //Return a string identifying you or your company.
        public override string AuthorName => "Thornton Tomasetti | CORE studio / AECtech";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "corestudio@thorntontomasetti.com";
    }
}