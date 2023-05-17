using Core.Models;

namespace Core.Utilities
{
    public delegate void UserFormClosedDone(object sender);
    public delegate void SubFormCreated<T>(object sender, FormArgs formArgs) where T : new();
    public delegate void CFLValueSelected(object sender, SAPbouiCOM.SBOItemEventArg pv, string id, string[,] table, int rowSelected);
}
