using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiniMAx
{
    internal class Agregar
    {
        public static void agregar()
        {
            if (ConexionSAP.Open())
            {
                SAPbobsCOM.UserTable oT = (SAPbobsCOM.UserTable)ConexionSAP.myCompany.UserTables.Item("AM_REABASTESIMIENTO");
                string path = @"C:\Users\PRIDE BEAVER\Documents\Resb.xlsx";
                SLDocument sl = new SLDocument(path);

                int iRow = 2;
                int x;
                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                {
                    string Code = sl.GetCellValueAsString(iRow, 1);
                    string Name = sl.GetCellValueAsString(iRow, 2);
                    string Articulo = sl.GetCellValueAsString(iRow, 3);
                    int Maximo = sl.GetCellValueAsInt32(iRow, 4);
                    int Minimo = sl.GetCellValueAsInt32(iRow, 5);
                    string Almacen = sl.GetCellValueAsString(iRow, 6);
                    string Ubicacion = sl.GetCellValueAsString(iRow, 7);

                    
                    oT.Code = Code;
                    oT.Name = Name;
                    oT.UserFields.Fields.Item("U_Articulo").Value = Articulo;
                    oT.UserFields.Fields.Item("U_Minimo").Value = Minimo;
                    oT.UserFields.Fields.Item("U_Maximo").Value = Maximo;
                    oT.UserFields.Fields.Item("U_Almacen").Value = Almacen;
                    oT.UserFields.Fields.Item("U_Ubicacion").Value = Ubicacion;
                    x=oT.Add();
                    if (x!=0)
                    {
                        string y=ConexionSAP.myCompany.GetLastErrorDescription();
                    }

                    Console.WriteLine(iRow);
                    iRow++;
                }

            }
        }
    }
}
