using SAPbobsCOM;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiniMAx
{
    internal class Eliminar
    {
        public static void eliminar()
        {
            //ConexionSAP.Close();

            if (ConexionSAP.Open())
            {
                SAPbobsCOM.UserTable oT = (SAPbobsCOM.UserTable)ConexionSAP.myCompany.UserTables.Item("AM_REABASTESIMIENTO");
                
                try
                {

                    


                    for(int x=1; x<=1280; x++)
                    {
                        try
                        {
                            oT.GetByKey(x.ToString());
                            oT.Remove();
                            Console.WriteLine(x);

                        }
                        catch (Exception)
                        {

                            
                        }
                    }
                    //oUserTable.Item("").
                    //if (oUserTable.Item("U_AM_REABASTECIMIENTO"))
                    //{
                    //    Console.WriteLine("SImon");
                    //    Console.ReadLine();
                    //}
                    //oUserTable.Item("U_AM_REABASTECIMIENTO");
                }
                catch (Exception EX)
                {

                }
                ConexionSAP.Close();
            }
        }
    }
}
