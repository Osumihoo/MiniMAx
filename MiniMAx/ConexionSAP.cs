using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiniMAx
{
    internal class ConexionSAP
    {
        public static SAPbobsCOM.Company myCompany = null;

        public static bool Open()
        {
            try
            {
                bool Respuesta = false;

                myCompany = new SAPbobsCOM.Company();
                myCompany.Server = "NDB@hanab1:30013";
                myCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                myCompany.CompanyDB = "SAPTESTAF";
                myCompany.UserName = "Desarrollo";
                myCompany.Password = "123456";
                myCompany.DbUserName = "SYSTEM"; //ConfigurationManager.AppSettings["bdUser"];
                myCompany.DbPassword = "B1AdminH2";
                myCompany.language = BoSuppLangs.ln_Spanish_La;

                //Globals.Ocompany = new SAPbobsCOM.Company();
                //Globals.Ocompany.Server = "NDB@hanab1:30013";
                //Globals.Ocompany.DbServerType = BoDataServerTypes.dst_HANADB;
                //Globals.Ocompany.CompanyDB = company;
                //Globals.Ocompany.UserName = usersap;//"manager";//"Sistema03";
                //Globals.Ocompany.Password = passsap; //"B1AdminH2";//"ti2021";
                //Globals.Ocompany.language = BoSuppLangs.ln_Spanish_La;


                int error = myCompany.Connect();


                //if (myCompany.Connect() == 0)
                //{
                //    Respuesta = true;
                //}


                if (error == 0)
                {
                    Respuesta = true;
                    //JavaScript.Alert("Conexión Exitosa a la Base de datos");
                    Console.WriteLine("Conexión Exitosa a la Base de datos");
                }
                else
                {
                    myCompany.Disconnect();
                    //JavaScript.Alert("Error - " + myCompany.GetLastErrorDescription().ToString());
                    Console.WriteLine("Error - " + myCompany.GetLastErrorDescription().ToString());

                }

                return Respuesta;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public static void Close()
        {
            myCompany.Disconnect();
        }
    }
}
