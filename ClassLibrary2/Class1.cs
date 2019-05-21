using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using Agresso.ServerExtension;
using Agresso.Interface.CommonExtension;

namespace Stange.Server
{
    [ServerProgram("VARL")] // identification attribute
    public class Standalone : ServerProgramBase
    {
        public override void Run()
        {
            string client = ServerAPI.Current.Parameters["client"];
            string days = ServerAPI.Current.Parameters["dager"];
     
            if (client == string.Empty)
                Me.StopReport("Klient ikke fylt ut");
            if (days == string.Empty)
                Me.StopReport("Dager ikke fylt ut");

            DataTable dataTable = new DataTable("Varsel");
            Dictionary<string, string> listtext = new Dictionary<string, string>();
            IServerDbAPI api = ServerAPI.Current.DatabaseAPI;
            IStatement sql1 = CurrentContext.Database.CreateStatement();
            sql1.Append("select beskrivelse, dato, dim_value from afxvarselliste where client = @client and dato >= @date_start and dato <= @date_end ");
            sql1["client"] = client;
            sql1["date_start"] = DateTime.Today.AddDays(Int32.Parse(days));
            sql1["date_end"] = DateTime.Today.AddDays(Int32.Parse(days) + 1);
            CurrentContext.Database.Read(sql1, dataTable);
            foreach(DataRow row in dataTable.Rows)
            {
                IStatement sql2 = CurrentContext.Database.CreateStatement();
                sql2.Append("select description from agldimvalue where client = client and dim_value = @resource_id and attribute_id = 'C0'");
                sql2["Client"] = client;
                sql2["resource_id"] = row["dim_value"];
                string name = "";
                CurrentContext.Database.ReadValue(sql2, ref name);
                string text = row["beskrivelse"] + " " + name + "\n";
                IStatement sql3 = CurrentContext.Database.CreateStatement();
                sql3.Append("select distinct a.rel_value, b.AvdelingLeder from aprposrelvalue a, ahi_idm_arbsttest b where a.client = @client  and a.rel_attr_id = 'MNAH' and a.resource_id = @resource_id and a.date_to > GETDATE() and a.rel_value = b.Avdeling and b.CLIENT= @client group by a.rel_value, b.AvdelingLeder");
                sql3["client"] = client;
                sql3["resource_id"] = row["dim_value"];
                DataTable managertable = new DataTable("manager");
                CurrentContext.Database.Read(sql3, managertable);
                foreach(DataRow row2 in managertable.Rows)
                {
                    string datetime = "" + row["dato"];
                    string textbody = "Ansattnr: " + row["dim_value"] + "\t Dato: " + datetime.Replace("00:00:00", "") + "\t Varsel: " + row["beskrivelse"] + " \t" + name +"\n";
                    string managerid = "" + row2["AvdelingLeder"];
                    if (listtext.ContainsKey(managerid))
                    {
                        listtext[managerid] = listtext[managerid] + textbody;
                    }
                    else
                    {
                        listtext.Add(managerid, textbody);
                    }
                }
                
              
                
            }
            foreach (var item in listtext)
            {
                IStatement sqlemail = CurrentContext.Database.CreateStatement();
                sqlemail.Append("select e_mail from agladdress where dim_value = @dim_value and client = @client and attribute_id = 'C0'");
                sqlemail["client"] = client;
                sqlemail["dim_value"] = item.Key;
                string email = "";
                CurrentContext.Database.ReadValue(sqlemail, ref email);

                if (ServerAPI.Current.SendMail(item.Value, "", "Varselliste", "Varselliste", "andre.sollie@stange.kommune.no", "agresso@hedmark-ikt.no"))
                {
                    Me.API.WriteLog("Epost sendt til : {0}", email);
                }
                else
                {
                    Me.API.WriteLog("Epost ikke sent til: {0}", email);
                }
            }

        }
        public override void End()
        {
            Me.API.WriteLog("Stopping report {0}", Me.ReportName);
        }
    }
}
