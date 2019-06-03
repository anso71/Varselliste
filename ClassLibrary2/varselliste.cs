using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using Agresso.ServerExtension;
using Agresso.Interface.CommonExtension;
using System.IO; 

namespace Stange.Server
{
    [ServerProgram("VARL")] // identification attribute
    public class Standalone : ServerProgramBase
    {
        
        public override void Run()
        {
            string client = ServerAPI.Current.Parameters["client"];
            string days = ServerAPI.Current.Parameters["dager"];
            string lonnmail = ServerAPI.Current.Parameters["epost"];
     
            if (client == string.Empty)
                Me.StopReport("Klient ikke fylt ut");
            if (days == string.Empty)
                Me.StopReport("Dager ikke fylt ut");

            DataTable dataTable = new DataTable("Varsel");
            Dictionary<string, string> listtext = new Dictionary<string, string>();
            IServerDbAPI api = ServerAPI.Current.DatabaseAPI;
            IStatement sql1 = CurrentContext.Database.CreateStatement();
            sql1.Append("select beskrivelse, dato, dim_value, stilling from afxvarselliste where client = @client and dato >= @date_start and dato < @date_end ");
            sql1["client"] = client;
            sql1["date_start"] = DateTime.Today.AddDays(Int32.Parse(days));
            sql1["date_end"] = DateTime.Today.AddDays(Int32.Parse(days)+1);
            Me.API.WriteLog("Varsel i perioden fra: {0}", DateTime.Today.AddDays(Int32.Parse(days)));
            Me.API.WriteLog("Til: {0}", DateTime.Today.AddDays(Int32.Parse(days)+1));



            CurrentContext.Database.Read(sql1, dataTable);
            foreach(DataRow row in dataTable.Rows)
            {
                IStatement sql2 = CurrentContext.Database.CreateStatement();
                sql2.Append("select description from agldimvalue where client = @client and dim_value = @resource_id and attribute_id = 'C0'");
                sql2["client"] = client;
                sql2["resource_id"] = row["dim_value"];
                string name = "";
                CurrentContext.Database.ReadValue(sql2, ref name);
                IStatement sql3 = CurrentContext.Database.CreateStatement();
                sql3.Append("select distinct a.rel_value, b.rel_value as AvdelingLeder from aprposrelvalue a, aglrelvalue b where a.client = @client and a.post_id = @stilling and a.rel_attr_id = 'MNAH' and a.resource_id = @resource_id and a.date_to > GETDATE() and a.rel_value = b.att_value and b.CLIENT= @client  and b.attribute_id  = 'MNAH' AND b.rel_attr_id = 'C0' group by a.rel_value, b.rel_value");
                sql3["client"] = client;
                sql3["resource_id"] = row["dim_value"];
                sql3["stilling"] = row["stilling"];
                DataTable managertable = new DataTable("manager");
                CurrentContext.Database.Read(sql3, managertable);
                foreach(DataRow row2 in managertable.Rows)
                {
                    string datetime = "" + row["dato"];
                    string resourceid = "" + row["dim_value"];
                    string stilling = "" + row["stilling"];
                    string varsel = "" + row["beskrivelse"];
                    name = WebUtility.HtmlEncode(name);
                    varsel = WebUtility.HtmlEncode(varsel);

                    //string textbody = "Ansattnr: " + (row["dim_value"]) + "</td><td>Dato:</td><td>" + datetime.Replace("00:00:00", "") + "</td><td>Varsel:</td><td>" + row["beskrivelse"] + "</td><td>" + name +"</td></tr>";
                    StringBuilder _sb = new StringBuilder();
                    _sb.Append("<tr><td>");
                    _sb.Append(string.Format(name.Replace("amp;","").ToFixedString(50))); //name
                    _sb.Append("</td><td>");
                    _sb.Append(string.Format(resourceid.ToFixedString(10))); //resource id
                    _sb.Append("</td><td>");
                    _sb.Append(string.Format(stilling.ToFixedString(10)));
                    _sb.Append("</td><td>");
                    _sb.Append(string.Format(datetime.Replace(" 00:00:00", "").ToFixedString(15))); //dato
                    _sb.Append("</td><td>");
                    _sb.Append(string.Format(varsel.ToFixedString(100))); //varel
                    _sb.Append("</td></tr>");

                    string managerid = "" + row2["AvdelingLeder"];
                    if (listtext.ContainsKey(managerid))
                    {
                        listtext[managerid] = listtext[managerid] + _sb.ToString() + "\n\r";
                    }
                    else
                    {
                        listtext.Add(managerid, _sb.ToString()+ "\n\r");
                    }
                }
                
              
                
            }
            StringBuilder _alltext = new StringBuilder();

            foreach (var item in listtext)
            {
                IStatement sqlemail = CurrentContext.Database.CreateStatement();
                sqlemail.Append("select e_mail from agladdress where dim_value = @dim_value and client = @client and attribute_id = 'C0'");
                sqlemail["client"] = client;
                sqlemail["dim_value"] = item.Key;
                string email = "";
                CurrentContext.Database.ReadValue(sqlemail, ref email);

                StringBuilder _sb2 = new StringBuilder();
                _sb2.Append("<html><head><style>#customers {  font-family: \"Trebuchet MS\", Arial, Helvetica, sans-serif;  border-collapse: collapse;  width: 100%;}#customers td, #customers th {  border: 1px solid #ddd;  padding: 8px;}#customers tr:nth-child(even){background-color: #f2f2f2;}#customers tr:hover {background-color: #ddd;}#customers th {  padding-top: 12px;  padding-bottom: 12px;  text-align: left;  background-color: #4CAF50;  color: white;}</style></head><body><H1>Varselliste</H1><table id =\"customers\"><tr><th><b>");
                _sb2.Append(string.Format("Name".ToFixedString(50))); //name
                _sb2.Append("</b></th><th><b>");
                _sb2.Append(string.Format("Ansattnr".ToFixedString(10))); //resource id
                _sb2.Append("</b></th><th><b>");
                _sb2.Append(string.Format("Stilling".ToFixedString(10)));
                _sb2.Append("</b></th><th><b>");
                _sb2.Append(string.Format("Dato".ToFixedString(15))); //dato
                _sb2.Append("</b></th><th><b>");
                _sb2.Append(string.Format("Varsel".ToFixedString(100))); //varsel
                _sb2.Append("</b></th></tr>");

                string emailtext = _sb2.ToString()  + item.Value + "</table></body></html>" ;
                string pathname = Path.Combine(Path.GetTempPath(), "Text.html");
                File.WriteAllText(pathname, emailtext);
                StringBuilder _Emailtext = new StringBuilder();
                _Emailtext.Append("Hei\r\nVedlagt fil innholder ansatte som ligger inne med varsel som inntreffer de neste 30 dager.\r\n");
                _Emailtext.Append("Gå inn på den ansatte og kontroller at opplysningen ligger riktig!\r\n");
                _Emailtext.Append("Mvh\r\nAgresso IntelAgent");
                if (ServerAPI.Current.SendMail(_Emailtext.ToString(), pathname, "Varselliste.html", "Varselliste ansatte", email,""))
                {
                    Me.API.WriteLog("Epost sendt til : {0}", email);
                    _alltext.Append("<br/>####### START ##############################<br/>");
                    _alltext.Append("Sendt til:<span/>");
                    _alltext.Append(email);
                    _alltext.Append(WebUtility.HtmlEncode("<br/>Med følgende text: <br/>"));
                    _alltext.Append(WebUtility.HtmlEncode(_Emailtext.ToString()));
                    _alltext.Append("<br/>");
                    _alltext.Append(emailtext);
                    _alltext.Append("<br/>####### SLUTT TEXT #####################################<br/>");
                }
                else
                {
                    Me.API.WriteLog("Epost ikke sent til: {0}", email);
                }
            }
            if (_alltext.ToString() != string.Empty)
            {

                string allpathname = Path.Combine(Path.GetTempPath(), "AllText.html");
                File.WriteAllText(allpathname, _alltext.ToString());
                if (ServerAPI.Current.SendMail("Se vedlegg på hva som ble sendt av varsel", allpathname, "AllVarselliste.html", "Varselliste samlet ansatte", lonnmail, ""))
                {
                    Me.API.WriteLog("Sendt samle mail");
                }
            }

        }
        
        public override void End()
        {
            Me.API.WriteLog("Stopping report {0}", Me.ReportName);
        }
    }
}
public static class StringExtensions
{
    /// <summary>
    /// Extends the <code>String</code> class with this <code>ToFixedString</code> method.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="length">The prefered fixed string size</param>
    /// <param name="appendChar">The <code>char</code> to append</param>
    /// <returns></returns>
    public static String ToFixedString(this String value, int length, char appendChar = ' ')
    {
        int currlen = value.Length;
        int needed = length == currlen ? 0 : (length - currlen);

        return needed == 0 ? value :
            (needed > 0 ? value + new string(' ', needed) :
                new string(new string(value.ToCharArray().Reverse().ToArray()).
                    Substring(needed * -1, value.Length - (needed * -1)).ToCharArray().Reverse().ToArray()));
    }
}