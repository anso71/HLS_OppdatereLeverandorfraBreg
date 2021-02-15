using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Data;
using Agresso.ServerExtension;
using Agresso.Interface.CommonExtension;

namespace HLS_OppdatereLevaBreg
{
    [ServerProgram("SUPREG2")]
    public class OppdateringLevBreg: ServerProgramBase
    {
        public override void Run()
        {
            string client = ServerAPI.Current.Parameters["client"];
            string url = ServerAPI.Current.Parameters["url"];
            string html = string.Empty;
            if (url == string.Empty)
                Me.StopReport("Ingen url");
            if (client == string.Empty)
                Me.StopReport("Ingen client");




            DataTable dataTable = new DataTable("Suppliers");
            IServerDbAPI api = ServerAPI.Current.DatabaseAPI;
            IStatement sql = CurrentContext.Database.CreateStatement();
            sql.Append("select apar_id, comp_reg_no from hls_supplierbreg where client = @client");
            sql["client"] = client;
            CurrentContext.Database.Read(sql, dataTable);
            
            foreach(DataRow row in dataTable.Rows)
            {
                try
                {
                    StringBuilder urlorg = new StringBuilder();
                    urlorg.Append(url);
                    urlorg.Append(row["comp_reg_no"].ToString().Trim());
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlorg.ToString());
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;
                    request.AutomaticDecompression = DecompressionMethods.GZip;
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader1 = new StreamReader(stream, Encoding.GetEncoding(1252)))
                    {
                        html = reader1.ReadToEnd();
                    }
                    Me.API.WriteLog("comp: {0}", row["comp_reg_no"]);
               
              

                JObject rss = JObject.Parse(html);

                string kode = (string)rss["naeringskode1"]["kode"];
                string kode2 = "";
                try
                {
                    kode2 = (string)rss["naeringskode2"]["kode"];
                }
                catch
                {
                    Me.API.WriteLog("ikke 2");
                }
                Me.API.WriteLog("kode {0}", kode);
                if (kode == String.Empty)
                    Me.API.WriteLog("ingen næringskode for org {0}", row["org_reg_no"]);
                else
                {
                    IStatement soksql = CurrentContext.Database.CreateStatement();
                    soksql.Append("select rel_value from aglrelvalue where client = @client and rel_attr_id = 'MNSV' and att_value = @apar_id and attribute_id = 'A5'");
                    soksql["client"] = client;
                    soksql["apar_id"] = row["apar_id"];
                    String verdi = "";
                    
                    Me.API.WriteLog(" verdi: {0} på kall {1}", verdi, soksql);
                    IStatement narsql = CurrentContext.Database.CreateStatement();
                    if (!CurrentContext.Database.ReadValue(soksql, ref verdi))
                    {
                        narsql.Append("insert into aglrelvalue(client,rel_attr_id, rel_value, att_val_from,att_val_to,date_from,date_to, attribute_id,att_value,last_update,percentage, user_id)");
                        narsql.Append(" values (@client,'MNSV',@kode,@apar_id,@apar_id,@datefrom, @dateto,'A5',@apar_id,getDate(),100,'BREGg')");
                        narsql["client"] = client;
                        narsql["kode"] = kode;
                        narsql["apar_id"] = row["apar_id"];
                        narsql["datefrom"] = Convert.ToDateTime("1900-01-01");
                        narsql["dateto"] = Convert.ToDateTime("2099-12-31");
                        Me.API.WriteLog("Legger inn for org: {0}  og kundenr: {1}", row["comp_reg_no"], row["apar_id"]);
                    }
                    else
                    {
                        narsql.Append("update aglrelvalue set rel_value  = @kode where client = @client and rel_attr_id = 'MNSV' and att_value = @apar_id and attribute_id = 'A5'");
                        narsql["client"] = client;
                        narsql["kode"] = kode;
                        narsql["apar_id"] = row["apar_id"];
                    }
                    if (!narsql.IsEmpty())
                    {
                        int tall = CurrentContext.Database.Execute(narsql);
                        Me.API.WriteLog("Legger inn {0} antall {1}", narsql,tall);
                    }
                    if (String.Equals(kode2,""))
                    {
                        Me.API.WriteLog("Tomt");
                    }
                    else
                    {
                        IStatement sok2sql = CurrentContext.Database.CreateStatement();
                        sok2sql.Append("select rel_value from aglrelvalue where client = @client and rel_attr_id = 'MNSX' and att_value = @apar_id and attribute_id = 'A5'");
                        sok2sql["client"] = client;
                        sok2sql["apar_id"] = row["apar_id"];
                        String verdi2 = "";
                        
                        IStatement nar2sql = CurrentContext.Database.CreateStatement();
                        if (!CurrentContext.Database.ReadValue(sok2sql, ref verdi2))
                        {
                            nar2sql.Append("insert into aglrelvalue(client,rel_attr_id, rel_value, att_val_from,att_val_to,date_from,date_to, attribute_id,att_value,last_update,percentage, user_id)");
                            nar2sql.Append(" values (@client,'MNSX',@kode,@apar_id,@apar_id,@datefrom, @dateto,'A5',@apar_id,getDate(),100,'BREGg')");
                            nar2sql["client"] = client;
                            nar2sql["kode"] = kode2;
                            nar2sql["apar_id"] = row["apar_id"];
                            nar2sql["datefrom"] = Convert.ToDateTime("1900-01-01");
                            nar2sql["dateto"] = Convert.ToDateTime("2099-12-31");
                            Me.API.WriteLog("Legger inn næ2 for org: {0}  og kundenr: {1}", row["comp_reg_no"], row["apar_id"]);
                        }
                        else
                        {
                            nar2sql.Append("update aglrelvalue set rel_value  = @kode where client = @client and rel_attr_id = 'MNSX' and att_value = @apar_id and attribute_id = 'A5'");
                            nar2sql["client"] = client;
                            nar2sql["kode"] = kode2;
                            nar2sql["apar_id"] = row["apar_id"];
                        }
                        if (!nar2sql.IsEmpty())
                            CurrentContext.Database.Execute(nar2sql);

                    }
                    
                }


             }
             catch
             {
                    Me.API.WriteLog("Finnes ikke comp {0}", row["comp_reg_no"]);

              }


            }

            
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
