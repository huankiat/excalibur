using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


namespace Excalibur.Models
{
   
    public class Channel
    {

        private string getChannelURL = "http://panoply-staging.herokuapp.com/api/channels.json";
        private string postChannelURL = "http://panoply-staging.herokuapp.com/api/channels/create_and_publish.json";
        private string channelDataURL = "http://panoply-staging.herokuapp.com/api/channels/";
        private string fileIDURL = "http://panoply-staging.herokuapp.com/api/spreadsheets.json";
        private string rePubURL = "http://panoply-staging.herokuapp.com/api/channels/";
        private string tokenURL = "http://panoply-staging.herokuapp.com/api/tokens.json";

        public JArray getAllChannels()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(getChannelURL);
            request.Method = "GET";
            request.Accept = "application/json";
            request.ContentType = "application/json";


            WebResponse response = request.GetResponse();
            using (Stream stream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                String responseString = reader.ReadToEnd();


                dynamic json = JValue.Parse(responseString);
                JArray datafeed = json.channels;
                reader.Close();
                response.Close();

                return datafeed;
            }
            
        }

        public string getChannelData(string channelID)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(channelDataURL + channelID + ".json");
            request.Method = "GET";
            request.Accept = "application/json";
            request.ContentType = "application/json";

            WebResponse response = request.GetResponse();
            using (Stream stream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                String responseString = reader.ReadToEnd();

                dynamic json = JValue.Parse(responseString);
                var channel = json.channel;
                string data = channel.value;
                reader.Close();
                response.Close();

                return data;
            }
        }

        public string publishChannel(string description, string value, int spreadsheet_id)
        {
            var jsonObject = new JObject();
            dynamic datafeed = jsonObject;
            datafeed.channel = new JObject();

            dynamic data = new JObject();
            data.description = description;
            data.value = value;

            datafeed.channel = data;
            datafeed.spreadsheet_id = spreadsheet_id;

            byte[] byteArray = Encoding.UTF8.GetBytes(datafeed.ToString());

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(postChannelURL);
            request.Method = "POST";
            request.Accept = "application/json";
            request.ContentType = "application/json";
            request.ContentLength = byteArray.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            //receive response
            WebResponse response = request.GetResponse();
            dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);

            string responseFromServer = reader.ReadToEnd();
            dynamic json = JValue.Parse(responseFromServer);
            string returnID = json.id;
            reader.Close();
            response.Close();

            return responseFromServer;
            

        }

        public string rePublishChannel(int channelID, string description, int value, int spreadsheet_id, bool to_replace)
        {
            var jsonObject = new JObject();
            dynamic datafeed = jsonObject;
            datafeed.channel = new JObject();

            dynamic data = new JObject();
            data.id = channelID;
            data.description = description;
            data.value = value;
            data.spreadsheet_id = spreadsheet_id;

            datafeed.channel = data;
            datafeed.can_override = to_replace;

            byte[] byteArray = Encoding.UTF8.GetBytes(datafeed.ToString());

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(rePubURL + channelID.ToString() + ".json");
            request.Method = "PUT";
            request.Accept = "application/json";
            request.ContentType = "application/json";
            request.ContentLength = byteArray.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            //receive response
            WebResponse response = request.GetResponse();
            dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);

            string responseFromServer = reader.ReadToEnd();
            reader.Close();
            response.Close();

            return responseFromServer.ToString();


        }

        public string getFileID(string filename)
        {
            var jsonObject = new JObject();
            dynamic datafeed = jsonObject;
            datafeed.spreadsheet = new JObject();

            dynamic data = new JObject();
            data.filename = filename;

            datafeed.spreadsheet = data;
            byte[] byteArray = Encoding.UTF8.GetBytes(datafeed.ToString());

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(fileIDURL);
            request.Method = "POST";
            request.Accept = "application/json";
            request.ContentType = "application/json";
            request.ContentLength = byteArray.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            //receive response
            WebResponse response = request.GetResponse();
            dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);

            string responseFromServer = reader.ReadToEnd();

            dynamic json = JValue.Parse(responseFromServer);
            string fileID = json.id.ToString();

            reader.Close();
            dataStream.Close();
            response.Close();

            return fileID;
        }

        public string checkFileID(Excel.Workbook wb)
        {
            string propertyName = "Excalibur ID";
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)wb.CustomDocumentProperties;
            
            foreach(Office.DocumentProperty prop in properties)
            {
                if (prop.Name.ToString() == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return "Nil";
            
        }

        public string getToken(string username, string password)
        {
            var jsonObject = new JObject();
            dynamic auth = jsonObject;
            auth.email = username;
            auth.password = password;

            byte[] byteArray = Encoding.UTF8.GetBytes(auth.ToString());

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(tokenURL);
            request.Method = "POST";
            request.Accept = "application/json";
            request.ContentType = "application/json";
            request.ContentLength = byteArray.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            //receive response
            WebResponse response = request.GetResponse();
            dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);

            string responseFromServer = reader.ReadToEnd();
            dynamic json = JValue.Parse(responseFromServer);
            string returnID = json.authentication_token;
            reader.Close();
            response.Close();

            return returnID;
        }



    }

    public class ExcelRange
    {
     
        public string cellName { get; set; }
        public string cellValue { get; set; }
        public Excel.Range rng { get; set; }

        public void nameCell()
        {
            this.rng.Name = this.cellName;
        }

        public void writeNameCell()
        {
            this.rng.Value = this.cellValue;
        }
        

    }

    public class AuthToken
    {
        private string authToken;
        Cookie excaliburCookie;

        public void setToken(string token)
        {
            authToken = token;
        }

    

        private void createCookie()
        {
            Cookie c = new Cookie("ExcaliburToken",authToken);
            excaliburCookie = c;
        }

        public CookieContainer createCookieContainer()
        {
            CookieContainer cc = new CookieContainer();
            createCookie();
            cc.Add(excaliburCookie);
            return cc;
        }


    }
}
