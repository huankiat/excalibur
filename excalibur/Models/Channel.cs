using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace Excalibur.Models
{
    public class Channel
    {
        public string getChannelURL = "http://panoply-staging.herokuapp.com/api/channels.json";
        public string postChannelURL = "http://panoply-staging.herokuapp.com/api/channels.json";
        public string channelDataURL = "http://panoply-staging.herokuapp.com/api/channels/";
        public string fileIDURL = "http://panoply-staging.herokuapp.com/api/spreadsheets.json";

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

        public string publishChannel(string description, string value)
        {
            var jsonObject = new JObject();
            dynamic datafeed = jsonObject;
            datafeed.channel = new JObject();

            dynamic data = new JObject();
            data.description = description;
            data.value = value;

            datafeed.channel = data;
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

            return responseFromServer;
            

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


    }
}
