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
using Microsoft.Win32;


namespace Excalibur.Models
{
    
    public class AuthChannel
    {
        private string authToken = "";

        public string tokenValue()
        {
            if (TokenStore.checkTokenInStore())
            {
                authToken = TokenStore.getTokenFromStore();
            }
            else
            {
                authToken = "";
            }

            return authToken;
        }

    }
    
    public class Channel
    {

        //All available channel endpoints
        private string getAllBroadcastsURL = "http://panoply-staging.herokuapp.com/api/broadcasts.json";
        private string getPermittedBroadcastsURL = "http://panoply-staging.herokuapp.com/api/broadcasts.json?filter=owned";
        private string postChannelURL = "http://panoply-staging.herokuapp.com/api/channels.json";
        private string channelDataURL = "http://panoply-staging.herokuapp.com/api/channels/";
        private string fileIDURL = "http://panoply-staging.herokuapp.com/api/spreadsheets.json";
        private string rePubURL = "http://panoply-staging.herokuapp.com/api/channels/";
        private string tokenURL = "http://panoply-staging.herokuapp.com/api/tokens.json";
        private string authToken = "";
        private dynamic channelJson;


        private String processclick_GET_request(string channelURL)
        {
            String responseTxt = "";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(channelURL);
            request.Method = "GET";
            request.Accept = "application/json";
            request.ContentType = "application/json";
            request.Headers[HttpRequestHeader.Authorization] = "Token token=" + this.authToken;
            
            WebResponse response = request.GetResponse();
            using (Stream stream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                String responseString = reader.ReadToEnd();
                responseTxt = responseString;
                reader.Close();
            }
            response.Close();

            return responseTxt;
        }

        private String processclick_POST_request(string channelURL, string datafeed, string method="POST")
        {
            String responseTxt = "";
            byte[] byteArray = Encoding.UTF8.GetBytes(datafeed);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(channelURL);
            request.Method = method;
            request.Accept = "application/json";
            request.ContentType = "application/json";
            request.Headers[HttpRequestHeader.Authorization] = "Token token=" + this.authToken;
            request.ContentLength = byteArray.Length;

            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();

            try
            {
                WebResponse response = request.GetResponse();
                dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);

                responseTxt = reader.ReadToEnd();
                reader.Close();
                response.Close();
            }
            catch (WebException ex)
            {
                HttpWebResponse webResponse = ex.Response as HttpWebResponse;
                if (webResponse.StatusCode == HttpStatusCode.NotFound)
                {
                    responseTxt = "404";
                }
                else if (webResponse.StatusCode == HttpStatusCode.Unauthorized)
                {
                    responseTxt = "401";
                }
                else if (webResponse.StatusCode == HttpStatusCode.Conflict)
                {
                    responseTxt = "409";
                }
                else if (webResponse.StatusCode == HttpStatusCode.Forbidden)
                {
                    responseTxt = "403";
                }

                webResponse.Close();
            }

            return responseTxt;
        }
        
        public void setAuthToken(string token)
        {
            authToken = token;
        }

        public string getChannelToken()
        {
            if (authToken == "")
            {
                return "No Token Assigned";
            }
            else
            {
                return authToken;
            }
        }


        public JArray getAllBroadcastsChannels()
        {
            String responseString = this.processclick_GET_request(getAllBroadcastsURL);
            dynamic json = JValue.Parse(responseString);
            JArray datafeed = json.broadcasts;
            return datafeed;
        }

        public JArray getPermittedBroadcastsChannels()
        {
            String responseString = this.processclick_GET_request(getPermittedBroadcastsURL);
            dynamic json = JValue.Parse(responseString);
            JArray datafeed = json.broadcasts;
            return datafeed;
        }

        public JArray getBroadcastsList(JArray datafeed)
        {
            JArray broadcastsList = new JArray();

            foreach (JObject json in datafeed)
            {
                dynamic output = new JObject();
                dynamic input = new JObject();
                input = json;
                output.id = input.id;
                output.description = input.description;
                broadcastsList.Add(output);               
            }

            return broadcastsList;

        }


        public JArray getBroadcastChannelList(string broadcastID, JArray datafeed)
        {
            JArray broadcastChannelList = new JArray();
            foreach (JObject json in datafeed)
            {
                dynamic broadcast = new JObject();
                broadcast = json;
                if (broadcast.id == broadcastID)
                {
                    broadcastChannelList = broadcast.channels;
                }
            }

            return broadcastChannelList;

        }

        public JArray filterPermittedChannels(int userID, JArray datafeed)
        {
            JArray filteredDataFeed = new JArray();
            foreach (JObject json in datafeed)
            {
                dynamic j = new JObject();
                j = json;
                if ((j.owner_id == userID) | (j.assignee_id == userID))
                {
                    filteredDataFeed.Add(json);
                }
            }

            return filteredDataFeed;
            

        }

        public JArray BroadcastList(int selector)
        {
            string filter = "";

            if (selector == 0)
            {
                filter = "owned";
            }
            else if (selector == 1)
            {
                filter = "followed";
            }

            String responseString = this.processclick_GET_request(this.getAllBroadcastsURL + "?filter=" + filter);
            dynamic json = JValue.Parse(responseString);
            JArray datafeed = json.broadcasts;
            return datafeed;
            
        }

        public void getChannelJson(string channelID)
        {
            String responseString = this.processclick_GET_request(this.channelDataURL + channelID + ".json");
            dynamic json = JValue.Parse(responseString);
            channelJson = json;
        }

        //get channel data from getChannelJson(channelID). 
        //If getChannelJson has not been called yet, call getChannelJson.
        public string getChannelData(string channelID)
        {
           
                getChannelJson(channelID);
         

            var channel = channelJson.channel;
            string data = channel.value;
            return data;
        }

        //get channel description
        public string getChannelDesc(string channelID)
        {
            
                getChannelJson(channelID);
            

            var channel = channelJson.channel;
            string data = channel.description;
            return data;

        }

        //get AssigneeIDs for a channel as int[]
        public int[] getAssigneeIDs(string channelID)
        {
         
                getChannelJson(channelID);
            
            
            var channel = channelJson.channel;
            List<int> assignee_id = new List<int>();
            JArray idArray = channel.assignee_id;

            foreach (dynamic j in idArray)
            {
                assignee_id.Add(j);
            }
            return assignee_id.ToArray();
        }

        //get Owner ID for a channel
        public int getOwnerID(string channelID)
        {
           
                getChannelJson(channelID);
            
            
            var channel = channelJson.channel;
            int data = channel.owner_id;
            return data;
        }

        //get spreadsheetID for a channel
        public int getChannelSpreadSheetID(string channelID)
        {
           
                getChannelJson(channelID);
          

            var channel = channelJson.channel;
            int data = channel.spreadsheet_id;
            return data;
        }
        

        /// <summary>
        /// Publish to a secured channel. Must call setAuthToken(string token) \
        /// first or else return error code 'no_token'
        /// </summary>
        /// <param name="description">channel description</param>
        /// <param name="value">channel value</param>
        /// <param name="spreadsheet_id">spreadsheet_id</param>
        /// <returns>channel ID if successful</returns>
        public string publishChannel(string description, string value, int spreadsheet_id, int broadcast_id)
        {
            string returnID;

            if (authToken != "")
            {
                var jsonObject = new JObject();
                dynamic datafeed = jsonObject;
                datafeed.channel = new JObject();

                dynamic data = new JObject();
                data.description = description;
                data.value = value;
                data.spreadsheet_id = spreadsheet_id;

                datafeed.channel = data;
                datafeed.broadcast = broadcast_id;

                String responseFromServer = this.processclick_POST_request(postChannelURL, datafeed.ToString());
                dynamic json = JValue.Parse(responseFromServer);
                returnID = json.id;
            }
            else
            {
                returnID = "no_token";
            }

            return returnID;
            

        }

        public string rePublishChannel(int channelID, string description, 
            string value, int spreadsheet_id, bool to_replace)
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
            datafeed.force = to_replace.ToString();

            string responseFromServer=this.processclick_POST_request(rePubURL+ channelID.ToString() + ".json",
                                      datafeed.ToString(),"PUT");

            return responseFromServer.ToString();
        }

        public string getSpreadSheetID(string filename)
        {
            var jsonObject = new JObject();
            dynamic datafeed = jsonObject;
            datafeed.spreadsheet = new JObject();

            dynamic data = new JObject();
            data.filename = filename;
            datafeed.spreadsheet = data;

            string responseFromServer = this.processclick_POST_request(fileIDURL, datafeed.ToString());
            dynamic json = JValue.Parse(responseFromServer);
            string fileID = json.id.ToString();         

            return fileID;
        }

        public string checkSpreadSheetID(Excel.Workbook wb)
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
            return "0";
            
        }

        public string getToken(string username, string password)
        {
            var jsonObject = new JObject();
            dynamic auth = jsonObject;
            auth.email = username;
            auth.password = password;

            string responseFromServer = this.processclick_POST_request(tokenURL, auth.ToString());
            dynamic json = JValue.Parse(responseFromServer);
            string returnID = json.authentication_token;                
           
            return returnID;
        }

        public string channelsRefresh(Excel.Workbook wb)
        {
            string message = "Error in refreshing";
            int channelID;
            string partial_name;
            string full_name;

            foreach (Excel.Name nRange in wb.Names)
            {
                full_name = nRange.Name.ToString();
                partial_name = full_name.Substring(4);
                channelID = Convert.ToInt32(partial_name);

                if (full_name.Substring(0, 3) == "SUB")
                {
                    nRange.RefersToRange.Value = this.getChannelData(channelID.ToString());
                    message = "Successful";
                }
                else if (full_name.Substring(0,3) == "PUB" & checkSpreadSheetID(wb) == "0")
                {                       
                    message = "You need to register this workbook";
                }
                        
                else if (full_name.Substring(0,3) == "PUB" & checkSpreadSheetID(wb) != "0")
                {                            
                    string r_value = nRange.RefersToRange.Value.ToString();
                    int fileID = Convert.ToInt32(checkSpreadSheetID(wb));
                    string cellDesc = this.getChannelDesc(channelID.ToString());
                    message = this.rePublishChannel(channelID, cellDesc, r_value, fileID, false);                           
                }
            }
            
            return message;
        }

    }

    public class ExcelOps
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

    //NOT USED - Use TokenStore instead//
//    public class AuthToken
//    {
//        private string authToken;
//        private Cookie excaliburCookie;
//        public static CookieContainer cContainer;
//        private Uri cURI;

//        public AuthToken()
//        {
//            cURI = new Uri("http://www.processclick.com/");
//        }

//        public void setCookieURI(Uri cookieURI)
//        {
//            cURI = cookieURI;
//        }

//        public void setToken(string token)
//        {
//            authToken = token;
//        }

//        public Cookie getCookie()
//        {
//            return excaliburCookie;
//        }
    

//        public void createCookieInContainer()
//        {
//            CookieContainer cc = new CookieContainer();
//            Cookie c = new Cookie();
//            cc.MaxCookieSize = 5000;

//            c.Value = authToken;
//            c.Name = "ExcaliburToken";
//            cc.Add(cURI, c);
//            excaliburCookie = c;
//            AuthToken.cContainer = cc;
//        }


//        public string readTokenFromStore()
//        {
//            string token = "X";
//            CookieCollection cookies = new CookieCollection();

//            cookies = AuthToken.cContainer.GetCookies(cURI);

//            if (cookies.Count == 0)
//            {
//                token = "No cookie detected";
//            }
//            else
//            {
//                for (int i = 0; i < cookies.Count; i++)
//                {
//                    Cookie c = cookies[i];
//                    if (c.Name == "ExcaliburToken")
//                    {
//                        token = c.Value;
//                    }

//                }
//            }
//            return token;
//        }

//    }
}


public static class TokenStore
{
    private static int daysToExpire = 15; //15 days for token to expire

    public static void addTokenToStore(string token)
    {
        if (Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur")==null)
        {
            Registry.CurrentUser.CreateSubKey("SOFTWARE\\Excalibur");
        }
        RegistryKey myKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur", true);
        myKey.SetValue("Token", token, RegistryValueKind.String);
        myKey.SetValue("TokenDate", DateTime.Today.ToShortDateString(), RegistryValueKind.String);

    }

    public static string getTokenFromStore()
    {
        RegistryKey myKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur", false);
        string myValue = (string)myKey.GetValue("Token");

        return myValue;
    }

    public static DateTime getTokenDateFromStore()
    {
        RegistryKey myKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur", false);
        DateTime myValue = (DateTime)Convert.ToDateTime(myKey.GetValue("TokenDate"));

        return myValue;
    }

    public static bool checkTokenInStore()
    {
        bool is_TokenInStore;

        if (Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur") == null)
        {
            is_TokenInStore = false;
        }
        else
        {
            RegistryKey myKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur", false);
            if (myKey.GetValue("Token").ToString() == ""| myKey.GetValue("Token")==null )
            {
                is_TokenInStore = false;
            }
            else
            {
                is_TokenInStore = true;
            }
        }
        return is_TokenInStore;
    }

    public static void checkTokenExpiry()
    {
        DateTime tokenDate;
        tokenDate = TokenStore.getTokenDateFromStore();
        if (DateTime.Today >= tokenDate.AddDays(daysToExpire))
        {
            RegistryKey myKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur", true);
            myKey.SetValue("Token", "", RegistryValueKind.String);
            myKey.SetValue("TokenDate", "", RegistryValueKind.String);
        }
    }

}
