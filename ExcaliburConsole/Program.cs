using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Net;
using Excalibur.Models;

namespace ExcaliburConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Channel ch = new Channel();
     
            string token = ch.getToken("huankiat@processclick.com", "password");
            
            string response = ch.publishChannel("ConsoleAuth", "123", 12);
            //string token = ch.getToken("huankiat@processclick.com", "password");
            //Console.Write("Token from Website: " + token + "\n");

            //at.setToken(token);
            //at.createCookieInContainer();
            //string txt = at.readTokenFromCookie();
            //Console.Write("Cookie Token: " + txt + "\n");
           
            
            Console.Write(response);

            Console.Read();
        }
    }
}
