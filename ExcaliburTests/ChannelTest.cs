using System;
using System.Net.Fakes;
using System.IO;
using System.Collections.Generic;
using NUnit.Framework;
using Excalibur;
using Excalibur.Models;
using Newtonsoft.Json.Linq;
using Microsoft.QualityTools.Testing.Fakes;
using Moq;

namespace ExcaliburTests
{
    [TestFixture]
    public class ChannelTest
    {
        [Test]
        public void GetAllChannelsTest()
        {
            using (ShimsContext.Create())
            {
                var responseStream = new FileStream("Fixtures/v1.channels.index.response", FileMode.Open);
                var responseShim = new ShimHttpWebResponse()
                {
                     GetResponseStream = () => responseStream
                };
                String actualMethod = "";
                var requestShim = new ShimHttpWebRequest()
                {
                    MethodSetString = (method) => { actualMethod = method; },
                    GetResponse = () => responseShim
                };

                String actualURL = "";
                ShimWebRequest.CreateString = (url) =>
                {
                    actualURL = url;
                    return requestShim;
                };

                Channel ch = new Channel();
                JArray json = ch.getAllBroadcastsChannels();
                
                StringAssert.AreEqualIgnoringCase("GET", actualMethod);
                StringAssert.AreEqualIgnoringCase("http://panoply-staging.herokuapp.com/api/channels.json", actualURL);
                Assert.AreEqual(2, json.Count);
            }
        }
    }
}
