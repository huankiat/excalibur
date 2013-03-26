using System;
using System.Net;
using System.IO;
using System.Collections.Generic;
using NUnit.Framework;
using Moq;
using Excalibur;
using Excalibur.Models;
using Newtonsoft.Json.Linq;

namespace ExcaliburTests
{
    [TestFixture]
    public class ChannelTest
    {
        [Test]
        public void GetAllChannelsTest()
        {
            var responseStream = new FileStream("Fixtures/v1.channels.index.response", FileMode.Open);

            var response = new Mock<HttpWebResponse>();
            response.Setup(m => m.GetResponseStream())
                .Returns(responseStream);

            var request = new Mock<HttpWebRequest>();            
            request.SetupSet(m => m.Method = "GET");
            request.Setup(m => m.GetResponse())
                .Returns(response.Object);

            var factory = new Mock<IWebRequestFactory>();
            factory.Setup(m => m.Create("http://panoply-staging.herokuapp.com/api/channels.json"))
                .Returns(request.Object);

            Channel ch = new Channel(factory.Object);
            JArray json = ch.getAllChannels();
            Assert.AreEqual(2, json.Count);
        }
    }
}
