using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excalibur.Models
{
    using NUnit.Framework;

    [TestFixture]
    public class ChannelTest
    {

        [Test]
        public void Publish()
        {
            Channel ch = new Channel();
            Assert.AreEqual(250, 1000);
        }
    }
}
