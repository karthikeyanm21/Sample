using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.Xml;
using Newtonsoft.Json;
using APLPX.Client;

namespace DemoTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //DataTable dt = new DataTable();
             AplBusinessLayer objBL=new AplBusinessLayer();
            APLPX.Client.localhost.StagingDbConfig stagingInfo =new APLPX.Client.localhost.StagingDbConfig();
            //stagingInfo.
            string result= objBL.TestConnection(stagingInfo);
        }
    }
}
