using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MarsFramework.Pages;
using NUnit.Framework;
#nullable disable

namespace MarsFramework.Test
{
    class Program
    {
        [TestFixture]
        [Category("Sprint1")]
        class User : Global.Base
        {
            [Test]
            public void a_TestAddShareSkill()
            {
                ShareSkill shareSkill = new ShareSkill();
                shareSkill.EnterShareSkill();
            }
            //[Test, Order(2)]
            //public void b_ViewManageShareSkill()
            //{
            //    //ManageListings manageListing = new ManageListings();
            //    //manageListing.viewListings();
            //    //string firstskilltitle = manageListing.getfirsttitle();
            //    //Assert.That(firstskilltitle == "Selenium", "title doest not match");

            //}

            //[Test, Order(3)]
            //public void c_editManageShareskill()
            //{
            //    //ManageListings managelistings = new ManageListings();
            //    ////managelistings.editListings();
            //    //string expectedtitle = managelistings.editListings();
            //    //string firstskilltitle = managelistings.getfirsttitle();
            //    //Assert.That(firstskilltitle == expectedtitle, "title doest not match");
            //}

            //[Test, Order(4)]
            //public void d_deleteListings()
            //{
            //    //ManageListings managelistings = new ManageListings();
            //    //string message = managelistings.deleteListings();
            //    //string actualmessage = managelistings.getpopupmessage();
            //    //Assert.AreEqual(message, actualmessage);

            //}
        }
    }
}
