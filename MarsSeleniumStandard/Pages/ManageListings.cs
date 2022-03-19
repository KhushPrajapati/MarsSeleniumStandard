using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MarsFramework.Global;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
#nullable disable

namespace MarsFramework.Pages
{
    internal class ManageListings
    {
        public ManageListings()
        {
            PageFactory.InitElements(Global.GlobalDefinitions.driver, this);

        }

        //Click on Manage Listings Link
        [FindsBy(How = How.LinkText, Using = "Manage Listings")]
        private IWebElement manageListingsLink { get; set; }

        //View the listing                 
        [FindsBy(How = How.XPath, Using = "//*[@id='listing-management-section']/div[2]/div[1]/div[1]/table/tbody/tr[1]/td[8]/div/button[1]/i")]
        private IWebElement view { get; set; }

        //view firstrecord
        [FindsBy(How = How.XPath, Using = "(//i[@class='skill-title'])[1]")]
        private IWebElement firstRecord { get; set; }

        //Firstskilltitle
        [FindsBy(How = How.XPath, Using = "//*[@id='service-detail-section']/div[2]/div/div[1]/div/div/div/div")]
        private IWebElement firstSkilltitle { get; set; }


        //Delete the listing
        [FindsBy(How = How.XPath, Using = "//table[1]/tbody[1]")]
        private IWebElement delete { get; set; }

        //Delete button of the first record in the Listing
        [FindsBy(How = How.XPath, Using = "//*[@id='listing-management-section']/div[2]/div[1]/div[1]/table/tbody/tr/td[8]/div/button[3]/i")]
        private IWebElement deletetherecord { get; set; }

        //YES button to delete
        [FindsBy(How = How.XPath, Using = "/html/body/div[2]/div/div[3]/button[2]")]
        private IWebElement yes { get; set; }


        //Edit the listing
        [FindsBy(How = How.XPath, Using = "(//i[@class='outline write icon'])[1]")]
        private IWebElement edit { get; set; }

        //Click on Yes or No
        [FindsBy(How = How.XPath, Using = "//div[@class='actions']")]
        private IWebElement clickActionsButton { get; set; }

        internal void viewListings()
        {
            Thread.Sleep(2000);
            //Click on ManageListing tab
            manageListingsLink.Click();
            Thread.Sleep(4000);
            view.Click();
        }
        public string getfirsttitle()
        {
            manageListingsLink.Click();
            Thread.Sleep(2000);
            view.Click();
            Thread.Sleep(2000);
            firstSkilltitle.Click();
            return firstSkilltitle.Text;
        }
        public string editListings()
        {
            Thread.Sleep(2000);
            manageListingsLink.Click();
            Thread.Sleep(4000);
            edit.Click();
            ShareSkill shareSkill = new ShareSkill();
            shareSkill.EditShareSkill();
            GlobalDefinitions.ExcelLib.PopulateInCollection(Base.ExcelPath, "EditShareSkill");
            Thread.Sleep(3000);
            string expectedtitle = GlobalDefinitions.ExcelLib.ReadData(2, "Title");
            return expectedtitle;

        }
        public string deleteListings()
        {
            Thread.Sleep(2000);
            manageListingsLink.Click();
            Thread.Sleep(2000);
            deletetherecord.Click();
            yes.Click();
            Thread.Sleep(2000);
            GlobalDefinitions.ExcelLib.PopulateInCollection(Base.ExcelPath, "EditShareSkill");
            //GlobalDefinitions.wait(1);
            Thread.Sleep(3000);
            string expectedtitle = GlobalDefinitions.ExcelLib.ReadData(2, "Title");
            string expectedmessage = expectedtitle + " has been deleted";
            return expectedmessage;

        }
        public string getpopupmessage()
        {
            //Deleted message
            String message = GlobalDefinitions.driver.FindElement(By.XPath(@"//div[@class='ns-box ns-growl ns-effect-jelly ns-type-success ns-show']")).Text;
            return message;
        }
    }
}
