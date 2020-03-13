using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;

namespace TaskRPA
{
    /// <summary>
    /// Class that allows to parse web page and get a list of elements.
    /// </summary>
    public class Parser
    {
        private readonly IWebDriver driver;
        private List<IWebElement> webElements;
        private List<Microwave> microwaves;

        public Parser(IWebDriver driver)
        {
            this.driver = driver;
        }

        public List<Microwave> Parse()
        {
            using (driver)
            {
                ManageDriver();
                MoveToPage();
                GetWebElements();
                GetMicrowaves();
            }

            return microwaves;
        }

        private void MoveToPage()
        {
            driver.Navigate().GoToUrl("https://onliner.by");
            driver.FindElement(By.XPath("//*[@class='b-top-menu']/div/nav/ul/li/a/span")).Click();
            driver.FindElement(By.XPath("//*[@data-id='3']")).Click();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//*[@id='container']/div/div[2]/div/div/div[1]/div[3]/div/div[3]/div[1]/div/div[6]")).Click();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//*[@id='container']/div/div[2]/div/div/div[1]/div[3]/div/div[3]/div[1]/div/div[6]/div[2]/div/a[1]")).Click();
            driver.FindElement(By.XPath("//*[@id='schema-order']/a")).Click();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//*[@id='schema-order']/div[2]/div/div[5]")).Click();
            Thread.Sleep(5000);
        }

        private void ManageDriver()
        {
            driver.Manage().Window.Maximize();
        }

        private void GetWebElements()
        {
            webElements = driver.FindElements(By.ClassName("schema-product__group")).ToList();
        }

        private void GetMicrowaves()
        {
            microwaves = new List<Microwave>();

            foreach (var a in webElements)
            {
                Microwave mv = new Microwave();

                mv.Title = a.FindElement(By.ClassName("schema-product__title")).Text.Replace("Микроволновая печь ", "");
                mv.Price = a.FindElements(By.ClassName("schema-product__price")).First().Text;
                mv.Href = a.FindElement(By.TagName("a")).GetAttribute("href");

                microwaves.Add(mv);
            }
        }
    }
}
