using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.DevTools.V148.Page;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using DevToolsSessionDomains = OpenQA.Selenium.DevTools.V148.DevToolsSessionDomains;

namespace TCore.WebControl
{
    public class StaleSafeAccess
    {
        private Func<IWebElement> m_refreshElement;
        private Func<IWebElement, bool> m_accessor;
        private int m_retryCount;

        public StaleSafeAccess(Func<IWebElement> refreshElement, Func<IWebElement, bool> accessor, int retryCount = 10)
        {
            m_refreshElement = refreshElement;
            m_accessor = accessor;
            m_retryCount = retryCount;
        }

        public bool Access()
        {
            int retries = 0;
            IWebElement element;

            while (true)
            {
                element = m_refreshElement();

                try
                {
                    return m_accessor(element);
                }
                catch (OpenQA.Selenium.StaleElementReferenceException)
                {
                    if (retries > m_retryCount)
                        throw;
                    retries++;
                }
            }
        }
    }

    public class WebControl
    {
        private IStatusReporter m_iStatusReporter;

        public IWebDriver Driver { get; }
        public string DownloadPath { get; set; }

        private const string CKE_Hack = @"
        console.log('loading...');
        if (window)
        {
            window.addEventListener('load',
                function()
                {
                    console.log('doc: window=' + window);
                    if (window.CKEDITOR)
                    {
                        console.log('doc: window.CKEDITOR=' + window.CKEDITOR);
                        if (window.CKEDITOR.config)
                        {
                            console.log('before: window.CKEDITOR.config.startupMode=' + window.CKEDITOR.config.startupMode);
                            window.CKEDITOR.config.startupMode = 'source';
                            console.log('after: window.CKEDITOR.config.startupMode=' + window.CKEDITOR.config.startupMode);
                        }
                    }
                });
        }";

        public WebControl(IStatusReporter iStatusReporter, bool fShowUI)
        {
            m_iStatusReporter = iStatusReporter;
            ChromeOptions options = new ChromeOptions();
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();

            if (fShowUI == false)
            {
                options.AddArguments("--window-size=1920,1080");
                options.AddArguments("--start-maximized");
                options.AddArgument("--headless");
                service.HideCommandPromptWindow = true;
            }

            options.AddArgument(
                "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36");

            DownloadPath = $"{Environment.GetEnvironmentVariable("TEMP")}\\arb-{Guid.NewGuid().ToString()}";
            Directory.CreateDirectory(DownloadPath);

            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("download.default_directory", DownloadPath);

            // Dear future self: When you get an exception here because the cromedriver doesn't match chrome, go to https://chromedriver.chromium.org/downloads
            Driver = new ChromeDriver(service, options);
            Driver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(10.0);
            Driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromMinutes(10.0);

            IDevTools devTools = Driver as IDevTools;
            DevToolsSession session = devTools.GetDevToolsSession();

            DevToolsSessionDomains domains = session.GetVersionSpecificDomains<DevToolsSessionDomains>();

            domains.Page.Enable(new OpenQA.Selenium.DevTools.V148.Page.EnableCommandSettings());
            domains.Page.AddScriptToEvaluateOnLoad(
                new AddScriptToEvaluateOnLoadCommandSettings()
                {
                    ScriptSource = "Object.defineProperty(navigator, 'webdriver', {get: ()=>undefined})"
                });
            domains.Page.AddScriptToEvaluateOnLoad(
                new AddScriptToEvaluateOnLoadCommandSettings()
                {
                    ScriptSource = CKE_Hack
                });
        }

#region Page Navigation

        /*----------------------------------------------------------------------------
            %%Function:FNavToPage
            %%Qualified:TCore.WebControl.WebControl.FNavToPage

        ----------------------------------------------------------------------------*/
        public bool FNavToPage(string sUrl)
        {
            MicroTimer timer = new MicroTimer();

            try
            {
                Driver.Navigate().GoToUrl(sUrl);
            }
            catch (Exception)
            {
                return false;
            }

            timer.Stop();
            m_iStatusReporter.LogData($"FNavToPage({sUrl}) elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            return true;
        }

        public bool FNavToPageIfNecessary(string sUrl)
        {
            if (string.Compare(Driver.Url, sUrl, StringComparison.CurrentCultureIgnoreCase) == 0)
                return true;

            MicroTimer timer = new MicroTimer();

            try
            {
                Driver.Navigate().GoToUrl(sUrl);
            }
            catch (Exception)
            {
                return false;
            }

            timer.Stop();
            m_iStatusReporter.LogData($"FNavToPage({sUrl}) elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            return true;
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForControl
            %%Qualified:TCore.WebControl.WebControl.WaitForControl
        ----------------------------------------------------------------------------*/
        public static bool WaitForControl(IWebDriver driver, IStatusReporter srpt, string sid)
        {
            if (sid == null)
                return true;

            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
            IWebElement element = wait.Until(theDriver => theDriver.FindElement(By.Id(sid)));

            return element != null;
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForControlName
            %%Qualified:TCore.WebControl.WebControl.WaitForControlName
        ----------------------------------------------------------------------------*/
        public static bool WaitForControlName(IWebDriver driver, IStatusReporter srpt, string name)
        {
            if (name == null)
                return true;

            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
            IWebElement element = wait.Until(theDriver => theDriver.FindElement(By.Name(name)));

            return element != null;
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForPageLoad
            %%Qualified:TCore.WebControl.WebControl.WaitForPageLoad
        ----------------------------------------------------------------------------*/
        public static void WaitForPageLoad(IStatusReporter srpt, IWebDriver driver, int maxWaitTimeInSeconds)
        {
            MicroTimer timer = new MicroTimer();

            string state = string.Empty;
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

                //Checks every 500 ms whether predicate returns true if returns exit otherwise keep trying till it returns ture
                wait.Until(
                    d =>
                    {
                        // use d instead of driver below?

                        try
                        {
                            state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                        }
                        catch (InvalidOperationException)
                        {
                            //Ignore
                        }
                        catch (NoSuchWindowException)
                        {
                            //when popup is closed, switch to last windows
                            driver.SwitchTo().Window(driver.WindowHandles[driver.WindowHandles.Count - 1]);
                        }

                        //In IE7 there are chances we may get state as loaded instead of complete
                        return state.Equals("complete", StringComparison.InvariantCultureIgnoreCase)
                            || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase);
                    });
            }
            catch (TimeoutException)
            {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            }
            catch (NullReferenceException)
            {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            }
            catch (WebDriverException)
            {
                if (driver.WindowHandles.Count == 1)
                {
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                }

                state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                if (!(state.Equals("complete", StringComparison.InvariantCultureIgnoreCase)
                        || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase)))
                    throw;
            }

            timer.Stop();
            srpt.LogData($"WaitForPageLoad elapsed: {timer.MsecFloat}", 1, MSGT.Body);
        }

        public void WaitForPageLoad(int maxWaitTimeInSeconds = 500) => WaitForPageLoad(m_iStatusReporter, Driver, maxWaitTimeInSeconds);

        public delegate bool WaitRepeatDelegate(IWebDriver driver);
        public delegate bool WaitUntilDelegate(IWebDriver driver);

        /*----------------------------------------------------------------------------
            %%Function:WaitForXpath
            %%Qualified:TCore.WebControl.WebControl.WaitForXpath
        ----------------------------------------------------------------------------*/
        public void WaitForXpath(string xpath, int msecTimeout = 500)
        {
            WaitForConditionDelegate(
                (d) => { return (d.FindElement(By.XPath(xpath)) != null); },
                1000);
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForConditionDelegate
            %%Qualified:TCore.WebControl.WebControl.WaitForConditionDelegate
        ----------------------------------------------------------------------------*/
        public void WaitForConditionDelegate(WaitUntilDelegate waitUntilDelegate, int msecTimeout = 500)
        {
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromMilliseconds(msecTimeout));

            wait.Until(
                (d) =>
                {
                    try
                    {
                        return waitUntilDelegate(d);
                    }
                    catch
                    {
                        return false;
                    }
                });
        }

        /*----------------------------------------------------------------------------
            %%Function:RepeatIfNotCondition
            %%Qualified:TCore.WebControl.WebControl.RepeatIfNotCondition
        ----------------------------------------------------------------------------*/
        public void RepeatIfNotCondition(WaitRepeatDelegate repeat, WaitUntilDelegate until, int repeatCount, int msecTimeout = 500)
        {
            WebDriverWait waitFirst = new WebDriverWait(Driver, TimeSpan.FromMilliseconds(10));
            WebDriverWait waitRepeat = new WebDriverWait(Driver, TimeSpan.FromMilliseconds(msecTimeout));

            bool fDone = false;

            fDone = waitFirst.Until(
                driver =>
                {
                    try
                    {
                        return until(driver);
                    }
                    catch
                    {
                        return false;
                    }
                });

            while (fDone && repeatCount > 0)
            {
                if (!repeat(Driver))
                    throw new Exception($"couldn't execute repeat of repeat if not");

                fDone = waitRepeat.Until(
                    driver =>
                    {
                        try
                        {
                            return until(driver);
                        }
                        catch
                        {
                            return false;
                        }
                    });

                repeatCount--;
            }

            if (!fDone)
                throw new Exception("repeat if not failed to complete condition");
        }

        /*----------------------------------------------------------------------------
            %%Function:GetElementBy
            %%Qualified:TCore.WebControl.WebControl.GetElementBy
        ----------------------------------------------------------------------------*/
        public IWebElement GetElementBy(OpenQA.Selenium.By by) => GetElementBy(Driver, by);

        /*----------------------------------------------------------------------------
            %%Function:GetElementBy
            %%Qualified:TCore.WebControl.WebControl.GetElementBy
        ----------------------------------------------------------------------------*/
        public static IWebElement GetElementBy(IWebDriver driver, OpenQA.Selenium.By by)
        {
            try
            {
                return driver.FindElement(by);
            }
            catch
            {
                return null;
            }
        }

        public static void WaitForCondition(
            IWebDriver driver, System.Func<OpenQA.Selenium.IWebDriver, ReadOnlyCollection<OpenQA.Selenium.IWebElement>> until,
            int msecTimeout = 500)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(msecTimeout));

            wait.Until(until);
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForCondition
            %%Qualified:TCore.WebControl.WebControl.WaitForCondition
        ----------------------------------------------------------------------------*/
        public void WaitForCondition(System.Func<OpenQA.Selenium.IWebDriver, ReadOnlyCollection<OpenQA.Selenium.IWebElement>> until, int msecTimeout = 500)
        {
            WaitForCondition(Driver, until, msecTimeout);
        }

        public static void WaitForCondition(
            IWebDriver driver, System.Func<OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement> until, int msecTimeout = 500)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(msecTimeout));

            wait.Until(until);
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForCondition
            %%Qualified:TCore.WebControl.WebControl.WaitForCondition
        ----------------------------------------------------------------------------*/
        public void WaitForCondition(System.Func<OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement> until, int msecTimeout = 500)
        {
            WaitForCondition(Driver, until, msecTimeout);
        }

        public static void WaitForCondition(IWebDriver driver, System.Func<OpenQA.Selenium.IWebDriver, bool> until, int msecTimeout = 500)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(msecTimeout));

            wait.Until(until);
        }

        /*----------------------------------------------------------------------------
            %%Function:WaitForCondition
            %%Qualified:TCore.WebControl.WebControl.WaitForCondition
        ----------------------------------------------------------------------------*/
        public void WaitForCondition(System.Func<OpenQA.Selenium.IWebDriver, bool> until, int msecTimeout = 500)
        {
            WaitForCondition(Driver, until, msecTimeout);
        }

#endregion

#region Individual Control Interaction

        /*----------------------------------------------------------------------------
            %%Function:FCheckForControlId
            %%Qualified:TCore.WebControl.WebControl.FCheckForControlId
        ----------------------------------------------------------------------------*/
        public static bool FCheckForControlId(IWebDriver driver, string sid)
        {
            IWebElement element;

            try
            {
                element = driver.FindElement(By.Id(sid));
            }
            catch (OpenQA.Selenium.NoSuchElementException)
            {
                return false;
            }

            return element != null;
        }

        public bool FCheckForControlId(string sid) => FCheckForControlId(Driver, sid);

        /*----------------------------------------------------------------------------
            %%Function:FClickControl
            %%Qualified:TCore.WebControl.WebControl.FClickControl
        ----------------------------------------------------------------------------*/
        public static bool FClickControl(IStatusReporter srpt, IWebDriver driver, IWebElement element, string sidWaitFor = null)
        {
            try
            {
                element?.Click();
            }
            catch (OpenQA.Selenium.WebDriverException e)
            {
                srpt.AddMessage($"Ignoring webdriver exception: {e.Message}");
            }
            catch (Exception e)
            {
                throw e;
            }

            if (sidWaitFor != null)
                return WaitForControl(driver, srpt, sidWaitFor);

            WaitForPageLoad(srpt, driver, 12000);
            return true;
        }

        public bool FCheckForXpath(string xpath)
        {
            try
            {
                Driver.FindElement(By.XPath(xpath));
                return true;
            }
            catch
            {
                return false;
            }
        }

        /*----------------------------------------------------------------------------
            %%Function:FClickControlByXpath
            %%Qualified:TCore.WebControl.WebControl.FClickControlByXpath
        ----------------------------------------------------------------------------*/
        public bool FClickControlByXpath(string xpath, bool ensureClickable = false, int msecTimeout = 500)
        {
            if (ensureClickable)
            {
                try
                {
                    WaitForCondition(ExpectedConditions.ElementToBeClickable(By.XPath(xpath)), msecTimeout);
                }
                catch
                {
                    return false;
                }
            }

            IWebElement element = Driver.FindElement(By.XPath(xpath));

            if (element == null)
                return false;

            return FClickControl(m_iStatusReporter, Driver, element);
        }

        /*----------------------------------------------------------------------------
            %%Function:FClickControlName
            %%Qualified:TCore.WebControl.WebControl.FClickControlName
        ----------------------------------------------------------------------------*/
        public bool FClickControlName(string sName, string sidWaitFor = null)
        {
            m_iStatusReporter.LogData($"FClickControlName {sName}", 5, MSGT.Body);

            return FClickControl(m_iStatusReporter, Driver, Driver.FindElement(By.Name(sName)));
        }

        /*----------------------------------------------------------------------------
            %%Function:FClickControlId
            %%Qualified:TCore.WebControl.WebControl.FClickControlId
        ----------------------------------------------------------------------------*/
        public bool FClickControlId(string sid, string sidWaitFor = null)
        {
            m_iStatusReporter.LogData($"FClickControlId {sid}", 5, MSGT.Body);

            return FClickControl(m_iStatusReporter, Driver, Driver.FindElement(By.Id(sid)));
        }

        /*----------------------------------------------------------------------------
            %%Function:FSetTextForInputControl
            %%Qualified:TCore.WebControl.WebControl.FSetTextForInputControl
        ----------------------------------------------------------------------------*/
        public static bool FSetTextForInputControl(IWebElement element, string sValue, bool fCheck)
        {
            if (element == null)
                return false;

            string sOriginalValue = fCheck ? element.GetAttribute("value") : null;

            try
            {
                element.Clear();
                element.SendKeys(sValue);
            }
            catch (Exception e)
            {
                return false;
            }


            if (fCheck)
                return String.Compare(sOriginalValue, sValue) != 0;

            return !fCheck;
        }

        /*----------------------------------------------------------------------------
            %%Function:FSetTextForAntInputControl
            %%Qualified:TCore.WebControl.WebControl.FSetTextForAntInputControl
        ----------------------------------------------------------------------------*/
        public static bool FSetTextForAntInputControl(IWebElement element, string sValue, bool fCheck)
        {
            if (element == null)
                return false;

            string sOriginalValue = fCheck ? element.GetAttribute("value") : null;

            try
            {
                element.SendKeys(Keys.Control + "A");
                element.SendKeys(Keys.Backspace);
                element.SendKeys(sValue);
                element.SendKeys(Keys.Return);
            }
            catch (Exception e)
            {
                return false;
            }


            if (fCheck)
                return String.Compare(sOriginalValue, sValue) != 0;

            return !fCheck;
        }

        /*----------------------------------------------------------------------------
            %%Function:FSetTextForAntInputControlId
            %%Qualified:TCore.WebControl.WebControl.FSetTextForAntInputControlId
        ----------------------------------------------------------------------------*/
        public static bool FSetTextForAntInputControlId(IWebDriver driver, string id, string sValue, bool fCheck) =>
            FSetTextForAntInputControl(driver.FindElement(By.Id(id)), sValue, fCheck);

        /*----------------------------------------------------------------------------
            %%Function:FSetTextForInputControlName
            %%Qualified:TCore.WebControl.WebControl.FSetTextForInputControlName
        ----------------------------------------------------------------------------*/
        public static bool FSetTextForInputControlName(IWebDriver driver, string sName, string sValue, bool fCheck) =>
            FSetTextForInputControl(driver.FindElement(By.Name(sName)), sValue, fCheck);

        /*----------------------------------------------------------------------------
            %%Function:FSetTextForInputControlId
            %%Qualified:TCore.WebControl.WebControl.FSetTextForInputControlId
        ----------------------------------------------------------------------------*/
        public static bool FSetTextForInputControlId(IWebDriver driver, string id, string sValue, bool fCheck, IWebElement parent = null)
        {
            if (parent != null)
                return FSetTextForInputControl(parent.FindElement(By.Id(id)), sValue, fCheck);

            return FSetTextForInputControl(driver.FindElement(By.Id(id)), sValue, fCheck);
        }

        public bool FSetTextForInputControlName(string sName, string sValue, bool fCheck) => FSetTextForInputControlName(Driver, sName, sValue, fCheck);

        public bool FSetTextForInputControlId(string id, string sValue, bool fCheck, IWebElement parent = null) =>
            FSetTextForInputControlId(Driver, id, sValue, fCheck, parent);

        public bool FSetTextForAntInputControlId(string id, string sValue, bool fCheck) => FSetTextForAntInputControlId(Driver, id, sValue, fCheck);

        /*----------------------------------------------------------------------------
            %%Function:FSetCheckboxControlNameVal
            %%Qualified:TCore.WebControl.WebControl.FSetCheckboxControlNameVal
        ----------------------------------------------------------------------------*/
        public static bool FSetCheckboxControlNameVal(IWebDriver driver, bool fChecked, string sName)
        {
            IWebElement element = driver.FindElement(By.Name(sName));
            string sValue = fChecked ? "true" : "false";
            string sActual = element.GetAttribute("checked");

            if (String.Compare(element.GetAttribute("checked"), sValue, true) != 0)
            {
                element.Click();
                return true;
            }

            return false;
        }

        public bool FSetCheckboxControlNameVal(bool fChecked, string sName) => FSetCheckboxControlNameVal(Driver, fChecked, sName);

        /*----------------------------------------------------------------------------
            %%Function: GetCurrentCheckedValueFromElement
            %%Qualified: TCore.WebControl.WebControl.GetCurrentCheckedValueFromElement
        ----------------------------------------------------------------------------*/
        private static string GetCurrentCheckedValueFromElement(IWebElement element)
        {
            string sActual = element.GetAttribute("checked");

            if (sActual == null)
                sActual = "false";

            return sActual;
        }

        private static bool FRetryWhenStale(IWebDriver driver, Func<IWebDriver, bool> func, TimeSpan timeout)
        {
            bool returnVal = false;

            WebDriverWait wait = new WebDriverWait(driver, timeout);

            wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
            wait.Until(
                d =>
                {
                    returnVal = func(d);
                    return true;
                });

            return returnVal;
        }

        /*----------------------------------------------------------------------------
            %%Function:FSetCheckboxControlIdVal
            %%Qualified:ArbWeb.ArbWebControl_Selenium.FSetCheckboxControlIdVal
        ----------------------------------------------------------------------------*/
        public static bool FSetCheckboxControlIdVal(IWebDriver driver, bool fChecked, string sid)
        {
            return new StaleSafeAccess(
                () => driver.FindElement(By.Id(sid)),
                element =>
                {
                    string sValue = fChecked ? "true" : "false";
                    string sActual = GetCurrentCheckedValueFromElement(element);

                    Console.WriteLine($"{sid} value before: {sActual}");

                    if (String.Compare(sActual, sValue, true) != 0)
                    {
                        element.Click();
                        sActual = GetCurrentCheckedValueFromElement(element);

                        Console.WriteLine($"{sid} value after: {sActual}");
                        return true;
                    }

                    return false;
                }).Access();
        }

        public bool FSetCheckboxControlIdVal(bool fChecked, string sid) => FSetCheckboxControlIdVal(Driver, fChecked, sid);

        /* S  G E T  C O N T R O L  V A L U E */
        /*----------------------------------------------------------------------------
            %%Function: SGetControlValue
            %%Qualified: ArbWeb.ArbWebControl.SGetControlValue
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static string GetValueForControlId(IWebDriver driver, string sId)
        {
            IWebElement control = driver.FindElement(By.Id(sId));

            return control.GetAttribute("value");
        }

        public IWebElement GetPreviousSiblingFromAntControlId(string sId, string siblingMatch)
        {
            try
            {
                string xpath = $"//*[@id='{sId}']/preceding::{siblingMatch}";

                return Driver.FindElement(By.XPath(xpath));
            }
            catch
            {
                return null;
            }
        }

        public IWebElement GetSiblingFromAntControlId(IWebDriver driver, string sId, string siblingMatch)
        {
            try
            {
                string xpath = $"//*[@id='{sId}']/../following-sibling::{siblingMatch}";

                return driver.FindElement(By.XPath(xpath));
            }
            catch
            {
                return null;
            }
        }

        /*----------------------------------------------------------------------------
            %%Function: GetSelectedValueFromAntControlId
            %%Qualified: TCore.WebControl.WebControl.GetSelectedValueFromAntControlId

            This will get the selected value for an Ant control id
        ----------------------------------------------------------------------------*/
        public string GetSelectedValueFromAntControlId(IWebDriver driver, string sId, string siblingMatch)
        {
            try
            {
                IWebElement element = GetSiblingFromAntControlId(driver, sId, siblingMatch);

                if (element == null)
                    return null;

                return element.Text;
            }
            catch
            {
                return null;
            }
        }

        public string GetTextFromControlId(IWebDriver driver, string sId)
        {
            try
            {
                IWebElement element = driver.FindElement(By.Id(sId));

                if (element == null)
                    return null;

                return element.Text;
            }
            catch
            {
                return null;
            }
        }

        public string GetValueForControlIdOrNull(string sId)
        {
            try
            {
                return GetValueForControlId(Driver, sId);
            }
            catch
            {
                return null;
            }
        }

        public string GetValueForControlId(string sId) => GetValueForControlId(Driver, sId);

        public static bool FSetTextAreaTextForControlName(IWebDriver driver, string sName, string sValue, bool fCheck)
        {
            IWebElement element = driver.FindElement(By.Name(sName));

            string sOriginal = null;

            if (fCheck)
                sOriginal = element.GetAttribute("value");

            element.Clear();
            element.SendKeys(sValue);

            if (fCheck)
                return String.Compare(sValue, sOriginal, true) != 0;

            return true;
        }

        public bool FSetTextAreaTextForControlName(string sName, string sValue, bool fCheck) => FSetTextAreaTextForControlName(Driver, sName, sValue, fCheck);

        public bool FSetTextAreaTextForControlAsChildOfDivId(string sDivId, string sValue, bool fCheck) =>
            FSetTextAreaTextForControlAsChildOfDivId(Driver, sDivId, sValue, fCheck);

        public static bool FSetTextAreaTextForControlAsChildOfDivId(IWebDriver driver, string sDivId, string sValue, bool fCheck)
        {
            IWebElement div = driver.FindElement(By.Id(sDivId));

            // and now get the first child
            IWebElement element = div.FindElement(By.TagName("textarea"));

            string sOriginal = null;

            if (fCheck)
                sOriginal = element.GetAttribute("value");

            element.Clear();
            string encoded = sValue.Replace("\"", "\\\"").Replace("\r\n", "\\r\\n");
            string js = $"arguments[0].value=\"{encoded}\"";
            driver.ExecuteJavaScript(js, element);

            if (fCheck)
                return String.Compare(sValue, sOriginal, true) != 0;

            return true;
        }

        public static bool FClickSubmitControlByValue(IWebDriver driver, string sValueText)
        {
            string sXPath;

            if (sValueText != null)
                sXPath = $"//input[@type='submit' and @value='{sValueText}']";
            else
                sXPath = "//input[@type='submit']";

            IWebElement element = driver.FindElement(By.XPath(sXPath));
            if (element == null)
                return false;

            element.Click();
            return true;
        }

        public bool FClickSubmitControlByValue(string sValueText) => FClickSubmitControlByValue(Driver, sValueText);

        public string GetOuterHtmlForControlByXPath(string sXPath)
        {
            return Driver.FindElement(By.XPath(sXPath)).GetAttribute("outerHTML");
        }

        public static string GetAttributeValueFromId(IWebDriver driver, string id, string attribute)
        {
            IWebElement element = driver.FindElement(By.Id(id));

            if (element == null) return null;

            return element.GetAttribute(attribute);
        }

        public string GetAttributeValueFromId(string id, string attribute)
        {
            return GetAttributeValueFromId(Driver, id, attribute);
        }

        public static string GetAttributeValueFromXpath(IWebDriver driver, string xpath, string attribute)
        {
            IWebElement element = driver.FindElement(By.XPath(xpath));

            if (element == null) return null;

            return element.GetAttribute(attribute);
        }

        public string GetAttributeValueFromXpath(string xpath, string attribute)
        {
            return GetAttributeValueFromXpath(Driver, xpath, attribute);
        }

#endregion

#region Select/Option Interaction

        /*----------------------------------------------------------------------------
            %%Function:GetOptionValueFromFilterOptionText
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionValueFromFilterOptionText
        ----------------------------------------------------------------------------*/
        public string GetOptionValueFromFilterOptionTextForControlName(string sName, string sOptionText)
        {
            m_iStatusReporter.LogData($"SGetSelectIDFromDoc for id {sName}", 3, MSGT.Body);

            string s = GetOptionValueForSelectControlNameOptionText(sName, sOptionText);

            m_iStatusReporter.LogData($"Return: {s}", 3, MSGT.Body);
            return s;
        }

        /*----------------------------------------------------------------------------
            %%Function:GetOptionTextFromOptionValue
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionTextFromOptionValue
        ----------------------------------------------------------------------------*/
        public static string GetOptionTextFromOptionValueForControlName(IWebDriver driver, IStatusReporter srpt, string sName, string sOptionValue)
        {
            IWebElement selectElement = driver.FindElement(By.Name(sName));

            Dictionary<string, string> mpValueText = GetOptionsValueTextMappingFromControl(selectElement, srpt);
            if (mpValueText.ContainsKey(sOptionValue))
                return mpValueText[sOptionValue];

            return null;
        }

        public string GetOptionTextFromOptionValueForControlName(string sName, string sOptionName) =>
            GetOptionTextFromOptionValueForControlName(Driver, m_iStatusReporter, sName, sOptionName);

        /*----------------------------------------------------------------------------
            %%Function:FSetSelectedOptionTextForControl
            %%Qualified:TCore.WebControl.WebControl.FSetSelectedOptionTextForControl
        ----------------------------------------------------------------------------*/
        public static bool FSetSelectedOptionTextForControl(IWebElement selectElement, IStatusReporter srpt, string sOptionText)
        {
            SelectElement select = new SelectElement(selectElement);
            string sOriginal = select.SelectedOption.Text;
            bool fChanged = false;

            if (String.Compare(sOriginal, sOptionText, true) != 0)
            {
                select.SelectByText(sOptionText);
                fChanged = true;
            }

            srpt.LogData($"Return: {fChanged}", 5, MSGT.Body);

            return fChanged;
        }

        /*----------------------------------------------------------------------------
            %%Function:FSetSelectControlText
            %%Qualified:ArbWeb.ArbWebControl_Selenium.FSetSelectControlText
        ----------------------------------------------------------------------------*/
        public static bool FSetSelectedOptionTextForControlId(IWebDriver driver, IStatusReporter srpt, string sid, string sOptionText)
        {
            srpt.LogData($"FSetSelectControlText for id {sid}", 5, MSGT.Body);

            return new StaleSafeAccess(
                () => driver.FindElement(By.Id(sid)),
                element => FSetSelectedOptionTextForControl(element, srpt, sOptionText)).Access();
        }

        public bool FSetSelectedOptionTextForControlId(string sid, string sValue) => FSetSelectedOptionTextForControlId(Driver, m_iStatusReporter, sid, sValue);

        public static bool FSetSelectedOptionTextForControlName(IWebDriver driver, IStatusReporter srpt, string sName, string sOptionText)
        {
            srpt.LogData($"FSetSelectControlText for id {sName}", 5, MSGT.Body);

            return FSetSelectedOptionTextForControl(driver.FindElement(By.Name(sName)), srpt, sOptionText);
        }

        public bool FSetSelectedOptionTextForControlName(string sName, string sOptionText) =>
            FSetSelectedOptionTextForControlName(Driver, m_iStatusReporter, sName, sOptionText);

        /*----------------------------------------------------------------------------
            %%Function:FSetSelectControlValue
            %%Qualified:ArbWeb.ArbWebControl_Selenium.FSetSelectControlValue
        ----------------------------------------------------------------------------*/
        public static bool FSetSelectedOptionValueForControlName(IWebDriver driver, IStatusReporter srpt, string sName, string sValue)
        {
            srpt.LogData($"FSetSelectControlValue for name {sName}", 5, MSGT.Body);

#if DOWAIT
            try
            {
                WaitForCondition(driver, ExpectedConditions.ElementToBeClickable(By.Name(sName)), 2000);
            }
            catch (Exception e)
            {
                srpt.AddMessage($"FSetSelectControlValue: couldn't find control {sName} to set value {sValue}: {e.Message}");
                return false;
            }
#endif
            return new StaleSafeAccess(
                () => driver.FindElement(By.Name(sName)),
                element =>
                {
                    SelectElement select = new SelectElement(element);

                    string sOriginal = select.SelectedOption.GetAttribute("value");
                    bool fChanged = false;

                    if (String.Compare(sOriginal, sValue, true) != 0)
                    {
                        try
                        {
                            select.SelectByValue(sValue);
                        }
                        catch
                        {
                            return false;
                        }

                        fChanged = true;
                    }

                    srpt.LogData($"Return: {fChanged}", 5, MSGT.Body);

                    return fChanged;
                }).Access();
        }

        public bool FSetSelectedOptionValueForControlName(string sName, string sValue) =>
            FSetSelectedOptionValueForControlName(Driver, m_iStatusReporter, sName, sValue);

        /*----------------------------------------------------------------------------
            %%Function:GetSelectedOptionTextFromSelectControl
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetSelectedOptionTextFromSelectControl
        ----------------------------------------------------------------------------*/
        public static string GetSelectedOptionTextFromSelectControlName(IWebDriver driver, string sName)
        {
            IWebElement selectElement = driver.FindElement(By.Name(sName));
            SelectElement select = new SelectElement(selectElement);

            return select.SelectedOption?.Text;
        }

        public string GetSelectedOptionTextFromSelectControlName(string sName) => GetSelectedOptionTextFromSelectControlName(Driver, sName);

        /*----------------------------------------------------------------------------
            %%Function:GetSelectedOptionValueFromSelectControlName
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetSelectedOptionValueFromSelectControlName
        ----------------------------------------------------------------------------*/
        public static string GetSelectedOptionValueFromSelectControlName(IWebDriver driver, string sName)
        {
            IWebElement selectElement = driver.FindElement(By.Name(sName));
            SelectElement select = new SelectElement(selectElement);

            return select.SelectedOption.GetAttribute("value");
        }

        public string GetSelectedOptionValueFromSelectControlName(string sName) => GetSelectedOptionValueFromSelectControlName(Driver, sName);

        /*----------------------------------------------------------------------------
            %%Function:GetOptionValueForSelectControlOptionText
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionValueForSelectControlOptionText
        ----------------------------------------------------------------------------*/
        public static string GetOptionValueForSelectControlNameOptionText(IWebDriver driver, string sName, string sOptionText)
        {
            IWebElement selectElement = driver.FindElement(By.Name(sName));
            Dictionary<string, string> mpValueText = GetOptionsValueTextMappingFromControl(selectElement, null);

            foreach (string sKey in mpValueText.Keys)
            {
                if (String.Compare(mpValueText[sKey], sOptionText, true) == 0)
                    return sKey;
            }

            return null;
        }

        public string GetOptionValueForSelectControlNameOptionText(string sName, string sOptionText) =>
            GetOptionValueForSelectControlNameOptionText(Driver, sName, sOptionText);

        /*----------------------------------------------------------------------------
            %%Function:MpGetSelectValuesFromControl
            %%Qualified:ArbWeb.ArbWebControl_Selenium.MpGetSelectValuesFromControl
        ----------------------------------------------------------------------------*/
        private static Dictionary<string, string> GetOptionsValueTextMappingFromControl(
            IWebElement selectElement,
            IStatusReporter srpt)
        {
            string sHtml = selectElement.GetAttribute("innerHTML");

            HtmlDocument html = new HtmlDocument();
            html.LoadHtml(sHtml);

            HtmlNodeCollection options = html.DocumentNode.SelectNodes("//option");
            Dictionary<string, string> mp = new Dictionary<string, string>();

            if (options != null)
            {
                foreach (HtmlNode option in options)
                {
                    string sValue = option.GetAttributeValue("value", null);
                    string sText = option.InnerText.Trim();

                    if (mp.ContainsKey(sValue))
                        srpt?.AddMessage(
                            $"How strange!  Option '{sValue}' shows up more than once in the options list!",
                            MSGT.Warning);
                    else
                        mp.Add(sValue, sText);
                }
            }

            return mp;
        }

        /* M P  G E T  S E L E C T  V A L U E S */
        /*----------------------------------------------------------------------------
            %%Function: MpGetSelectValues
            %%Qualified: ArbWeb.AwMainForm.MpGetSelectValues
            %%Contact: rlittle

            for a given <select name=$sName><option value=$sValue>$sText</option>...

            Find the given sName select object. Then add a mapping of
            $sText -> $sValue to a dictionary and return it.
        ----------------------------------------------------------------------------*/
        public static Dictionary<string, string> GetOptionsValueTextMappingFromControlId(IWebDriver driver, IStatusReporter srpt, string sid)
        {
            MicroTimer timer = new MicroTimer();

            Dictionary<string, string> mp = GetOptionsValueTextMappingFromControl(driver.FindElement(By.Id(sid)), srpt);

            timer.Stop();
            srpt.LogData($"MpGetSelectValues({sid}) elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            return mp;
        }

        public Dictionary<string, string> GetOptionsValueTextMappingFromControlId(string sid) =>
            GetOptionsValueTextMappingFromControlId(Driver, m_iStatusReporter, sid);

        /*----------------------------------------------------------------------------
            %%Function:GetOptionsTextValueMappingFromControl
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionsTextValueMappingFromControl
        ----------------------------------------------------------------------------*/
        private static Dictionary<string, string> GetOptionsTextValueMappingFromControl(
            IWebElement selectElement,
            IStatusReporter srpt)
        {
            string sHtml = selectElement.GetAttribute("innerHTML");

            HtmlDocument html = new HtmlDocument();
            html.LoadHtml(sHtml);

            HtmlNodeCollection options = html.DocumentNode.SelectNodes("//option");
            Dictionary<string, string> mp = new Dictionary<string, string>();

            if (options != null)
            {
                foreach (HtmlNode option in options)
                {
                    string sValue = option.GetAttributeValue("value", null);
                    string sText = option.InnerText.Trim();

                    if (mp.ContainsKey(sText))
                        srpt?.AddMessage(
                            $"How strange!  Option '{sText}' shows up more than once in the options list!",
                            MSGT.Warning);
                    else
                        mp.Add(sText, sValue);
                }
            }

            return mp;
        }

        /*----------------------------------------------------------------------------
            %%Function:GetOptionsTextValueMappingFromControlId
            %%Qualified:ArbWeb.ArbWebControl_Selenium.GetOptionsTextValueMappingFromControlId
        ----------------------------------------------------------------------------*/
        public static Dictionary<string, string> GetOptionsTextValueMappingFromControlId(IWebDriver driver, IStatusReporter srpt, string sid)
        {
            MicroTimer timer = new MicroTimer();
            Dictionary<string, string> mp = new Dictionary<string, string>();

            if (new StaleSafeAccess(
                    () => driver.FindElement(By.Id(sid)),
                    element =>
                    {
                        mp = GetOptionsTextValueMappingFromControl(element, srpt);

                        timer.Stop();
                        srpt.LogData($"GetOptionsTextValueMappingFromControlId({sid}) elapsed: {timer.MsecFloat}", 1, MSGT.Body);
                        return true;
                    }).Access())
            {
                return mp;
            }

            throw new Exception("stalesafeaccess did not succeed");
        }

        public Dictionary<string, string> GetOptionsTextValueMappingFromControlId(string sid) =>
            GetOptionsTextValueMappingFromControlId(Driver, m_iStatusReporter, sid);

        public static Dictionary<string, string> GetOptionsTextValueMappingFromControlName(IWebDriver driver, IStatusReporter srpt, string sName)
        {
            MicroTimer timer = new MicroTimer();

            Dictionary<string, string> mp = GetOptionsTextValueMappingFromControl(driver.FindElement(By.Name(sName)), srpt);

            timer.Stop();
            srpt.LogData($"GetOptionsTextValueMappingFromControlId({sName}) elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            return mp;
        }

        public Dictionary<string, string> GetOptionsTextValueMappingFromControlName(string sName) =>
            GetOptionsTextValueMappingFromControlName(Driver, m_iStatusReporter, sName);

        /*----------------------------------------------------------------------------
            %%Function:FResetMultiSelectOptions
            %%Qualified:ArbWeb.ArbWebControl_Selenium.FResetMultiSelectOptions

            Uncheck all of the items for this multiselect control
        ----------------------------------------------------------------------------*/
        public static bool FResetMultiSelectOptionsForControlName(IWebDriver driver, string sName)
        {
            return
                new StaleSafeAccess(
                        () => driver.FindElement(
                            By.Name(sName)),
                        element =>
                        {
                            SelectElement select = new SelectElement(element);

                            select.DeselectAll();

                            return true;
                        })
                   .Access();
        }

        public bool FResetMultiSelectOptionsForControlName(string sName) => FResetMultiSelectOptionsForControlName(Driver, sName);

        // if fValueIsValue == false, then sValue is the "text" of the option control
        /* F  S E L E C T  M U L T I  S E L E C T  O P T I O N */
        /*----------------------------------------------------------------------------
            %%Function: FSelectMultiSelectOption
            %%Qualified: ArbWeb.ArbWebControl.FSelectMultiSelectOption
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static bool FSelectMultiSelectOptionValueForControlName(IWebDriver driver, string sName, string sValue)
        {
            IWebElement selectElement = driver.FindElement(By.Name(sName));
            SelectElement select = new SelectElement(selectElement);

            try
            {
                select.SelectByValue(sValue);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool FSelectMultiSelectOptionValueForControlName(string sName, string sValue) =>
            FSelectMultiSelectOptionValueForControlName(Driver, sName, sValue);

#endregion

#region File Downloader

        public class FileDownloader
        {
            public delegate void StartDownload();

            private readonly string m_sExpectedFullName;
            private readonly string[] m_expectedFullNameTemplates;

            private readonly string m_sTargetFile;
            private StartDownload m_startDownload;

            public FileDownloader(WebControl webControl, string[] expectedFileTemplates, string targetFile, StartDownload startDownload)
            {
                m_expectedFullNameTemplates = new string[expectedFileTemplates.Length];

                for (int i = 0; i < expectedFileTemplates.Length; i++)
                    m_expectedFullNameTemplates[i] = Path.Combine(webControl.DownloadPath, expectedFileTemplates[i]);

                if (targetFile == null)
                {
                    m_sTargetFile = Path.Combine(
                        webControl.DownloadPath,
                        $"{System.Guid.NewGuid().ToString()}.{Path.GetExtension(expectedFileTemplates[0])}");
                }
                else
                {
                    m_sTargetFile = targetFile;
                }

                // make sure that file doesn't already exist
                string sFile = GetDownloadedFileFromTemplates();

                if (sFile != null)
                {
                    throw new Exception(
                        $"File {sFile} already exists! our temp download directory should start out empty! Someone not cleaning up?");
                }

                m_startDownload = startDownload;
            }

            public FileDownloader(WebControl webControl, string expectedFile, string targetFile, StartDownload startDownload)
            {
                m_sExpectedFullName = Path.Combine(webControl.DownloadPath, expectedFile);
                if (targetFile == null)
                {
                    m_sTargetFile = Path.Combine(
                        webControl.DownloadPath,
                        $"{System.Guid.NewGuid().ToString()}.{Path.GetExtension(expectedFile)}");
                }
                else
                {
                    m_sTargetFile = targetFile;
                }

                // make sure that file doesn't already exist
                if (File.Exists(m_sExpectedFullName))
                {
                    throw new Exception(
                        $"File {m_sExpectedFullName} already exists! our temp download directory should start out empty! Someone not cleaning up?");
                }

                m_startDownload = startDownload;
            }

            string GetDownloadedFileFromTemplates()
            {
                if (m_expectedFullNameTemplates == null)
                {
                    if (File.Exists(m_sExpectedFullName))
                        return m_sExpectedFullName;

                    return null;
                }

                int c = 0;

                while (c < 50)
                {
                    foreach (string s in m_expectedFullNameTemplates)
                    {
                        string sFull = String.Format(s, c);
                        if (File.Exists(sFull))
                            return sFull;
                    }

                    c++;
                }

                return null;
            }

            public string GetDownloadedFile()
            {
                string sFile = null;

                m_startDownload();

                // now wait for the file to be available and non-zero
                int cRetry = 500;
                while (--cRetry > 0)
                {
                    Thread.Sleep(100);

                    sFile = GetDownloadedFileFromTemplates();

                    if (sFile != null)
                    {
                        FileInfo info = new FileInfo(sFile);

                        if (info.Length > 0)
                            break;
                    }
                }

                if (cRetry <= 0)
                    throw new Exception("file never downloaded?");

                File.Move(sFile, m_sTargetFile);
                return m_sTargetFile;
            }
        }

#endregion
    }
}
