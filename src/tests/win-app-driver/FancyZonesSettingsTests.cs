using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;

namespace PowerToysTests
{
    [TestClass]
    public class FancyZonesSettingsTests : PowerToysSession
    {
        private string _settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft/PowerToys/FancyZones/settings.json");
        private string _initialSettings;
        private JObject _initialSettingsJson;

        private static WindowsElement _saveButton;
        private static Actions _scrollDown;
        private static Actions _scrollUp;

        private static void Init()
        {
            OpenSettings();
            ShortWait();

            OpenFancyZonesSettings();

            _saveButton = session.FindElementByName("Save");
            Assert.IsNotNull(_saveButton);

            WindowsElement powerToysWindow = session.FindElementByXPath("//Window[@Name=\"PowerToys Settings\"]");
            Assert.IsNotNull(powerToysWindow);
            _scrollUp = new Actions(session).MoveToElement(_saveButton).MoveByOffset(0, _saveButton.Rect.Height).ContextClick()
                .SendKeys(OpenQA.Selenium.Keys.PageUp + OpenQA.Selenium.Keys.PageUp);
            Assert.IsNotNull(_scrollUp);
            _scrollDown = new Actions(session).MoveToElement(_saveButton).MoveByOffset(0, _saveButton.Rect.Height).ContextClick()
                .SendKeys(OpenQA.Selenium.Keys.PageDown + OpenQA.Selenium.Keys.PageDown);
            Assert.IsNotNull(_scrollDown);
        }

        private static void OpenFancyZonesSettings()
        {
            WindowsElement fzNavigationButton = session.FindElementByXPath("//Button[@Name=\"FancyZones\"]");
            Assert.IsNotNull(fzNavigationButton);

            fzNavigationButton.Click();
            fzNavigationButton.Click();

            ShortWait();
        }

        private JObject getProperties()
        {
            JObject settings = JObject.Parse(File.ReadAllText(_settingsPath));
            return settings["properties"].ToObject<JObject>();
        }

        private T getPropertyValue<T>(string propertyName)
        {
            JObject properties = getProperties();
            return properties[propertyName].ToObject<JObject>()["value"].Value<T>();
        }

        private T getPropertyValue<T>(JObject properties, string propertyName)
        {
            return properties[propertyName].ToObject<JObject>()["value"].Value<T>();
        }
        
        private void ScrollDown()
        {
            _scrollDown.Perform();
        }

        private void ScrollUp()
        {
            _scrollUp.Perform();
        }

        private void SaveChanges()
        {
            string isEnabled = _saveButton.GetAttribute("IsEnabled");
            Assert.AreEqual("True", isEnabled);

            _saveButton.Click();

            isEnabled = _saveButton.GetAttribute("IsEnabled");
            Assert.AreEqual("False", isEnabled);
        }

        private void SaveAndCheckOpacitySettings(WindowsElement editor, int expected)
        {
            Assert.AreEqual(expected.ToString() + "\r\n", editor.Text);

            SaveChanges();
            ShortWait();

            int value = getPropertyValue<int>("fancyzones_highlight_opacity");
            Assert.AreEqual(expected, value);
        }

        private void SetOpacity(WindowsElement editor, string key)
        {
            editor.Click(); //activate
            editor.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace); //clear previous value
            editor.SendKeys(key);
            editor.SendKeys(OpenQA.Selenium.Keys.Enter); //confirm changes
        }

        private void TestRgbInput(string name)
        {
            WindowsElement colorInput = session.FindElementByXPath("//Edit[@Name=\"" + name + "\"]");
            Assert.IsNotNull(colorInput);

            colorInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            colorInput.SendKeys("0");
            colorInput.SendKeys(OpenQA.Selenium.Keys.Enter);
            Assert.AreEqual("0\r\n", colorInput.Text);
            
            string invalidSymbols = "qwertyuiopasdfghjklzxcvbnm,./';][{}:`~!@#$%^&*()_-+=\"\'\\";
            foreach (char symbol in invalidSymbols)
            {
                colorInput.SendKeys(symbol.ToString() + OpenQA.Selenium.Keys.Enter);
                Assert.AreEqual("0\r\n", colorInput.Text);
            }

            string validSymbols = "0123456789";
            foreach (char symbol in validSymbols)
            {
                colorInput.SendKeys(symbol.ToString() + OpenQA.Selenium.Keys.Enter);
                Assert.AreEqual(symbol.ToString() + "\r\n", colorInput.Text);
                colorInput.SendKeys(OpenQA.Selenium.Keys.Backspace);
            }

            //print zero first
            colorInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            colorInput.SendKeys("0");
            colorInput.SendKeys("1");
            Assert.AreEqual("1\r\n", colorInput.Text);

            //too many symbols
            colorInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            colorInput.SendKeys("1");
            colorInput.SendKeys("2");
            colorInput.SendKeys("3");
            colorInput.SendKeys("4");
            Assert.AreEqual("123\r\n", colorInput.Text);

            //too big value
            colorInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            colorInput.SendKeys("555");
                        
            Actions action = new Actions(session); //reset focus from input
            action.MoveToElement(colorInput).MoveByOffset(0, colorInput.Rect.Height).Click().Perform();

            Assert.AreEqual("255\r\n", colorInput.Text);
        }

        private void ClearInput(WindowsElement input)
        {
            input.Click();
            input.SendKeys(OpenQA.Selenium.Keys.Control + "a");
            input.SendKeys(OpenQA.Selenium.Keys.Backspace);
        }

        private void TestHotkey(WindowsElement input, int modifierKeysState, string key, string keyString)
        {
            BitArray b = new BitArray(new int[] { modifierKeysState });
            int[] flags = b.Cast<bool>().Select(bit => bit ? 1 : 0).ToArray();

            Actions action = new Actions(session).MoveToElement(input).Click();
            string expectedText = "";
            if (flags[0] == 1)
            {
                action.KeyDown(OpenQA.Selenium.Keys.Command);
                expectedText += "Win + ";
            }
            if (flags[1] == 1)
            {
                action.KeyDown(OpenQA.Selenium.Keys.Control);
                expectedText += "Ctrl + ";
            }
            if (flags[2] == 1)
            {
                action.KeyDown(OpenQA.Selenium.Keys.Alt);
                expectedText += "Alt + ";
            }
            if (flags[3] == 1)
            {
                action.KeyDown(OpenQA.Selenium.Keys.Shift);
                expectedText += "Shift + ";
            }

            expectedText += keyString + "\r\n";

            action.SendKeys(key + key);
            action.MoveByOffset(0, (input.Rect.Height / 2) + 10).ContextClick();
            if (flags[0] == 1)
            {
                action.KeyUp(OpenQA.Selenium.Keys.Command);
            }
            if (flags[1] == 1)
            {
                action.KeyUp(OpenQA.Selenium.Keys.Control);
            }
            if (flags[2] == 1)
            {
                action.KeyUp(OpenQA.Selenium.Keys.Alt);
            }
            if (flags[3] == 1)
            {
                action.KeyUp(OpenQA.Selenium.Keys.Shift);
            }
            action.Perform();

            SaveChanges();
            ShortWait();

            //Assert.AreEqual(expectedText, input.Text);

            JObject props = getProperties();
            JObject hotkey = props["fancyzones_editor_hotkey"].ToObject<JObject>()["value"].ToObject<JObject>();
            Assert.AreEqual(flags[0] == 1, hotkey.Value<bool>("win"));
            Assert.AreEqual(flags[1] == 1, hotkey.Value<bool>("ctrl"));
            Assert.AreEqual(flags[2] == 1, hotkey.Value<bool>("alt"));
            Assert.AreEqual(flags[3] == 1, hotkey.Value<bool>("shift"));
            //Assert.AreEqual(keyString, hotkey.Value<string>("key"));
        }

        [TestMethod]
        public void FancyZonesSettingsOpen()
        {
            WindowsElement fzTitle = session.FindElementByName("FancyZones Settings");
            Assert.IsNotNull(fzTitle);
        }
        /*
        [TestMethod]
        public void EditorOpen()
        {
            session.FindElementByXPath("//Button[@Name=\"Edit zones\"]").Click();
            ShortWait();

            WindowsElement editorWindow = session.FindElementByName("FancyZones Editor");
            Assert.IsNotNull(editorWindow);

            editorWindow.SendKeys(OpenQA.Selenium.Keys.Alt + OpenQA.Selenium.Keys.F4);
        }
        */
        /*
         * click each toggle,
         * save changes,
         * check if settings are changed after clicking save button
         */
        [TestMethod]
        public void TogglesSingleClickSaveButtonTest()
        {
            List<WindowsElement> toggles = session.FindElementsByXPath("//Pane[@Name=\"PowerToys Settings\"]/*[@LocalizedControlType=\"toggleswitch\"]").ToList();
            Assert.AreEqual(8, toggles.Count);

            List<bool> toggleValues = new List<bool>();
            foreach (WindowsElement toggle in toggles)
            {
                Assert.IsNotNull(toggle);
                
                bool isOn = toggle.GetAttribute("Toggle.ToggleState") == "1";
                toggleValues.Add(isOn);

                toggle.Click();

                SaveChanges();
                ShortWait();
            }
            
            //check saved settings
            JObject savedProps = getProperties();
            Assert.AreNotEqual(toggleValues[0], getPropertyValue<bool>(savedProps, "fancyzones_shiftDrag"));
            Assert.AreNotEqual(toggleValues[1], getPropertyValue<bool>(savedProps, "fancyzones_overrideSnapHotkeys"));
            Assert.AreNotEqual(toggleValues[2], getPropertyValue<bool>(savedProps, "fancyzones_zoneSetChange_flashZones"));
            Assert.AreNotEqual(toggleValues[3], getPropertyValue<bool>(savedProps, "fancyzones_displayChange_moveWindows"));
            Assert.AreNotEqual(toggleValues[4], getPropertyValue<bool>(savedProps, "fancyzones_zoneSetChange_moveWindows"));
            Assert.AreNotEqual(toggleValues[5], getPropertyValue<bool>(savedProps, "fancyzones_virtualDesktopChange_moveWindows"));
            Assert.AreNotEqual(toggleValues[6], getPropertyValue<bool>(savedProps, "fancyzones_appLastZone_moveWindows"));
            Assert.AreNotEqual(toggleValues[7], getPropertyValue<bool>(savedProps, "use_cursorpos_editor_startupscreen"));
        }

        /*
         * click each toggle twice,
         * save changes,
         * check if settings are unchanged after clicking save button
         */
        [TestMethod]
        public void TogglesDoubleClickSave()
        {
            List<WindowsElement> toggles = session.FindElementsByXPath("//Pane[@Name=\"PowerToys Settings\"]/*[@LocalizedControlType=\"toggleswitch\"]").ToList();
            Assert.AreEqual(8, toggles.Count);

            List<bool> toggleValues = new List<bool>();
            foreach (WindowsElement toggle in toggles)
            {
                Assert.IsNotNull(toggle);
                
                bool isOn = toggle.GetAttribute("Toggle.ToggleState") == "1";
                toggleValues.Add(isOn);

                toggle.Click();
                toggle.Click();
            }
            
            SaveChanges();
            ShortWait();

            JObject savedProps = getProperties();
            Assert.AreEqual(toggleValues[0], getPropertyValue<bool>(savedProps, "fancyzones_shiftDrag"));
            Assert.AreEqual(toggleValues[1], getPropertyValue<bool>(savedProps, "fancyzones_overrideSnapHotkeys"));
            Assert.AreEqual(toggleValues[2], getPropertyValue<bool>(savedProps, "fancyzones_zoneSetChange_flashZones"));
            Assert.AreEqual(toggleValues[3], getPropertyValue<bool>(savedProps, "fancyzones_displayChange_moveWindows"));
            Assert.AreEqual(toggleValues[4], getPropertyValue<bool>(savedProps, "fancyzones_zoneSetChange_moveWindows"));
            Assert.AreEqual(toggleValues[5], getPropertyValue<bool>(savedProps, "fancyzones_virtualDesktopChange_moveWindows"));
            Assert.AreEqual(toggleValues[6], getPropertyValue<bool>(savedProps, "fancyzones_appLastZone_moveWindows"));
            Assert.AreEqual(toggleValues[7], getPropertyValue<bool>(savedProps, "use_cursorpos_editor_startupscreen"));
        }

        [TestMethod]
        public void HighlightOpacitySetValue()
        {
            WindowsElement editor = session.FindElementByName("Zone Highlight Opacity (%)");
            Assert.IsNotNull(editor);

            SetOpacity(editor, "50");
            SaveAndCheckOpacitySettings(editor, 50);

            SetOpacity(editor, "-50");
            SaveAndCheckOpacitySettings(editor, 0);

            SetOpacity(editor, "200");
            SaveAndCheckOpacitySettings(editor, 100);

            //for invalid input values previously saved value expected
            SetOpacity(editor, "asdf"); 
            SaveAndCheckOpacitySettings(editor, 100); 
            
            SetOpacity(editor, "*");
            SaveAndCheckOpacitySettings(editor, 100); 
            
            SetOpacity(editor, OpenQA.Selenium.Keys.Return);
            SaveAndCheckOpacitySettings(editor, 100);

            Clipboard.SetText("Hello, clipboard");
            SetOpacity(editor, OpenQA.Selenium.Keys.Control + "v");
            SaveAndCheckOpacitySettings(editor, 100);
        }

        [TestMethod]
        public void HighlightOpacityIncreaseValue()
        {
            WindowsElement editor = session.FindElementByName("Zone Highlight Opacity (%)");
            Assert.IsNotNull(editor);

            SetOpacity(editor, "99");
            SaveAndCheckOpacitySettings(editor, 99);

            System.Drawing.Rectangle editorRect = editor.Rect;
            
            Actions action = new Actions(session);
            action.MoveToElement(editor).MoveByOffset(editorRect.Width / 2 + 10, -editorRect.Height / 4).Perform();
            ShortWait();

            action.Click().Perform();
            Assert.AreEqual("100\r\n", editor.Text);
            SaveAndCheckOpacitySettings(editor, 100);

            action.Click().Perform();
            Assert.AreEqual("100\r\n", editor.Text);
            SaveAndCheckOpacitySettings(editor, 100);
        }

        [TestMethod]
        public void HighlightOpacityDecreaseValue()
        {
            
            WindowsElement editor = session.FindElementByName("Zone Highlight Opacity (%)");
            Assert.IsNotNull(editor);

            SetOpacity(editor, "1");
            SaveAndCheckOpacitySettings(editor, 1);

            System.Drawing.Rectangle editorRect = editor.Rect;

            Actions action = new Actions(session);
            action.MoveToElement(editor).MoveByOffset(editorRect.Width / 2 + 10, editorRect.Height / 4).Perform();
            ShortWait();

            action.Click().Perform();
            Assert.AreEqual("0\r\n", editor.Text);
            SaveAndCheckOpacitySettings(editor, 0);

            action.Click().Perform();
            Assert.AreEqual("0\r\n", editor.Text);
            SaveAndCheckOpacitySettings(editor, 0);
        }

        [TestMethod]
        public void HighlightOpacityClearValueButton()
        {
            WindowsElement editor = session.FindElementByName("Zone Highlight Opacity (%)");
            Assert.IsNotNull(editor);

            editor.Click(); //activate
            AppiumWebElement clearButton = editor.FindElementByName("Clear value");
            Assert.IsNotNull(clearButton);
            
            /*element is not pointer- or keyboard interactable.*/
            Actions action = new Actions(session);
            action.MoveToElement(clearButton).Click().Perform();

            Assert.AreEqual("\r\n", editor.Text);
        }

        //in 0.15.2 sliders cannot be found by inspect.exe
        /*
        [TestMethod]
        public void HighlightColorSlidersTest()
        {
            ScrollDown();

            WindowsElement saturationAndBrightness = session.FindElementByName("Saturation and brightness");
            WindowsElement hue = session.FindElementByName("Hue");
            WindowsElement hex = session.FindElementByXPath("//Edit[@Name=\"Hex\"]");
            WindowsElement red = session.FindElementByXPath("//Edit[@Name=\"Red\"]");
            WindowsElement green = session.FindElementByXPath("//Edit[@Name=\"Green\"]");
            WindowsElement blue = session.FindElementByXPath("//Edit[@Name=\"Blue\"]");

            Assert.IsNotNull(saturationAndBrightness);
            Assert.IsNotNull(hue);
            Assert.IsNotNull(hex);
            Assert.IsNotNull(red);
            Assert.IsNotNull(green);
            Assert.IsNotNull(blue);

            System.Drawing.Rectangle satRect = saturationAndBrightness.Rect;
            System.Drawing.Rectangle hueRect = hue.Rect;

            //black on the bottom
            new Actions(session).MoveToElement(saturationAndBrightness).ClickAndHold().MoveByOffset(0, satRect.Height / 2).Click().Perform();
            ShortWait();

            Assert.AreEqual("0\r\n", red.Text);
            Assert.AreEqual("0\r\n", green.Text);
            Assert.AreEqual("0\r\n", blue.Text);
            Assert.AreEqual("000000\r\n", hex.Text);

            SaveChanges();
            ShortWait();            
            Assert.AreEqual("#000000", getPropertyValue<string>("fancyzones_zoneHighlightColor"));

            //white in left corner
            new Actions(session).MoveToElement(saturationAndBrightness).ClickAndHold().MoveByOffset(-(satRect.Width/2), -(satRect.Height / 2)).Click().Perform();
            Assert.AreEqual("255\r\n", red.Text);
            Assert.AreEqual("255\r\n", green.Text);
            Assert.AreEqual("255\r\n", blue.Text);
            Assert.AreEqual("ffffff\r\n", hex.Text);

            SaveChanges();
            ShortWait();
            Assert.AreEqual("#ffffff", getPropertyValue<string>("fancyzones_zoneHighlightColor"));

            //color in right corner
            new Actions(session).MoveToElement(saturationAndBrightness).ClickAndHold().MoveByOffset((satRect.Width / 2), -(satRect.Height / 2)).Click()
                .MoveToElement(hue).ClickAndHold().MoveByOffset(-(hueRect.Width / 2), 0).Click().Perform();
            Assert.AreEqual("255\r\n", red.Text);
            Assert.AreEqual("0\r\n", green.Text);
            Assert.AreEqual("0\r\n", blue.Text);
            Assert.AreEqual("ff0000\r\n", hex.Text);

            SaveChanges();
            ShortWait();
            Assert.AreEqual("#ff0000", getPropertyValue<string>("fancyzones_zoneHighlightColor"));
        }
        
        [TestMethod]
        public void HighlightColorTest()
        {
            ScrollDown();

            WindowsElement saturationAndBrightness = session.FindElementByName("Saturation and brightness");
            WindowsElement hue = session.FindElementByName("Hue");
            WindowsElement hex = session.FindElementByXPath("//Edit[@Name=\"Hex\"]");

            Assert.IsNotNull(saturationAndBrightness);
            Assert.IsNotNull(hue);
            Assert.IsNotNull(hex);

            hex.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            hex.SendKeys("63c99a");
            new Actions(session).MoveToElement(hex).MoveByOffset(0, hex.Rect.Height).Click().Perform();

            Assert.AreEqual("Saturation 51 brightness 79", saturationAndBrightness.Text);
            Assert.AreEqual("152", hue.Text);

            SaveChanges();
            ShortWait();
            Assert.AreEqual("#63c99a", getPropertyValue<string>("fancyzones_zoneHighlightColor"));
        }
        */

        [TestMethod]
        public void HighlightRGBInputsTest()
        {
            ScrollDown();

            TestRgbInput("Red");
            TestRgbInput("Green");
            TestRgbInput("Blue"); 
        }

        [TestMethod]
        public void HighlightHexInputTest()
        {
            ScrollDown();

            WindowsElement hexInput = session.FindElementByXPath("//Edit[@Name=\"Hex\"]");
            Assert.IsNotNull(hexInput);
            
            hexInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            
            string invalidSymbols = "qwrtyuiopsghjklzxvnm,./';][{}:`~!#@$%^&*()_-+=\"\'\\";
            foreach (char symbol in invalidSymbols)
            {
                hexInput.SendKeys(symbol.ToString());
                Assert.AreEqual("", hexInput.Text.Trim());
            }

            string validSymbols = "0123456789abcdef";
            foreach (char symbol in validSymbols)
            {
                hexInput.SendKeys(symbol.ToString());
                Assert.AreEqual(symbol.ToString(), hexInput.Text.Trim());
                hexInput.SendKeys(OpenQA.Selenium.Keys.Backspace);
            }
            
            //too many symbols
            hexInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            hexInput.SendKeys("000000");
            hexInput.SendKeys("1");
            Assert.AreEqual("000000\r\n", hexInput.Text);

            //short string
            hexInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            hexInput.SendKeys("000");
            new Actions(session).MoveToElement(hexInput).MoveByOffset(0, hexInput.Rect.Height).Click().Perform();
            Assert.AreEqual("000000\r\n", hexInput.Text);

            hexInput.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Backspace);
            hexInput.SendKeys("1234");
            new Actions(session).MoveToElement(hexInput).MoveByOffset(0, hexInput.Rect.Height).Click().Perform();
            Assert.AreEqual("112233\r\n", hexInput.Text);
        }

        [TestMethod]
        public void ExcludeApps()
        {
            WindowsElement input = session.FindElementByXPath("//Edit[contains(@Name, \"exclude\")]");
            Assert.IsNotNull(input);
            ClearInput(input);

            string inputValue;

            //valid
            inputValue = "Notepad\nChrome";
            input.SendKeys(inputValue);
            SaveChanges();
            ClearInput(input);
            ShortWait();
            Assert.AreEqual(inputValue, getPropertyValue<string>("fancyzones_excluded_apps"));

            //invalid
            inputValue = "Notepad Chrome";
            input.SendKeys(inputValue);
            SaveChanges();
            ClearInput(input);
            ShortWait();
            Assert.AreEqual(inputValue, getPropertyValue<string>("fancyzones_excluded_apps"));

            inputValue = "Notepad,Chrome";
            input.SendKeys(inputValue);
            SaveChanges();
            ClearInput(input);
            ShortWait();
            Assert.AreEqual(inputValue, getPropertyValue<string>("fancyzones_excluded_apps"));

            inputValue = "Note*";
            input.SendKeys(inputValue);
            SaveChanges();
            ClearInput(input);
            ShortWait();
            Assert.AreEqual(inputValue, getPropertyValue<string>("fancyzones_excluded_apps"));

            inputValue = "Кириллица";
            input.SendKeys(inputValue);
            SaveChanges();
            ClearInput(input);
            ShortWait();
            Assert.AreEqual(inputValue, getPropertyValue<string>("fancyzones_excluded_apps"));
        }

        [TestMethod]
        public void ExitDialogSave()
        {
            WindowsElement toggle = session.FindElementByXPath("//Pane[@Name=\"PowerToys Settings\"]/*[@LocalizedControlType=\"toggleswitch\"]");          
            Assert.IsNotNull(toggle);

            bool initialToggleValue = toggle.GetAttribute("Toggle.ToggleState") == "1";

            toggle.Click();
            CloseSettings();
            WindowsElement exitDialog = session.FindElementByName("Changes not saved");
            Assert.IsNotNull(exitDialog);

            exitDialog.FindElementByName("Save").Click();

            //check if window still opened
            WindowsElement powerToysWindow = session.FindElementByXPath("//Window[@Name=\"PowerToys Settings\"]");
            Assert.IsNotNull(powerToysWindow); 

            //check settings change
            JObject savedProps = getProperties();
            
            Assert.AreNotEqual(initialToggleValue, getPropertyValue<bool>(savedProps, "fancyzones_shiftDrag"));
            
            //return initial app state
            toggle.Click(); 
        }

        [TestMethod]
        public void ExitDialogExit()
        {
            WindowsElement toggle = session.FindElementByXPath("//Pane[@Name=\"PowerToys Settings\"]/*[@LocalizedControlType=\"toggleswitch\"]");
            Assert.IsNotNull(toggle);

            bool initialToggleValue = toggle.GetAttribute("Toggle.ToggleState") == "1";
            
            toggle.Click();
            CloseSettings();

            WindowsElement exitDialog = session.FindElementByName("Changes not saved");
            Assert.IsNotNull(exitDialog);

            exitDialog.FindElementByName("Exit").Click();

            //check if window still opened
            try 
            {
                WindowsElement powerToysWindow = session.FindElementByXPath("//Window[@Name=\"PowerToys Settings\"]");
                Assert.IsNull(powerToysWindow);
            }
            catch(OpenQA.Selenium.WebDriverException)
            {
                //window is no longer available, which is expected
            }

            //return initial app state
            Init();

            //check settings change
            JObject savedProps = getProperties();
            Assert.AreEqual(initialToggleValue, getPropertyValue<bool>(savedProps, "fancyzones_shiftDrag"));
        }

        [TestMethod]
        public void ExitDialogCancel()
        {
            WindowsElement toggle = session.FindElementByXPath("//Pane[@Name=\"PowerToys Settings\"]/*[@LocalizedControlType=\"toggleswitch\"]");
            Assert.IsNotNull(toggle);

            toggle.Click();
            CloseSettings();
            WindowsElement exitDialog = session.FindElementByName("Changes not saved");
            Assert.IsNotNull(exitDialog);

            exitDialog.FindElementByName("Cancel").Click();

            //check if window still opened
            WindowsElement powerToysWindow = session.FindElementByXPath("//Window[@Name=\"PowerToys Settings\"]");
            Assert.IsNotNull(powerToysWindow);

            //check settings change
            JObject savedProps = getProperties();
            JObject initialProps = _initialSettingsJson["properties"].ToObject<JObject>();
            Assert.AreEqual(getPropertyValue<bool>(initialProps, "fancyzones_shiftDrag"), getPropertyValue<bool>(savedProps, "fancyzones_shiftDrag"));

            //return initial app state
            toggle.Click();
            SaveChanges();
        }

        [TestMethod]
        public void ConfigureHotkey()
        {
            WindowsElement input = session.FindElementByXPath("//Edit[contains(@Name, \"hotkey\")]");
            Assert.IsNotNull(input);

            for (int i = 0; i < 16; i++)
            {
                TestHotkey(input, i, OpenQA.Selenium.Keys.End, "End");
            }
        }

        [TestMethod]
        public void ConfigureLocalSymbolHotkey()
        {
            WindowsElement input = session.FindElementByXPath("//Edit[contains(@Name, \"hotkey\")]");
            Assert.IsNotNull(input);
            TestHotkey(input, 0, "ё", "Ё");
        }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            Setup(context);
            Init();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            CloseSettings();

            try
            {
                WindowsElement exitDialogButton = session.FindElementByName("Exit");
                if (exitDialogButton != null)
                {
                    exitDialogButton.Click();
                }
            }
            catch(OpenQA.Selenium.WebDriverException)
            {
                //element couldn't be located
            }

            TearDown();
        }

        [TestInitialize]
        public void TestInitialize()
        {
            try
            {
                _initialSettings = File.ReadAllText(_settingsPath);
                _initialSettingsJson = JObject.Parse(_initialSettings);
            }
            catch (System.IO.FileNotFoundException)
            {
                _initialSettings = "";
            }
        }

        [TestCleanup]
        public void TestCleanup()
        {
            ScrollUp();

            if (_initialSettings.Length > 0)
            {
                File.WriteAllText(_settingsPath, _initialSettings);
            }            
        }
    }
}
