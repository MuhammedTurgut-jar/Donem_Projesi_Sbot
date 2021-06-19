using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.Globalization;

namespace Sbot
{

    class driver
    {
        private static string user, pass;
        private List<string> ad;
        private List<string> soyad;
        private List<string> eposta;
        private List<string> kullaniciAdi;
        private List<string> pozisyon;
        private List<string> yetki;
        private List<string> basvuru;

        private static string url = "https://crmbeyazmasa.ibb.gov.tr/epublicsector_trk/start.swe?SWECmd=Start&SWEHo=crmbeyazmasa.ibb.gov.tr";
        
      
        public static Boolean KullaniciEkle(string user, string pass, List<string>ad, List<string> soyad, List<string> eposta, List<string> kullaniciAdi)
        {
            
            IWebDriver driver;
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl(url);
            Actions action = new Actions(driver);

            // ***** Login Case ***** 
            driver.FindElement(By.Id("s_swepi_1")).SendKeys(user);
            driver.FindElement(By.Id("s_swepi_2")).SendKeys(pass);
            driver.FindElement(By.Id("s_swepi_22")).Click();
            // ***** Login Case *****
           //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //string ready = (String)js.ExecuteScript("return document.readyState");
            
            try
            {
                driver.FindElement(By.ClassName("siebui-error")).Text.ToString();
                driver.Quit();
                return false;
            }
            catch
            {

            // ***** Yönetim - Kullanıcı ******
                tr:
                try { driver.FindElement(By.XPath("//*[text()='Yönetim - Kullanıcı']")).Click(); }
                catch { goto tr; }


            // ***** Yönetim - Kullanıcı ******


            // ***** Kullanıcı Ekleme ******
                tr2:
                    try { driver.FindElement(By.Id("s_1_1_11_0_Ctrl")).Click(); }
                    catch { goto tr2; }
              

                for (int i=0;i<eposta.Count;i++)
                     {
                    

                    tr3:
                    try { driver.FindElement(By.Id("1_First_Name")).SendKeys(ad[i]);}
                    catch {goto tr3;}

                    driver.FindElement(By.Id("1_s_1_l_Last_Name")).Click();
                    driver.FindElement(By.Id("1_Last_Name")).SendKeys(soyad[i]);

                    driver.FindElement(By.Id("1_s_1_l_Login_Name")).Click();
                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(kullaniciAdi[i]);

                    driver.FindElement(By.Id("1_s_1_l_IBB_Assign_SR_User")).Click();
                    driver.FindElement(By.Id("1_IBB_Assign_SR_User")).SendKeys("MUHAMMED.TURGUT");

                    driver.FindElement(By.Id("1_s_1_l_IBB_Location")).Click();
                    driver.FindElement(By.Id("1_IBB_Location")).SendKeys("ALO 153 ÇAĞRI MERKEZİ");

                    driver.FindElement(By.Id("1_s_1_l_EMail_Addr")).Click();
                    driver.FindElement(By.Id("1_EMail_Addr")).SendKeys(eposta[i]);

                    if (i < 1)
                    {
                        //Agent Yetki Ekleme
                        driver.FindElement(By.Id("1_s_1_l_Responsibility")).Click();
                        driver.FindElement(By.Id("s_1_2_48_0_icon")).Click();
                    tr4:
                        try { driver.FindElement(By.Id("s_5_1_10_0_Ctrl")).Click(); }
                        catch { goto tr4; }
                        driver.FindElement(By.Id("1_Description")).SendKeys("Agent");
                        driver.FindElement(By.Id("s_5_1_8_0_Ctrl")).Click();
                        driver.FindElement(By.Id("s_4_1_79_0_Ctrl")).Click();
                    tr5:
                        try { driver.FindElement(By.Id("1_s_4_l_SSA_Primary_Field")).Click(); }
                        catch { goto tr5; }
                        driver.FindElement(By.ClassName("siebui-icon-checkbox-checked")).Click();

                        //Admin Yetki Kaldırma
                        driver.FindElement(By.Id("2_s_4_l_SSA_Primary_Field")).Click();
                        driver.FindElement(By.Id("s_4_1_80_0_Ctrl")).Click();

                        //Yetki pop up tamam butonu
                        driver.FindElement(By.Id("s_4_1_82_0_Ctrl")).Click();

                        //Agent Pozisyon Ekleme

                        driver.FindElement(By.Id("1_s_1_l_Position")).Click();
                        driver.FindElement(By.Id("s_1_2_49_0_icon")).Click();
                    tr6:
                        try { driver.FindElement(By.Id("s_7_1_10_0_Ctrl")).Click(); }
                        catch { goto tr6; }
                    tr7:
                        try { driver.FindElement(By.Id("1_s_7_l_Name")).Click(); }
                        catch { goto tr7; }
                        driver.FindElement(By.Id("1_Name")).SendKeys("IBB CAGRI MERKEZI AGENT 001");
                        driver.FindElement(By.Id("s_7_1_8_0_Ctrl")).Click();
                        driver.FindElement(By.Id("s_6_1_79_0_Ctrl")).Click();
                        driver.FindElement(By.Id("s_6_1_82_0_Ctrl")).Click();
                    }
                    //Agent Pozisyon Ekleme


                    driver.FindElement(By.Id("1_s_1_l_Employment_Status")).Click();
                    driver.FindElement(By.Id("1_Employment_Status")).SendKeys("Aktif");

                    if (i < eposta.Count - 1)
                    {
                        driver.FindElement(By.Id("1_s_1_l_Last_Name")).Click();
                        driver.FindElement(By.Id("1_s_1_l_Last_Name")).SendKeys(OpenQA.Selenium.Keys.Control + "b");
                    }
                   
                }
            
                // ***** Kullanıcı Ekleme ******
            }
            driver.FindElement(By.Id("1_s_1_l_Last_Name")).Click();
            driver.FindElement(By.Id("1_s_1_l_Last_Name")).SendKeys(OpenQA.Selenium.Keys.Control + "s");

            MessageBox.Show("İşlem tamamlandı!");
            driver.Quit();
            return true;
        }


        public static Boolean KullaniciSil(string user, string pass, List<string> kullaniciAdi)
        {

            IWebDriver driver;
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl(url);
            Actions action = new Actions(driver);

            // ***** Login Case ***** 
            driver.FindElement(By.Id("s_swepi_1")).SendKeys(user);
            driver.FindElement(By.Id("s_swepi_2")).SendKeys(pass);
            driver.FindElement(By.Id("s_swepi_22")).Click();
            // ***** Login Case *****
            //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //string ready = (String)js.ExecuteScript("return document.readyState");

            try
            {
                driver.FindElement(By.ClassName("siebui-error")).Text.ToString();
                driver.Quit();
                return false;
            }
            catch
            {

            // ***** Yönetim - Kullanıcı ******
            tr:
                try { driver.FindElement(By.Id("ui-id-184")).Click(); }
                catch (Exception e)
                {
                    e.Data.Clear();
                    goto tr; }


            // ***** Yönetim - Kullanıcı ******


            // ***** Kullanıcı Silme ******
            
                for (int i = 0; i < kullaniciAdi.Count; i++)
                {
                tr2:
                    try { driver.FindElement(By.Id("s_1_1_10_0_Ctrl")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr2; }

                tr3:
                    try { driver.FindElement(By.Id("1_s_1_l_Login_Name")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr3; }

                    
                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(kullaniciAdi[i]);

                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(OpenQA.Selenium.Keys.Enter);

                    driver.FindElement(By.Id("1_s_1_l_Employment_Status")).Click();

                tr4:
                    try { driver.FindElement(By.Id("1_Employment_Status")).SendKeys("Pasif"); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr4;
                    }
                    


                }

                // ***** Kullanıcı Silme ******
            }
            driver.FindElement(By.Id("1_s_1_l_Last_Name")).Click();
            driver.FindElement(By.Id("1_s_1_l_Last_Name")).SendKeys(OpenQA.Selenium.Keys.Control + "s");

            MessageBox.Show("İşlem tamamlandı!");
            driver.Quit();
            return true;
        }

        public static Boolean PozisyonEkle(string user, string pass, List<string> kullaniciAdi, List<string> pozisyon)
        {

            IWebDriver driver;
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl(url);
            Actions action = new Actions(driver);

            // ***** Login Case ***** 
            driver.FindElement(By.Id("s_swepi_1")).SendKeys(user);
            driver.FindElement(By.Id("s_swepi_2")).SendKeys(pass);
            driver.FindElement(By.Id("s_swepi_22")).Click();
            // ***** Login Case *****
            //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //string ready = (String)js.ExecuteScript("return document.readyState");

            try
            {
                driver.FindElement(By.ClassName("siebui-error")).Text.ToString();
                driver.Quit();
                return false;
            }
            catch
            {

            // ***** Yönetim - Kullanıcı ******
            tr:
                try { driver.FindElement(By.XPath("//*[text()='Yönetim - Kullanıcı']")).Click(); }
                catch (Exception e)
                {
                    e.Data.Clear();
                    goto tr; }


                // ***** Yönetim - Kullanıcı ******

                // ***** Pozisyon Ekleme ******

                for (int i = 0; i < kullaniciAdi.Count; i++)
                {
                tr2:
                    try { driver.FindElement(By.Id("s_1_1_10_0_Ctrl")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr2; }

                tr3:
                    try { driver.FindElement(By.Id("1_s_1_l_Login_Name")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr3; }

                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(kullaniciAdi[i]);

                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(OpenQA.Selenium.Keys.Enter);

                    driver.FindElement(By.Id("1_s_1_l_Position")).Click();
                    driver.FindElement(By.Id("s_1_2_49_0_icon")).Click();

                tr6:
                    try { driver.FindElement(By.XPath("//button[@title='Pozisyon Ekle:Sorgula']")).Click();}
                    
                    catch (Exception e)
                    {
                        e.Data.Clear(); 
                        goto tr6; }
                tr7:
                    try { driver.FindElement(By.Id("1_s_5_l_Name")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr7; }
                    driver.FindElement(By.Id("1_Name")).SendKeys(pozisyon[i]);
                    driver.FindElement(By.Id("s_5_1_8_0_Ctrl")).Click();
                    driver.FindElement(By.Id("s_4_1_79_0_Ctrl")).Click();
                    driver.FindElement(By.Id("s_4_1_82_0_Ctrl")).Click();

                   

                    driver.FindElement(By.XPath("//*[text()='Yönetim - Kullanıcı']")).Click();
                }

                // ***** Pozisyon Ekleme ******
            }

            MessageBox.Show("İşlem tamamlandı!");
            driver.Quit();
            return true;
        }

        public static Boolean YetkiEkle(string user, string pass, List<string> kullaniciAdi, List<string> yetki)
        {

            IWebDriver driver;
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl(url);
            Actions action = new Actions(driver);

            // ***** Login Case ***** 
            driver.FindElement(By.Id("s_swepi_1")).SendKeys(user);
            driver.FindElement(By.Id("s_swepi_2")).SendKeys(pass);
            driver.FindElement(By.Id("s_swepi_22")).Click();
            // ***** Login Case *****
            //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //string ready = (String)js.ExecuteScript("return document.readyState");

            try
            {
                driver.FindElement(By.ClassName("siebui-error")).Text.ToString();
                driver.Quit();
                return false;
            }
            catch
            {

            // ***** Yönetim - Kullanıcı ******
            tr:
                try { driver.FindElement(By.XPath("//*[text()='Yönetim - Kullanıcı']")).Click(); }
                catch (Exception e)
                {

                    e.Data.Clear(); 
                    goto tr; }


                // ***** Yönetim - Kullanıcı ******



                for (int i = 0; i < kullaniciAdi.Count; i++)
                {
                tr2:
                    try { driver.FindElement(By.Id("s_1_1_10_0_Ctrl")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear(); 
                        goto tr2; }

                tr3:
                    try { driver.FindElement(By.Id("1_s_1_l_Login_Name")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr3; }


                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(kullaniciAdi[i]);

                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(OpenQA.Selenium.Keys.Enter);

                    driver.FindElement(By.Id("1_s_1_l_Responsibility")).Click();
                    driver.FindElement(By.Id("s_1_2_48_0_icon")).Click();
                tr4:
                    try { driver.FindElement(By.Id("s_5_1_10_0_Ctrl")).Click(); }
                    catch (Exception e)
                    {
                       e.Data.Clear();
                        goto tr4; }
                tr5:
                    try
                    { driver.FindElement(By.Id("1_Description")).SendKeys(yetki[i]); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr5;
                    }
            driver.FindElement(By.Id("s_5_1_8_0_Ctrl")).Click();
                    driver.FindElement(By.Id("s_4_1_79_0_Ctrl")).Click();

                    driver.FindElement(By.Id("s_4_1_82_0_Ctrl")).Click();

                    driver.FindElement(By.Id("1_s_1_l_Last_Name")).Click();
                    // driver.Navigate().Refresh();

                    driver.FindElement(By.XPath("//*[text()='Yönetim - Kullanıcı']")).Click();
                }

               
            }
        

            MessageBox.Show("İşlem tamamlandı!");
            driver.Quit();
            return true;
        }

        public static Boolean BasvuruTansfer(string user, string pass, List<string> kullaniciAdi, List<string> basvuru)
        {

            IWebDriver driver;
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl(url);
            Actions action = new Actions(driver);

            // ***** Login Case ***** 
            driver.FindElement(By.Id("s_swepi_1")).SendKeys(user);
            driver.FindElement(By.Id("s_swepi_2")).SendKeys(pass);
            driver.FindElement(By.Id("s_swepi_22")).Click();
            // ***** Login Case *****
            //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //string ready = (String)js.ExecuteScript("return document.readyState");

            try
            {
                driver.FindElement(By.ClassName("siebui-error")).Text.ToString();
                driver.Quit();
                return false;
            }
            catch
            {

                for (int i = 0; i < basvuru.Count; i++)
                {

                tr:
                    try { driver.FindElement(By.XPath("//*[text()='Başvuru']")).Click(); }
                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr; }

                tr5:
                    try { driver.FindElement(By.XPath("//button[@title='Ara:Git']")).Click(); }

                    catch (Exception e)
                    {
                        e.Data.Clear(); 
                        goto tr5; }

                tr2:
                    try { driver.FindElement(By.XPath("//button[@title='Başvurular:Sorgula']")).Click(); }
                    catch (Exception e){
                        e.Data.Clear();
                        goto tr2; }

                tr8:
                    try
                    {
                        driver.FindElement(By.Name("SR_Number")).SendKeys(basvuru[i]);
                    }

                    catch (Exception e)
                    {
                        e.Data.Clear();
                        goto tr8;
                    }

                    driver.FindElement(By.Name("SR_Number")).SendKeys(OpenQA.Selenium.Keys.Enter);
                tr3:
                    try
                    {
                        driver.FindElement(By.XPath("//button[@title='Başvuru:Sorumlu Kullanıcı Güncelle']")).Click();
                    }

                    catch (Exception e)
                    {

                        e.Data.Clear();

                        goto tr3;
                    };

                   tr4:
                    try { driver.FindElement(By.XPath("//button[@title='Sorumlu Kullanıcı Seçiniz:Sorgula']")).Click(); }
                    catch (Exception e) {
                        
                        e.Data.Clear();
                       
                        goto tr4; 
                    };

                    tr7:
                    try { 
                    driver.FindElement(By.Id("1_s_4_l_Login_Name")).Click();
                    }
                    catch (Exception e)
                    {

                        e.Data.Clear();

                        goto tr7;
                    }

                    driver.FindElement(By.Id("1_Login_Name")).SendKeys(kullaniciAdi[i]);

                    driver.FindElement(By.XPath("//button[@title='Sorumlu Kullanıcı Seçiniz:Git']")).Click();

                    tr9:
                    try { driver.FindElement(By.XPath("//button[@title='Sorumlu Kullanıcı Seçiniz:Tamam']")).Click(); }

                    catch (Exception e)
                    {

                        e.Data.Clear();

                        goto tr9;
                    }

        }

            }
       

            MessageBox.Show("İşlem tamamlandı!");
            driver.Quit();
            return true;
        }

    }

}
