using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace OfficeKeys
{
    public class KeysLoader
    {
        private string _cookies = string.Empty;

        public List<Product> Load(string cookies)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(
                delegate
                {
                    return true;
                });

            //_cookies = cookies;
            _cookies = Regex.Replace(cookies, "muxcan=[a-zA-Z0-9]+", "");

            List<Product> products = new List<Product>();

            //string url = "https://office.microsoft.com/MyAccount.aspx?fromAR=1";
            string url = "https://stores.office.com/myaccount/home.aspx?lc=en-us&ui=en-US&rs=en-US&ad=CZ&fromAR=1";
            HttpWebRequest wr = CreateWebRequest(url);

            WebResponse response = wr.GetResponse();

            string responseContent;
            using (StreamReader sr = new StreamReader(response.GetResponseStream()))
            {
                responseContent = sr.ReadToEnd();
            }
            //string corrId = "";
            Match m;
            //m = Regex.Match(responseContent, "g_strCorrId='([^']+)");
            //if (m.Success)
            //{
            //    corrId = m.Groups[1].Value;
            //}

            string mux = "";
            m = Regex.Match(responseContent, "var CanaryGuid='([0-9a-zA-Z]+)';");
            if (m.Success)
            {
                mux = m.Groups[1].Value;
            }

            wr = CreateWebRequest(url, "muxcan=" + mux);
            wr.Method = "POST";
            wr.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
            wr.Referer = "https://stores.office.com/myaccount/home.aspx?lc=en-us&ui=en-US&rs=en-US&ad=US&fromAR=1";
            wr.Headers["X-Requested-With"] = "XMLHttpRequest";
            Stream reqStream = wr.GetRequestStream();
            string postData = "load=true&clienttzo=-60";
            byte[] postArray = Encoding.ASCII.GetBytes(postData);
            reqStream.Write(postArray, 0, postArray.Length);
            reqStream.Close();

            response = wr.GetResponse();
            using (StreamReader sr = new StreamReader(response.GetResponseStream()))
            {
                responseContent = sr.ReadToEnd();

            }

            Match ul = Regex.Match(responseContent, "<ul id=\"myaccountUl\">.*?</ul>");
            if (ul.Success)
            {
                //List<Product> products = new List<Product>();
                string ulHtml = ul.Groups[0].Value;

                string[] lis = ulHtml.Split(new string[] { "<li id=\"myaccountLi" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string li in lis)
                {
                    Match link = Regex.Match(li, "<a id=\"[^\"]+\" href=\"([^\"]+)\">Install from a disc");

                    if (link.Success)
                    {
                        Match productName = Regex.Match(li, "<h3 id=\"muxlt\" class=\"oxmuxlt\">([^<]+)</h3>");
                        Match language = Regex.Match(li, "Language: <span id=\"muxilang00\">([^<]+)</span>");
                        Match added = Regex.Match(li, "<h3 id=\"oxmuxoffersubtitle\" class=\"oxmuxlst\">[^.]+. Added to account on [A-Za-z]+, ([^.]+).</h3>");


                        Product product = new Product();
                        product.Link = link.Groups[1].Value;
                        if (productName.Success)
                            product.Name = productName.Groups[1].Value;
                        if (language.Success)
                            product.Language = language.Groups[1].Value;
                        if (added.Success)
                        {
                            product.Added = added.Groups[1].Value;

                            try
                            {
                                product.AddedDate = DateTime.Parse(product.Added);
                            }
                            catch { }
                        }

                        products.Add(product);
                    }
                }
            }

//            MatchCollection coll = Regex.Matches(responseContent, "<a id=\"[^\"]+\" href=\"([^\"]+)\">Install from a disc");
//            string[] urls = coll.Cast<Match>().Select(match => match.Groups[1].Value).ToArray();
            string corrId = "";

            foreach(Product product in products)
            {
                //wr = CreateWebRequest(HttpUtility.HtmlDecode(u));
                wr = CreateWebRequest(HttpUtility.HtmlDecode(product.Link));
                response = wr.GetResponse();

                using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                {
                    responseContent = sr.ReadToEnd();
                }

                m = Regex.Match(responseContent, "g_strCorrId='([^']+)");
                if (m.Success)
                {
                    corrId = m.Groups[1].Value;
                }

                string canary = "";
                m = Regex.Match(response.Headers[HttpResponseHeader.SetCookie], "muxcan=([^;]+)");
                if (m.Success)
                {
                    canary = m.Groups[1].Value;
                }

                wr = CreateWebRequest(product.Link + "&corr=" + corrId);
                wr.Method = "POST";
                wr.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                wr.Headers["X-Requested-With"] = "XMLHttpRequest";
                reqStream = wr.GetRequestStream();
                postData = "load=true&clienttzo=-60";
                postArray = Encoding.ASCII.GetBytes(postData);
                reqStream.Write(postArray, 0, postArray.Length);
                reqStream.Close();

                response = wr.GetResponse();
                using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                {
                    responseContent = sr.ReadToEnd();
                }

                //POST https://office.microsoft.com/cs-cz/InstallFromDisc.aspx?subsid=uiOPcwAAAAAAAA0A&entitleid=uiOPcwAAAAAAAA0A%5FA9A0ED46 HTTP/1.1

                string productName = "";
                m = Regex.Match(responseContent, "<h3 id=\"muxlt\" class=\"oxmuxlt\">([^<]+)");
                if (m.Success)
                {
                    productName = m.Groups[1].Value;
                }

                //string canary = "";
                m = Regex.Match(response.Headers[HttpResponseHeader.SetCookie], "muxcan=([^;]+)");
                if (m.Success)
                {
                    canary = m.Groups[1].Value;
                }

                string subsid = "";
                m = Regex.Match(product.Link, "subsid=([^&]+)");
                if (m.Success)
                {
                    subsid = m.Groups[1].Value;
                }

                wr = CreateWebRequest(HttpUtility.HtmlDecode(product.Link), "muxcan=" + canary);
                wr.Method = "POST";
                wr.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                wr.Headers["X-Requested-With"] = "XMLHttpRequest";
                wr.Referer = HttpUtility.HtmlDecode(product.Link);
                reqStream = wr.GetRequestStream();
                postData = "muxaction=getProductKey&subsid=" + subsid + "&canaryguid=" + canary;
                postArray = Encoding.ASCII.GetBytes(postData);
                reqStream.Write(postArray, 0, postArray.Length);
                reqStream.Close();

                response = wr.GetResponse();
                using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                {
                    responseContent = sr.ReadToEnd();

                    m = Regex.Match(responseContent, "^([A-Za-z0-9]+)-([A-Za-z0-9]+)-([A-Za-z0-9]+)-([A-Za-z0-9]+)-([A-Za-z0-9]+)$");
                    if (m.Success)
                    {
                        product.Id = subsid;
                        product.Key = responseContent;
                        //Product p = new Product()
                        //{
                        //    Id = subsid,
                        //    Name = productName,
                        //    Key = responseContent,
                        //};

                        //products.Add(p);
                    }
                }
            }

            return products;
        }

        private HttpWebRequest CreateWebRequest(string url, string appendCookie = null)
        {
            WebRequest request = WebRequest.Create(url);

            HttpWebRequest wr = request as HttpWebRequest;

            wr.Host = "stores.office.com";
            wr.Headers.Add(HttpRequestHeader.AcceptLanguage, "cs-CZ,cs;q=0.8");
            //wr.Connection = "keep-alive";
            wr.Accept = "*/*";
            wr.UserAgent = "Mozilla/5.0 (Windows NT 6.4; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2249.0 Safari/537.36";
            wr.Referer = "https://login.live.com/ppsecure/post.srf?wa=wsignin1.0&rpsnv=12&ct=1417783837&rver=6.1.6206.0&wp=MBI_SSL&wreply=https:%2F%2Foffice.microsoft.com%2Fwlid%2Fauthredir.aspx%3Furl%3Dhttps%253A%252F%252Foffice%252Emicrosoft%252Ecom%252FMyAccount%252Easpx%26hurl%3DAEBFB4399BF80E8C9B7B53E730F0E7A44A5BC73E%26ipt%3D1&lc=1029&id=34134&bk=1417783867&uaid=088e02f760c449c0a26e941ad6e48d59";
            //wr.Headers.Add(HttpRequestHeader.Cookie, "lc=cs-CZ+CZ:99; _DetectCookies=Y; WT_NVR_RU=0=msdn:1=:2=; MUID=139A49C582496448044F4EC28649627E; MC1=GUID=74ed361e7029564f8c163183e1496417&HASH=1e36&LV=201412&V=4&LU=1417783582988; A=I&I=AxUFAAAAAADnBwAAk+HAG9FJXDgK9iqqZ3PTCw!!&V=4; omniID=4f270c71_6091_4892_a595_b70d262bd755; s_cc=true; s_sq=%5B%5BB%5D%5D; TocPosition=1; MS0=bd32350219c441a2b010322ad445994d; RPSOOnline=FABCARTmLF7/9zpgYvJ7Vwc033zfJwtYPwNmAAAEgAAACP0xlsJSCBX0AAEVetdSYhvE4Shzss8g1kb/Au7V%2BJXJTd2s89L3K/ZVUflDp8M0m4xCXUaC0r6T3%2BBETp3Gds6%2BtIzufGs5c5CLiBrLunD4YGpZjiniuxOzc6nyl%2BM96kXM583JqEH%2BBzTk17jOWQYZK8IkOfXmanR2tzir1kA3HOA0a3wjYmEzEowcoMG%2BZ6XdmI4GBl5FqlLfRVXLWfl/OSBWcMvuQTwjxT0PnYi9sMXzTTI5beShLUdeE83gojZanNZHYONBFCEEMsd5UuGURE0WyYAnkFeD5WEdMXPPWP2IqmZM7WgyH%2BMdulIGuBmN2wnEFSkjxudgAqeSbgGh3DRyEdVVuTGvFADLKn%2BbuInT1bnPZNGr/PyUzU5ilw%3D%3D; RPSSecOOnline=FABCARTmLF7/9zpgYvJ7Vwc033zfJwtYPwNmAAAEgAAACBA4Z6wwtJ3vAAGUbPU0LSC0ZP4cI5oSXUlX83ntb1zsfRdX/gQjIQFARNzzoGNDC4CkxN6u4%2B2siMr7ph3m/bAiiehqfk05o7pk6xtlORnSeOXANTdYoNMqJb7idgS2F/gsO45h0M21mMNr0IOV4Pvzr1ip6ZtXcyV077fhmLTpQOXgMvfZSHWuyVwoYZqTT8aYH74nAxFt1aQylOJoMmLdtTqECk7fwIJb1O/pFFVlfMYkNQ72vUee3c2qeRzRcnpryZiOqvggMu%2B9ZMOFDqIWWnB//WahKWa9746iZUROf0gQoLfFKYm3TXdRtaa0r0YKFxZBXLCGafy55yR75xCuO/ywaS7gsLqqFAD7sM3vMuUuacHLtt/8QHYWYVr2Tw%3D%3D; RPSClearCT=EgCKAQMAAAAEgAAAC4AAd8Ehc88sXwWTdY%2BtCUyBMar499Y3AmdG2dAijPmGh01t6HaqoPnDferSsFRFgV0OFmvRSOoi5srGpzC717jvD/dNrYS3SEh0Ee%2Bmgu3cIAcWbSs4QOksAb75JlglG9dDCyfvccFFi9JDr3vkIGdaWYoA3WRYf2MbHHZa4zpBLnf5AGQA%2BQD9vwMAHHFIwKaqgVSmqoFUVoUAAAoQIAAQFgBjei1pdEBhcHMtaG9sZGluZy5jb20AWwAAImN6LWl0JWFwcy1ob2xkaW5nLmNvbUBwYXNzcG9ydC5jb20AAAACQ1oABTExMDAwAAAAAAgJAgAAciVmQAAGQQADQVBTAAdIT0xESU5HBQwAAAAAAAAAAAAAAAC9NlTT0mxLpAAApqqBVKdR%2BFQAAAAAAAAAAAAAAAAOADg2LjQ5LjE4NS4xMzgABAEAAAAAAAAAAAAAAAABBAAAAAAAAAAAAAAAAAAAAPJ1qwWAfm1oAAAAAAAAAAAAAAAAAAAAAA%3D%3D; ANON=A=98F14BC2A4A7C8E245C59364FFFFFFFF&E=104b&W=1; NAP=V=1.9&E=ff1&C=cSwtIikfdvVM8FXJRwNI8SyTXd8lpR_uh4PrIEbxm4iV0IXL6h-KOQ&W=1; ODCSSLAuth=t=2014-12-05 12:52:56Z&h=UY/SN/bplA2EugLHxPjygQUKJeI=; ul=1; WT_FPC=id=ba854c3c-6538-4e98-8075-993459cc9fcd:lv=1417751578170:ss=1417749995585; MSFPC=ID=74ed361e7029564f8c163183e1496417&CS=3&LV=201412&V=1; muxcan=8b7df068e9364da0975a70b40887ffa5");
            //wr.Headers.Add(HttpRequestHeader.Cookie, "lc=cs-CZ+CZ:99; _DetectCookies=Y; WT_FPC=id=ddd00cae-1101-4452-a546-0fe29fbdbde4:lv=1417763652113:ss=1417763652113; RPSOOnline=FABCARTmLF7/9zpgYvJ7Vwc033zfJwtYPwNmAAAEgAAACLOi5AJtYslKAAEChqK4WbXswibtk/5nMP8WdGC1TOBXDrTu9SiOYbE/cK7Ea62h6iLMVlrGn5ehK10uwMIHZ7XI1C8AO45Nex2%2B0uopPhaaiZoyUGdRTCAKwm/evT3eInLSuLUv2MsZWCt7FJU2YlPGP7n9MASTRLSL%2Bm5WfGYL9fouaM/qDhvsmTYuE8VrMyZxoN1HWlv3ZoAm%2B1pmd5japXupd2jalmJHCcaDGBpytofdqZucKMzb3xP%2BgowWcUjdBKOiDtsAOjJB6m3HWPa7o1jXCpSEgipjWt9vkwMTwN90G0RN5lU4L8%2BU1IwOSwGQyIWQ5Q0HXenLHZwZ4S3fcmLX5/0bSzQ4FACHquXJlveLxFoZIm7cGQp9lDpePw%3D%3D; RPSSecOOnline=FABCARTmLF7/9zpgYvJ7Vwc033zfJwtYPwNmAAAEgAAACM2b7%2BP1YUxrAAH7sFbGQc9l1J2S6167KdwenW6A23Y6Fu%2BpZUYnyTvPedEAQ5%2BLbYbDjG0W2%2B9fAdQaWEAq12sN1t4U3hjFGONcrjd36Tfi0WnehUbzK1EDT75g8Jq3wYIbSle/pL1Lzq2OaGJWuXuWg1dk6R9fh/m8e%2BFma3ArGR3IWoinX4TRyty8A8THvlVWJLXP4pGHyvhvBRiMONQrLjWgK37m5oOzp6aB8YtbilaKMBhMymzwJHfokOfq3TBKgU7nHvKrv4csh0Oep37PgUD7IkyQltpQq75/yLm8AzSN6ULH/V%2Bm0sIsOT9ga%2Ba9iLGgujEVS30PwNsPaGgAFsW9nvBCJCWWFABDY6/Rq4BVmiauSy82ybWOWx7QrQ%3D%3D; RPSClearCT=EgCFAQMAAAAEgAAAC4AAu8EVGMKcR1kVWg2w6K7p1tfj2k62SL2fJioATH5xElqAvMUG2ErcS12XzxBnzu77ktsSPjmcZrFrNFkzXBzejugW11dmYsJy6dTa%2BrfhQ/alZ0H5y5Cgmj8bo0GOc7%2B/VXpTXNOcfK3dAWj0LdnSLQ9mlOEqLMrYRw5fNJ8uM4n0AGQA9AABQAMAJUuoy97ZgVTe2YFUVoUAAAoQIAAQFgBoby1pdEBhcHMtaG9sZGluZy5jb20AVgAAImhvLWl0JWFwcy1ob2xkaW5nLmNvbUBwYXNzcG9ydC5jb20AAAABQ1oABTExMDAwAAAAAAQFAgAAfv11QAAGQQADQVBTAAJJVAUMAAAAAAAAAAAAAAAAAM45aUfN1cAAAN7ZgVTfgPhUAAAAAAAAAAAAAAAADgA4Ni40OS4xODUuMTM4AAQBAAAAAAAAAAAAAAAAAQQAAAAAAAAAAAAAAAAAAADyZXPCfb04AwAAAAAAAAAAAAAAAAAAAAA%3D; ANON=A=6DEF26AB8F95B5F49494659DFFFFFFFF&E=104b&W=1; NAP=V=1.9&E=ff1&C=z6ZfSJmsOVc6vn5-0jl72tkZ7-Ys9pkLxZZsJgaFPFE5zsSobB9hcA&W=1; ODCSSLAuth=t=2014-12-05 16:14:23Z&h=su47UHMzjUXU8k+fTl+crun0MNI=; muxcan=50582522c2724b23b5df36ce01046941; ul=1");

            string cookies = _cookies;

            if (appendCookie != null)
            {
                cookies = appendCookie + "; " + cookies;
            }

            wr.Headers.Add(HttpRequestHeader.Cookie, cookies);
            return wr;
        }
    }
}
