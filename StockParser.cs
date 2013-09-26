using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Google.GData.Client;
using Google.GData.Spreadsheets;


namespace GoogleStockParser
{
    public class StockParser
    {
        Spreadsheets _spreadsheets;
        public XDocument Informations;

        public StockParser()
        {
            try
            {
                _spreadsheets = new Spreadsheets("your gmail adress", "your gmail password");
                Informations = new XDocument(new XElement("STOCK_REQUEST"));
            }
            catch (GDataRequestException) { throw new NoConnectionException(); }
            catch (ClientFeedException) { throw new NoConnectionException(); }
        }

        //format : "HO;IBM;AAPL", "01/01/13;01/05/13;DAILY"
        public List<Stock_History> GetStockHistory(string parameters)
        {
            List<string> parameter = parameters.Split('&').ToList();
            for (int i = 0; i < parameter.Count; i++) parameter[i] = "&" + parameter[i];

            if (parameter.Count != 2) throw new ParameterException();
            if (!parameter.First().StartsWith("&q") && !parameter.Last().StartsWith("&q")) throw new ParameterException();
            if (!parameter.First().StartsWith("&p") && !parameter.Last().StartsWith("&p")) throw new ParameterException();

            Dictionary<string, string> dico = new Dictionary<string, string>();

            foreach (string p in parameter)
            {
                if (dico.Keys.Contains(p.Split('=').First())) throw new ParameterException();
                else dico.Add(p.Split('=').First(), p.Split('=').Last());
            }
            try
            {
                return GetStockHistory(dico["&q"], dico["&p"]);
            }
            catch (GDataRequestException) { throw new NoConnectionException(); }
            catch (ClientFeedException) { throw new NoConnectionException(); }
        }
        public List<Stock_History> GetStockHistory(string stocks, string attributes)
        {
            List<string> stock_list = new List<string>();
            List<string> attribute_list = new List<string>();
            List<Historical_Data> historical_data = new List<Historical_Data>();

            foreach (string stock in stocks.Split(';'))
                stock_list.Add(stock);

            foreach (string attribute in attributes.Split(';'))
                attribute_list.Add(attribute);

            AttributeValidation(attribute_list);


            List<Stock_History> stock_history = new List<Stock_History>();

            Spreadsheet spreadsheet = _spreadsheets.At("Historical_Data");
            spreadsheet.Clear();

            foreach (string stock in stock_list)
            {
                List<List<string>> data = spreadsheet.Add_Retrieve("A1",
                    String.Format("=EXPAND(GoogleFinance(\"{0}\";\"all\";\"{1}\"; \"{2}\";\"{3}\"))", stock, attribute_list[0], attribute_list[1], attribute_list[2]),
                    stock, 100, 100);

                List<Historical_Data> data_list = new List<Historical_Data>();


                foreach (List<string> list in data)
                {
                    data_list.Add(new Historical_Data(list[0], list[1], list[2], list[3], list[4], list[5]));
                }

                stock_history.Add(new Stock_History(stock, data_list));
            }
            spreadsheet.Clear();
            XElement xmlHistory = new XElement("Stock_History_Request", new XAttribute("date", DateTime.Now.ToString("G")));
            foreach (Stock_History s_history in stock_history)
                xmlHistory.Add(s_history._XMLContent);

            Informations.Root.Add(xmlHistory);
            return stock_history;
        }
        public XElement GetLiveForex()
        {
            XElement data = default(XElement);
            try { data = XElement.Load("http://rates.fxcm.com/RatesXML"); }
            catch (WebException) { throw new NoConnectionException(); }

            XElement forexXml = new XElement("Forex_Request");
            forexXml.Add(data);
            Informations.Root.Add(forexXml);
            return forexXml;
        }
        public void SetHeader(string stocks)
        {
            List<string> stock_list = new List<string>();
            foreach (string stock in stocks.Split(';'))
                stock_list.Add(stock);

            try
            {
                _spreadsheets.At("Stock_Data").DeleteLastWorksheet();
                _spreadsheets.At("Stock_Data").Add("Sheet1", 100, 100);
                Worksheet wk = _spreadsheets.At("Stock_Data").Last();

                wk.Edit("A1", "name");
                wk.Edit("B1", "price");
                wk.Edit("C1", "currency");
                wk.Edit("D1", "change");
                wk.Edit("E1", "changepct");
                wk.Edit("F1", "volume");
                wk.Edit("G1", "marketcap");
                wk.Edit("H1", "delay");
                wk.Edit("I1", "pe");
                wk.Edit("J1", "eps");
                wk.Edit("K1", "high52");
                wk.Edit("L1", "low52");

                if (stocks.Length > 0)
                    SetStockData(stocks);
            }
            catch (GDataRequestException) { throw new NoConnectionException(); }
            catch (ClientFeedException) { throw new NoConnectionException(); }

        }
        public void SetStockData(string stocks)
        {

            List<string> stock_list = new List<string>();
            foreach (string stock in stocks.Split(';'))
                stock_list.Add(stock);

            stock_list.Sort();
            try
            {

                List<string> already_exist = AlreadyExist(stock_list);

                Worksheet wk = _spreadsheets.At("Stock_Data").Last();

                for (int i = 0; i < stock_list.Count; i++)
                {
                    if (already_exist.Contains(stock_list[i])) continue;
                    else
                    {
                        List<string> row = new List<string>();

                        row.Add(stock_list[i]);
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"price\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"currency\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"change\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"changepct\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"volume\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"marketcap\")/1000", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"datadelay\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"pe\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"eps\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"high52\")", stock_list[i]));
                        row.Add(String.Format("=GoogleFinance(\"{0}\";\"low52\")", stock_list[i]));

                        wk.Add_Row(row);
                    }
                }

                DeleteInvalidEntries();
            }
            catch (GDataRequestException) { throw new NoConnectionException(); }
            catch (ClientFeedException) { throw new NoConnectionException(); }
        }
        public List<Stock_Data> GetStockdata()
        {

            try
            {
                Worksheet wk = _spreadsheets.At("Stock_Data").Last();

                List<List<string>> all_data = wk.Get_Rows();
                List<Stock_Data> data_list = new List<Stock_Data>();

                //Multithread ?
                foreach (List<string> list in all_data)
                {
                    try
                    {
                        data_list.Add(new Stock_Data(list[0], list[1], list[2], list[3], list[4], list[5],
                            list[6], list[7], list[8], list[9], list[10], list[11]));
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        Console.WriteLine("Header Exception ! reSetting the header...");
                        List<List<string>> entries = wk.Get_Rows();
                        string data = String.Empty;

                        foreach (List<string> entry in entries)
                            data += (entry[0] + ";");

                        data = data.Substring(0, data.Length - 1);

                        SetHeader(data);/*throw new InvalidHeaderException();*/
                    }
                }
                XElement XMLStockData = new XElement("Stock_Data_Request", new XAttribute("date", DateTime.Now.ToString("G")));
                foreach (Stock_Data data in data_list)
                {
                    XMLStockData.Add(data._XMLContent);
                }

                Informations.Root.Add(XMLStockData);
                return data_list;
            }
            catch (GDataRequestException) { throw new NoConnectionException(); }
            catch (ClientFeedException) { throw new NoConnectionException(); }

        }
        public void DeleteRows(string stocks)
        {
            try
            {
                Worksheet wk = _spreadsheets.At("Stock_Data").Last();

                List<string> stock_list = new List<string>();
                foreach (string stock in stocks.Split(';'))
                    stock_list.Add(stock);

                stock_list.Sort();

                List<int> positions = new List<int>();
                string cols = "A1";
                int nb = 1;
                while (wk.At(cols) != null)
                {
                    if (stock_list.Contains(wk.At(cols).Entry.Value)) positions.Add(nb);
                    cols = "A" + (++nb);
                }

                ListFeed listFeed = wk.Get_List_Feed();
                List<ListEntry> rows = new List<ListEntry>();

                foreach (int v in positions)
                    rows.Add((ListEntry)listFeed.Entries[v - 2]);

                foreach (ListEntry row in rows)
                    row.Delete();
            }
            catch (GDataRequestException) { throw new NoConnectionException(); }
            catch (ClientFeedException) { throw new NoConnectionException(); }
        }

        private List<string> AlreadyExist(List<string> stocks)
        {
            List<string> already_exist = new List<string>();
            Worksheet wk = _spreadsheets.At("Stock_Data").Last();

            string cols = "A1";
            int col_size = 1;
            while (wk.At(cols) != null)
            {
                if (stocks.Contains(wk.At(cols).Entry.Value)) already_exist.Add(wk.At(cols).Entry.Value);
                col_size++;
                cols = "A" + col_size.ToString();
            }

            return already_exist;
        }
        private void AttributeValidation(List<string> attributes)
        {

            if (!attributes.Contains("DAILY") && !attributes.Contains("MONTHLY"))
                throw new ParameterException();

            CultureInfo provider = new CultureInfo("en");
            string format = "mm/dd/yy";
            DateTime date1 = new DateTime(), date2 = new DateTime();
            try
            {
                date1 = DateTime.ParseExact(attributes[0], format, provider);
                date2 = DateTime.ParseExact(attributes[1], format, provider);
            }
            catch (Exception) { throw new ParameterException(); }

            if (DateTime.Compare(date1, date2) >= 0)
                throw new ParameterException();
            if (DateTime.Compare(date1, DateTime.Now) >= 0)
                throw new ParameterException();
            if (DateTime.Compare(date2, DateTime.Now) >= 0)
                throw new ParameterException();


        }

        private void DeleteInvalidEntries()
        {

            Worksheet wk = _spreadsheets.At("Stock_Data").Last();
            List<List<string>> all_data = wk.Get_Rows();
            string entryToDelete = String.Empty;

            foreach (List<string> list in all_data)
                if (list[2] == "#N/A" && list[3] == "#N/A" && list[4] == "#N/A") entryToDelete += list[0] + ";";

            if (entryToDelete.Length > 0)
                entryToDelete = entryToDelete.Substring(0, entryToDelete.Length - 1);

            DeleteRows(entryToDelete);

        }
    }


    public struct Stock_History
    {
        public string _name;
        public List<Historical_Data> _data;
        public XElement _XMLContent;

        public Stock_History(string name, List<Historical_Data> data)
        {
            _name = name;
            _data = data;
            _XMLContent = null;
            _XMLContent = Get_XmlElement();

        }

        private XElement Get_XmlElement()
        {
            XElement data = new XElement("Stock_History", new XAttribute("name", _name));
            foreach (Historical_Data h_data in _data)
                data.Add(h_data._XMLContent);


            return data;
        }

    }
    public struct Historical_Data
    {
        public string _date;
        public string _open;
        public string _close;
        public string _high;
        public string _low;
        public string _volume;

        public XElement _XMLContent;

        public Historical_Data(string date, string open, string close, string high, string low, string volume)
        {
            _date = date;
            _open = open;
            _close = close;
            _high = high;
            _low = low;
            _volume = volume;
            _XMLContent = null;
            _XMLContent = Get_XmlElement();

        }


        private XElement Get_XmlElement()
        {
            XElement data =
                new XElement("data",
                    new XAttribute("date", _date),
                    new XAttribute("open", _open),
                    new XAttribute("close", _close),
                    new XAttribute("high", _high),
                    new XAttribute("low", _low),
                    new XAttribute("volume", _volume)
                    );
            return data;
        }



    }
    public struct Stock_Data
    {
        public string _date;
        public string _tickerSymbol;
        public string _price;
        public string _currency;
        public string _change;
        public string _changepct;
        public string _volume;
        public string _marketCap;
        public string _delay;
        public string _priceToEarnings;
        public string _erningsPerShare;
        public string _52WeekHigh;
        public string _52WeekLow;

        public XElement _XMLContent;

        public Stock_Data(string tickerSymbol, string price, string currency, string change, string changepct, string volume, string marketCap, string delay, string pe, string eps, string Wkhigh, string wkLow)
        {
            _date = DateTime.Now.ToString("G");
            _tickerSymbol = tickerSymbol;
            _price = price;
            _currency = currency;
            _change = change;
            _changepct = changepct;
            _volume = volume;
            _marketCap = marketCap;
            _delay = delay;
            _priceToEarnings = pe;
            _erningsPerShare = eps;
            _52WeekHigh = Wkhigh;
            _52WeekLow = wkLow;

            _XMLContent = null;
            _XMLContent = Get_XmlElement();
        }


        private XElement Get_XmlElement()
        {
            XElement data =
                new XElement("data",
                    new XAttribute("ticker_symbol", _tickerSymbol),
                    new XElement("date", _date),
                    new XElement("price", _price),
                    new XElement("currency", _currency),
                    new XElement("change", _change),
                    new XElement("change_pct", _changepct),
                    new XElement("volume", _volume),
                    new XElement("market_cap", _marketCap),
                    new XElement("delay", _delay),
                    new XElement("price_to_earnings", _priceToEarnings),
                    new XElement("earnings_per_share", _erningsPerShare),
                    new XElement("high52", _52WeekHigh),
                    new XElement("low52", _52WeekLow)
                    );

            return data;
        }


    }


}
