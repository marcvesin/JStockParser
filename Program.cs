using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Google.GData.Client;
using Google.GData.Spreadsheets;


namespace GoogleStockParser
{
    class Program
    {
        static void Main(string[] args)
        {
            StockParser parser = new StockParser();
            List<Stock_History> l = parser.GetStockHistory("q=OR&p=11/01/12;01/12/12;DAILY");
            //parser.GetStockdata();
            //Console.WriteLine(parser.Informations);
            using (StreamWriter wr = new StreamWriter("lolilol.txt"))
            { wr.WriteLine(parser.Informations); }



            //using (StreamWriter wr = new StreamWriter("lol.txt"))
            //{
            //    wr.WriteLine(parser.Informations);
            //}
            //XmlDocument doc = new XmlDocument();
            //doc.Load("lol.txt");
            //XmlNode root = doc.DocumentElement;
            //XmlNodeList list = root.SelectNodes("Stock_History_Request/Stock_History");

            //Dictionary<string,XmlNodeList> data = new Dictionary<string,XmlNodeList>();

            //foreach (XmlNode a in list)
            //{
            //    string keyName = a.Attributes[0].Value;
            //    XmlNodeList list2 = root.SelectNodes("Stock_History_Request/Stock_History/"+keyName);
            //    data.Add(keyName, list2);
            //}



            
            //Console.WriteLine(parser.Informations);
            //q=HO;IBM&p=01/01/13;01/05/13;DAILY"
        }

    }
}


//Stock_Parser test = new Stock_Parser();

//using (StreamWriter writer = new StreamWriter("Output.txt"))
//{
//    //////get an historical data////////////// Format : Try_Get_Stock_History("HO;AAPL;...", "01/0/13;01/05/13;DAILY")////////
//    Console.WriteLine(test.Try_Get_Stock_History("HO;IBM;AAPL", "01/01/13;01/05/13;DAILY"));

//    ////////set a stock data////////////// Format : Set_Stock_Data("HO;kikoo;AAPL;IBM;...) //////
//    //test.Set_Stock_Data("HO;kikoo;AAPL;IBM");
//    /////// get all stock data////////////
//    //Console.WriteLine(test.Get_Stock_data());
//}