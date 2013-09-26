using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleStockParser
{
    public class StockException : Exception
    { }


    public class ParameterException : StockException
    {
        public ParameterException() : base()
        {
            HelpLink = "Wrong parameters. A valid format is  \"mm/dd/yy;mm/dd/yy;timespan\"  with: " +
                "\t=> date1 < date2  with date1 and date2 <= today" +
                "\n\t=> timespan = MONTHLY or DAILY";
        }

    }
    public class InvalidHeaderException : StockException
    {
        public InvalidHeaderException()
            : base()
        {
            HelpLink = "Non valid Header.";
        }
    
    }
    public class InvalidSpreadsheetException : StockException
    {
        public InvalidSpreadsheetException()
            : base()
        {
            HelpLink = "The SpreadSheet doesn't exist.";
        }
    }

    public class AlreadyExistException : StockException
    {
        public AlreadyExistException()
            : base()
        {
            HelpLink = "This name already exist, please choose an other one.";
        }


    }
    public class InvalidWorksheetException : StockException
    {

        public InvalidWorksheetException()
            : base()
        {
            HelpLink = "The Worksheet doesn't exist.";
        }
    
        
    
    }
    public class NoConnectionException : StockException
    {
        public NoConnectionException() : base()
        {
            HelpLink = "No connection found for the request.";
        }
    
    }

}
