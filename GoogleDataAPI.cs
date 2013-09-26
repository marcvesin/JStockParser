using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleStockParser
{
   public class Spreadsheets
    {
        SpreadsheetsService _service;
        List<Spreadsheet> _spreadsheets;
        Dictionary<string, int> _name;
        public SpreadsheetFeed Entry;

        public Spreadsheets(string username, string password, string name = "MySpreadsheetIntegration-v1")
        {
            _service = new SpreadsheetsService(name);
            _service.setUserCredentials(username, password);

            Entry = Get_Spreadsheets(_service);

            _spreadsheets = new List<Spreadsheet>();
            _name = new Dictionary<string, int>();
            int nb = 0;
            foreach (SpreadsheetEntry entry in Entry.Entries)
            {
                _spreadsheets.Add(new Spreadsheet(entry, _service));
                _name.Add(entry.Title.Text, nb);
                nb++;
            }
        }
        public void Display()
        {
            int nb = 0;

            foreach (Spreadsheet spreadsheet in _spreadsheets)
            {
                Console.WriteLine();
                Console.WriteLine(nb + ": " + spreadsheet.Entry.Title.Text);
                spreadsheet.Display();
                nb++;
            }
        
        }
        private static SpreadsheetFeed Get_Spreadsheets(SpreadsheetsService service)
        {
            SpreadsheetQuery query = new SpreadsheetQuery();
            return service.Query(query);
        }
        public Spreadsheet At(int number)
        {
            try { return _spreadsheets[number]; }
            catch (KeyNotFoundException) { throw new InvalidSpreadsheetException(); }
        }
        public Spreadsheet At(string name) {
            try { return _spreadsheets[_name[name]]; }
            catch (KeyNotFoundException) { throw new InvalidSpreadsheetException(); }
        }
        public Spreadsheet Last()
        {
            return _spreadsheets.Last();
        }
    
    }
   public class Spreadsheet
    {
        SpreadsheetsService _service;
        List<Worksheet> _spreadsheet;
        Dictionary<string, int> _name;
        public SpreadsheetEntry Entry;
        


        public Spreadsheet(SpreadsheetEntry entry,SpreadsheetsService service) {
            Entry = entry;
            _service = service;

            _spreadsheet = new List<Worksheet>();
            _name = new Dictionary<string, int>();

            WorksheetFeed feed = Entry.Worksheets;
            int nb = 0;
            foreach (WorksheetEntry wk in feed.Entries)
            {
                _name.Add(wk.Title.Text+"-"+nb, nb);
                _spreadsheet.Add(new Worksheet(wk, _service));
                nb++;
            }

        }
        public void Display()
        {
            WorksheetFeed wsFeed = Entry.Worksheets;
            int nb = 0;
            foreach (WorksheetEntry worksheet in wsFeed.Entries)
            {
                Console.WriteLine("\t" + nb + " : " + String.Format("{0} - rows:{1}  cols:{2}", worksheet.Title.Text, worksheet.Rows, worksheet.Cols));
                nb++;
            }
        }

        public List<List<string>> Add_Retrieve(string zone, string function,string title, uint rows, uint cols)
        {
            Add(title, rows, cols);
            At(title).Edit(zone, function);
            List<List<string>> data = At(title).Get_Rows();
            return data;
        }
        public WorksheetEntry Add(string title, uint rows, uint cols) 
        {

            if (_name.ContainsKey(title))
                throw new AlreadyExistException();

            WorksheetEntry wk = new WorksheetEntry();
            wk.Title.Text = title;
            wk.Rows = rows;
            wk.Cols = cols;
            
            _service.Insert((WorksheetFeed)Entry.Worksheets,wk);
            Update();

            return wk;
        }
        public void DeleteWorksheet(int number)
        {
            WorksheetFeed wsFeed = Entry.Worksheets;
            WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[number];
            worksheet.Delete();
            Update();        
        }
        //public void DeleteWorksheet(string title)
        //{
        //    WorksheetFeed wsFeed = Entry.Worksheets;
        //    WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[number];
        //    worksheet.Delete();
        //    Update();
        //}




        public void Clear()
        { 
            WorksheetFeed wsFeed = Entry.Worksheets;
            int s = _spreadsheet.Count;
            //for(int i = 1; i < s; i++)
            //{
            //    WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[i];
            //       worksheet.Delete();
            //        Update();
            
            //}
            Parallel.For(1, s, i =>
            {
                WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[i];
                worksheet.Delete();
                Update();
            });
                

        
        
        }
        public void DeleteLastWorksheet()
        {
            WorksheetFeed wsFeed = Entry.Worksheets;
            int last = wsFeed.Entries.Count-1;

            if (last < 0) { Console.WriteLine("Nothing to delete"); }
            else
            {
                DeleteWorksheet(last);
                Update();
            }
        }
        
        public Worksheet At(int number)
        {
            try { return _spreadsheet[number]; }
            catch (KeyNotFoundException){throw new InvalidWorksheetException();} 
        
        }
        public Worksheet At(string name)
        {
            try{return _spreadsheet[_name[name]];}
            catch (KeyNotFoundException) { throw new InvalidWorksheetException(); } 
            
        }
        public Worksheet Last()
        {
            return _spreadsheet[_spreadsheet.Count - 1];
        }


        private void Update()
        {
            _spreadsheet.Clear();
            _name.Clear();
            WorksheetFeed feed = Entry.Worksheets;
            foreach (WorksheetEntry wk in feed.Entries)
            {
                _spreadsheet.Add(new Worksheet(wk, _service));
                if (!_name.ContainsKey(wk.Title.Text)) _name.Add(wk.Title.Text, feed.Entries.Count - 1);
            }
        }
    }
   public class Worksheet
    {
        SpreadsheetsService _service;
        Dictionary<string,Cell> _worksheet;
        public WorksheetEntry Entry;
        

        public Worksheet(WorksheetEntry entry, SpreadsheetsService service) { 

            Entry = entry;
            _service = service;
            
            CellFeed cellFeed = Get_CellFeed();
            if (cellFeed != null)
            {
                _worksheet = new Dictionary<string, Cell>();
                foreach (CellEntry cell in cellFeed.Entries)
                    _worksheet.Add(cell.Title.Text, new Cell(cell, _service));
            }
        
        }
        public void Display() {
            Console.WriteLine(String.Format("{0} - rows:{1}  cols:{2}", Entry.Title.Text, Entry.Rows, Entry.Cols)); }
        public Cell At(string zone)
        {
            foreach (Cell cell in _worksheet.Values)
                if (cell.Entry.Title.Text == zone)
                    return cell;

            return null;
        }
        public Cell Edit(string zone, string value)
        {
            CellFeed feed = Get_CellFeed(true);
            foreach (CellEntry cell in feed.Entries)
                if (cell.Title.Text == zone)
                {
                    cell.InputValue = value;
                    cell.Update();
                    Update();
                }
            return _worksheet[zone];
        }
        public ListFeed Get_List_Feed()
        {
            AtomLink listFeedLink = Entry.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);
            ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
            return _service.Query(listQuery);
        }
        private CellFeed Get_CellFeed(bool returnEmpty = false)
        {
            //try { if (Entry.CellFeedLink == null) Console.WriteLine("null"); }
            //catch (Exception e) { return null; }

            CellQuery cellQuery = new CellQuery(Entry.CellFeedLink);
            if (returnEmpty) cellQuery.ReturnEmpty = ReturnEmptyCells.yes;

            CellFeed cellFeed = _service.Query(cellQuery); 
            
            return cellFeed;

        }
        private void Update()
        {
            CellFeed feed = Get_CellFeed();
            _worksheet.Clear();
            foreach (CellEntry cell in feed.Entries)
                _worksheet.Add(cell.Title.Text, new Cell(cell, _service));
        }
        public void Display_Rows()
        {
            foreach (ListEntry row in Get_List_Feed().Entries)
                foreach (ListEntry.Custom element in row.Elements)
                    Console.WriteLine(element.Value);

        }
        public List<List<string>> Get_Rows()
        {
            List<List<string>> rows = new List<List<string>>();

            foreach (ListEntry row in Get_List_Feed().Entries)
            {
                List<string> values = new List<string>();

                foreach (ListEntry.Custom element in row.Elements)
                    values.Add(element.Value);

                rows.Add(values);
            }

            return rows;
        }
        public List<string> Get_First_Row()
        {
            List<string> first_row = new List<string>();
            string cols = "A1";
            while (At(cols) != null)
            {
                first_row.Add(At(cols).Entry.Value);
                char next_col = (char)(Convert.ToUInt16(cols[0])+1);
                cols = next_col.ToString() + "1";
            }
            return first_row;
        
        }
        public void Add_Row(List<string> values)
        {
            ListFeed feed = Get_List_Feed();
            ListEntry row = new ListEntry();

            List<string> first_row = Get_First_Row();
            for (int i = 0; i < values.Count && i < first_row.Count; i++)
            {
                row.Elements.Add(new ListEntry.Custom() { LocalName = first_row[i], Value = values[i] });
            }

            _service.Insert(feed, row);
            Update();
        
        }
    
    }
   public class Cell
    {
        SpreadsheetsService _service;
        public CellEntry Entry;

        public Cell(CellEntry entry, SpreadsheetsService service) { Entry = entry; _service = service; }
        public Cell() { Entry = null; }
        public void Display()
        {
            if(Entry != null)
                Console.WriteLine(Get_Position(Entry) + " : " + Entry.Value);
        }
        static private string Get_Position(CellEntry cell)
        {
            return cell.Id.Uri.Content.Substring(cell.Id.Uri.Content.LastIndexOf("/") + 1);
        }
   
    
    }
    

}
