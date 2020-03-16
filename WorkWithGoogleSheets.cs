// before using google apis we need to download the JSON licence
// how to use :) 
// https://www.twilio.com/blog/2017/03/google-spreadsheets-and-net-core.html?utm_source=youtube&utm_medium=video&utm_campaign=google-sheets-dotnet
// https://www.youtube.com/watch?v=afTiNU6EoA8&list=LLVHPD40ves6a773yXkhb5dA&index=5&t=627s 
 

// importing / dowloading the needed packeges 
using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace training
{
    class Program // general class 
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Legislators";
        static readonly string SpreadsheetId = "1FTCc621vMskmgoRr5oNG0VPtGkwc-nTSMjNFlRg1ymk";   // taken from the google spreadsheets url in the middle 
        static readonly string sheet = "congress";  // sheet name that we are using 
        static SheetsService service;
        static void Main(string[] args)  // run file function 
        {
            Console.WriteLine("Hello World!");
            GoogleCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))  // open access to the spreadsheets with the api credentials 
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // running functions 
             
            CreateEntry();   // running the writing data function 
            UpdateEntry();   // running the update cell function 
            DeleteEntry();   // running delete data cell function 
            ReadEntries();   // running the reading data function
        }

        // reading data
        static void ReadEntries()     // reading from spreadsheet 
        {
            var range = $"{sheet}!A1:A20";   // select the square range that we want to print should be identical to the row num
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            Console.WriteLine(values);
       
            foreach (var row in values)  // double loop to print all the cells in the range 
            {
                foreach (var i in row)
                {
                    Console.WriteLine(i);
                }
               
            }
        
        }

        // creating data 
        static void CreateEntry() 
        { 
            var range = $"{sheet}!A16:A16";  // select the square range that we want to write into 
            var valueRange = new ValueRange(); 

            // creating the string or list we want to write in
            var oblist = new List<object>() {"Alex Dezho is in the house"};  
            valueRange.Values = new List<IList<object>> { oblist };

            // appending data to the cell , range , shpreadsheet id, range 
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }

        // updating 

        static void UpdateEntry()
        {
            var range = $"{sheet}!A6:A6";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "wow" };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();
        }

        // deleting cell 

        static void DeleteEntry()
        {
            var range = $"{sheet}!A16:A16";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var deleteReponse = deleteRequest.Execute();
        }


    }
}
