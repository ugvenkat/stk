using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;

class Program
{
    static async Task Main(string[] args)
    {
        string inputFilePath = @"C:\BarchartData\Technical-Google-Download\watchlist-feb17-intraday-02-20-2025_night.xlsx"; // Update with actual file path

        string bearerToken = "exysysysysysys"; // Replace with actual token
        int numberOfWeeks = 4; // Example: Get data for the next 4 Fridays

        var tickers = ReadTickersFromExcel(inputFilePath);
        var callData = new List<object[]>();
        var putData = new List<object[]>();

        var fridayDates = GetNextFridays(numberOfWeeks);

        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {bearerToken}");

            foreach (var ticker in tickers)
            {
                if (ticker != "Symbol")
                {
                    try
                    {
                        decimal currentPriceDecimal = await GetCurrentStockPrice(client, ticker);
                        int roundedPrice = (int)currentPriceDecimal;

                        foreach (var strikeDate in fridayDates)
                        {
                            var callOption = await GetOptionData(client, ticker, strikeDate.ToString("yyyy-MM-dd"), "call", roundedPrice + 3);
                            var putOption = await GetOptionData(client, ticker, strikeDate.ToString("yyyy-MM-dd"), "put", roundedPrice - 3);

                            if (callOption != null && callOption.AskPrice > 0.50m)
                                callData.Add(new object[] { ticker, strikeDate.ToString("yyyy-MM-dd"), currentPriceDecimal, callOption.StrikePrice, callOption.BidPrice, callOption.AskPrice });

                            if (putOption != null && putOption.AskPrice > 0.50m)
                                putData.Add(new object[] { ticker, strikeDate.ToString("yyyy-MM-dd"), currentPriceDecimal, putOption.StrikePrice, putOption.BidPrice, putOption.AskPrice });
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing ticker {ticker}: {ex.Message}");
                    }
                }
            }
        }

        WriteToExcel(inputFilePath, "CALL Options", callData);
        WriteToExcel(inputFilePath, "PUT Options", putData);
    }

    static List<string> ReadTickersFromExcel(string filePath)
    {
        var tickers = new List<string>();
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1);
            foreach (var row in worksheet.RowsUsed())
                tickers.Add(row.Cell(1).GetString());
        }
        return tickers;
    }

    static async Task<decimal> GetCurrentStockPrice(HttpClient client, string ticker)
    {
        string url = $"https://api.robinhood.com/quotes/{ticker}/";
        var response = await client.GetStringAsync(url);
        var json = JObject.Parse(response);
        return json["last_trade_price"].Value<decimal>();
    }

    static async Task<OptionData> GetOptionData(HttpClient client, string ticker, string expiration, string type, decimal strikePrice)
    {
        string url = $"https://api.robinhood.com/options/instruments/?chain_symbol={ticker}&expiration_dates={expiration}&state=active&type={type}";
        var response = await client.GetStringAsync(url);
        var json = JObject.Parse(response);

        foreach (var option in json["results"])
        {
            if (option["strike_price"].Value<decimal>() == strikePrice)
            {
                string marketDataUrl = $"https://api.robinhood.com/marketdata/options/{option["id"].Value<string>()}/";
                var marketDataResponse = await client.GetStringAsync(marketDataUrl);
                var marketJson = JObject.Parse(marketDataResponse);
                return new OptionData
                {
                    StrikePrice = strikePrice,
                    BidPrice = marketJson["bid_price"].Value<decimal>(),
                    AskPrice = marketJson["ask_price"].Value<decimal>()
                };
            }
        }
        return null;
    }

    static void WriteToExcel(string filePath, string sheetName, List<object[]> data)
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheets.Contains(sheetName) ? workbook.Worksheet(sheetName) : workbook.AddWorksheet(sheetName);
            worksheet.Clear();
            worksheet.Row(1).Cell(1).InsertData(new List<object[]> { new object[] { "Ticker", "Expiration Date", "Current Stock Price", "Strike Price", "Bid Price", "Ask Price" } });
            for (int i = 0; i < data.Count; i++)
                worksheet.Row(i + 2).Cell(1).InsertData(new List<object[]> { data[i] });
            workbook.Save();
        }
    }

    static List<DateTime> GetNextFridays(int weeks)
    {
        var fridays = new List<DateTime>();
        var today = DateTime.Today;
        var dayOfWeek = (int)today.DayOfWeek;
        var daysUntilFriday = (DayOfWeek.Friday - today.DayOfWeek + 7) % 7;
        var firstFriday = today.AddDays(daysUntilFriday);

        for (int i = 0; i < weeks; i++)
        {
            fridays.Add(firstFriday.AddDays(i * 7));
        }
        return fridays;
    }
}

class OptionData
{
    public decimal StrikePrice { get; set; }
    public decimal BidPrice { get; set; }
    public decimal AskPrice { get; set; }
}