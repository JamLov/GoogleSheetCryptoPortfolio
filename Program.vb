imports System
imports System.IO
imports System.net
imports System.Collections
imports System.Collections.Generic
imports Google.Apis.Services
imports Google.Apis.Sheets.v4
imports Google.Apis.Sheets.v4.Data
imports Google.Apis.Auth.OAuth2
imports newtonsoft.json
imports Newtonsoft.Json.Linq

Module Program
	readonly property Scopes as string() = { SheetsService.Scope.Spreadsheets }
	readonly property ApplicationName as string = "CryptoPortfolio_Logger"
	readonly property SpreadsheetId as string = "1Q_2IIJHYOualB2RBwbYVvyq16FnqDydCP4COGjok2tA"
	property service as SheetsService
	
    Sub Main(args As String())
        Console.WriteLine("Hello World!")
		
		'Retrieve credentials from Google JSON object
		dim credential as GoogleCredential
		using stream as new FileStream("client_secret.json", FileMode.Open, FileAccess.Read)
		    credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
		end using

		'//Create Google Sheets API service.
		service = new SheetsService(new BaseClientService.Initializer() with {
			.HttpClientInitializer = credential,
			.ApplicationName = ApplicationName
		})
		
		'Begin Logging
		ReadEntries()
    End Sub
	
	'Performs logging on multiple sheets within the opened spreadsheet
	sub ReadEntries()
		doSheet( "JPL Portfolio", "B18:M", "USD" )
		doSheet( "JvdS Portfolio", "B17:M", "EUR")
	end sub
	
	'Retrieves the currencies from Cryptowatch APIs, downloads the relevant range and adds a row
	'if necessary.  Also logs the 2-hourly rate for comparison to the current rate
	' 1. Downloads the current rate for each currency (using the @curr base rate, usd or eur for example) and stamps
	'    this in the spreadsheet in column K4
	' 2. Then downloads the whole history range from the spreadsheet, for historical data.  
	' 3. Parse this into a list of HistoryRow objects, for ease of use
	' 4. If it has been over 2 hours since the last update, add a new row with the current
	'    values and add a final 'up to date' row using the formulae for the sheet to get values from 
	'    row K
	' 5. Also if >2 hours have gone by, update the 2 hourly rate in column M so that our
	'    green/red indicators on the current rate are up to date
	sub DoSheet(  sheetName as string, historyrange as string, curr as string )
	
		CheckCurrencies(SheetName, curr)

	    dim range as string = $"{sheetName}!{historyrange}"
		console.writeline("*****************")
		console.writeline(" Sheet Range: " & range)
		console.writeline("******************")
		
		'### Get the initial range
		dim request as SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.Get(SpreadsheetId, range)
		dim response as Object = request.Execute()
		
		'### Values = raw objects
	    dim values as IList(of IList(of Object)) = response.Values
		
		'### parse values into nice list of typed objects
		dim allHistoryRows as new list(of HistoryRow)
	    if (values isnot nothing andalso values.Count > 0) then 
	        for each row as list(of Object) in values
	        	try
					Console.WriteLine(renderRow(row))
				catch e as Exception
					Console.WriteLine(e.ToString())
				end try
				try
					allHistoryRows.add( new HistoryRow(row) )
				catch e as Exception
					Console.WriteLine(e.ToString())
				end try
	        next
	    else
	        Console.WriteLine("No data found.")
	    end if
	
		console.writeline(" ************************** ")	
		'### Get the Last item (which is the CURRENT row...) and create a new
		'### row which will be our new 'fixed' penultimate row in the final list
		dim lastItem as HistoryRow = allHistoryRows.Last()
		dim newItem as HistoryRow = lastItem.Clone()
		newItem.notes = "Auto Added : " & datetime.now()
		
		'### remove the last item so that we dont destroy the 'current' row..
		allHistoryRows.removeAt( allhistoryrows.count-1 )
		
		'### Has it been more than 6 hours since the last update? If so, add...
		dim TimeSinceLastUpdate as Timespan = DateTime.Now().Subtract(allHistoryRows.last().dt)
		console.writeline( "Minutes since last update... " & timeSincelastUpdate.totalMinutes() )
		if timeSincelastUpdate.TotalMinutes >= 120 then
			ArchiveCurrencies(SheetName, curr)
			allHistoryRows.add( newItem )
			UpdateSpreadsheet( range, allHistoryRows )
		else
			console.writeline("not Updating")
		end if
		console.writeline("**********")
	end sub
	
	private sub UpdateSpreadSheet( range as string, allHistoryRows as list(of HistoryRow) )
		for each r as HistoryRow in allHistoryRows
			console.writeline( RenderRow(r.asObj) )
		next
		
		dim ValRange as new ValueRange()
		valRange.Values = HistoryRowsToObjectList( allHistoryRows )
		'### add in "Current" row...
		valRange.Values.add( GetCurrentRow() )
		
		dim updateRequest = service.Spreadsheets.Values.Update(valRange, SpreadsheetId, range)
	    updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
	    dim appendReponse = updateRequest.Execute()
		
	end sub
	
	private sub ArchiveCurrencies(sheetname as string, curr as string)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!M4", GetCoinMarketCapURL(curr, "bitcoin"), curr)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!M5", GetCoinMarketCapURL(curr, "ethereum"), curr)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!M6", GetCoinMarketCapURL(curr, "ripple"), curr)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!M7", GetCoinMarketCapURL(curr, "litecoin"), curr)
		UpdateCurrencyBitStamp( $"{sheetname}!M4", GetBitStampURL(curr, "btc"), curr, "btc")
		UpdateCurrencyBitStamp( $"{sheetname}!M5", GetBitStampURL(curr, "eth"), curr, "eth")
		UpdateCurrencyBitStamp( $"{sheetname}!M6", GetBitStampURL(curr, "xrp"), curr, "xrp")
		UpdateCurrencyBitStamp( $"{sheetname}!M7", GetBitStampURL(curr, "ltc"), curr, "ltc")
	end sub
	
	private sub CheckCurrencies(sheetname as string, curr as string)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!K4", GetCoinMarketCapURL(curr, "bitcoin"), curr)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!K5", GetCoinMarketCapURL(curr, "ethereum"), curr)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!K6", GetCoinMarketCapURL(curr, "ripple"), curr)
		'UpdateCurrencyCoinMarketCap( $"{sheetname}!K7", GetCoinMarketCapURL(curr, "litecoin"), curr)
		UpdateCurrencyBitStamp( $"{sheetname}!K4", GetBitStampURL(curr, "btc"), curr, "btc")
		UpdateCurrencyBitStamp( $"{sheetname}!K5", GetBitStampURL(curr, "eth"), curr, "eth")
		UpdateCurrencyBitStamp( $"{sheetname}!K6", GetBitStampURL(curr, "xrp"), curr, "xrp")
		UpdateCurrencyBitStamp( $"{sheetname}!K7", GetBitStampURL(curr, "ltc"), curr, "ltc")
	end sub
	
	private function GetBitStampURL(curr as string, coinCode as string) as string
		return $"https://api.cryptowat.ch/markets/bitstamp/{coinCode}{curr}/price"
	end function
	private function GetCoinMarketCapURL(curr as string, coinName as string) as string
		return $"https://api.coinmarketcap.com/v1/ticker/{coinName}/?convert={curr}"
	end function
	
	private sub UpdateCurrencyCoinMarketCap( range as string, URL as string, curr as string )
		dim stringResponse as string = ""
		using wc as new WebClient() 
			stringResponse = wc.DownloadString(url)
		end using
		
		dim d as object  = Newtonsoft.Json.Linq.JArray.Parse(stringResponse)
		dim currentUSD as string = d(0)("price_" & curr.tolower()).toString()
		dim name as string = d(0)("name").tostring()
		dim symbol as string = d(0)("symbol").tostring()
		
		console.writeline( $"{name}/{symbol} : '{currentUSD}'" )
		DoFinalUpdate(range, currentUSD)
	end sub
	private sub UpdateCurrencyBitStamp( range as string, URL as string, curr as string, coin as string )
		dim stringResponse as string = ""
		using wc as new WebClient() 
			stringResponse = wc.DownloadString(url)
		end using
		
		dim d as object  = Newtonsoft.Json.Linq.JObject.Parse(stringResponse)
		dim currentUSD as string = d("result")("price").toString() 
		console.writeline( $"{coin} : '{currentUSD}'" )
		doFinalUpdate(range, currentUSD)
	end sub
	
	sub DoFinalUpdate(range as string, currentUSD as string)
		dim i as integer = 0
		if Decimal.TryParse( currentUSD, i ) then 
			dim valueRange as new ValueRange()
			
			dim oblist as new List(of object)
			oblist.add(currentUSD)
			
			dim rows as new List(of IList(of object))
			rows.add(oblist)
			
			valueRange.Values = rows
		 
			dim updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range)
			updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
			dim appendReponse = updateRequest.Execute()
		else
			console.writeline("NOT NUMERIC")
		end if 
	end sub
	
	private function GetCurrentRow() as IList(of object)
		return { "=NOW()", "=G4", "=H4", "=G5", "=H5", "=G6", "=H6", "=G7", "=H7", "=H9", "=H10", ""}
	end function
	
	private function HistoryRowsToObjectList( rows as list(of HistoryRow) ) as iList(of iList(of Object) )
		dim values as new List(of iList(of Object))
		for each hr as HistoryRow in rows
			values.add( hr.asObj() )
		next
		return values
	end function
	
	private function renderRow( row as list(of Object) ) as string
		dim x as string = ""
		for each columnval as string in row
			x &= columnval & " | "
		next
		return x
	end function
		
	class HistoryRow
		public dt as DateTime
		public amnt_BTC as decimal
		public usd_BTC as decimal
		public amnt_ETH as decimal
		public usd_ETH as decimal
		public amnt_XRP as decimal
		public usd_XRP as decimal
		public amnt_LTC as decimal
		public usd_LTC as decimal
		public total_USD as decimal
		public total_GBP as decimal
		public notes as string
		private function CurrencyStringToDecimal( o as object ) as decimal
			if o.tostring() = "" then o = "0"
			return o.toString().replace("$","").replace("£","").replace("€", "")
		end function
		public sub new(row as IList(of Object))
			me.dt = DateTime.Parse(row(0))
			me.amnt_BTC = CurrencyStringToDecimal(row(1))
			me.usd_BTC = CurrencyStringToDecimal(row(2))
			me.amnt_ETH = CurrencyStringToDecimal(row(3))
			me.usd_ETH = CurrencyStringToDecimal(row(4))
			me.amnt_XRP = CurrencyStringToDecimal(row(5))
			me.usd_XRP = CurrencyStringToDecimal(row(6))
			me.amnt_LTC = CurrencyStringToDecimal(row(7))
			me.usd_LTC = CurrencyStringToDecimal(row(8))
			me.total_USD = CurrencyStringToDecimal(row(9))
			me.total_GBP = CurrencyStringToDecimal(row(10))
			if row.count > 11 then me.notes = row(11)
		end sub
		public function asObj() as list(of Object)
			dim l as List(of Object) 
			l = { dt.toString("yyyy-MM-dd HH:mm"), 
				amnt_BTC, usd_BTC, 
				amnt_ETH, usd_ETH,
				amnt_XRP, usd_XRP,
				amnt_LTC, usd_LTC,
				total_USD, total_GBP,
				notes
			}.toList()
			return l
		end function
		public function clone() as HistoryRow
			return new HistoryRow( asObj() )
		end function
	end class
End Module
