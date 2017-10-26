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
imports Microsoft.Extensions.Configuration

Module Program
	readonly property Scopes as string() = { SheetsService.Scope.Spreadsheets }
	readonly property ApplicationName as string = "CryptoPortfolio_Logger"
	property service as SheetsService
	
	property CurrencyPairs as List(of CryptoRate)
	property EURUSDRate as decimal = 0
	
	property Configuration as IConfigurationRoot

    Sub Main(args As String())
        
		dim builder as new ConfigurationBuilder() 
		builder.SetBasePath(Directory.GetCurrentDirectory())
        builder.AddJsonFile("appconfig.json")

        Configuration = builder.Build()
		
		console.writeline("Loaded Configuration")
		
		dim sheetList as new list(of SheetDetail)
		Configuration.GetSection("sheets").Bind(SheetList)
		for each sd as SheetDetail in SheetList
			console.writeline( "Loaded sheet from Config: " & sd.id & " - " & sd.tabname & " - " & sd.basecurrency)
		next
		
		Console.writeline("Connect Google...")
		dim credential as GoogleCredential
		using stream as new FileStream("client_secret.json", FileMode.Open, FileAccess.Read)
		    credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
		end using

		'//Create Google Sheets API service.
		service = new SheetsService(new BaseClientService.Initializer() with {
			.HttpClientInitializer = credential,
			.ApplicationName = ApplicationName
		})
		
		Console.writeline("Get EURUSD")
		EURUSDRate = GetEURUSD()
		
		Console.writeline("Get Crypto")
		LoadAllCurrencies()
		
		'Begin Logging
		ReadEntries(sheetList)
    End Sub
	
	private function GetEURUSD() as decimal
		'### Get the initial range
		dim request as SpreadsheetsResource.ValuesResource.GetRequest = service.Spreadsheets.Values.Get("1Q_2IIJHYOualB2RBwbYVvyq16FnqDydCP4COGjok2tA", "JPL Portfolio!K9")
		dim response as Object = request.Execute()
		
		'### Values = raw objects
	    dim values as IList(of IList(of Object)) = response.Values
		console.writeline( "EURUSD: " & values(0)(0) )
		return values(0)(0)
	end function
	
	private sub LoadAllCurrencies()
		CurrencyPairs = new list(of CryptoRate)
		currencyPairs.add( DownloadCryptoRate(Crypto.BTC, FIAT.USD) )
		currencyPairs.add( DownloadCryptoRate(Crypto.BTC, FIAT.EUR) )
		currencyPairs.add( DownloadCryptoRate(Crypto.ETH, FIAT.USD) )
		currencyPairs.add( DownloadCryptoRate(Crypto.ETH, FIAT.EUR) )
		currencyPairs.add( DownloadCryptoRate(Crypto.LTC, FIAT.USD) )
		currencyPairs.add( DownloadCryptoRate(Crypto.LTC, FIAT.EUR) )
		currencyPairs.add( DownloadCryptoRate(Crypto.XRP, FIAT.USD) )
		currencyPairs.add( DownloadCryptoRate(Crypto.XRP, FIAT.EUR) )
		currencyPairs.add( DownloadCryptoRate(Crypto.OMG, FIAT.USD, "bitfinex") ) '## bitstamp doesn't support this currency
	end sub
	
	private function DownloadCryptoRate( c as crypto, f as fiat, optional market as string = "bitstamp") as CryptoRate
		dim cr as new CryptoRate
		cr.cryptocode = c
		cr.fiatcode = f
		cr.rate = getCurrencyValue(c, f, market)
		console.writeline( c.tostring.tolower & "/" & f.tostring.tolower & ": " & cr.rate)
		return cr
	end function
	private function getCurrencyValue(c as Crypto, f as fiat, market as string) as decimal
		dim theURL as string = GetCryptoWatchURL( f.tostring.toLower(), c.toString.toLower(), market)
		dim stringResponse as string = ""
		using wc as new WebClient() 
			stringResponse = wc.DownloadString(theURL)
		end using
		
		dim d as object  = Newtonsoft.Json.Linq.JObject.Parse(stringResponse)
		dim currentVal as decimal = d("result")("price").toString() 
		return currentVal
	end function
	
	sub ReadEntries(sheetList as list(of SheetDetail))
		console.writeline("Reading sheets...")
		for each sd as SheetDetail in SheetList
			doSheet(sd)
		next
	end sub
	
	sub doSheet( sd as SheetDetail )
		console.writeline( sd.id & "-----" & sd.tabname & " --- " & "B18:O" & " --- " & sd.basecurrency )
		doSheet( sd.id, sd.tabname, "B18:O", sd.basecurrency )
	end sub
	
	sub DoSheet( spreadsheetID as string, sheetName as string, historyrange as string, mainFiat as FIAT )
		
		CheckCurrencies(spreadsheetid, SheetName, mainFIAT)

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
			ArchiveCurrencies(Spreadsheetid, SheetName, mainFIAT)
			allHistoryRows.add( newItem )
			UpdateSpreadsheet( spreadsheetid, range, allHistoryRows )
		else
			console.writeline("not Updating")
		end if
		console.writeline("**********")
	end sub
	
	private sub UpdateSpreadSheet( spreadsheetid as string, range as string, allHistoryRows as list(of HistoryRow) )
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
	
	private sub ArchiveCurrencies(spreadsheetid as string, sheetname as string, mainFIAT as FIAT)
		UpdateCurrency( spreadsheetid, $"{sheetname}!M4", GetCryptoRate(mainFIAT, crypto.btc) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!M5", GetCryptoRate(mainFIAT, crypto.eth) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!M6", GetCryptoRate(mainFIAT, crypto.xrp) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!M7", GetCryptoRate(mainFIAT, crypto.omg) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!M8", GetCryptoRate(mainFIAT, crypto.ltc) )
	end sub
	
	private sub CheckCurrencies(spreadsheetid as string, sheetname as string, mainFIAT as FIAT)
		'dim btc_curr as CryptoRate = 
		UpdateCurrency( spreadsheetid, $"{sheetname}!K4", GetCryptoRate(mainFIAT, crypto.btc) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!K5", GetCryptoRate(mainFIAT, crypto.eth) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!K6", GetCryptoRate(mainFIAT, crypto.xrp) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!K7", GetCryptoRate(mainFIAT, crypto.omg) )
		UpdateCurrency( spreadsheetid, $"{sheetname}!K8", GetCryptoRate(mainFIAT, crypto.ltc) )
	end sub
	
	private function GetCryptoRate(f as Fiat, c as Crypto) as CryptoRate
		'Get the rate for the f/c pair, but if not available in EUR then get the USD rate and convert
		'based on current USD/EUR rate
		console.writeline($"Cached {f.tostring.tolower} / {c.tostring.tolower} ...")
		dim theRate as CryptoRate = CurrencyPairs.Find(function(x) x.CryptoCode = c and x.Fiatcode = f)
		if theRate is nothing andalso f = fiat.EUR then
			'try getting the USD rate 
			dim USDRate = GetCryptoRate(Fiat.USD, c)
			if USDRate isnot nothing then
				theRate = new CryptoRate
				therate.fiatcode = f
				therate.cryptocode = c
				therate.rate = USDRate.rate / EURUSDRate
				console.writeline($"Could not get Eur Rate... got this instead: {therate.rate}")
			end if
		end if
		console.writeline($"{therate.rate}")
		return therate
	end function
	
	private function GetCryptoWatchURL(curr as string, coinCode as string, optional market as string = "bitstamp") as string
		return $"https://api.cryptowat.ch/markets/{market}/{coinCode}{curr}/price"
	end function
	
	private sub UpdateCurrency( spreadsheetid as string, range as string, val as CryptoRate )
		doFinalUpdate(spreadsheetid, range, val.rate)
	end sub
	
	sub DoFinalUpdate(spreadsheetid as string, range as string, currentVal as decimal)
		dim valueRange as new ValueRange()
			
		dim oblist as new List(of object)
		oblist.add(currentVal)
			
		dim rows as new List(of IList(of object))
		rows.add(oblist)
			
		valueRange.Values = rows
		 
		dim updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range)
		updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED
		dim appendReponse = updateRequest.Execute()
	end sub
	
	private function GetCurrentRow() as IList(of object)
		return { "=NOW()", "=G4", "=H4", "=G5", "=H5", "=G6", "=H6", "=G8", "=H8", "=G7", "=H7", "=H10", "=H11", ""}
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
	
	class SheetDetail
		public property id as string
		public property tabname as string
		public property basecurrency as FIAT
	end class
	class CryptoRate
		public property CryptoCode as Crypto
		public property FIATCode as FIAT
		public property rate as decimal
	end class
	enum Crypto
		OMG
		BTC
		ETH
		LTC
		XRP
	end enum
	enum FIAT
		USD
		EUR
	end enum
	
	
	class HistoryRow
		public property dt as DateTime
		public property amnt_BTC as decimal
		public property usd_BTC as decimal
		public property amnt_ETH as decimal
		public property usd_ETH as decimal
		public property amnt_XRP as decimal
		public property usd_XRP as decimal
		public property amnt_LTC as decimal
		public property usd_LTC as decimal
		public property amnt_OMG as decimal
		public property usd_OMG as decimal
		public property total_USD as decimal
		public property total_GBP as decimal
		public property notes as string
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
			me.amnt_OMG = CurrencyStringToDecimal(row(9))
			me.usd_OMG = CurrencyStringToDecimal(row(10))
			me.total_USD = CurrencyStringToDecimal(row(11))
			me.total_GBP = CurrencyStringToDecimal(row(12))
			if row.count > 13 then me.notes = row(13)
		end sub
		public function asObj() as list(of Object)
			dim l as List(of Object) 
			l = { dt.toString("yyyy-MM-dd HH:mm"), 
				amnt_BTC, usd_BTC, 
				amnt_ETH, usd_ETH,
				amnt_XRP, usd_XRP,
				amnt_LTC, usd_LTC,
				amnt_OMG, usd_OMG,
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
