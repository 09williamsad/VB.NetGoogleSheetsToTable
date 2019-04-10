Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
Public Sub googleSheetGet()	
    Static Dim Scopes As String() = {SheetsService.Scope.Spreadsheets} 'If changing the scope then delete  App_Data\MyGoogleStorage\.credentials\sheets.googleapis.com-dotnet-quickstart.json
    Dim ApplicationName As String = "Google Sheets API .NET Quickstart" 'Tutorial name for .net api

    Dim location = Server.MapPath("client_secret.json") 'Designate file with sheets api key and use it to setup authentication
    Using stream = New FileStream(location, FileMode.Open, FileAccess.Read) 'Read file and setup credentials for sheets api
        Dim credPath As String = System.Web.HttpContext.Current.Server.MapPath("/App_Data/MyGoogleStorage")
        credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json")
        Dim credential As UserCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, New FileDataStore(credPath, True)).Result
    End Using
    
    Dim service = New SheetsService(New BaseClientService.Initializer() With {.HttpClientInitializer = credential, .ApplicationName = ApplicationName}) 'Create Google Sheets API service.
	Dim spreadsheetId as String = "" 'ID of spread sheet to get data from
	Dim subsheetID as String = "Sheet1" 'ID of the subsheet to get data from
	Dim range As [String] = subsheetID & "!A1:Z300" 'Data range to get, gets data from column A row 1 to column Z row 300
	
	Dim request As SpreadsheetsResource.ValuesResource.GetRequest = Service.Spreadsheets.Values.[Get](spreadsheetId, range) 'Make the request
	Dim response1 As ValueRange = request.Execute() 'Execute the request and get data
	Dim values As IList(Of IList(Of [Object])) = response1.Values
	Table1.Rows.Clear() 'Clear table with id Table1 to account for it having data after refreshing

	If values IsNot Nothing AndAlso values.Count > 0 Then 'If data is not nothing and is more than 0
		For Each row As IList In values 'For each row pulled from sheets
			If row.Count >= 1 Then 'If there is a row
				Dim tRow As New TableRow() 'Make new table row
				Table1.Rows.Add(tRow) 'Add row to table
				
				Dim tCell1 As New TableCell() 'Make new cell
				tRow.Cells.Add(tCell1) 'Add tcell1 to row
				Try
					tCell1.Text = row(0) 'Set the text of cell 1 to be the text of the column 0 (A) cell
					Catch ArgumentOutOfRangeException As Exception 'Will catch out of range if empty/blank
				End Try
				
				Dim tCell2 As New TableCell() 'Make new cell
				tRow.Cells.Add(tCell2) 'Add tcell2 to row
				Try
					tCell2.Text = row(1) 'Set the text of cell 2 to be the text of the column 1 (B) cell
				Catch ArgumentOutOfRangeException As Exception 'Will catch out of range if empty/blank
				End Try
				
				'Set the start and end ranges of the further columns you want to get.
				Dim currentColNum As Integer = 2 'Column C
				Dim lastColNum as Integer = 10 'Column J
				Do Until currentColNum > lastColNum 'Loop through the sheet data and make a table cell for each column cell
					tCell3 = New TableCell()
					tRow.Cells.Add(tCell3)
					Try
						tCell3.Text = row(currentColNum)
					Catch ArgumentOutOfRangeException As Exception 'Will catch out of range if empty/blank
					End Try
					currentColNum = currentColNum + 1
				Loop 'Repeat until desired columns have been added.
			End If
		Next
	End If
End Sub
