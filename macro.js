(function()
{
    // Config options
    var startRowNum = 2;
    var endRowNum = 10;
    
    // Varibale setup
    var oWorksheet = Api.GetActiveSheet();
    var baseUrl = "https://rest.coinapi.io/";
    // Insert your coinapi.io key below
    var apikey = "?apikey=KEY-HERE";
    var xmlhttp = new XMLHttpRequest();
    
    for (i = startRowNum; i < (endRowNum + 1); i++) {
        var columnAValue = oWorksheet.GetRange("A" + i);
        var columnCValue = oWorksheet.GetRange("C" + i);
        var getCoinName = columnAValue.GetText();
        var endpoint = "v1/exchangerate/" + getCoinName + "/USD";
        var link = baseUrl + endpoint + apikey; // Build req URL
        
        columnCValue.SetValue(""); // Clear old data
        xmlhttp.open("GET", link, false);
        xmlhttp.send();
        
        var coinVal = JSON.parse(xmlhttp.responseText).rate; // Parse JSON
        columnCValue.SetValue(coinVal);
        console.log(coinVal);
    }
}
)();
