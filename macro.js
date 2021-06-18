(function()
{
    // Config options
    var debug = true;
    var startRowNum = 2;
    var endRowNum = 9;
    
    // Varibale setup
    var oWorksheet = Api.GetActiveSheet();
    var baseUrl = "https://rest.coinapi.io/";
    var apikey = "?apikey=KEY-HERE";
    var xmlhttp = new XMLHttpRequest();
    
    for (i = startRowNum; i < (endRowNum + 1); i++) {
        // Grab A startRowNum from sheet
        var columnAValue = oWorksheet.GetRange("A" + i);
        var columnCValue = oWorksheet.GetRange("C" + i);
        var getCoinName = columnAValue.GetText();
        var endpoint = "v1/exchangerate/" + getCoinName + "/USD";
        var link = baseUrl + endpoint + apikey; // Build req URL
        
        xmlhttp.open("GET", link, false);
        xmlhttp.send();
        
        var coinVal;
        if(xmlhttp.status == 200) {
            columnCValue.SetValue(""); // Clear old data
            coinVal = JSON.parse(xmlhttp.responseText).rate; // Parse JSON
            columnCValue.SetValue(coinVal);
        } else {
            console.log(xmlhttp.status);
        }
        if(debug) {
            console.log("Coin: " + getCoinName + ", HTTP Status: " + xmlhttp.status + ", USD: " + coinVal);
            //console.log("Rate Limiting (used/total): " + xmlhttp.getResponseHeader("X-RateLimit-Used") + "/" + xmlhttp.getResponseHeader("X-RateLimit-Limit"));
        }
    }
}
)();
