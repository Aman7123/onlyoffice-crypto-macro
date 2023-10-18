(async function() {
    // Config options
    var startRowNum = 2;
    var endRowNum = 9;

    // Variable setup
    var oWorksheet = Api.GetActiveSheet();
    var baseUrl = "https://rest.coinapi.io/";
    var apikey = "?apikey=APIKEY";

    async function fetchRate(i) {
        // Grab A startRowNum from sheet
        var columnAValue = oWorksheet.GetRange("A" + i);
        var columnCValue = oWorksheet.GetRange("C" + i);
        var getCoinName = columnAValue.GetText();
        var endpoint = "v1/exchangerate/" + getCoinName + "/USD";
        var link = baseUrl + endpoint + apikey; // Build req URL
        
        try {
            let response = await fetch(link);

            if(response.ok) {
                let data = await response.json();
                var coinVal = data.rate;
                columnCValue.SetValue(""); // Clear old data
                columnCValue.SetValue(coinVal);
                console.log("Coin: " + getCoinName + ", HTTP Status: " + response.status + ", USD: " + coinVal);
            } else {
                console.error("HTTP Error: " + response.status);
            }
        } catch(error) {
            console.error(error);
        }
    }

    for (var i = startRowNum; i < (endRowNum + 1); i++) {
        await fetchRate(i);  // Wait for each request to complete before proceeding to the next
    }
})();