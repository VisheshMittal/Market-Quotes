/// <reference path../../Scripts/App.js" />

(function () {
    "use strict";
    
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {


$('#get-text').click(getTextFromDocument);      
        });
    }

})();
function getTextFromDocument() {

    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        { valueFormat: "unformatted", filterType: "all" },

        function (asyncResult) {
            showStockData(asyncResult.value);
        });

}
function showStockData(symbol){
    // Yahoo YQL - http://developer.yahoo.com/yql/ 
var yql = 'select * from yahoo.finance.quotes where symbol in (\'' + symbol + '\')';
var queryURL = 'https://query.yahooapis.com/v1/public/yql?q=' + yql + '&format=json&env=http%3A%2F%2Fdatatables.org%2Falltables.env&callback=?';

$.getJSON(queryURL, function(results) {
if(results.query.count > 0)
{
var quotes = results.query.results.quote;

$('#stock-name').text(quotes.Name);
$('#prev-close').text(quotes.PreviousClose);
$('#open').text(quotes.Open);
$('#bid').text(quotes.Bid);
$('#ask').text(quotes.Ask);
$('#target-est').text(quotes.OneyrTargetPrice);
$('#days-range').text(quotes.DaysRange);
$('#volume').text(quotes.Volume);
$('#avg-volume').text(quotes.AverageDailyVolume);
$('#market-cap').text(quotes.MarketCapitalization);
$('#pe-ratio').text(quotes.PERatio);
$('#earnings').text(quotes.EarningsShare);
$('#yield').text(quotes.DividendYield);

}

});

}