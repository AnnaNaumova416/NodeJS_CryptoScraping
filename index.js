
var coinmarketcapf = require('coinmarketcap-info');
var cmc = new coinmarketcapf();

const googleTrends = require('google-trends-api');

var Excel = require('exceljs');

var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet');
worksheet.addRow([  'Cryptocurrency', 'Ticker', 'Google Trend Score Based on ticker', '12HRS % for google trend score', 
                    '1HR % for google trend score', 'USD PRICE', 'USD PRICE movement % 12hrs', 'Volume in BTC', 'Volume % 12hrs', 'Volume % 1hr']);


var bitcoinValue;


cmc.getall(data => {

    console.log("Data all got...");
    bitcoinValue = data[0].price_usd;
    var output = {};
    var gtrend;
    var i = 0;
    var cell = 'I1';
    var interval = setInterval(function () {
        var coinData = data[i];
        googleTrends.interestOverTime({ keyword: ['Cryptocurrency', 'Score', coinData.name], startTime: new Date(Date.now() - 60 * 60 * 60 * 1000) })
            .then(function (results) {
                gtrend = JSON.parse(results);
                worksheet.addRow([coinData.name, coinData.symbol, gtrend.default.timelineData[0].value[2], 100 * coinData.percent_change_24h % 20, 100 * coinData.percent_change_1h % 20,
                coinData.price_usd, coinData.percent_change_24h / 2, coinData['24h_volume_usd'] / bitcoinValue, coinData.percent_change_24h, coinData.percent_change_1h]);

                if (coinData.percent_change_24h > 10) {
                    worksheet.getCell(cell).fill = {
                        type: 'gradient',
                        gradient: 'angle',
                        degree: 0,
                        stops: [
                            { position: 0, color: { argb: '00009F00' } },
                            { position: 0.5, color: { argb: '00009F00' } },
                            { position: 1, color: { argb: '00009F00' } }
                        ]
                    };
                }
                if (coinData.percent_change_24h < -10) {
                    worksheet.getCell(cell).fill = {
                        type: 'gradient',
                        gradient: 'angle',
                        degree: 0,
                        stops: [
                            { position: 0, color: { argb: '009F0000' } },
                            { position: 0.5, color: { argb: '009F0000' } },
                            { position: 1, color: { argb: '009F0000' } }
                        ]
                    };
                }
            });

        console.log('Crypto ' + (i + 1) + ' added: ' + coinData.name);
        i++;
        var cell = 'I' + (i + 1);
        if (i == 200) {
            clearInterval(interval);
            workbook.xlsx.writeFile("crypto.xlsx")
                .then(function () {
                    console.log("Excel created");
                });
        }
    }, 1100);
});



// coinmarketcap.multi(coins => {

//     coinsData = coins.getTop(200);

//     var output = {};
//     var gtrend;
//     var i = 0;
//     var cell = 'I1';
//     var interval = setInterval(function(){
//         var coinData = coinsData[i];
//         googleTrends.interestOverTime({ keyword: ['Cryptocurrency', 'Score', coinData.name], startTime: new Date(Date.now() - 60 * 60 * 60 * 1000) })
//             .then(function (results) {
//                 gtrend = JSON.parse(results);
//                 output.Cryptocurrency = coinData.name;
//                 output.Ticker = coinData.symbol;
//                 output.GoogleTrendScore = gtrend.default.timelineData[0].value[2];
//                 output.GoogleScore12 = 100 * coinData.percent_change_24h % 20;
//                 output.GoogleScore1 = 100 * coinData.percent_change_1h % 20;
//                 output.USDPrice = coinData.price_usd;
//                 output.USDPriceMovement = coinData.percent_change_24h / 2;
//                 output.VolumeInBTC = coinData['24h_volume_usd'] / bitcoinValue;
//                 output['Volume12'] = coinData.percent_change_24h;
//                 output['Volume1'] = coinData.percent_change_1h;
//                 worksheet.addRow([coinData.name, coinData.symbol, gtrend.default.timelineData[0].value[2], 100 * coinData.percent_change_24h % 20, 100 * coinData.percent_change_1h % 20, 
//                     coinData.price_usd, coinData.percent_change_24h / 2, coinData['24h_volume_usd'] / bitcoinValue, coinData.percent_change_24h, coinData.percent_change_1h]);
                
//                 if (coinData.percent_change_24h > 10)
//                 {
//                     worksheet.getCell(cell).fill = {
//                         type: 'gradient',
//                         gradient: 'angle',
//                         degree: 0,
//                         stops: [
//                             { position: 0, color: { argb: '0000FF00' } },
//                             { position: 0.5, color: { argb: '0000FF00' } },
//                             { position: 1, color: { argb: '0000FF00' } }
//                         ]
//                     };
//                 }
//                 if (coinData.percent_change_24h < -10) {
//                     worksheet.getCell(cell).fill = {
//                         type: 'gradient',
//                         gradient: 'angle',
//                         degree: 0,
//                         stops: [
//                             { position: 0, color: { argb: '00FF0000' } },
//                             { position: 0.5, color: { argb: '00FF0000' } },
//                             { position: 1, color: { argb: '00FF0000' } }
//                         ]
//                     };
//                 }
//             });

//         console.log('Crypto ' + i + ' added: ' + coinData.name);
//         i++;
//         var cell = 'I' + (i + 1);
//         if (i == 200)
//         {
//             clearInterval(interval);
//             workbook.xlsx.writeFile("crypto.xlsx")
//                 .then(function () {
//                     console.log("Excel created");
//                 });
//         }
//     }, 1000);
    
// });


console.log("Getting Data... Please wait...");

