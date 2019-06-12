const WCU = require('./libs/WCU');

WCU.LoadTable('tables/test.xlsx');

WCU.ForSheet("Sheet1", function(row, i, sheet) {
    console.log(i, row);
});
