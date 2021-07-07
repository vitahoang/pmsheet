const { visitNode } = require("typescript");
var moment = require('moment'); // require
moment().format(); 


var date = moment();
console.log(date);
var monday = date.isoWeekday(1).add(1, 'week');
console.log(monday);
var newday = monday.format('dddd DD MMM');
console.log(newday);