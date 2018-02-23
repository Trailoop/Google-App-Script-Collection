function date_by_subtracting_days(date, days) {
    return new Date(
        date.getFullYear(), 
        date.getMonth(), 
        date.getDate() - days,
        date.getHours(),
        date.getMinutes(),
        date.getSeconds(),
        date.getMilliseconds()
    );
}


function getEmail() {

var today = new Date();
var beforeWeek = date_by_subtracting_days(today, 365);
Logger.log(beforeWeek);



  
}
