function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [
    {
      name : "Fetch Prices",
      functionName : "main"
    }];  
    sheet.addMenu("Scraper", entries);
  };
  
  var list_states = [['NSW', 'NSW1'], ['QLD', 'QLD1'], ['SA', 'SA1'], ['TAS', 'TAS1'], ['VIC', 'VIC1']];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  function main(){
    var response1 = UrlFetchApp.fetch('https://www.asxenergy.com.au/'); 
    fetchPrices(response1);
    FetchFutureOptsTrades(response1);
    sendMail();
    fetchFuturesData();
    // AEMO restested 01122021
    fetchAemo(); // AEMO Blocked Google crawlers. tested another scraper = failed. tested chrome browser = success
  }
  
  function fetchHtml() { //capture screenshot then save to Drive
    const folder_id = "0B5dc9j91k5sEb1FmNnRtMWtrZ2c";
    var folder = DriveApp.getFolderById(folder_id);
    var response = UrlFetchApp.fetch("https://www.asxenergy.com.au/");
    var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
    Logger.log(formattedDate);
    var contents=  response.getContentText();
    folder.createFile(formattedDate+".html", contents, MimeType.HTML);
  }
  
  function fetchFuturesData() {
    const skipCol = 9; // from 5 to 6 02042020 // from 6 to 7 01102020 // from 7 to 8 06042021 // from 8 to 9 12102021
    var list_states = [['NSW', 'N'], ['VIC', 'V'], ['QLD', 'Q'], ['SA', 'S']]; // BStr and PkStr fetch from /mobile/futures?[state][data]
    const list_data = [['BStr', 'H'], ['PkStr', 'D']];
    
    var settle_regex = /<td[^>]*?>([^(?:Settle)]+)<\/td>\s*<\/tr>/ig;
    var all_data = []; // dump all_data with proper encapsulation
    for (var state in list_states){
      var state_array = [];
      for (var data in list_data){
        var data_array = [];
        var response = UrlFetchApp.fetch('https://www.asxenergy.com.au/mobile/futures?'+list_states[state][1]+list_data[data][1]);
        var content = response.getContentText();
        while(row = settle_regex.exec( content)) {         
          if(row && row[1]) {
            data_array.push(row[1]);
          }
        }
        state_array.push(data_array);
      }
      all_data.push(state_array);
    }
    
    var maxLength = all_data[0][0].length; // number of columns
    
    var response = UrlFetchApp.fetch('https://www.asxenergy.com.au/futures_au/dataset'); // CapStr doesn't have its own page so we're fetching from /dataset
    var content = response.getContentText();
    var capstr_regex = /CapStr(?:\s*<[^>]+>)+<\/td>([\w\W]*?)<\/table>/ig;
    settle_regex = /<td[^>]*?>([^(?:Settle)]+)<\/td>\s*<\/tr>/ig;
    var capStr = [];
    while(row2 = capstr_regex.exec( content)) {
      var data_array = [];
      if(row2 && row2[1]) {
        var datacontent = row2[1];
        while(row3 = settle_regex.exec(datacontent)) {         
          if(row3 && row3[1]) {
            data_array.push(row3[1]);
          }
        }
      }
      if(data_array.length < maxLength) {
        for (var diff = data_array.length; diff < maxLength; diff++) { data_array.push(''); }
      }
      capStr.push(data_array);
      all_data[capStr.indexOf(data_array)].push(data_array); // insert CapStr with BStr and PkStr data
    }
  
    var all_sum = [];
    for (var y = 0; y < all_data.length; y++) {
      var baseCap = [];
      for (var z = 0; z < all_data[y][0].length; z++) {
        baseCap.push(Number(all_data[y][0][z]) + Number(all_data[y][2][z]));
      }
      all_sum.push(baseCap); // insert BaseCap with CapStr, BStr, PkStr
    }
    for (var y = 0; y < all_data.length; y++) {
      var peakCap = [];
      for (var z = 0; z < all_data[y][1].length; z++) {
        peakCap.push(Number(all_data[y][1][z]) + Number(all_data[y][2][z]));
      }
      all_sum.push(peakCap); // insert PeakCap with CapStr, BStr, PkStr
    }
   Logger.log(all_sum);
    //var ss = SpreadsheetApp.getActiveSpreadsheet();
    var FUTURES_SHEET  = ss.getSheetByName('FUTURES')
    var date= new Date();
    var row = FUTURES_SHEET.getLastRow()+1;
  
    var rowEntry = []; // dump all_data in a single row
    
    for (var s in all_data) { // each state (NSW, QLD, VIC, SA)
      for (var d in all_data[s]) { // each data (BStr, PStr, CapStr, BaseCap, PeakCap)
        for (var skip = 0; skip < skipCol; skip++) { all_data[s][d].unshift(''); } // skip columns
        for (var y in all_data[s][d]) { rowEntry.push(all_data[s][d][y]); } // each column (FY,CY)
      }
    }
    /*for each (var sum in all_sum) {
      for (var skip = 0; skip < skipCol; skip++) { sum.unshift(''); }
      for each (var s1 in sum) { rowEntry.push(s1); }
    }*/
    rowEntry.unshift(date);
    
    FUTURES_SHEET.getRange(row, 1, 1, rowEntry.length).setValues([rowEntry]);
  }
  
  function chunk(arr, n) {
    return arr.slice(0,(arr.length+n-1)/n|0).
    map(function(c,i) { return arr.slice(n*i,n*i+n); });
  }
  
  function sumBaseCap( b, c ){
    var b2=b.map(Number);
    Logger.log(b2);
    var c2=c.map(Number);
    Logger.log(c2);
      return b2.map( function (num, idx) {
        return num + c2[idx];
      });
  }
  
  function fetchWEPIdata() {  
    var WEPI_DATA = [];
    var response = UrlFetchApp.fetch('https://www.asxenergy.com.au');
    var content = response.getContentText();
    const data_block_reg = /class\=\"instrument\">2019<\/td>([\w\W]*?)<\/tr>/i;  
    var data_block = data_block_reg.exec(content);
    const values_reg = /numeric">\s*([^<]*?)</ig;  
    if(data_block && data_block[1]) {
        while(value = values_reg.exec( data_block[1])) { 
          Logger.log(value[1] );
          WEPI_DATA.push(value[1]);
        }
    }
    return WEPI_DATA;
  }
  
  
  
  function baseFuturesData() {
    var BASE_DATA = [];
    var response = UrlFetchApp.fetch('https://www.asxenergy.com.au');
    var content = response.getContentText();
    const row_block_reg = /class\=\"instrument\">\d+<\/td>([\w\W]*?)<\/tr>/ig;  
    var row_block;
    
    while ( row_block = row_block_reg.exec(content) ) {
      if(row_block && row_block[1]) {
       var values_reg = /numeric">\s*([^<]*?)</ig;  
        var rowdata= [];
        while(value = values_reg.exec( row_block[1])) { 
          Logger.log(value[1] );
          rowdata.push(value[1]);
        }
      }
      BASE_DATA.push([rowdata[0], rowdata[2],rowdata[3],rowdata[1]])
      
    }
    Logger.log(BASE_DATA)
    return BASE_DATA;
  }
  
  
  function fetchPrices(response1) {
    //var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet =  ss.getSheetByName('INPUT');
    var sheet1 = ss.getSheetByName('PRICES');  
    var lrow = sheet1.getLastRow();
    var lrec = sheet1.getRange(lrow, 1, 1);  
    //var response = UrlFetchApp.fetch('http://d-cyphatrade.com.au/mobile/closing-prices');
  var response = UrlFetchApp.fetch('https://www.asxenergy.com.au/mobile/closing-prices');
    var document = Xml.parse(response.getContentText(), true);  
    var data = [];  
    var prices = evaluate(document.getElement(), 'body/div/table/1');
    
    var date = '';
    try { 
      closeDate = evaluate(document.getElement(), 'body/h4')[0];  
      date = closeDate.getText();  
    }
    catch(e) {
    
    };
    
    
    var dateIdx = date.indexOf('COB ');  
    date = date.substring(dateIdx + 8);  
    var lrecdate = lrec.getValue();
    Logger.log(lrecdate);
    
    var formattedDate = Utilities.formatDate(lrecdate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "d MMM YYYY");  
    
    Logger.log('Check 1:'+date);
    Logger.log('Check 2:'+formattedDate);
    
    //if (date === formattedDate) {
    //  return;
    //}  
   
    //data.push(date);  
    data.push(Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy"));
    
    var markets = ['NSW','QLD','SA','VIC'];  
    var wepi = evaluate(prices, 'thead/tr/td');  
    //var wepi = fetchWEPIdata();
    
    for (var c = 0; c<markets.length;c++){ //add each state
      data.push(wepi[c+1].getText());
    } 
  
    var baseFutures =  evaluate(document.getElement(), 'body/div/table/2/tr');
    
    var baseFuturesSkipYear = [2013,2014,2015,2016,2017,2018,2019,2020,2021];
    
    for ( var i = 0; i<baseFuturesSkipYear.length;i++ ){ //skip each year
      for ( var c = 0; c<markets.length;c++ ){ //skip each state
        data.push('');
      }
    }
  
    var baseFutures = baseFuturesData();
    for (var i = 0; i < baseFutures.length; i++) {
  
      for (var c = 0; c<markets.length;c++){ //add each state
        data.push( baseFutures[i][c] );
      }
      
      Logger.log(data)
      
    }
    
    
    var response1 = UrlFetchApp.fetch('https://www.asxenergy.com.au/');  
    var forecast_block_reg = /<div\s*class="dataset"\s*id="forecast">([\w\W]*?)<\/div>/i
    var forecast_block = forecast_block_reg.exec(response1.getContentText());
    if(forecast_block && forecast_block[1])
    {
      var data_reg= /<tr><td>(?:Sydney|Melbourne|Brisbane|Adelaide|Hobart)<\/td><td\s*class="dataset-numeric">([^<]*?)</ig
      while(matches = data_reg.exec( forecast_block[1] )) {   
        matches[1] = matches[1].replace('&deg;C','');
        
        data.push(matches[1]);     
      }
    }
    //FetchFutureOptsTrades(response1);
    var openInterst = evaluate(document.getElement(), 'body/div/table/3/tr');  
    for (var i = 0; i < openInterst.length; i++) {
      var p = openInterst[i].getElements('td');
      
      for (var c = 0; c<markets.length;c++){ //add each state
        data.push(p[c+1].getText());
      }
    }
  
    var indicesSkipYear = [2014,2015,2016,2017,2018,2019,2020];
  
    try { 
      var indices = evaluate(document.getElement(), 'body/div/table/4/tr');
      
      for (var c = 0; c<indicesSkipYear.length;c++){ //skip each state
        data.push(''); // skip eastern
        data.push(''); // skip national
      }
      
      for (var i = 0; i < indices.length; i++) {
        var p = indices[i].getElements('td');
        data.push(p[2].getText());
        data.push(p[4].getText());
      }
    }
    catch(e) {
    
    };
    sheet1.getRange(lrow + 1, 1, 1, data.length).setValues([data]);
    var email = sheet.getRange(1,2,1).getValue();  
  }
  
  function evaluate(element, path) {
    var paths = path.split('/');  
    for (var i=0; i < paths.length; i++) {
      if(!element) {
        break;
      }    
      if (!isNaN(paths[i])) {
        element = element[paths[i]];
      }
      else {
        if (element instanceof Array) {
          element = element[0];
        }      
        if (element) {
          element = element.getElements(paths[i]);
        }
      }
    }  
    return element;
  }
  
  function sendMail() {
    
    generateCharts();
    
    var emailAddress = 'energyops@fortiserve.com.au';    
    var bcc = 'commercial@makeitcheaper.com.au'
    var bcc = 'tenders@leadingedgeenergy.com.au';
    // QLD Feature -  https://docs.google.com/spreadsheet/oimg?key=0AobmIMcMSTPAdEU2QkU1Nnd0YndrT19WOFkwVGZVc1E&oid=16&zx=cv35qd1n6o3e  
    // NSW WEF - https://docs.google.com/spreadsheet/oimg?key=0AobmIMcMSTPAdEU2QkU1Nnd0YndrT19WOFkwVGZVc1E&oid=4&zx=pxolna4vywk  
    // WEPI - https://docs.google.com/spreadsheet/oimg?key=0AobmIMcMSTPAdEU2QkU1Nnd0YndrT19WOFkwVGZVc1E&oid=2&zx=cz0a3m816m21  
    // SA Future - https://docs.google.com/spreadsheet/oimg?key=0AobmIMcMSTPAdEU2QkU1Nnd0YndrT19WOFkwVGZVc1E&oid=17&zx=z0nsnxekn5a4  
    // VIC Feature - https://docs.google.com/spreadsheet/oimg?key=0AobmIMcMSTPAdEU2QkU1Nnd0YndrT19WOFkwVGZVc1E&oid=19&zx=evfj2yy8od31
    // NSW Futures -  '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=733756293&format=image" />'
    // QLD Futures - '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=1979846902&format=image" />'
    // SA Futures - '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=277166968&format=image" />'
    // VIC Futures '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=131454501&format=image" />'
    // '<img width="643" src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=741351003&format=image" />'
    // WEPI - '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=1002594169&format=image" />'
    
    var message = "<HTML><BODY>"   
    //+ '<img src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=733756293&format=image" />'
    //+ '<img src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=1979846902&format=image" />'
    //+ '<img src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=277166968&format=image" />'
    //+ '<img src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=131454501&format=image" />'
     + '<img width="643" src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=741351003&format=image" />'
    //+ '<img src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=1002594169&format=image" />'
    + "</BODY></HTML>";
    
    
    // SECOND EMAIL FOR PRICE MOVEMENT
    Logger.log('sending email');
  //disable email
    //MailApp.sendEmail(emailAddress, "Daily ASX Energy Report - Price Movement",'', {htmlBody: message,cc:bcc});  
    
    
    
    var emailAddress = 'brokers@energin.co';
    
    
     var message = "<HTML><BODY>"   
    +  '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=733756293&format=image" />'
    + '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=1979846902&format=image" />'
    + '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=277166968&format=image" />'
    + '<img src="https://docs.google.com/spreadsheets/d/e/2PACX-1vS8lbyE6XJQ9vhv2X3nFBOY2Te770wfvvNgL_Ug865HEI3dVsuwEt3q2lvbCpyGGAGx0koxRMtODgfx/pubchart?oid=131454501&format=image" />'
    // + '<img width="643" src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=741351003&format=image" />'
    //+ '<img src="https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubchart?oid=1002594169&format=image" />'
    + "</BODY></HTML>";
    
    
    
    Logger.log('sending email');
  //disable email
    //MailApp.sendEmail(emailAddress, "Daily ASX Energy Report",'', {htmlBody: message,cc:bcc});
    
    
    
    var emailAddress = 'ibsales@leadingedgeenergy.com.au';
    
    
     var message = "<HTML><BODY>"
    + 'Click the link below to check our Price Movement Report: <br />'
    + 'https://docs.google.com/spreadsheets/d/1Q-oO9ng_bSDUjyIpkDGtQqUN8RZoEq0WkNT_01kPAPE/pubhtml?gid=402878118&single=true'
    + "</BODY></HTML>";
    
    
    
    Logger.log('sending email');
  //disable email
    //MailApp.sendEmail(emailAddress, "Price Movement Report",'', {htmlBody: message,cc:bcc});
  }
  
  function updateChartData(){
    //var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[1];
    var charts = sheet.getCharts();
    for (var i in charts) {
      var chart = charts[i];
      Logger.log(i);
    }
  }
  
  function fetchAemo(){
    var sheet = ss.getSheetByName('AEMO Spot Prices')
    var sheet_lrow = sheet.getLastRow()
    var sheet_lrec = sheet.getRange(sheet_lrow, 1);    
    var sheet_lrecdate = sheet_lrec.getValue(); 
    
    var sheet_formattedDate = '';
    try {
      sheet_formattedDate = Utilities.formatDate(sheet4_lrecdate, "GMT", "dd/MM/yyyy");
    }
    catch(e) {
      sheet_formattedDate = sheet_lrecdate; 
    }
    var MILLIS_YESTERDAY = 1000 * 60 * 60 * 24;
    var now = new Date();
    var fetchdate = sheet_formattedDate = Utilities.formatDate(new Date(now.getTime() - MILLIS_YESTERDAY), "GMT", "yyyy/MM/dd");
    var fetchdate_match = sheet_formattedDate = Utilities.formatDate(new Date(now.getTime() - MILLIS_YESTERDAY), "GMT", "yyyy/MM/dd");
    var fetchyear = sheet_formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy");
    var fetchmonth = sheet_formattedDate = Utilities.formatDate(new Date(), "GMT", "MM");
    
    var fetchurl = 'https://www.aemo.com.au/aemo/data/nem/averageprices/WEB_AVERAGE_PRICE_DAY_'+fetchyear+fetchmonth+'.csv';
    Logger.log("fetchurl: "+fetchurl);
    
    var response2 = UrlFetchApp.fetch(fetchurl, {'muteHttpExceptions':true});
    try{
      var csv = Utilities.parseCsv(response2);
      Logger.log(csv);
      Logger.log("csv (0) "+csv[0]);
      Logger.log("csv (1)(0) "+csv[1][0]);
      
      if(sheet_lrecdate != fetchdate_match){
        var returnArr = [fetchdate_match];
  
        var nowDate = Utilities.formatDate(now,"GMT","yyyy/MM/dd");
        for (var state in list_states){
          try{
            for (var csvState in csv){
              Logger.log("target: "+fetchdate);
              Logger.log(csv[csvState][0]);
              Logger.log("csv: "+csv[csvState][0].substring(0,10));
              if(list_states[state][1] == csv[csvState][1] && csv[csvState][0].indexOf(fetchdate) > -1){
                Logger.log(csv[csvState][2]);
                Logger.log(csv[csvState][3]);
                returnArr.push(csv[csvState][2]);
                returnArr.push(csv[csvState][3]);
                break;
              }
            }
          }
          catch(e){
            Logger.log("fetch CSV length does not match. Please update parsing.")
          }
        }
        Logger.log([returnArr]);
        sheet.getRange(sheet_lrow+1, 1, 1, returnArr.length).setValues([returnArr]);
      }
    }
    catch(e){
      Logger.log("Scraping AEMO website failed. Please update fetch URL.");
    }
  }
  
  function FetchFutureOptsTrades(response) {     
    var response = UrlFetchApp.fetch('https://www.asxenergy.com.au/');    
    //var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet2 = ss.getSheetByName('Future/Opts Trades')
    var cdate =  new Date();
    var cdate_formattedDate = Utilities.formatDate(cdate, "GMT", "dd/MM/yyyy");  
    if(cdate_formattedDate) {     
      var messages_block_regex = /<div\s*id="home-messages">([\w\W]*?)<\/div>/i;  
      var messages_block = messages_block_regex.exec(response.getContentText());
      if(messages_block && messages_block[1]){    
        var message_reg = /<li>\s*<span[^>]*?>([^<]*?)<\/span>(?:\s*<[^>]*?>\s*)+([^<]*?)\s*<\/a>\s*([^<]*?)</ig;      
        while(message = message_reg.exec( messages_block[1])) {        
          if(message[1] && message[3]){
            var sheet2_lrow = sheet2.getLastRow();
            var sheet2_date = sheet2.getRange(sheet2_lrow, 2).getValue();     
            var sheet2_time = sheet2.getRange(sheet2_lrow, 3).getValue();  
            Logger.log('Checking the FetchFutureOptsTrades date function');
            Logger.log('sheet2_date:'+sheet2_date);
            Logger.log('sheet2_time:'+sheet2_time);        
            var sheet2_formattedDate = '';
            try {
              sheet2_formattedDate = Utilities.formatDate(sheet2_date, "GMT", "dd/MM/yyyy");
            }
            catch(e) {
              sheet2_formattedDate = sheet2_date;
            }
            if(sheet2_formattedDate!=cdate_formattedDate && sheet2_time!=message[1]){
              Logger.log('Going to update');
              sheet2.getRange(sheet2_lrow+1,2).setValue(cdate_formattedDate);
              sheet2.getRange(sheet2_lrow+1,3).setValue(message[1]);
              sheet2.getRange(sheet2_lrow+1,4).setValue(message[2]+' '+message[3]);
            }         
          }        
        }
      }
    }    
  }
  
  
  
  
  function dynamicBuildChart() {
    var NSW_FEATURES = ['A','F','J','N','R','AW'];  
    var data_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRICES');  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Charts');  
    var lastrow = data_sheet.getLastRow();  
    //var ss = SpreadsheetApp.getActiveSpreadsheet();  
    var values = [];
    for(var c in NSW_FEATURES) {    
      var range_notation = NSW_FEATURES[c]+"1:"+NSW_FEATURES[c]+lastrow;   
      var range = data_sheet.getRange(range_notation);
      values.push(range.getValues())
    }  
    var dataTable = Charts.newDataTable()
    var values = {};
    for(var c in NSW_FEATURES) {    
      var range_notation = NSW_FEATURES[c]+"1:"+NSW_FEATURES[c]+lastrow;
      Logger.log(range_notation);
      if(NSW_FEATURES[c]=='A') {
        dataTable.addColumn(Charts.ColumnType.DATE, data_sheet.getRange(NSW_FEATURES[c]+"1").getValue())
      }
      else{    
        dataTable.addColumn(Charts.ColumnType.NUMBER, data_sheet.getRange(NSW_FEATURES[c]+"1").getValue())
       }
      values[NSW_FEATURES[c]] = data_sheet.getRange(range_notation).getValues();
    }
  
    for(var i=1;i<lastrow;i++) {     
      var rows=[];
      for(var c in NSW_FEATURES) {          
        rows.push(values[NSW_FEATURES[c]][i][0])
      }   
      dataTable.addRow(rows)
    }
    var textStyleBuilder = Charts.newTextStyle().setColor('#0000FF').setFontSize(8)
    var style = textStyleBuilder.build();  
    var chartImage = Charts.newLineChart()  
    .setLegendTextStyle(style)
    .setTitle('NSW FEATURES')
    .setDataTable(dataTable)  
    .setDimensions(1500, 500)
    .build();  
    sheet.insertImage(chartImage.getAs('image/png'), 1, 1)
    /*
    var mB = "<h2> Testing Mail </h2>";
    MailApp.sendEmail({
      to: "johnpeterdinesh@gmail.com",
      subject: "Chart",
      htmlBody: mB,
      inlineImages:{
        chartImg: chartImage.getAs('image/png')
      }
    });  
    */
  }
  
  
  function generateCharts() {
    var NSW_FEATURES_COLS = ['A','F','J','N','R','V'];  
    var QLD_FEATURES_COLS = ['A','G','K','O','S','W'];  
    var SA_FEATURES_COLS = ['A','H','L','P','T','X'];  
    var VIC_FEATURES_COLS = ['A','I','M','Q','U','Y']; 
    
    var WEPI_COLS = ['A','B','C','D','E'];  
    
    var NSW_WEPI_COLS = ['A','B','F','J','N','R','V']; 
    var QLD_WEPI_COLS = ['A','C','G','K','O','S','W'];   
    var SA_WEPI_COLS = ['A','D','H','L','P','T','X'];  
    var VIC_WEPI_COLS = ['A','E','I','M','Q','U','Y'];  
  
    dynamicBuildEmbedChart('NSW Wholesale Electricity Futures',NSW_FEATURES_COLS,'NSW Futures');
    
    dynamicBuildEmbedChart('QLD Wholesale Electricity Futures',QLD_FEATURES_COLS,'QLD Futures');
    dynamicBuildEmbedChart('SA Wholesale Electricity Futures',SA_FEATURES_COLS,'SA Futures');
    dynamicBuildEmbedChart('VIC Wholesale Electricity Futures',VIC_FEATURES_COLS,'VIC Futures');
    dynamicBuildEmbedChart('WEPI-NEM',WEPI_COLS,'WEPI');
    dynamicBuildEmbedChart('NSW - WEPI',NSW_WEPI_COLS,'NSW - WEPI');
    dynamicBuildEmbedChart('VIC - WEPI',QLD_WEPI_COLS,'VIC - WEPI');
    dynamicBuildEmbedChart('SA - WEPI',SA_WEPI_COLS,'SA - WEPI');   
    dynamicBuildEmbedChart('QLD - WEPI',VIC_WEPI_COLS,'QLD - WEPI');
   
  }
  
  function dynamicBuildEmbedChart(title,cols,sheet) {
    
    
    var data_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRICES');  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);  
    var charts = sheet.getCharts();
    
    var chart  = data_sheet.newChart().setChartType(Charts.ChartType.LINE);  
    Logger.log(title);
    if(charts.length>0) {
      //sheet.removeChart(charts[0])
      chart = charts[0].modify();
      Logger.log('Chart found');
      Logger.log(chart.getRanges().length);
      var ranges = chart.getRanges();
      for(var r=0;r<ranges.length;r++) {
        chart.removeRange(ranges[r])
      }
    }
    Logger.log(chart.getRanges().length);
    var lastrow = data_sheet.getLastRow();    
    
    for(var c in cols) {    
      var range_notation = cols[c]+"1:"+cols[c]+(lastrow);
      Logger.log(range_notation);
      chart.addRange(data_sheet.getRange(range_notation))
    }
    chart.setOption('title', title);
    chart.setPosition(1, 1, 0, 0);  
    chart.setOption('width', 1200)
    chart.setOption('height', 640)
    chart.build();
    if(charts.length==0) {
      sheet.insertChart(chart.build());   
    }
  }
  
  function inlineImage() {
     var googleLogoUrl = "http://www.google.com/intl/en_com/images/srpr/logo3w.png";
     var youtubeLogoUrl =
           "https://developers.google.com/youtube/images/YouTube_logo_standard_white.png";
     var googleLogoBlob = UrlFetchApp
                            .fetch(googleLogoUrl)
                            .getBlob()
                            .setName("googleLogoBlob");
     var youtubeLogoBlob = UrlFetchApp
                             .fetch(youtubeLogoUrl)
                             .getBlob()
                             .setName("youtubeLogoBlob");
     //disable email
    /*MailApp.sendEmail({
       to: "em2data2@gmail.com",
       subject: "Logos",
       htmlBody: "inline Google Logo<img src='cid:googleLogo'> images! <br>" +
                 "inline YouTube Logo <img src='cid:youtubeLogo'>",
       inlineImages:
         {
           googleLogo: googleLogoBlob,
           youtubeLogo: youtubeLogoBlob
         }
     });*/
   }
  
   function test10(){
     var testdate = ss.getSheetByName("AEMO Spot Prices").getCurrentCell().getValue();
     var MILLIS_YESTERDAY = 1000 * 60 * 60 * 24;
     var now = new Date();
     var fetchdate = sheet_formattedDate = Utilities.formatDate(new Date(testdate.getTime() - MILLIS_YESTERDAY), "GMT", "yyyy/MM/dd");
     var fetchdate_match = sheet_formattedDate = Utilities.formatDate(new Date(testdate.getTime() - MILLIS_YESTERDAY), "GMT", "yyyy/MM/dd");
     Logger.log(fetchdate);
     Logger.log(fetchdate_match);
   }
  
   function fetchPricesWithAve(response1) {
      //var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet =  ss.getSheetByName('INPUT');
      var sheet1 = ss.getSheetByName('Copy of PRICES');  
      var lrow = sheet1.getLastRow();
      var lrec = sheet1.getRange(lrow, 1, 1);  
      //var response = UrlFetchApp.fetch('http://d-cyphatrade.com.au/mobile/closing-prices');
    var response = UrlFetchApp.fetch('https://www.asxenergy.com.au/mobile/closing-prices');
      var document = Xml.parse(response.getContentText(), true);  
      var data = [];  
      var prices = evaluate(document.getElement(), 'body/div/table/1');
      
      var date = '';
      try { 
        closeDate = evaluate(document.getElement(), 'body/h4')[0];  
        date = closeDate.getText();  
      }
      catch(e) {
      
      };
      
      
      var dateIdx = date.indexOf('COB ');  
      date = date.substring(dateIdx + 8);
      var lrecdate = lrec.getValue();
      Logger.log(lrecdate);
      
      var formattedDate = Utilities.formatDate(lrecdate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "d MMM YYYY");  
      
      Logger.log('Check 1:'+date);
      Logger.log('Check 2:'+formattedDate);
      
      if (date === formattedDate) {
       return;
      }  
     
      //data.push(date);  
      data.push(Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy"));
      
      var markets = ['NSW','QLD','SA','VIC'];  
      var wepi = evaluate(prices, 'thead/tr/td');  
      //var wepi = fetchWEPIdata();
      
      for (var c = 0; c<markets.length;c++){ //add each state
        data.push(wepi[c+1].getText());
      } 
    
      var baseFutures =  evaluate(document.getElement(), 'body/div/table/2/tr');
      
      var baseFuturesSkipYear = [2013,2014,2015,2016,2017,2018,2019,2020,2021];
      
      for ( var i = 0; i<baseFuturesSkipYear.length;i++ ){ //skip each year
        for ( var c = 0; c<markets.length;c++ ){ //skip each state
          data.push('');
        }
      }
  
      var baseFutures = baseFuturesData();
      var dataAveBaseFutures = [];
      var sumNSW = 0, sumQLD = 0, sumSA = 0, sumVIC = 0;
  
      for (var i = 0; i < baseFutures.length; i++) {
    
        for (var c = 0; c<markets.length;c++){ //add each state
          data.push( baseFutures[i][c] );
  
          var parsedValue = parseInt(baseFutures[i][c]);
          if(c === 0){ sumNSW += parsedValue; }
          if(c === 1){ sumQLD += parsedValue; }
          if(c === 2){ sumSA += parsedValue; }
          if(c === 3){ sumVIC += parsedValue; }
        }
        // Logger.log(data);
      }
  
      //Get the total Ave of basefuture values in each regions
      var aveNSW = sumNSW / markets.length;
      var aveQLD = sumQLD / markets.length;
      var aveSA = sumSA / markets.length;
      var aveVIC = sumVIC / markets.length;
      dataAveBaseFutures.push(aveNSW, aveQLD, aveSA, aveVIC); //push averaget to all array
      Logger.log(dataAveBaseFutures); //log all average
      sheet1.getRange(lrow + 1, 99, 1, dataAveBaseFutures.length).setValues([dataAveBaseFutures]).setNumberFormat("0.00"); //set ave value from CU to CX
      
      var response1 = UrlFetchApp.fetch('https://www.asxenergy.com.au/');  
      var forecast_block_reg = /<div\s*class="dataset"\s*id="forecast">([\w\W]*?)<\/div>/i
      var forecast_block = forecast_block_reg.exec(response1.getContentText());
      if(forecast_block && forecast_block[1])
      {
        var data_reg= /<tr><td>(?:Sydney|Melbourne|Brisbane|Adelaide|Hobart)<\/td><td\s*class="dataset-numeric">([^<]*?)</ig
        while(matches = data_reg.exec( forecast_block[1] )) {   
          matches[1] = matches[1].replace('&deg;C','');
          
          data.push(matches[1]);     
        }
      }
      //FetchFutureOptsTrades(response1);
      var openInterst = evaluate(document.getElement(), 'body/div/table/3/tr');  
      for (var i = 0; i < openInterst.length; i++) {
        var p = openInterst[i].getElements('td');
        
        for (var c = 0; c<markets.length;c++){ //add each state
          data.push(p[c+1].getText());
        }
      }
    
      var indicesSkipYear = [2014,2015,2016,2017,2018,2019,2020];
    
      try { 
        var indices = evaluate(document.getElement(), 'body/div/table/4/tr');
        
        for (var c = 0; c<indicesSkipYear.length;c++){ //skip each state
          data.push(''); // skip eastern
          data.push(''); // skip national
        }
        
        for (var i = 0; i < indices.length; i++) {
          var p = indices[i].getElements('td');
          data.push(p[2].getText());
          data.push(p[4].getText());
        }
      }
      catch(e) {
      
      };
      sheet1.getRange(lrow + 1, 1, 1, data.length).setValues([data]);
      var email = sheet.getRange(1,2,1).getValue();  
    }