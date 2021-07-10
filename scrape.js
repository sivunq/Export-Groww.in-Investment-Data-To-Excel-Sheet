const axios = require('axios');
const Excel = require('exceljs');

let i,value,datas,prevRow,newRow,total,sheet,worksheet,workbook,prevRowValues,today,dd,mm,yyyy,sipMfCodes,liqMfCodes,Values,items,stocks,price,sortedStocks,buyPrice,shares,change,prevDate,count;

const loginUrl="https://groww.in/v1/api/user/v1/login";
const mfUrl="https://groww.in/v1/api/portfolio/v1/dashboard";
const stocksUrl="https://groww.in/v1/api/stocks/v1/holding";
const emailId="<Your-EmailId>";
const password="<Your-Password>";

//your investments codes
sipMfCodes = ["120505", "120503"]; 
liqMfCodes = ["118560","120513"];

//get today's date
today = new Date();
dd = String(today.getDate()).padStart(2, '0');
mm = String(today.getMonth() + 1).padStart(2, '0');
yyyy = today.getFullYear();
today = dd+"-"+mm+"-"+yyyy;

//function to calculate mutual fund values
mfCalculation =(datas,mfCodes,workbook,sheet) => {
	Values=[];
	Values.push(today);
	total=0;
	for (scheme of mfCodes){
		value= (datas[scheme][0]["current_value"])-(datas[scheme][0]["amount_invested_tax"]);
		value=Math.round((value + Number.EPSILON) * 100) / 100;
		Values.push(value);
		total+=value;
	}
	total=Math.round((total + Number.EPSILON) * 100) / 100;
	Values.push(total,0);
	cellStyling(Values,workbook,sheet);
}

//function to calculate each stock's current value
stocksCalculation = (items,workbook,sheet) => {
	stocks={};
	for (item in items) {
		company= String(items[item]["companyShortName"]);
		price= items[item]["livePriceDto"]["ltp"]
		stocks[company] = price;
	}
	sortedStocks = Object.keys(stocks).sort().reduce((acc, key) => ({...acc, [key]: stocks[key]}), {}); //sort mfs by name
	
	worksheet= workbook.getWorksheet(sheet);
	buyPrice = worksheet.getRow(2).values;
	buyPrice= buyPrice.slice(2,buyPrice.length-2); //only take number values
	shares = worksheet.getRow(3).values;
	shares= shares.slice(2,shares.length-2);
	
	i=0;
	total=0;
	Values=[];
	Values.push(today);
	for (company in sortedStocks){
		value= (sortedStocks[company]- buyPrice[i])*shares[i];
		value=Math.round((value + Number.EPSILON) * 100) / 100;
		Values.push(value);
		total+=value;
		i++;
	}
	total=Math.round((total + Number.EPSILON) * 100) / 100;
	Values.push(total,0);
	cellStyling(Values,workbook,sheet);
}

//style cell depending on its values compared to last day's data
cellStyling = (Values,workbook,sheet) => {
	worksheet= workbook.getWorksheet(sheet);
	count=worksheet.rowCount;
	prevDate=worksheet.getRow(count).values[1];
	
	prevRow=[];
	if(prevDate == today)prevRow = worksheet.getRow(count-1).values;
	else prevRow = worksheet.getRow(count).values;

	prevRow=prevRow.slice(1,prevRow.length);
	
	change= Values[Values.length-2]-prevRow[prevRow.length-2];  //DayChange
	change=Math.round((change + Number.EPSILON) * 100) / 100;
	Values[Values.length-1] = change;

	if(prevDate == today)worksheet.spliceRows(count,1,Values);
	else worksheet.addRow(Values);
	
	newRow = worksheet.lastRow;
	i=1;
	newRow.eachCell((cell, colNumber) => {
		cell.border = {top: {style:'thin'},left: {style:'thin'},bottom: {style:'thin'},right: {style:'thin'}};
		if(colNumber >1 && colNumber < Values.length-1){
			if(cell.value < prevRow[i])
				cell.fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'FCE4D6'},bgColor:{argb:'FCE4D6'}}; //lightPink
			else 
				cell.fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'A9D08E'},bgColor:{argb:'A9D08E'}}; //lightGreen
			i++;
		}else
			cell.fill = {type: 'pattern',pattern:'solid',fgColor:{argb:'70AD47'},bgColor:{argb:'70AD47'}}; //darkGreen
	});
}

async function getData() { 
	//MutualFunds
	 await axios(mfUrl,config)
		.then(response => {
			datas=response.data.mf_dashboard.investments.internal_investment.portfolio_schemes;
			mfCalculation(datas,sipMfCodes,workbook,'MutualFunds');
			mfCalculation(datas,liqMfCodes,workbook,'LiquidFunds');
		}).catch(err => console.error(err));
	console.log("Mutual Funds Done.... ");
	//Stocks
	await axios(stocksUrl,config)
		.then(response => {
			items = response.data.holdingList;
			stocksCalculation(items,workbook,'CurrentStocks');
		}).catch(err => console.error(err));
	console.log("Stocks Done.... ");

	workbook.xlsx.writeFile('FinTrack.xlsx');
	}

//login into groww.in profile
axios.post(loginUrl, {"userId": emailId, "password": password})
  .then(response => {
	const config = { headers: { Authorization: `Bearer ${response.data.authToken}`}};
	workbook = new Excel.Workbook();
	workbook.xlsx.readFile('FinTrack.xlsx')
	.then(() => {
		getData();
	}).catch(err => console.error(err));
  }).catch(err => console.error(err));
