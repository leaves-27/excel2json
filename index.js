const XLSX = require('xlsx'); 
const fs = require('fs');

const url = '/Users/leaves/Desktop/2019-8-1小程序/shops.xls';
const workbook = XLSX.readFile(url);
// 获取 Excel 中所有表名
const sheetNames = workbook.SheetNames;
const getJsonOfsheet = (worksheet)=>{
	const headers = {};
	const data = [];
	const keys = Object.keys(worksheet);
	keys.filter(k => k[0] !== '!').forEach(k => {
	    let col = k.substring(0, 1);
	    let row = parseInt(k.substring(1));
	    let value = worksheet[k].v;
	    if (row === 1) {
	        headers[col] = value;
	        return;
	    }
	    if (!data[row]) {
	        data[row] = {};
	    }
	    data[row][headers[col]] = value;
	});
	// 剔除为Null的值
	const newData = [];
	data.forEach((item)=>{
		if(item) {
			newData.push({
				id: item['国际编码'],
				name: item['商品名称'],
				price: item['活动价'],
				other: item['活动备注'],
				time: item['活动期限']
			})
		}
	});
	return JSON.stringify(newData);
}
sheetNames.forEach((item)=>{
	// 根据表名获取对应某张表
	const worksheet = workbook.Sheets[item];
	try {
		fs.writeFileSync(`./dist/${item}.json`, getJsonOfsheet(worksheet));
	} catch(err){
		console.log('写文件出错', err);
	}
	
})




