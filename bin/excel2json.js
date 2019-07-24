#! /usr/bin/env node
const fs = require('fs');
const XLSX = require('xlsx');
const getCustomerKeyItem = (item, theme) => { // '国际编码'
	const data = {};
	Object.keys(theme).forEach((subItem)=>{
		data[theme[subItem]] = item[subItem];
	})
	return data;
}
const generateJsonByPathAndOutDir = (path, outDir, themePath) => {
	const workbook = XLSX.readFile(path);
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
				let newItem;
				if(themePath){
					const theme = require(themePath);
					newItem = getCustomerKeyItem(item, theme);
				} else {
					newItem = item;
				}
				newData.push(item)
			}
		});
		return JSON.stringify(newData);
	}
	sheetNames.forEach((item)=>{
		// 根据表名获取对应某张表
		const worksheet = workbook.Sheets[item];
		try {
			fs.writeFileSync(`${outDir}/${item}.json`, getJsonOfsheet(worksheet));
		} catch(err){
			console.log('写文件出错', err);
		}
	})
} 

var program = require('commander');
 
program
  .option('-s, --source', '指定要转换的源文件')
  .option('-o, --out', '指定要输出的目录')
  .option('-t, --theme', '指定要采用的主题文件路径')
  .action(function (options1, options2, options3) {
  	console.log('options1', options1);
  	console.log('options2', options2);
  	console.log('options1-1', !options1);
    if(!options1){
    	console.log('请指定要转换的源文件');
    	return;
    }
    if(!options2){
		console.log('请指定要输出的目录');
    	return;
    }
    console.log('options3:', typeof options3 === 'string' ? options3 : '');
    generateJsonByPathAndOutDir(options1, options2, typeof options3 === 'string' ? options3 : '');
  })
  .parse(process.argv);