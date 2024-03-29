#! /usr/bin/env node
const fs = require('fs');
const XLSX = require('xlsx');
const getCustomerKeyItem = (item, theme) => { // '国际编码'
	const data = {};
	const themeKeys = Object.keys(theme);
	themeKeys.forEach((subItem, index)=>{
		data[theme[subItem]] = item[subItem];
	})
	return data;
}
const generateJsonByPathAndOutDir = (path, outDir, themePath) => {
	const workbook = XLSX.readFile(path);
	// 获取 Excel 中所有表名
	const sheetNames = workbook.SheetNames;
	const getJsonOfsheet = (worksheet, sheetName, categoryName)=>{
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
		const theme = require(themePath);
		const themeKeys = Object.keys(theme);
		const id = themeKeys.find((key)=>{
			return theme[key] === 'id';
		});

		data.forEach((item, index)=>{
			if(item) {
				const newItem = !!themePath ? getCustomerKeyItem(item, theme) : item;
				newData.push(newItem);
			}
		});
		return newData.map((item, index)=>{
			if (!id){
				return {
					id: 'code_' + categoryName + '_' + sheetName + '_' + (index + 1),
					...item
				};
			} else {
				return item;
			}
		});
	}
	const fileNameIndex = path.lastIndexOf('/') + 1;
	const fileName = path.substring(fileNameIndex);
	const name = fileName.split('.')[0];
	const json = {};
	sheetNames.forEach((item, index)=>{
		const worksheet = workbook.Sheets[item];
		json[item] = getJsonOfsheet(worksheet, item, name);
	});
	try {
		fs.writeFileSync(`${outDir}/${name}.json`, JSON.stringify(json));
	} catch(err){
		console.log('写文件出错', err);
	}
}

var program = require('commander');

program
  .option('-s, --source', '指定要转换的源文件')
  .option('-o, --out', '指定要输出的目录')
  .option('-t, --theme', '指定要采用的主题文件路径')
  .action(function (options1, options2, options3) {
    if(!options1){
    	console.log('请指定要转换的源文件');
    	return;
    }
    if(!options2){
		console.log('请指定要输出的目录');
    	return;
    }
    generateJsonByPathAndOutDir(options1, options2, typeof options3 === 'string' ? options3 : '');
  })
  .parse(process.argv);
