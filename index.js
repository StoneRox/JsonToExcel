let getArguments = require('./CI-arguments');
let fs = require('fs');
var excel = require('excel4node');

// Create a new instance of a Workbook class
var workbook = new excel.Workbook();



let commands = [
	{command: 'pathToJSON', alias: 'P', description: 'Here should be the path to the json file.'},
	{command: 'selectProperty', alias: 'sp', description: 'Here a property of the json can be selected to be used\ninstead of the whole json. You have access to [json].YOUR_PROPERTY.\nExaple: if the json is {"somekey":{"someotherkey":["test"]}}\nand you use this option like -sp somekey.someotherkey in result\n["test"] will be used as json to work with.'},
	{command: 'outputFileName', alias: 'ofn', description: 'Here should be name of the excel file.\nIf no name is provided the name will be the current date and time of creation.'},
	{command: 'help', alias: 'h', description: 'This option will show the help section.'}
];

let config = getArguments(commands);

if (config.help) {
	commands.forEach((cmd) => {
		console.log();
		console.log("\x1b[32m",`command: --${cmd.command} or -${cmd.alias}\n${cmd.description}`,"\x1b[37m");
	})
	process.exit(1);
}

let jsonFile = JSON.parse(fs.readFileSync(config.pathToJSON, { encoding: 'utf8' }));

if (config.selectProperty) {
	jsonFile = eval('jsonFile.' + config.selectProperty);
}
jsonFile = [].concat(jsonFile);
let headings = [];

if (jsonFile instanceof Array) {
	jsonFile.forEach((element, worksheetIndex) => {
		// Add Worksheets to the workbook
		let worksheet = workbook.addWorksheet('Sheet ' + worksheetIndex);
		headings = [];
		if (element.constructor === Object || element instanceof Array) {
			Object.keys(element).forEach((key, i) => {
				let currentColumn = 1;
				// First prop : Row heading
				setHeading(worksheet,worksheet.lastUsedRow, currentColumn,key);
				if (element[key].constructor === Object || element[key] instanceof Array) {
					Object.keys(element[key]).forEach((key2,i2) => {						
						let secondValue = element[key][key2];
						// Second prop : Column heading
						setHeading(worksheet,1, currentColumn + i2 + 1,key2);
						let nextRow = i2 > 0 ? 2 : worksheet.lastUsedRow + 1;
						if (secondValue.constructor === Object || secondValue instanceof Array) {
							Object.keys(element[key][key2]).forEach((key3,i3) => {
								let thirdValue = element[key][key2][key3];
								// Third prop : Row heading + \n
								setHeading(worksheet,nextRow++, 1,key3);	
								if (thirdValue.constructor === Object || thirdValue instanceof Array) {
									Object.keys(thirdValue).forEach((key4,i4) => {
										// Fourth prop: Row heading -> row values									
										setHeading(worksheet,nextRow, 1,key4);								
										setString(worksheet,nextRow, worksheet.lastUsedCol,thirdValue[key4]);
										nextRow++;
									});
								} else {
									setString(worksheet,nextRow - 1, worksheet.lastUsedCol,thirdValue);
								}
							});
						} else {
							setString(worksheet,nextRow, worksheet.lastUsedCol,secondValue);							
						}
					});
				} else {
					setString(worksheet,worksheet.lastUsedRow, worksheet.lastUsedCol + 1,element[key]);
				}
			});
		}
	})
}
let fileName = (config.outputFileName || new Date().toLocaleString().replace(/[^\w]/g,''))+'.xlsx';
workbook.write(fileName);
console.log("\x1b[32m","Done!","\x1b[37m");
console.log("\x1b[32m",`${fileName} has been created!`,"\x1b[37m");
function setHeading (worksheet, row, col, heading) {
	if (!headings.find((h) => h.row === row && h.col === col && h.heading === heading)	) {
		worksheet.cell(row, col).string(heading);
		headings.push({
			row,
			col,
			heading
		});		
		//console.log(row,col,heading);
	}
}

function setString (worksheet, row, col, string) {
	worksheet.cell(row, col).string(JSON.stringify(string));
	//console.log(row,col,string);
}