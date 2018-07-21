let getArguments = require('./CI-arguments');
let fs = require('fs');
let excel = require('excel4node');

// Create a new instance of a Workbook class
let workbook = new excel.Workbook();

let dryRunContent = [];

let commands = [
	{command: 'pathToJSON', alias: 'P', description: 'After the command should be one space and the path to the json file.\nThis command is required.'},
	{command: 'selectProperty', alias: 'sp', description: 'After the command should be one space and a property of the json\ncan be selected to be used instead of the whole json.\nYou have access to [json].YOUR_PROPERTY.\nExaple: If the json is {"somekey":{"someotherkey":["test"]}}\nand you use this option like -sp somekey.someotherkey in result\n["test"] will be used as json to work with.'},
	{command: 'outputFileName', alias: 'ofn', description: 'After the command should be one space and the name of the excel file.\nIf no name is provided the name will be the current date and time of creation.'},
	{command: 'help', alias: 'h', description: 'This option will show the help section.'},
	{command: 'dry-run', alias: 'dr', description: 'This log the cells info without creating the file.\nThis command must be used last in order to work.'}
];

let config = getArguments(commands);

if (config.help || config.pathToJSON === undefined) {
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
let sheetIndex;
let maxRowIndexUsed = 0;
let columnHeadingRow = 0;
if (jsonFile instanceof Array) {
	jsonFile.forEach((element, worksheetIndex) => {
		// Add Worksheets to the workbook
		let worksheet = workbook.addWorksheet('Sheet ' + worksheetIndex);
		headings = [];
		sheetIndex = worksheetIndex;
		dryRunContent[sheetIndex] = [];
		if (element.constructor === Object || element instanceof Array) {
			Object.keys(element).forEach((key, i) => {
				let currentColumn = 1;
				// First prop : groupHeading
				setHeading(worksheet,maxRowIndexUsed + 1, currentColumn,key);
				columnHeadingRow = worksheet.lastUsedRow;				
				if (element[key].constructor === Object || element[key] instanceof Array) {
					Object.keys(element[key]).forEach((key2,i2) => {
						let secondValue = element[key][key2];
						let lastUsedRow = i2 < 1 ? worksheet.lastUsedRow : 1;
						// Second prop : columnHeading
						let nextCol = worksheet.lastUsedCol + 1;
						setHeading(worksheet, columnHeadingRow, nextCol,key2);
						let nextRow = i2 > 0 ? columnHeadingRow + 1 : worksheet.lastUsedRow + 1;
						if (secondValue.constructor === Object || secondValue instanceof Array) {
							Object.keys(element[key][key2]).forEach((key3,i3) => {
								let thirdValue = element[key][key2][key3];
								// Third prop : rowHeading
								setHeading(worksheet,nextRow++, 1,key3);
								if (thirdValue.constructor === Object || thirdValue instanceof Array) {
									Object.keys(thirdValue).forEach((key4,i4) => {
										// Fourth prop: rowSubHeading -> cell values
										setHeading(worksheet,nextRow, 1,key4);
										setString(worksheet,nextRow, nextCol,thirdValue[key4]);   
										nextRow++;
									});
								} else {
									setString(worksheet,nextRow - 1, worksheet.lastUsedCol, thirdValue);
								}
							});
						} else {
							setString(worksheet,maxRowIndexUsed+1, i2+2,secondValue);
						}
					});
				} else {
					setString(worksheet,maxRowIndexUsed, worksheet.lastUsedCol + 1,element[key]);
				}
			});
		}
	})
}
let fileName = (config.outputFileName || new Date().toLocaleString().replace(/[^\w]/g,''))+'.xlsx';
if (config["dry-run"]) {
	for (let i = 0; i < dryRunContent.length; i++) {
		console.log('\nsheet' + i);
		for (let j = 1; j <= dryRunContent[i].length; j++) {
			let rowContent = dryRunContent[i].filter((cell) => cell.row === j);
			if (!rowContent.length > 0) {
				break;
			} else {
				let prevCol;
				let log = rowContent.sort((a,b) => {
					if (a.col > b.col) {
						return 1;
					}
					if (a.col < b.col) {
						return -1;
					}
					return 0;
				}).map((cell, cellIndex) => {
					let result = '';		
					if(cellIndex === 0) {
						prevCol = 0;
					}
					if (!isNaN(prevCol)) {
						let emptyCellsCount = cell.col - prevCol - 1;
						result += '[]'.repeat(emptyCellsCount < 0 ? 0 : emptyCellsCount);
					}
					prevCol = cell.col;
					return result + '[' + cell.content + ']';
				});
				console.log(log.join(' '));
			}
		}
	}
	process.exit(1);
}

workbook.write(fileName);
console.log("\x1b[32m","Done!","\x1b[37m");
console.log("\x1b[32m",`${fileName} has been created!`,"\x1b[37m");
function setHeading (worksheet, row, col, heading, force) {
	if (!headings.find((h) => h.row === row && h.col === col && h.content === heading) || force) {
		worksheet.cell(row, col).string(heading);
		let cell = {
			content: heading,
			row,
			col
		}
		headings.push(cell);
		worksheet.lastUsedCol = col;
		if(row > maxRowIndexUsed){
			maxRowIndexUsed = row;
		}
		dryRunContent[sheetIndex].push(cell);
	}
}

function setString (worksheet, row, col, string) {
	if (dryRunContent[sheetIndex].find((h) => h.row === row && h.col === col)) {
		col++;
	}
	if (typeof string !== 'string') {
		string = JSON.stringify(string);
	}
	worksheet.cell(row, col).string(string);
	worksheet.lastUsedCol = col;
	if(row > maxRowIndexUsed){
		maxRowIndexUsed = row;
	}
	let cell = {
			content: string,
			row,
			col
		}
	dryRunContent[sheetIndex].push(cell);
}
