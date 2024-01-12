const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const prompt = require('prompt-sync')();
const fs = require('fs');
const path = require('path');
const readXlsxFile = require('read-excel-file/node');

let rootPath = path.resolve();
let inputDirectoryPath = path.join(rootPath, 'input');
let outputDirectoryPath = path.join(rootPath, 'output');
let docsInputPath = path.join(inputDirectoryPath, 'docs');
let excelsInputPath = path.join(inputDirectoryPath, 'excels');
let showDevDebug = false;

function logDevInfo(message) {
	if (showDevDebug) {
		console.log('[DEV-INFO] ' + message);
	}
}

function directoryHasFolder(directoryPath, folderName) {
	try {
		let items = getFolderContent(directoryPath);
		let hasFolder = items.filter(item => item == folderName).length == 1;
		return hasFolder;
	} catch {
		return false;
	}
}

function getFolderContent(directoryPath) {
	try {
		return fs.readdirSync(directoryPath);
	} catch {
		return null;
	}
}

function ensureNecessaryFilesExists() {
	logDevInfo('Validating necessary folders');

	let foldersToCreate = [
		{
			directoryPath: rootPath,
			folderName: 'input',
		},
		{
			directoryPath: rootPath,
			folderName: 'output',
		},
		{
			directoryPath: inputDirectoryPath,
			folderName: 'docs',
		},
		{
			directoryPath: inputDirectoryPath,
			folderName: 'excels',
		},
	];

	foldersToCreate.forEach(folderToCreateObj => {
		if (
			!directoryHasFolder(
				folderToCreateObj.directoryPath,
				folderToCreateObj.folderName
			)
		) {
			logDevInfo(
				`Directory: ${folderToCreateObj.directoryPath} does not have folder: ${folderToCreateObj.folderName}, creating new one.`
			);
			fs.mkdirSync(
				path.join(folderToCreateObj.directoryPath, folderToCreateObj.folderName)
			);
		}
	});

	logDevInfo('Finished necessary folders validation');
}

function getValidFiles() {
	const docsFileInputContent = getFolderContent(docsInputPath)
		.filter(fileName => fileName.match('.docx'))
		.map(fileName => {
			return {
				file: fileName.replace('.docx', ''),
			};
		});

	const excelsFileInputContent = getFolderContent(excelsInputPath)
		.filter(fileName => fileName.match('.xlsx'))
		.map(fileName => {
			return {
				file: fileName.replace('.xlsx', ''),
			};
		});

	const validMatchedFiles = docsFileInputContent.filter(docsFileNameObj => {
		return (
			excelsFileInputContent.filter(e => e.file == docsFileNameObj.file)
				.length == 1
		);
	});

	return validMatchedFiles;
}

function getChoiceFromValidFiles(validFiles) {
	console.log('Arquivos validos encontrados:');
	let message = '';

	validFiles.forEach((e, i) => {
		message += `${i} - ${e.file}\n`;
	});

	console.log(message);
	console.log('Selecione utilizando o número identificador do arquivo:');

	let input = prompt();

	let inputNumber = Number(input);

	if (!Number.isInteger(inputNumber)) {
		logDevInfo(`Number: ${input} is not valid`);
		throw new Error('Input inserido não é um número válido\n');
	}

	let validFilesLength = validFiles.length;

	if (inputNumber < 0 || inputNumber > validFilesLength - 1) {
		throw new Error(
			'Input inserido não faz parte dos números de arquivos válidos'
		);
	}

	let chosenFile = validFiles[inputNumber].file;

	logDevInfo('Chosen file: ' + chosenFile);

	return chosenFile;
}

async function getExcelJsonDataListFromChoice(choice) {
	const excelFilePath = path.join(excelsInputPath, choice + '.xlsx');
	const excelJsonDataList = [];
	const concatenators = [];

	let xlsxObject = await readXlsxFile(excelFilePath);
	let headers = xlsxObject[0].map(header => {
		let result = header;
		const indexLength = result.length - 1;

		if (result[indexLength] == '_') {
			result = result.substring(0, indexLength);
			concatenators.push(result);
		}

		return result;
	});

	let headersCount = headers.length;

	let hasAllFieldsFulfilled = xlsxObject.every((row, index) => {
		const hasRowFulfilled = row.every(elem => {
			if (elem == null || elem == undefined) return false;

			return elem.trim() != '';
		});

		if (row.length == headersCount && hasRowFulfilled) {
			return true;
		}

		throw new Error(`Linha ${index + 1} está com uma linha faltando valor(es)`);
	});

	xlsxObject.shift();

	logDevInfo('Has all fields fulfilled: ' + hasAllFieldsFulfilled);

	xlsxObject.forEach(row => {
		let jsonData = {};
		headers.forEach((header, index) => {
			jsonData[header] = row[index];
		});
		excelJsonDataList.push(jsonData);
	});

	return [excelJsonDataList, concatenators];
}

function processDocxFiles(excelJsonDataList, concatenators, fileName) {
	let docxInputFilePath = path.join(docsInputPath, fileName + '.docx');

	excelJsonDataList.forEach(excelJsonData => {
		const docxContent = fs.readFileSync(docxInputFilePath, 'binary');
		const zipDocx = new PizZip(docxContent);
		const doc = new Docxtemplater(zipDocx, {
			paragraphLoop: true,
			linebreaks: true,
		});

		doc.render(excelJsonData);

		const buf = doc.getZip().generate({
			type: 'nodebuffer',
			compression: 'DEFLATE',
		});

		let concatenatedFileName = `${fileName}`;

		concatenators.forEach(concatenator => {
			concatenatedFileName += `_${excelJsonData[concatenator]}`;
		});

		fs.writeFileSync(
			path.resolve(outputDirectoryPath, concatenatedFileName + '.docx'),
			buf
		);
	});
}

async function executeProgram() {
	logDevInfo('Starting program execution');
	try {
		ensureNecessaryFilesExists();
		let validFiles = getValidFiles();
		let fileNameChoice = getChoiceFromValidFiles(validFiles);
		let [excelJsonDataList, concatenators] =
			await getExcelJsonDataListFromChoice(fileNameChoice);
		processDocxFiles(excelJsonDataList, concatenators, fileNameChoice);
	} catch (error) {
		prompt(error.message);
	}
}

executeProgram();