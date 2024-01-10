const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const prompt = require('prompt-sync')();
const fs = require('fs');
const path = require('path');

let rootPath = path.resolve();
let inputDirectoryPath = path.join(rootPath, 'input');

function logDevInfo(message) {
	console.log('[DEV-INFO] ' + message);
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
	let docsInputPath = path.join(inputDirectoryPath, 'docs');
	let excelsInputPath = path.join(inputDirectoryPath, 'excels');

	const docsFileInputContent = getFolderContent(docsInputPath)
		.filter(fileName => fileName.match('.docx'))
		.map(fileName => {
			let formatedFileName = fileName.replace('.docx', '');
			return {
				filename: formatedFileName,
				path: path.join(docsInputPath, fileName),
			};
		});

	const excelsFileInputContent = getFolderContent(excelsInputPath)
		.filter(fileName => fileName.match('.xlsx'))
		.map(fileName => {
			let formatedFileName = fileName.replace('.xlsx', '');
			return {
				filename: formatedFileName,
				path: path.join(excelsInputPath, fileName),
			};
		});

	const validMatchedFiles = docsFileInputContent.filter(docsFileNameObj => {
		return (
			excelsFileInputContent.filter(e => e.filename == docsFileNameObj.filename)
				.length == 1
		);
	});

	return validMatchedFiles;
}

// function generateOutputFile(file) {
// 	let payloadData = fs
// 		.readFileSync(path.resolve(rootPath, payloadFolderName, 'data.txt'))
// 		.toString();
// 	let payloadJsonData;
// 	try {
// 		payloadJsonData = JSON.parse(payloadData);
// 	} catch (error) {
// 		throw new Error(
// 			'Erro: não foi possivel converter o arquivo para um JSON válido, garanta que o arquivo data.txt tenha um formato válido seguindo a seguinte estrutura: { "nome_da_variavel": "valor_da_variavel", "nome_da_variavel_2", "valor_da_variavel_2" }, site de referencia: https://support.oneskyapp.com/hc/en-us/articles/208047697-JSON-sample-files'
// 		);
// 	}
// 	try {
// 		const docxContent = fs.readFileSync(file.filePath, 'binary');
// 		const zipDocx = new PizZip(docxContent);
// 		const doc = new Docxtemplater(zipDocx, {
// 			paragraphLoop: true,
// 			linebreaks: true,
// 		});
// 		doc.render(payloadJsonData);
// 		const buf = doc.getZip().generate({
// 			type: 'nodebuffer',
// 			compression: 'DEFLATE',
// 		});
// 		fs.writeFileSync(
// 			path.resolve(rootPath, outputFolderName, file.fileName),
// 			buf
// 		);
// 	} catch (error) {
// 		throw new Error(error);
// 	}
// }

function executeProgram() {
	logDevInfo('Starting program execution');
	try {
		ensureNecessaryFilesExists();
		let validFiles = getValidFiles();
	} catch (error) {
		prompt(error.message);
	}
}

executeProgram();
