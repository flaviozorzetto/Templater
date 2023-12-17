const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const prompt = require('prompt-sync')();
const fs = require('fs');
const path = require('path');

let inputFolderName = 'input';
let outputFolderName = 'output';
let payloadFolderName = 'payload';

let rootPath = path.resolve();

function createFolder(folderName) {
	let directoryContent = fs.readdirSync(rootPath);
	if (!directoryContent.filter(f => f == folderName).length == 1) {
		fs.mkdirSync(folderName);
	}
}

function ensureNecessaryFilesExists() {
	createFolder(inputFolderName);
	createFolder(outputFolderName);
	createFolder(payloadFolderName);

	let payloadFolderContent = fs.readdirSync(
		path.resolve(rootPath, payloadFolderName)
	);

	if (payloadFolderContent.filter(f => f == 'data.txt').length == 0) {
		throw new Error(
			"Erro: nenhum arquivo 'data.txt' foi encontrado dentro da pasta de payload"
		);
	}
}

function getFilesList() {
	let validDocxFilePaths = [];
	const docxRegexMatcher = /(\.)(\w*)$/g;
	const files = fs.readdirSync(path.resolve(rootPath, 'input'));

	files.forEach((file, index) => {
		let match = file.match(docxRegexMatcher)[0];
		if (match == '.docx') {
			const validDocxFilePath = path.resolve(rootPath, 'input', file);
			validDocxFilePaths.push({
				fileName: file,
				filePath: validDocxFilePath,
				index,
			});
		}
	});

	if (validDocxFilePaths.length == 0) {
		throw new Error(
			'Erro: não existe nenhum arquivo do tipo .docx na pasta de input, insira um valido e execute de novo o programa'
		);
	}

	return validDocxFilePaths;
}

function throwInvalidNumberException() {
	throw new Error('Erro: selecione um número válido');
}

function selectFileFromList(files) {
	let message = 'Selecione um dos arquivos a serem formatados: ';

	files.forEach(file => {
		message += `${file.index} - ${file.fileName}; `;
	});

	let promptInput = prompt(message);
	if (promptInput.length == 0) {
		throwInvalidNumberException();
	}

	let selectedInput = Number(promptInput);
	if (isNaN(selectedInput)) {
		throwInvalidNumberException();
	}

	let fileToReturn = files[selectedInput];
	if (!fileToReturn) {
		throwInvalidNumberException();
	}

	return fileToReturn;
}

function generateOutputFile(file) {
	let payloadData = fs
		.readFileSync(path.resolve(rootPath, payloadFolderName, 'data.txt'))
		.toString();
	let payloadJsonData;
	try {
		payloadJsonData = JSON.parse(payloadData);
	} catch (error) {
		throw new Error(
			'Erro: não foi possivel converter o arquivo para um JSON válido, garanta que o arquivo data.txt tenha um formato válido seguindo a seguinte estrutura: { "nome_da_variavel": "valor_da_variavel", "nome_da_variavel_2", "valor_da_variavel_2" }, site de referencia: https://support.oneskyapp.com/hc/en-us/articles/208047697-JSON-sample-files'
		);
	}
	try {
		const docxContent = fs.readFileSync(file.filePath, 'binary');
		const zipDocx = new PizZip(docxContent);
		const doc = new Docxtemplater(zipDocx, {
			paragraphLoop: true,
			linebreaks: true,
		});
		doc.render(payloadJsonData);
		const buf = doc.getZip().generate({
			type: 'nodebuffer',
			compression: 'DEFLATE',
		});
		fs.writeFileSync(
			path.resolve(rootPath, outputFolderName, file.fileName),
			buf
		);
	} catch (error) {
		throw new Error(error);
	}
}

function executeProgram() {
	console.log('Starting program execution');
	try {
		ensureNecessaryFilesExists();
		let files = getFilesList();
		let selectedFile = selectFileFromList(files);
		generateOutputFile(selectedFile);
	} catch (error) {
		prompt(error.message);
	}
}

executeProgram();
