const prompt = require('prompt-sync')();
const fs = require('fs');
const path = require('path');

let validDocxFilePaths = [];
const docxRegexMatcher = /(\.)(\w*)$/g;

function executeProgram() {
	const files = fs.readdirSync(path.resolve(__dirname, 'input'));

	files.forEach(file => {
		let match = file.match(docxRegexMatcher)[0];
		if (match == '.docx') {
			const validDocxFilePath = path.resolve(__dirname, 'input', file);
			validDocxFilePaths.push({ fileName: file, filePath: validDocxFilePath });
		}
	});

	if (validDocxFilePaths.length == 0) {
		throw new Error(
			'Não existe nenhum arquivo do tipo .docx na pasta de input'
		);
	}

	console.log(validDocxFilePaths);
}

executeProgram();
