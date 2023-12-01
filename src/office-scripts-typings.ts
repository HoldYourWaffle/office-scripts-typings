#!/usr/bin/env node

import * as fs from 'fs/promises'
import * as fs_ from 'fs'
import * as path from 'path'

import { Command } from 'commander'
import { dirname } from 'path';
import { fileURLToPath } from 'url';
import ts from 'typescript'

const program = new Command();
program
	.name('office-scripts-typings')
	.version(JSON.parse(await fs.readFile(findPackageJson(), 'utf-8')).version);


program
	.command('generate')
	.description('generate typings for office-scripts')
	.option('-o, --output <dir>', "output directory, make sure this folder is included in your tsconfig typeRoots", '@types/office-scripts')
	.action(async ({ output }) => {
		if (!fs_.existsSync(output)) {
			await fs.mkdir(output, { recursive: true });
		}
		await generateTypings(output);
	});


program.parse(process.argv);


async function generateTypings(outputDir: string) {
	// TODO test correct typing behavior (somehow)

	console.log('Downloading excel.d.ts...');
	const excelDts = await fetch('https://raw.githubusercontent.com/OfficeDev/office-scripts-docs-reference/main/generate-docs/script-inputs/excel.d.ts');
	if (!excelDts.ok) {
		throw new Error(`Failed to download excel.d.ts: ${excelDts.statusText}`);
	}
	fs.writeFile(path.resolve(outputDir, 'excel.d.ts'), await excelDts.text());

	console.log('Extracting types from lib.dom.d.ts...');
	const domDtsPath = new URL(import.meta.resolve('typescript/lib/lib.dom.d.ts'));
	const domDts = ts.createSourceFile('lib.dom.d.ts', await fs.readFile(domDtsPath, 'utf8'), ts.ScriptTarget.ESNext, false, ts.ScriptKind.TS);
	fs.writeFile(path.resolve(outputDir, 'misc.d.ts'), extractTypeScriptNodes(domDts, ['Console', 'console']));

	console.log('Creating index.d.ts...');
	fs.writeFile(path.resolve(outputDir, 'index.d.ts'), '/// <reference types="./excel.d.ts" />\n/// <reference types="./misc.d.ts" />\n');

	console.log('Done!');
}


// Adapted from https://github.com/microsoft/TypeScript/wiki/Using-the-Compiler-API#re-printing-sections-of-a-typescript-file
function extractTypeScriptNodes(sourceFile: ts.SourceFile, identifiers: string[]): string {
	// TODO figure out minimum typescript version for peerDependencies (3.7?)
	const nodes: Record<string, ts.Node> = {};

	ts.forEachChild(sourceFile, node => {
		let name = '';

		if (ts.isInterfaceDeclaration(node)) {
			name = node.name.text
		} else if (ts.isVariableStatement(node)) {
			name = node.declarationList.declarations[0]!.name.getText(sourceFile);
		}

		if (identifiers.includes(name)) {
			nodes[name] = node;
		}
	});

	const notFound = identifiers.filter(id => !(id in nodes));
	if (notFound.length > 0) {
		throw new Error(`Could not find identifiers: ${notFound.join(', ')}.`);
	}

	const printer = ts.createPrinter({ newLine: ts.NewLineKind.LineFeed });
	return Object.values(nodes).map(node => printer.printNode(ts.EmitHint.Unspecified, node, sourceFile)).join('\n\n');
}


function findPackageJson(): string {
	const __dirname = dirname(fileURLToPath(import.meta.url));
	
	for (let parentIndex = 0; parentIndex < 10; parentIndex++) {
		const parentDir = path.resolve(__dirname, '../'.repeat(parentIndex));
		const packageJsonPath = path.resolve(parentDir, './package.json');
		if (fs_.existsSync(packageJsonPath)) {
			return packageJsonPath;
		}
	}
	throw new Error(`Could not find package.json for ${__dirname}`);
}
