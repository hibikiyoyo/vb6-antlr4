import { VisualBasic6Lexer, VisualBasic6Parser } from "vb6-antlr4";
import { CharStreams, CommonTokenStream } from "antlr4ts";
import * as fs from "fs";
import * as path from "path";

const extensionsList = ['vb', 'frm', 'cls', 'bas', 'vbw', 'vbp'];

// Function to parse a VB6 file and save its parse tree to a text file
function parseAndSaveTree(filePath, projectFolder, outputFolder) {
    let extension = filePath.split('.').pop()

    if (!extensionsList.includes(extension)) {
        return
    }

    const extensionsRegex = new RegExp(`\\.(${extensionsList.join('|')})$`, 'i');


    let input = fs.readFileSync(filePath, "utf8");
    let inputStream = new CharStreams.fromString(input);

    let lexer = new VisualBasic6Lexer(inputStream);
    let tokenStream = new CommonTokenStream(lexer);
    let parser = new VisualBasic6Parser(tokenStream);
    let tree = parser.startRule();


    // Generate the output path based on the original structure
    const relativePath = path.relative(projectFolder, filePath);
    const outputPath = path.join(outputFolder, relativePath.replace(extensionsRegex, "_" + extension + "_tree.txt"));

    // Ensure the directory exists
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });

    // Save the tree to the output path
    fs.writeFileSync(outputPath, tree.toStringTree(parser), "utf8");
    console.log(`Parse tree saved for ${filePath} to ${outputPath}`);
}

// Recursive function to find VB-related files in all subdirectories
function getAllFiles(directory) {
    let files = [];
    const items = fs.readdirSync(directory, { withFileTypes: true });

    items.forEach(item => {
        const itemPath = path.join(directory, item.name);
        if (item.isDirectory()) {
            files = files.concat(getAllFiles(itemPath)); // Recurse into subdirectory
        } else {
            files.push(itemPath); // Add matching file
        }
    });

    return files;
}


/// Main function to parse all VB files in a folder
function parseProjectFolder(projectFolder, outputFolder) {
    // Ensure the output folder exists
    if (!fs.existsSync(outputFolder)) {
        fs.mkdirSync(outputFolder);
    }

    const files = getAllFiles(projectFolder);
    saveTheStructureOfProject(files, outputFolder)
    console.log(files)
    files.forEach(file => parseAndSaveTree(file, projectFolder, outputFolder));
}

function saveTheStructureOfProject(files, outputFolder) {
    const content = files.join('\n');

    const outputPath = path.join(outputFolder, "structure.txt");
    console.log(outputPath)
    // Ensure the directory exists
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });

    // Write the content to a text file
    fs.writeFile(outputPath, content, (err) => {
    if (err) {
        console.error('Error writing file:', err);
    } else {
        console.log('File has been written successfully!');
    }
    });
}
// Usage

const projectFolder = "vbFile"; // Path to your VB project folder
const outputFolder = "output"; // Folder to save the output trees
parseProjectFolder(projectFolder, outputFolder);

