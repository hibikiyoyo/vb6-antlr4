// var vb6Antlr4 = require("vb6-antlr4")
import { VisualBasic6Lexer, VisualBasic6Parser } from "vb6-antlr4";
import { ANTLRInputStream, CommonTokenStream } from "antlr4ts";

// const antlr4 = require('antlr4');
let inputStream = new ANTLRInputStream(`
Private Sub Command1_Click ()
   Text1.Text = "Hello, world!"
End Sub
`);
let lexer = new VisualBasic6Lexer(inputStream);
let tokenStream = new CommonTokenStream(lexer);
let parser = new VisualBasic6Parser(tokenStream);
 
let tree = parser.startRule();
console.log(tree.toStringTree(parser))