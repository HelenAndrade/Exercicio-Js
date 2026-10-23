/*Tarefa JS e excel
  Hélen Andrade
*/

const XLSX = require("xlsx");

const arquivo = XLSX.readFile("produtos.xlsx");
const planilha = arquivo.Sheets[arquivo.SheetNames[0]];
const produtos = XLSX.utils.sheet_to_json(planilha);

let valorEstoque = 0;
let precisamRepor = 0;

for (const produto of produtos) {
    produto["Valor Total"] = produto["Preço"] * produto["Quantidade"];

    if (produto["Quantidade"] < 5) {
        produto["Situação"] = "Repor";
        precisamRepor++;
    } else {
        produto["Situação"] = "OK";
    }

    valorEstoque = valorEstoque + produto["Valor Total"];
}

const novaPlanilha = XLSX.utils.json_to_sheet(produtos);
const novoArquivo = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(novoArquivo, novaPlanilha, "Resultado");
XLSX.writeFile(novoArquivo, "resultado.xlsx");

const valorFormatado = valorEstoque.toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
});

console.log("Processamento concluído!");
console.log("");
console.log(`Valor total do estoque: R$ ${valorFormatado}`);
console.log(`Produtos que precisam de reposição: ${precisamRepor}`);
