import * as fs from 'fs';
import * as XLSX from 'xlsx';

// Lê a string base64 do arquivo base64.txt
let base64String: string;
try {
  base64String = fs.readFileSync('output_base64.txt', 'utf8').trim();
  // base64String = fs.readFileSync('base64.txt', 'utf8').trim();
} catch (err) {
  console.error('Erro ao ler o arquivo base64:', err);
  process.exit(1);
}

// Decodifica a string base64 para um buffer
const buffer: Buffer = Buffer.from(base64String, 'base64');

// Converte o buffer para uma planilha Excel
let workbook: XLSX.WorkBook;
try {
  workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
} catch (err) {
  console.error('Erro ao converter o buffer para planilha Excel:', err);
  process.exit(1);
}

// Gera um timestamp para adicionar ao nome do arquivo
const timestamp: number = Date.now();

// Nome do arquivo de saída com timestamp
const outputFileName: string = `resultado_${timestamp}.xlsx`;

// Salva a planilha Excel no disco com o nome único
try {
  XLSX.writeFile(workbook, outputFileName);
  console.log(`Planilha Excel salva com sucesso como ${outputFileName}!`);
} catch (err) {
  console.error('Erro ao salvar a planilha Excel:', err);
  process.exit(1);
}