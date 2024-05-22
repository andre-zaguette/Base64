import * as fs from 'fs';
import * as XLSX from 'xlsx';

// Defina o caminho para o arquivo de saída base64
const base64OutputFilePath = 'output_base64.txt';

// Defina o caminho para o arquivo Excel
const excelFilePath = 'resultado_1716392722324.xlsx';

try {
  // Lê o arquivo Excel do disco
  const workbook: XLSX.WorkBook = XLSX.readFile(excelFilePath,{cellDates: true  } );
  console.log(workbook)
  // Converte a planilha Excel para um buffer
  const buffer: Buffer = XLSX.write(workbook, { type: 'buffer', cellDates: true  });

  // Codifica o buffer para uma string base64 com tipo MIME
  const base64String: string = buffer.toString('base64', + ';type=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  // Salva a string base64 em um arquivo de texto
  fs.writeFileSync(base64OutputFilePath, base64String, 'utf8');

  console.log(`Arquivo base64 salvo com sucesso como ${base64OutputFilePath}!`);
} catch (error) {
  console.error('Erro ao processar o arquivo Excel:', error);
}