const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const folderPath = __dirname; // Ajuste conforme necessário

// Função para decodificar Quoted-printable
function decodeQuotedPrintable(str) {
    // Remove soft line breaks (= no final da linha)
    str = str.replace(/=\r?\n/g, '');
    // Substitui =XX por o caractere correspondente
    return str.replace(/=([A-Fa-f0-9]{2})/g, function(match, hex) {
        return String.fromCharCode(parseInt(hex, 16));
    });
}

function gerarExcelContatos() {
    const pastaNumeros = path.join(folderPath, 'numeros');
    const arquivos = fs.readdirSync(pastaNumeros).filter(f => f.endsWith('.txt'));
    const excelData = [["Nome", "Número", "Arquivo", "Repetido"]];

    // Para contar repetições
    const numeroCount = {};

    // Primeiro, colete todos os números e conte as ocorrências
    const contatosTemp = [];
    arquivos.forEach(arquivo => {
        const filePath = path.join(pastaNumeros, arquivo);
        const linhas = fs.readFileSync(filePath, 'utf-8').split('\n');
        let nome = '';
        linhas.forEach(linha => {
            if (linha.startsWith('FN:') || linha.startsWith('FN;')) {
                let valor = linha.split(':')[1]?.trim() || '';
                if (/QUOTED-PRINTABLE/i.test(linha)) {
                    valor = decodeQuotedPrintable(valor);
                }
                nome = valor;
            }
            if (linha.startsWith('TEL')) {
                let numero = linha.split(':')[1]?.replace(/\D/g, '');
                if (numero) {
                    contatosTemp.push([nome, numero, arquivo]);
                    numeroCount[numero] = (numeroCount[numero] || 0) + 1;
                }
            }
            if (linha.startsWith('END:VCARD')) {
                nome = '';
            }
        });
    });

    // Agora, monte a planilha sinalizando os repetidos
    contatosTemp.forEach(([nome, numero, arquivo]) => {
        const repetido = numeroCount[numero] > 1 ? "SIM" : "";
        excelData.push([nome, numero, arquivo, repetido]);
    });

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Contatos");
    XLSX.writeFile(wb, path.join(folderPath, 'contatos_com_arquivo.xlsx'));
    console.log('Arquivo "contatos_com_arquivo.xlsx" criado com os contatos!');
}


gerarExcelContatos();