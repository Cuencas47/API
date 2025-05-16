import express from 'express';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import fs from 'fs/promises';
import 'dotenv/config'

const app = express();
const PORT = process.env.PORT;

// Carrega as credenciais e inicializa o JWT
async function criarJWT() {
  const raw = await fs.readFile('./credenciais.json', 'utf8');
  const creds = JSON.parse(raw);

  return new JWT({
    email: creds.client_email,
    key: creds.private_key,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive.file',
    ],
  });
}

// Função principal para ler os dados da planilha
async function acessarPlanilha() {
  try {
    const jwt = await criarJWT();
    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, jwt);      
    await doc.loadInfo();                                                             //Lê os dados da planilha

    const sheet = doc.sheetsByIndex[0];                                               //Acessa a primeira Aba
    console.log(`Título da aba: ${sheet.title}`);                                     

    await sheet.loadCells('A1:Z123');                                                  //Celulas acessadas
    const rows = await sheet.getRows();
    const header = sheet.headerValues;

    const meses = header.slice(2, 25);   // Colunas B a S - Acessa o cabeçalho

    const resultado = [];
  
    for (let rowIndex = 0; rowIndex < rows.length && rowIndex < 123; rowIndex++) {
      const row = rows[rowIndex];
      const cod_setor = row._rawData[0];
      const setor = row._rawData[1];
      if (!setor) continue;

      for (let colMes = 2; colMes <= 25; colMes++) {
        if (colMes % 2 !== 0) continue;
        const data = meses[colMes - 2];
        const valor = row._rawData[colMes];

    // Procurar se já existe no resultado (mesmo setor + data)
    let item = resultado.find(
      r => r.cod_setor === cod_setor && r.data === data
    );

    if (!item) {
      item = {
        cod_setor,
        setor,
        data,
        valor_absenteismo: null,
        valor_rotatividade: null
      };
      resultado.push(item);
    }

    if (rowIndex < 61) {
      item.valor_absenteismo = valor;
    } else if (rowIndex >= 62) {
      item.valor_rotatividade = valor;
    }
  }
}

  return resultado;

  } catch (error) {
    console.error('Erro ao acessar a planilha:', error);
    return { erro: 'Falha ao acessar a planilha'};
  }
}

const dados = await acessarPlanilha();

// Rota que carrega e retorna os dados da planilha
app.get('/', async (req, res) => {
  const dados = await acessarPlanilha();
  res.json(dados);
});

app.listen(PORT, () => {
  console.log(`O servidor está rodando na porta ${PORT}`);
});


