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
    await doc.loadInfo();

    const sheet = doc.sheetsByIndex[0];
    console.log(`Título da aba: ${sheet.title}`);

    await sheet.loadCells('A1:Z61');
    const rows = await sheet.getRows();
    const header = sheet.headerValues;

    const resultado = [];

    const meses = header.slice(1, 19);   // Colunas B a S
    const motivos = header.slice(20, 26); // Colunas U a Z

    for (let rowIndex = 0; rowIndex < rows.length && rowIndex < 61; rowIndex++) {
      const row = rows[rowIndex];
      const setor = row._rawData[0];
      if (!setor) continue;

      for (let colMes = 1; colMes <= 18; colMes++) {
        if (colMes % 2 === 0) continue; // Apenas colunas ímpares
        const mes = meses[colMes - 1];
        const valorMes = row._rawData[colMes];

        for (let colMotivo = 20; colMotivo <= 25; colMotivo++) {
          const motivo = motivos[colMotivo - 20];
          const valorMotivo = row._rawData[colMotivo];

          resultado.push({
            setor,
            mes,
            valorMes,
            motivo,
            valorMotivo,
          });
        }
      }
    }

    return resultado;

  } catch (error) {
    console.error('Erro ao acessar a planilha:', error);
    return { erro: 'Falha ao acessar a planilha' };
  }
}

// Rota que carrega e retorna os dados da planilha
app.get('/', async (req, res) => {
  const dados = await acessarPlanilha();
  res.json(dados);
});

app.listen(PORT, () => {
  console.log(`O servidor está rodando na porta ${PORT}`);
});
