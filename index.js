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

    await sheet.loadCells('A1:Z61');                                                  //Celulas acessadas
    const rows = await sheet.getRows();
    const header = sheet.headerValues;

    const meses = header.slice(2, 19);   // Colunas B a S                             //Acessa o cabeçalho
    const motivos = header.slice(21, 25); // Colunas V a Z

    const setores = [];
    const absenteismo = [];
    const motAbs = []; //Motivo do absenteismo

    for (let rowIndex = 0; rowIndex < rows.length && rowIndex < 61; rowIndex++) {
      const row = rows[rowIndex];
      const cod_setor = row._rawData[0];
      const setor = row._rawData[1];
      if (!setor) continue;
      setores.push({ cod_setor, setor});

      for (let colMes = 2; colMes <= 18; colMes++) {
        if (colMes % 2 !== 0) continue;
        const data = meses[colMes - 2];
        const valorAb = row._rawData[colMes];
        absenteismo.push({cod_setor, data, valorAb});
      }    
        
        for (let colMotivo = 21; colMotivo <= 25; colMotivo++) {
        const motivo = motivos[colMotivo - 21];
        const valorMAb = row._rawData[colMotivo];
        motAbs.push({cod_setor, motivo, valorMAb});   
                
    }   
  }
    return {
      setores,
      absenteismo,
      motAbs
    }

  } catch (error) {
    console.error('Erro ao acessar a planilha:', error);
    return { erro: 'Falha ao acessar a planilha'};
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
