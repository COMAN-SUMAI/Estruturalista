/**
 * API pública SOMENTE DE LEITURA do painel COMAN / SUMAI.
 * ATENÇÃO: pessoas na internet conseguem consultar esses campos.
 * Não devolve remetentes, notas técnicas, observações ou e-mails.
 * Implante como Web App: Executar como EU; acesso QUALQUER PESSOA.
 */
const SPREADSHEET_ID = '1DixWa79scXUArhq5fJD4kMJDep0eNHmvZTsyLbMyns8';
const SHEET_NAME = 'Controle de avaliações';

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || '');
  if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput('Callback inválido').setMimeType(ContentService.MimeType.TEXT);
  }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Aba não encontrada');
    const last = sheet.getLastRow();
    const values = last > 1 ? sheet.getRange(2, 1, last - 1, 13).getValues() : [];
    const iso = value => {
      if (value instanceof Date) return Utilities.formatDate(value, 'America/Bahia', 'yyyy-MM-dd');
      let s = String(value || '');
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s.slice(6) + '-' + s.slice(3,5) + '-' + s.slice(0,2);
      return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
    };
    const rows = values.map((v,i) => ({
      row:i+2,
      date:iso(v[0]),
      unit:String(v[2] || '').slice(0,150),
      service:String(v[3] || '').replace(/\s+/g, ' ').slice(0,280),
      req:String(v[4] || '-').slice(0,55),
      priority:Number((String(v[5] || '').match(/[123]/) || [0])[0]),
      rank:v[6] === '' ? null : Number(v[6]) || null,
      formal:String(v[9] || '').trim().toLowerCase() === 'sim'
    })).filter(x=>x.unit);
    const payload = {rows, updatedAt:Utilities.formatDate(new Date(),'America/Bahia','dd/MM/yyyy HH:mm')};
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(payload).replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029') + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (error) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify({error:'Erro de leitura da planilha'}) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}
