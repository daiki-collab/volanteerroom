import type { APIRoute } from 'astro';
import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

const EXCEL_FILE_PATH = path.join(process.cwd(), 'data', 'registrations.xlsx');

interface RegistrationData {
  name: string;
  affiliation: string;
  studentId: string;
  email: string;
  phone: string;
}

function ensureDataDirectory() {
  const dataDir = path.join(process.cwd(), 'data');
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
  }
}

function createExcelFile() {
  const headers = ['登録日時', '名前', '所属', '学生番号', 'メールアドレス', '電話番号'];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers]);
  XLSX.utils.book_append_sheet(wb, ws, 'registrations');
  XLSX.writeFile(wb, EXCEL_FILE_PATH);
}

function addRegistrationToExcel(data: RegistrationData) {
  let wb: XLSX.WorkBook;
  
  if (fs.existsSync(EXCEL_FILE_PATH)) {
    wb = XLSX.readFile(EXCEL_FILE_PATH);
  } else {
    ensureDataDirectory();
    createExcelFile();
    wb = XLSX.readFile(EXCEL_FILE_PATH);
  }
  
  const wsName = 'registrations';
  const ws = wb.Sheets[wsName];
  
  // Get existing data
  const existingData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
  
  // Add new row
  const timestamp = new Date().toLocaleString('ja-JP');
  const newRow = [
    timestamp,
    data.name,
    data.affiliation,
    data.studentId,
    data.email,
    data.phone
  ];
  
  existingData.push(newRow);
  
  // Create new worksheet with updated data
  const newWs = XLSX.utils.aoa_to_sheet(existingData);
  wb.Sheets[wsName] = newWs;
  
  // Save the file
  XLSX.writeFile(wb, EXCEL_FILE_PATH);
}

export const POST: APIRoute = async ({ request }) => {
  try {
    const data: RegistrationData = await request.json();
    
    // Validate required fields
    if (!data.name || !data.affiliation || !data.studentId || !data.email || !data.phone) {
      return new Response(JSON.stringify({ error: '必須フィールドが入力されていません' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      });
    }
    
    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(data.email)) {
      return new Response(JSON.stringify({ error: 'メールアドレスの形式が正しくありません' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      });
    }
    
    // Add registration to Excel file
    addRegistrationToExcel(data);
    
    return new Response(JSON.stringify({ message: '登録が完了しました' }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' }
    });
    
  } catch (error) {
    console.error('Registration error:', error);
    return new Response(JSON.stringify({ error: '内部サーバーエラーが発生しました' }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' }
    });
  }
};