import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor() { }

  fetchDataFromExcel(file: File): Promise<any[]> {
    return new Promise((resolve, reject) => {
      const fileReader = new FileReader();

      fileReader.onload = (e) => {
        const arrayBuffer: ArrayBuffer | null = e.target?.result as ArrayBuffer;
        if (arrayBuffer) {
          console.log('ArrayBuffer:', arrayBuffer);
          const data = this.parseExcelData(arrayBuffer);
          console.log('Parsed Data:', data);
          resolve(data);
        } else {
          reject(new Error('Failed to read file.'));
        }
      };

      fileReader.onerror = (error) => {
        console.error('FileReader Error:', error);
        reject(new Error('FileReader Error'));
      };

      fileReader.readAsArrayBuffer(file);
    });
  }

  private parseExcelData(arrayBuffer: ArrayBuffer): any[] {
    const data: any[] = [];

    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    XLSX.utils.sheet_to_json(worksheet, { raw: true }).forEach((row: any) => {
      data.push(row);
    });

    return data;
  }
}