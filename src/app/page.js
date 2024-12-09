"use client"
import Link from 'next/link';
import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const ExcelHandler = () => {
  const [data, setData] = useState([]);
  const [headers, setHeaders] = useState([]);

  // Function to handle file upload
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
      const binaryStr = event.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      setHeaders(jsonData[0] || []);
      setData(jsonData.slice(1));
    };

    reader.readAsBinaryString(file);
  };

  // Function to handle file download
  const handleFileDownload = () => {
    // Get all input values from the table
    const updatedData = data.map(row => 
      headers.map((_, colIdx) => {
        // If the cell is empty or undefined, use the previous non-empty value
        if (row[colIdx] === '' || row[colIdx] === undefined) {
          let lastValue = '';
          // Look for the last non-empty value in the same column
          for (let i = 0; i < data.length; i++) {
            if (data[i][colIdx] !== '' && data[i][colIdx] !== undefined) {
              lastValue = data[i][colIdx];
            }
            if (i >= data.indexOf(row)) break;
          }
          return lastValue;
        }
        return row[colIdx];
      })
    );

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...updatedData]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    XLSX.writeFile(workbook, 'output.xlsx');
  };

  // Function to update cell data
  const handleInputChange = (rowIdx, colIdx, value) => {
    const updatedData = [...data];
    updatedData[rowIdx][colIdx] = value;
    setData(updatedData);
  };
  let old=""

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel Upload and Download</h1>
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
        style={{ marginBottom: '20px' }}
      />

      {data.length > 0 && (
        <table border="1" style={{ marginBottom: '20px' }}>
          <thead>
            <tr>
              {headers.map((header, idx) => (
                <th key={idx}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.map((row, rowIdx) => (
              <tr key={rowIdx}>
                {headers.map((_, colIdx) => {
                  if(row[colIdx]!='' && row[colIdx]!=undefined){
                    old=row[colIdx]
                  }
                  return(
                  <td key={colIdx}>
                    <input
                      type="text"
                      className='border border-gray-300 text-black rounded-md p-2'
                      value={row[colIdx] || old}
                      onChange={(e) =>
                        handleInputChange(rowIdx, colIdx, e.target.value)
                      }
                    />
                  </td>
                  )
                }
                )}
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {data.length > 0 && (
        <button onClick={handleFileDownload}>Download Excel</button>
      )}
      by <Link target='_blank' href={'https://dishantkapoor.com'}>Dishant Kapoor</Link>
    </div>
  );
};

export default ExcelHandler;
