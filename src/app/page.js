"use client"
import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import Link from "next/link";

const Home = () => {
  const [data, setData] = useState([]);
  const [filename, setFilename] = useState("output.xlsx");
  const [column, setColumn] = useState("");

  // Handle file upload
  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const binaryStr = e.target.result;
        const workbook = XLSX.read(binaryStr, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        let headers=Object.keys(sheetData[0]);
        let dummyData=[];
        let dummyOld={};
        console.log(headers,"headers")
        console.log(sheetData,"sheetData",sheetName,"sheetName")
        for(let i=0;i<sheetData.length;i++){
          let row={};
          for(let j=0;j<headers.length;j++){
            if(sheetData[i][headers[j]]!="" && sheetData[i][headers[j]]!=undefined){
              row[headers[j]]=sheetData[i][headers[j]];
              dummyOld[j]=sheetData[i][headers[j]];
            }else{
              if(headers[j]==column){
                row[headers[j]]=dummyOld[j];
              }else{
                row[headers[j]]="";
              }
            }
          }
          dummyData.push(row);
        }
        console.log(dummyData)
        setData(dummyData);
      };
      reader.readAsBinaryString(file);
    }
  };

  // Handle data editing
  const handleEdit = (rowIndex, key, value) => {
    const updatedData = data.map((row, index) => {
      if (index === rowIndex) {
        return { ...row, [key]: value };
      }
      return row;
    });
    setData(updatedData);
  };

  // Handle empty cells
  const handleEmptyCells = () => {
    const updatedData = data.map((row, rowIndex) => {
      const updatedRow = { ...row };
      Object.keys(row).forEach((key) => {
        if (row[key] === "" || row[key] === undefined) {
          // Find the nearest above non-empty cell value
          let value = "";
          for (let i = rowIndex - 1; i >= 0; i--) {
            if (data[i][key] !== "" && data[i][key] !== undefined) {
              value = data[i][key];
              break;
            }
          }
          updatedRow[key] = value;
        }
      });
      return updatedRow;
    });
    setData(updatedData);
  };

  // Handle file download
  const handleFileDownload = () => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = {
      Sheets: { data: worksheet },
      SheetNames: ["data"],
    };
    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([excelBuffer], { type: "application/octet-stream" });
    saveAs(blob, filename);
  };

  return (
    <div style={{ padding: "20px" }} className="bg-black rounded-md p-4">
      <h1 className="text-2xl font-bold text-white">Excel File Upload and Edit</h1>
      <p className="text-sm text-gray-500">Fill the empty cells with the nearest above non-empty cell value</p>
      <p className="text-sm font-bold text-white">Instructions:</p>
      <p className="text-sm text-white">1. Enter the column name to fill the empty cells</p>
      <p className="text-sm text-white">2. Upload the Excel file</p>
      <p className="text-sm text-white">3. Keep the column name on first row</p>
      <p className="text-sm text-white">4. Click on the &apos;Download Edited File&apos; button to download the edited file</p>
      <input type='text' className='border border-gray-300 text-black rounded-md p-2' value={column} name='column' placeholder='Enter column name' onChange={(e)=>setColumn(e.target.value)}/>
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
        className="border border-gray-300 text-white rounded-md p-2 ml-3"
        style={{ marginBottom: "20px" }}
      />
      <table border="1" style={{ marginBottom: "20px", width: "100%" }}>
        <thead>
          <tr>
            {data.length > 0 && Object.keys(data[0]).map((key) => <th key={key}>{key}</th>)}
          </tr>
        </thead>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {Object.keys(data[0]).map((key) => (
                (row[key]!="")?(
                <td key={key}>
                  <input
                    type="text"
                    value={row[key]}
                    className='border border-gray-300 text-black rounded-md p-2'
                    onChange={(e) => handleEdit(rowIndex, key, e.target.value)}
                  />
                </td>
              ):(
                <td key={key}>{row[key]} No Data</td>
              )))}
            </tr>
          ))}
        </tbody>
      </table>
      <input
        type="text"
        value={filename}
        onChange={(e) => setFilename(e.target.value)}
        placeholder="Enter filename"
        style={{ marginBottom: "20px" }}
        className="border border-gray-300 text-black rounded-md p-2 my-3"
      />
      <br />
      <button className="bg-blue-500 text-white rounded-md p-2" onClick={handleFileDownload}>Download Edited File</button><br/>
      by <Link target='_blank' href={'https://dishantkapoor.com'}>Dishant Kapoor</Link>
    </div>
  );
};

export default Home;


