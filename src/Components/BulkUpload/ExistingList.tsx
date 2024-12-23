import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/fields";
import styles from "../../Components/BulkUpload/ExistingList.module.scss";
import * as XLSX from "xlsx"; // For parsing Excel and CSV files

const ExistingList: React.FC = () => {
  const [lists, setLists] = useState<any[]>([]); // State to store the lists
  const [selectedList, setSelectedList] = useState<string>(""); // Selected list ID
  const [listColumns, setListColumns] = useState<string[]>([]); // Columns of the selected list
  const [fileData, setFileData] = useState<any[]>([]); // Uploaded file data
  const [] = useState<string[]>([]); // Headers from the uploaded file
  const [showPopup, setShowPopup] = useState<boolean>(false); // Popup state

  // Fetch all lists when the component mounts
  useEffect(() => {
    sp.web.lists
      .filter("BaseTemplate eq 100") // Fetch only custom lists
      .get()
      .then((listData) => {
        const filteredLists = listData.map((list) => ({
          id: list.Id,
          title: list.Title,
        }));
        setLists(filteredLists);
      });
  }, []);

  // Fetch columns of the selected list when it changes
  useEffect(() => {
    if (selectedList) {
      sp.web.lists
        .getById(selectedList)
        .fields.filter("Hidden eq false and ReadOnlyField eq false") // Exclude hidden and read-only fields
        .get()
        .then((fields) => {
          const columnNames = fields.map((field) => field.Title);
          setListColumns(columnNames);
        });
    } else {
      setListColumns([]);
    }
  }, [selectedList]);

  const handleDropdownChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setSelectedList(event.target.value);
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
  
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  
        // Get headers from the first row
        const headers = jsonData[0] as string[];
  
        // Validate file headers with list columns
        const isValid = headers.every((header) => listColumns.includes(header)); // Check if headers exist in listColumns
        if (!isValid) {
          setShowPopup(true); // Show popup if validation fails
        } else {
          setFileData(jsonData);
        }
      };
  
      reader.readAsArrayBuffer(file);
    }
  };
  

  const closePopup = () => {
    setShowPopup(false);
  };

  return (
    <div className={styles.dropdowncontainer}>
      <h1 className={styles.dropdownheader}>Select a List</h1>
      <select
        className={styles.dropdownselect}
        value={selectedList}
        onChange={handleDropdownChange}
      >
        <option value="">-- Select a List --</option>
        {lists.map((list) => (
          <option key={list.id} value={list.id}>
            {list.title}
          </option>
        ))}
      </select>

      {selectedList && (
        <div className={styles.fileuploadcontainer}>
          <h2>Upload File (.xlsx or .csv)</h2>
          <input
            type="file"
            accept=".xlsx, .csv"
            onChange={handleFileUpload}
            className="file-input"
          />
        </div>
      )}

      {fileData.length > 0 && (
        <div className={styles.tablecontainer}>
          <h2>Uploaded File Data</h2>
          <table className={styles.datatable}>
            <thead>
              <tr>
                {fileData[0].map((header: string, index: number) => (
                  <th key={index}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {fileData.slice(1).map((row: any[], rowIndex: number) => (
                <tr key={rowIndex}>
                  {row.map((cell: any, cellIndex: number) => (
                    <td key={cellIndex}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {showPopup && (
        <div className={styles.popup}>
          <div className={styles.popupContent}>
            <h2>Validation Error</h2>
            <p>The uploaded file's columns do not match the selected list's columns. Please check and try again.</p>
            <button onClick={closePopup}>Close</button>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExistingList;
