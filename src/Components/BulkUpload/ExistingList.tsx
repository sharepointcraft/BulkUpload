import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/fields";
import styles from "../../Components/BulkUpload/ExistingList.module.scss";
import * as XLSX from "xlsx";


const ExistingList: React.FC = () => {
  const [lists, setLists] = useState<any[]>([]);
  const [selectedList, setSelectedList] = useState<string>("");
  const [listColumns, setListColumns] = useState<string[]>([]);
  const [fileData, setFileData] = useState<any[]>([]);
  const [attachments, setAttachments] = useState<{ [index: number]: File | null }>({});
  const [showPopup, setShowPopup] = useState<boolean>(false);

  useEffect(() => {
    sp.web.lists
      .getByTitle("BulkUpload_Central_List")
      .items.select("List_Name")
      .get()
      .then((items) => {
        const uniqueLists = items.map((item) => ({
          id: item.List_Name, // Assuming List_Name is unique
          title: item.List_Name,
        }));
  
        setLists(uniqueLists);
      })
      .catch((error) => {
        console.error("Error fetching lists:", error);
      });
  }, []);
  

  useEffect(() => {
    if (selectedList) {
      sp.web.lists
        .getById(selectedList)
        .fields.filter("Hidden eq false and ReadOnlyField eq false")
        .get()
        .then((fields) => {
          const columnNames = fields.map((field) => field.Title);
          setListColumns(columnNames);
        });
    } else {
      setListColumns([]);
    }
  }, [selectedList]);

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const handleDropdownChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setSelectedList(event.target.value);
    setFileData([]);
    setAttachments({});
    
    // Reset file input
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };
  
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    
    const file = event.target.files?.[0];
    setFileData([]);
    if (file) {
      const reader = new FileReader();

      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json<any>(worksheet, { header: 1 });

        const headers = jsonData[0] as string[];
        const isValid = headers.every((header) => listColumns.includes(header));

        if (!isValid) {
          setShowPopup(true);
          setFileData([]);
        } else {
          setShowPopup(false);
          const formattedData = jsonData.slice(1).map((row) => {
            const rowObject: any = {};
            headers.forEach((header, index) => {
              rowObject[header] = row[index];
            });
            return rowObject;
          });
          setFileData(formattedData);
        }
      };

      reader.readAsArrayBuffer(file);
    }
  };

  const handleAttachmentUpload = (rowIndex: number) => {
    const attachmentInput = document.createElement("input");
    attachmentInput.type = "file";
    attachmentInput.accept = ".xlsx, .csv, .pdf, .docx";
    attachmentInput.onchange = (event) => {
      const file = (event.target as HTMLInputElement).files?.[0];
      if (file) {
        setAttachments((prevAttachments) => ({
          ...prevAttachments,
          [rowIndex]: file,
        }));
      }
    };
    attachmentInput.click();
  };

  const handleAttachmentCancel = (rowIndex: number) => {
    setAttachments((prevAttachments) => ({
      ...prevAttachments,
      [rowIndex]: null,
    }));
  };

  const closePopup = () => {
    setShowPopup(false);
  };

  const handleSubmit = async () => {
    if (!selectedList || fileData.length === 0) {
      console.error("No list selected or no file data available.");
      alert("Please select a list and upload valid data before submitting.");
      return;
    }
  
    try {
      for (const row of fileData) {
        const itemId = row.ID; // Assuming the Excel file contains an 'ID' column
        
        if (itemId) {
          try {
            // Try updating the existing item by ID
            console.log(`Updating item with ID: ${itemId}`);
            await sp.web.lists.getById(selectedList).items.getById(itemId).update(row);
          } catch (updateError) {
            console.warn(`Item with ID ${itemId} not found. Adding new item.`);
            // If the ID doesn't exist, add the item as new
            await sp.web.lists.getById(selectedList).items.add(row);
          }
        } else {
          // Add new item if no ID is present
          console.log("Adding new item:", row);
          await sp.web.lists.getById(selectedList).items.add(row);
        }
      }
  
      alert("Data submitted successfully!");
    } catch (error) {
      console.error("Error updating the list:", error);
      alert("An error occurred while updating the list. Please try again.");
    }
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
          <h2>Upload File</h2>
          <input
            type="file"
            accept=".xlsx, .csv"
            onChange={handleFileUpload}
            ref={fileInputRef} // Add ref here
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
                <th>Attachment</th>
                {Object.keys(fileData[0]).map((header, index) => (
                  <th key={index}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {fileData.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  <td>
                    {attachments[rowIndex] ? (
                      <div className={styles.attachmentInfo}>
                        <span>{attachments[rowIndex]?.name}</span>
                        <button
                          className={styles.cancelButton}
                          onClick={() => handleAttachmentCancel(rowIndex)}
                        >
                          Cancel
                        </button>
                      </div>
                    ) : (
                      <button
                        className={styles.attachmentButton}
                        onClick={() => handleAttachmentUpload(rowIndex)}
                      >
                        Attach
                      </button>
                    )}
                  </td>
                  {Object.keys(row).map((key, cellIndex) => (
                    <td key={cellIndex}>{row[key]}</td>
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
            <p>
              The uploaded file's columns do not match the selected list's
              columns. Please check and try again.
            </p>
            <button onClick={closePopup}>Close</button>
          </div>
        </div>
      )}

      {selectedList && fileData.length > 0 && (
        <button
          className={styles.submitButton}
          onClick={handleSubmit}
        >
          Submit
        </button>
      )}
    </div>
  );
};

export default ExistingList;
