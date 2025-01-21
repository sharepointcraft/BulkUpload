import * as React from "react";
import * as XLSX from "xlsx";
import { Link } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from "../../Components/BulkUpload/NewList.module.scss";
import SuccessPopUp from "../SuccessPopUp/SuccessPopUp";
//import TableSection from "../TableSection/TableSection"; // Adjust the import path
import { validateColumns, createSharePointList, createDocumentLibrary, addDataToList } from "../Back_SubmitButton/Functions";
import { useNavigate } from "react-router-dom";
import {
  Dialog,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
} from "@fluentui/react";
import ErrorPopup from "../ErrorComponent/ErrorPopup";
import BackSubmitButtons from "../Back_SubmitButton/BackSubmitButton";

interface INewListProps {
  context: WebPartContext;
}

const NewList: React.FC<INewListProps> = ({ context }) => {
  const [tableData, setTableData] = React.useState<string[][]>([]);
  const [tableHeaders, setTableHeaders] = React.useState<string[]>([]);
  const [columnTypes, setColumnTypes] = React.useState<string[]>([]);
  const [uniqueId, setUniqueId] = React.useState<string | null>(null);
  const [selectedColumnIndex, setSelectedColumnIndex] = React.useState<number | null>(null);
  const [listName, setListName] = React.useState<string>("");
  const [showTable, setShowTable] = React.useState(true); // State to control table visibility
  const [showButtons, setShowButtons] = React.useState(false); // State to control back and validate/submit visibility
  const [showSuccessPopup, setShowSuccessPopup] = React.useState(false); // State to control progress popup
  const [popupMessage, setPopupMessage] = React.useState(""); // State to control message in progress popup
  const [isDialogVisible, setIsDialogVisible] = React.useState(false); // State to control confirmation popup
  const [showSuccessIcon, setShowSuccessIcon] = React.useState(true);
  const [isPopupOpen, setIsPopupOpen] = React.useState(false);
  const [errorPopupMessage, setErrorPopupMessage] = React.useState("");
  const [createDocLib, setCreateDocLib] = React.useState("no");
  const [recordFiles, setRecordFiles] = React.useState<Record<string, File>>({}); // Object to store files for each record
  console.log(recordFiles);
  console.log(`Selected column index: ${selectedColumnIndex}`);


  const siteUrl =
    context.pageContext.web.absoluteUrl ||
    "https://realitycraftprivatelimited.sharepoint.com/sites/BulkUpload";


  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    setUniqueId(""); // Reset the unique ID
    setSelectedColumnIndex(null);
    setShowButtons(true);

    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target?.result;
        if (typeof data === "string" || data instanceof ArrayBuffer) {
          const workbook = XLSX.read(data, { type: "binary" });
          const sheetName = workbook.SheetNames[0];
          const sheetData = XLSX.utils.sheet_to_json(
            workbook.Sheets[sheetName],
            { header: 1 }
          ) as string[][];

          const [headers, ...rows] = sheetData;

          // Format rows (removed date detection logic)
          const formattedRows = rows.map(
            (row) => row.map((cell) => cell) // No date conversion here
          );

          // Set headers and data
          setTableHeaders(headers as string[]);
          setTableData(formattedRows);

          // Infer column types based on data
          const inferredTypes = headers.map((header, colIndex) => {
            // Check all rows for the given column to determine the column type
            const firstRow = formattedRows.map((row) => row[colIndex]);

            // If any cell in this column exceeds 255 characters, it's "Multiple Line of text"
            const isMultipleLineText = firstRow.some(
              (cell) => typeof cell === "string" && cell.length > 255
            );

            if (isMultipleLineText) {
              return "Multiple Line of text"; // If any cell has more than 255 characters
            } else if (firstRow.some((cell) => !isNaN(Number(cell)))) {
              return "Number"; // If the first row value is numeric
            } else {
              return "Single line of text"; // Default to text if none of the conditions match
            }
          });

          setColumnTypes(inferredTypes); // Set the inferred column types
        }
      };
      reader.readAsBinaryString(file);
    }
  };

  const handleColumnTypeChange = (index: number, type: string) => {
    const newColumnTypes = [...columnTypes];
    newColumnTypes[index] = type;
    setColumnTypes(newColumnTypes);
  };

  const handleUniqueIdChange = (index: number) => {
    setSelectedColumnIndex(index); // Track the selected radio button column index
    const selectedColumnValues = tableData.map(row => row[index]); // Get all values in the selected column
    setUniqueId(JSON.stringify(selectedColumnValues)); // Store column values instead of column name
    console.log(`Selected column values: ${selectedColumnValues}`);

  };

  const getRequestDigest = async (): Promise<string> => {
    const response = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  };

  // Method for Confirmation popup Yes button
  const handleDialogYes = () => {
    setShowTable(!showTable);
    setIsDialogVisible(false);
  };

  // Method for Confirmation popup No button
  const handleDialogNo = () => {
    setIsDialogVisible(false);
  };

  // Handle file selection and set the file in state
  const handleFileSelection = (selectedFile: File, record: string) => {
    console.log(`Selected file:`, selectedFile);
    setRecordFiles((prev) => ({ ...prev, [record]: selectedFile })); // Store file for each record
  };


  const navigate = useNavigate(); // React Router hook for navigation

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const resetForm = () => {
    // Reset file input
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }

    // Reset state variables
    setTableData([]);
    setTableHeaders([]);
    setColumnTypes([]);
    setUniqueId(null);
    setListName("");
    setShowButtons(false);
    setShowTable(true);

    // Navigate to the upload page
    navigate("/newlist");
  };

  return (
    <div className={styles.mainBox}>
      {/* Success Popup */}
      {showSuccessPopup && (
        <SuccessPopUp
          showSuccessPopup={showSuccessPopup}
          showSuccessIcon={showSuccessIcon}
          popupMessage={popupMessage}
          setShowSuccessPopup={setShowSuccessPopup}
          setShowSuccessIcon={setShowSuccessIcon}
          resetForm={resetForm}
        />
      )}

      {/* Error Popup */}
      {isPopupOpen && (
        <ErrorPopup
          isOpen={isPopupOpen}
          message={errorPopupMessage}
          onClose={() => setIsPopupOpen(false)} // Close the popup
        />
      )}

      {/* Confirmation popup after validation */}
      <Dialog
        hidden={!isDialogVisible}
        onDismiss={handleDialogNo}
        dialogContentProps={{
          title: "Confirmation",
          subText: "Ready to proceed with validation?",
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={handleDialogYes} text="Yes" />
          <DefaultButton onClick={handleDialogNo} text="No" />
        </DialogFooter>
      </Dialog>

       {/* Home Button */}
       <div className={`${styles.homeBtn}`}>
        <button>
          <Link to="/">
            {" "}
            <img
              src={require("../../../src/webparts/bulkUpload/assets/Homeicon.png")}
              alt="Bulk-Upload-home-icon Image"
            />
          </Link>
        </button>
        {/* Title Of the Page */}
        <div className={`${styles.InnerBox}`}>
          <h1>Bulk Upload</h1>
        </div>
        <button onClick={resetForm}>
          <img
            src={require("../../../src/webparts/bulkUpload/assets/circular.png")}
            alt="Bulk-Upload-Reset-home-icon Image"
          />
        </button>
      </div>
        <div className={`${styles.homeInner}`}>
      {!showTable ? (
        <div className={styles["form-group"]}>
          <label htmlFor="listName">List Name:</label>
          <input
            type="text"
            id="listName"
            placeholder="Enter list name"
            value={listName}
            onChange={(e) => setListName(e.target.value)}
          />
          <div className={styles["radio-group"]}>
            <label>Would you like to attach a file to the respected data?</label>
            <div className={`${styles.yesNobtn}`}>
            <div>
              <input
                type="radio"
                id="createDocLibYes"
                name="createDocLib"
                value="yes"
                checked={createDocLib === "yes"}
                onChange={(e) => setCreateDocLib(e.target.value)}
              />
              <label htmlFor="createDocLibYes">Yes</label>
            </div>
            <div>
              <input
                type="radio"
                id="createDocLibNo"
                name="createDocLib"
                value="no"
                checked={createDocLib === "no"}
                onChange={(e) => setCreateDocLib(e.target.value)}
              />
              <label htmlFor="createDocLibNo">No</label>
            </div>
            </div>
          </div>
        </div>
      ) : (
        <div className={styles["form-group"]}>
          <label htmlFor="fileUpload">Upload File:</label>
          <input
            type="file"
            id="fileUpload"
            accept=".xlsx, .xls, .csv"
            onChange={handleFileUpload}
            ref={fileInputRef} // Attach the ref here
          />
          <i
            className={styles.uploadinfo}
            data-tooltip="Use single-cell column titles, include only one sheet, and ensure all columns have names."
          >
            i
          </i>
        </div>
      )}

      {/* Table Display */}
      {tableData.length > 0 && (
        <div className={styles.tableContainer}>
          <div className={styles.verticalTableWrapper}>
            {showTable ? (
              <table className={styles.verticalTable}>
                <thead>
                  <tr>
                    <th className={styles.uniqueID}>
                      Unique ID
                      <i
                        className={`${styles.infoIconID}`}
                        data-tooltip="Select a column with no duplicate or repeated values as the unique ID"
                      >
                        i
                      </i>
                    </th>
                    <th>Column Names</th>
                    <th className={styles.columnType}>
                      Column Type
                      <i
                        className={`${styles.infoIconCT}`}
                        data-tooltip="Specify the type of data for this column (e.g., text, number, date)."
                      >
                        i
                      </i>
                    </th>
                    <th>Sample Data 1</th>
                    <th>Sample Data 2</th>
                  </tr>
                </thead>

                <tbody>
                  {tableHeaders.map((header, index) => (
                    <tr key={index}>
                      <td className={`${styles.radioCenter}`}>
                        <input
                          type="radio"
                          name="uniqueId"
                          checked={selectedColumnIndex === index} // Keep track of the selected column
                          onChange={() => handleUniqueIdChange(index)}
                        />
                      </td>
                      <td>{header}</td>
                      <td>
                        <select
                          value={columnTypes[index]}
                          onChange={(e) => handleColumnTypeChange(index, e.target.value)}
                        >
                          <option value="Single line of text">Single line of text</option>
                          <option value="Multiple Line of text">Multiple Line of text</option>
                          <option value="Number">Number</option>
                          <option value="Currency">Currency</option>
                          <option value="DateTime">Date</option>
                        </select>
                        {columnTypes[index] === 'Single line of text' && (
                          <div className={`${styles.infomessage}`}>255 characters limit</div>
                        )}
                        {columnTypes[index] === 'Multiple Line of text' && (
                          <div className={`${styles.infomessage}`}>Multiple lines allowed.</div>
                        )}
                        {columnTypes[index] === 'Number' && (
                          <div className={`${styles.infomessage}`}>Enter a number (no symbols).</div>
                        )}
                        {columnTypes[index] === 'DateTime' && (
                          <div className={`${styles.infomessage}`}>Select a date (MM/DD/YYYY).</div>
                        )}
                        {columnTypes[index] === 'Currency' && (
                          <div className={`${styles.infomessage}`}>Enter a currency value.</div>
                        )}
                      </td>
                      <td>{tableData[0]?.[index] || ''}</td>
                      <td>{tableData[1]?.[index] || ''}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <table className={styles.dataTable}>
                <thead>
                  <tr>
                    {tableHeaders.map((header, index) => (
                      <th key={index}>{header}</th>
                    ))}
                    {createDocLib === 'yes' && <th>Attachment</th>}
                  </tr>
                </thead>
                <tbody>
                  {tableData.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, cellIndex) => (
                        <td key={cellIndex}>{cell}</td>
                      ))}
                      {createDocLib === 'yes' && (
                        <td>
                          <input
                            type="file"
                            onChange={(e) => {
                              const file = e.target.files?.[0]; // Safely get the first file
                              if (file && selectedColumnIndex !== null && selectedColumnIndex !== undefined) {
                                handleFileSelection(file, row[selectedColumnIndex]); // Ensure row[0] is a unique identifier
                              }
                            }}
                          />
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}
  </div>
      {/* Back and Submit Buttons */}
      {showButtons && (
        <BackSubmitButtons
          showButtons={showButtons}
          showTable={showTable}
          validateColumns={() => validateColumns(tableHeaders, tableData, columnTypes, selectedColumnIndex,setErrorPopupMessage, setIsPopupOpen)}
          createSharePointList={() => createSharePointList(siteUrl, listName, tableHeaders, columnTypes, setErrorPopupMessage, setIsPopupOpen, getRequestDigest)}
          createDocumentLibrary={() => {
            const selectedColumnValues = uniqueId ? JSON.parse(uniqueId) : [];
            return createDocumentLibrary(getRequestDigest, listName, siteUrl, selectedColumnValues, recordFiles)
          }}
          addDataToList={() => addDataToList(tableData, tableHeaders, columnTypes, siteUrl, listName, getRequestDigest)}
          createDocLib={createDocLib}
          setIsDialogVisible={setIsDialogVisible}
          setPopupMessage={setPopupMessage}
          setShowSuccessPopup={setShowSuccessPopup}
          setErrorPopupMessage={setErrorPopupMessage}
          setIsPopupOpen={setIsPopupOpen}
          setShowTable={setShowTable}
          setShowSuccessIcon={setShowSuccessIcon}
        />
      )}
    </div>

  );
};
export default NewList;
