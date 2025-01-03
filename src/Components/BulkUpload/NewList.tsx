import * as React from "react";
import * as XLSX from "xlsx";
import { Link } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from "../../Components/BulkUpload/NewList.module.scss";
import {
  Dialog,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
} from "@fluentui/react";
import ErrorPopup from "../ErrorComponent/ErrorPopup";

interface INewListProps {
  context: WebPartContext;
}

const NewList: React.FC<INewListProps> = ({ context }) => {
  const [tableData, setTableData] = React.useState<string[][]>([]);
  const [tableHeaders, setTableHeaders] = React.useState<string[]>([]);
  const [columnTypes, setColumnTypes] = React.useState<string[]>([]);
  const [uniqueId, setUniqueId] = React.useState<string | null>(null);
  const [listName, setListName] = React.useState<string>("");
  const [showTable, setShowTable] = React.useState(true); // State to control table visibility
  const [showButtons, setShowButtons] = React.useState(false); // State to control back and validate/submit visibility
  const [showSuccessPopup, setShowSuccessPopup] = React.useState(false); // State to control progress popup
  const [progress, setProgress] = React.useState(0); // State to control width of progesss bar
  const [popupMessage, setPopupMessage] = React.useState(''); // State to control message in progress popup
  const [isDialogVisible, setIsDialogVisible] = React.useState(false); // State to control confirmation popup
  const [showSuccessIcon, setShowSuccessIcon] = React.useState(true);
  const [isPopupOpen, setIsPopupOpen] = React.useState(false);
  const [errorPopupMessage, setErrorPopupMessage] = React.useState('');
  const [createDocLib, setCreateDocLib] = React.useState("no");


  const siteUrl =
    context.pageContext.web.absoluteUrl ||
    "https://realitycraftprivatelimited.sharepoint.com/sites/BulkUpload";


  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    setUniqueId(""); // Reset the unique ID
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
    setUniqueId(tableHeaders[index]);
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

  const validateColumns = async () => {
    const invalidCells: { row: number; col: string; issue: string }[] = [];

    // Check if column headers contain special characters
    const specialCharPattern = /[^a-zA-Z0-9_ ]/; // Allow letters, numbers, underscores, and spaces
    tableHeaders.forEach((header, index) => {
      if (specialCharPattern.test(header)) {
        invalidCells.push({
          row: 0, // Header row
          col: header,
          issue: "Contains special characters",
        });
      }
    });

    tableData.forEach((row, rowIndex) => {
      columnTypes.forEach((type, colIndex) => {
        let cellValue = row[colIndex];

        if (type === "Number" && isNaN(Number(cellValue))) {
          // Collect invalid cells for Number columns
          invalidCells.push({
            row: rowIndex + 1,
            col: tableHeaders[colIndex],
            issue: "Expected a number",
          });
        } else if (type === "Single line of text") {
          if (typeof cellValue === "number") {
            // Convert numbers to strings for text columns
            row[colIndex] = String(cellValue); // Convert number to string
          }

          // Check character length for Single line of text
          if (cellValue.length > 255) {
            invalidCells.push({
              row: rowIndex + 1,
              col: tableHeaders[colIndex],
              issue: "Exceeded 255 character limit",
            });
          }
        }
      });
    });

    if (invalidCells.length > 0) {
      const message = `
                        <strong>Invalid data found in the following cells:</strong><br/>
                      
                         ${invalidCells.slice(0, 5)
          .map((cell) => `<li>Row ${cell.row}, Column ${cell.col}: ${cell.issue}</li>`)
          .join('')}
                        ${invalidCells.length > 5 ? '<button id="moreErrorsButton">More Errors</button>' : ''}
                      `;

      // alert(message);
      setErrorPopupMessage(message);
      setIsPopupOpen(true)

      // Generate the message for the new tab
      const allErrorsMessage = `
                                <strong>All Errors:</strong><br/>
                                <ul>
                                ${invalidCells
          .map((cell) => `<li>Row ${cell.row}, Column ${cell.col}: ${cell.issue}</li>`)
          .join('')}
                                </ul>
                              `;

      // Add an event listener for the "More Errors" button
      setTimeout(() => {
        const moreErrorsButton = document.getElementById('moreErrorsButton');
        if (moreErrorsButton) {
          moreErrorsButton.addEventListener('click', () => {
            const newTab = window.open('', '_blank');
            if (newTab) {
              newTab.document.write(`
          <html>
            <head>
              <title>All Errors</title>
            </head>
            <body>
              ${allErrorsMessage}
            </body>
          </html>
        `);
              newTab.document.close();
            }
          });
        }
      }, 0);

      return false; // Return false if data is invalid
    }


    return true; // Return true if all data is valid
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


  const createSharePointList = async (): Promise<boolean> => {
    if (!listName || !uniqueId) {
      //alert("Please provide a list name and select a unique ID.");
      setErrorPopupMessage('Please provide a list name and select a unique ID.');
      setIsPopupOpen(true);
      return false; // Return false to indicate failure
    }

    try {
      const requestDigest = await getRequestDigest(); // Fetch digest dynamically

      // Create the list
      const listPayload = {
        __metadata: { type: "SP.List" },
        Title: listName,
        BaseTemplate: 100, // Custom List
      };

      const response = await fetch(`${siteUrl}/_api/web/lists`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": requestDigest,
        },
        body: JSON.stringify(listPayload),
      });

      if (!response.ok) {
        throw new Error(`Error creating list: ${response.statusText}`);
      }

      // Add columns to the list
      for (let i = 0; i < tableHeaders.length; i++) {
        const fieldType = columnTypes[i];
        let metadataType = "SP.Field"; // Default metadata type
        let fieldTypeKind = 2; // Default to Single Line of Text
        let additionalProperties = {}; // Default additional properties

        switch (fieldType) {
          case "DateTime":
            fieldTypeKind = 4; // DateTime Field
            metadataType = "SP.FieldDateTime"; // Use the correct type for DateTime
            additionalProperties = {
              DisplayFormat: 0, // 0 for DateOnly
            };
            break;
          case "Number":
            fieldTypeKind = 9; // Number Field
            break;
          case "Currency":
            fieldTypeKind = 10; // Currency Field
            break;
          case "Single line of text":
          default:
            fieldTypeKind = 2; // Single Line of Text
        }
        await fetch(
          `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": requestDigest,
            },
            body: JSON.stringify({
              __metadata: { type: metadataType },
              Title: tableHeaders[i],
              FieldTypeKind: fieldTypeKind,
              ...additionalProperties,
            }),
          }
        );

        // Add the column to the default view
        await fetch(
          `${siteUrl}/_api/web/lists/getbytitle('${listName}')/defaultview/viewfields/addviewfield('${tableHeaders[i]}')`,
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": requestDigest,
            },
          }
        );
      }

      //alert("List created successfully!");
      return true; // Return true to indicate success
    } catch (error) {
      console.error(error);
      //alert("Error creating the SharePoint list or adding data.");
      return false; // Return false if an error occurred
    }
  };

  const createDocumentLibrary = async (): Promise<boolean> => {
    try {
      const requestDigest = await getRequestDigest(); // Fetch digest dynamically

      const libraryPayload = {
        __metadata: { type: "SP.List" },
        Title: `${listName}_Documents`,
        BaseTemplate: 101, // Document Library
      };

      const response = await fetch(`${siteUrl}/_api/web/lists`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": requestDigest,
        },
        body: JSON.stringify(libraryPayload),
      });

      if (!response.ok) {
        throw new Error(`Error creating library: ${response.statusText}`);
      }

      //alert("Document library created successfully!");
      return true;
    } catch (error) {
      console.error(error);
      //alert("Error creating the document library.");
      return false;
    }
  };



  const addDataToList = async (): Promise<boolean> => {
    const requestDigest = await getRequestDigest(); // Fetch digest dynamically
    let allDataAddedSuccessfully = true; // Flag to track overall success

    const parseCurrency = (value: string | number): number => {
      if (typeof value === "string") {
        return Number(value.replace(/[^0-9.]/g, "")); // Remove '$' and other non-numeric characters
      }
      return Number(value);
    };

    // Iterate over each row of data
    for (const row of tableData) {
      const itemPayload: Record<string, any> = {};

      // Map headers and cell values to payload
      tableHeaders.forEach((header, index) => {
        const internalColumnName = header
          .replace(/\s+/g, "_x0020_") // Replace spaces with _x0020_
          .replace(/\//g, "_x002f_"); // Replace slashes with _x002f_
        const cellValue = row[index]; // Get the cell value
        if (columnTypes[index] === "Currency") {
          const numericValue = parseCurrency(cellValue);
          itemPayload[internalColumnName] = numericValue;
        } else {
          itemPayload[internalColumnName] = cellValue; // Map each header to its corresponding cell value
        }
      });

      // Try to add data to the list
      try {
        const response = await fetch(
          `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": requestDigest,
            },
            body: JSON.stringify({
              __metadata: { type: "SP.ListItem" },
              ...itemPayload,
            }),
          }
        );

        if (!response.ok) {
          throw new Error(`Error adding item to list: ${response.statusText}`);
        }
      } catch (error) {
        console.error("Failed to add data to list:", error);
        allDataAddedSuccessfully = false; // Mark failure
      }
    }

    // Return true if all items were added successfully, otherwise false
    return allDataAddedSuccessfully;
  };

  return (
    <div className={styles.mainBox}>

      {/* Success Popup */}
      {showSuccessPopup && (
        <div className={`${styles.successPopup}`}>
          <div className={`${styles.popupContent}`} style={{ borderColor: showSuccessPopup ? 'yellow' : 'green', borderWidth: '2px', borderStyle: 'solid', }}>

            {showSuccessIcon ? (
              <div className={`${styles.circularProgress}`}>
                <div className={`${styles.loadingSpinner}`}></div>
                <div className={`${styles.progressText}`}>{progress}%</div>
              </div>
            ) : (
              <span className={`${styles["success-icon"]}`}>✔</span>
            )}
            <p>{popupMessage}</p>
          </div>
        </div>
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
          subText: "Do you want to display the data table?",
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
          <Link to="/"> <img
            src={require("../../../src/webparts/bulkUpload/assets/Homeicon.png")}
            alt="Bulk-Upload-home-icon Image"
          /></Link>
        </button>
      </div>

      {/* Title Of the Page */}
      <div className={`${styles.InnerBox}`}>
        <h1>Bulk Upload</h1>
      </div>

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
            <label>Do you want to create a document library?</label>
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
      ) : (
        <div className={styles["form-group"]}>
          <label htmlFor="fileUpload">Upload File:</label>
          <input
            type="file"
            id="fileUpload"
            accept=".xlsx, .xls, .csv"
            onChange={handleFileUpload}
          />
          <i className={styles.uploadinfo} data-tooltip="Excel Columns should be in first Line">i</i>
        </div>
      )}

       
      {/* Table Display */}
      {tableData.length > 0 && (
        <div className={styles.tableContainer}>
          {/* Vertical Table */}
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

                    <th>
                      Column Names

                    </th>
                    <th className={styles.columnType}>
                      Column Type
                      <i
                        className={`${styles.infoIconCT}`}
                        data-tooltip="Specify the type of data for this column (e.g., text, number, date)."
                      >
                        i
                      </i>
                    </th>
                    <th>
                      Sample Data 1
                      {/* <i
                        className={`${styles.infoIcon}`}
                        title="An example of data for this column."
                      >
                       ℹ️
                      </i> */}
                    </th>
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
                          checked={uniqueId === header}
                          onChange={() => handleUniqueIdChange(index)}
                        />
                      </td>
                      <td>{header}</td>
                      <td>
                        <select
                          value={columnTypes[index]}
                          onChange={(e) =>
                            handleColumnTypeChange(index, e.target.value)
                          }
                        >
                          <option value="Single line of text">
                            Single line of text
                          </option>
                          <option value="Multiple Line of text">
                            Multiple Line of text
                          </option>
                          <option value="Number">Number</option>
                          <option value="Currency">Currency</option>
                          <option value="DateTime">Date</option>
                        </select>
                        {columnTypes[index] === "Single line of text" && (
                          <div className={`${styles.infomessage}`}>
                            {/* <span className={`${styles.infoicon}`}>ℹ️</span> */}
                            255 characters limit
                          </div>
                        )}
                        {columnTypes[index] === "Multiple Line of text" && (
                          <div className={`${styles.infomessage}`}>
                            {/* <span className={`${styles.infoicon}`}>ℹ️</span> */}
                            Multiple lines allowed.
                          </div>
                        )}
                        {columnTypes[index] === "Number" && (
                          <div className={`${styles.infomessage}`}>
                            {/* <span className={`${styles.infoicon}`}>ℹ️</span> */}
                            Enter a number (no symbols).
                          </div>
                        )}
                        {columnTypes[index] === "DateTime" && (
                          <div className={`${styles.infomessage}`}>
                            {/* <span className={`${styles.infoicon}`}>ℹ️</span> */}
                            Select a date (MM/DD/YYYY).
                          </div>
                        )}
                        {columnTypes[index] === "Currency" && (
                          <div className={`${styles.infomessage}`}>
                            {/* <span className={`${styles.infoicon}`}>ℹ️</span> */}
                            Enter a currency value.
                          </div>
                        )}
                      </td>
                      <td>{tableData[0]?.[index] || ""}</td>
                      <td>{tableData[1]?.[index] || ""}</td>
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
                    {/* Conditionally render the 'Attachment' header based on radio selection */}
                    {createDocLib === "yes" && <th>Attachment</th>}
                  </tr>
                </thead>
                <tbody>
                  {tableData.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, cellIndex) => (
                        <td key={cellIndex}>{cell}</td>
                      ))}
                      {/* Conditionally render the 'Attachment' column based on radio selection */}
                      {createDocLib === "yes" && (
                        <td>
                          <input type="file" />
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

      {/* Back and Submit Buttons */}
      {showButtons && <div className={`${styles.backSubmitbtn}`}>
        <div className={`${styles.backBtn}`}>
          <button>
            <Link to="/selectlisttype">Back</Link>
          </button>
        </div>
        <div className={`${styles.validateBtn}`}>
          {showTable ? (
            <button
              onClick={async () => {
                const isValid = await validateColumns();
                if (isValid) {
                  setIsDialogVisible(true);
                } else {
                  //alert("Validation failed. Please correct the data.");
                }
              }}
            >
              Validate
            </button>
          ) : (
            <button
              onClick={async () => {
                try {
                  // Step 1: Create SharePoint List
                  setPopupMessage('Creating SharePoint list...');
                  if (createDocLib === "yes") {
                    setProgress(35);
                  } else {
                    setProgress(50);
                  }

                  setShowSuccessPopup(true);

                  const isListCreationSuccess = await createSharePointList();
                  if (!isListCreationSuccess) {
                    //setProgress(0);
                    //setPopupMessage('Failed to create SharePoint list.');
                    //await new Promise((resolve) => setTimeout(resolve, 1000)); // Show for 3 seconds
                    setShowSuccessPopup(false);
                    setErrorPopupMessage('Failed to create SharePoint list.');
                    setIsPopupOpen(true);
                    return; // Stop the process
                  }
                  await new Promise((resolve) => setTimeout(resolve, 1000)); // Show for 3 seconds



                  if (createDocLib === "yes") {
                    // Step 2: Create Document Library
                    setPopupMessage('Creating document library...');
                    setProgress(70);

                    const isLibraryCreationSuccess = await createDocumentLibrary();
                    if (!isLibraryCreationSuccess) {
                      setShowSuccessPopup(false);
                      setErrorPopupMessage('Failed to create document library.');
                      setIsPopupOpen(true);
                      return; // Stop the process
                    }
                  }

                  await new Promise((resolve) => setTimeout(resolve, 1000)); // Show for 3 seconds

                  // Step 3: Add Data to List
                  setPopupMessage('Submitting data...');
                  setProgress(100);

                  const isDataSubmissionSuccess = await addDataToList();
                  if (!isDataSubmissionSuccess) {
                    setShowSuccessPopup(false);
                    setErrorPopupMessage('Failed to submit data.');
                    setIsPopupOpen(true);
                    return; // Stop the process
                  }
                  await new Promise((resolve) => setTimeout(resolve, 1000)); // Show for 3 seconds


                  setPopupMessage('Data successfully submitted.');
                  setShowSuccessIcon(false); // Show the success icon
                  await new Promise((resolve) => setTimeout(resolve, 2000)); // Show for 3 seconds

                  // Hide popup after completion
                  setShowSuccessPopup(false);
                  setShowTable(false);
                  
                } catch (error) {
                  setErrorPopupMessage(`An unexpected error occurred. ${error.message} `);
                  setIsPopupOpen(true);
                }
              }}
            >
              Submit
            </button>

          )}
        </div>
      </div>
      }
    </div>
  );
};
export default NewList;
