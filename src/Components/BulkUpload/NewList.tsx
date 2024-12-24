import * as React from "react";
import * as XLSX from "xlsx";
import { Link } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from "../../Components/BulkUpload/NewList.module.scss";

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
  const [showButtons, setShowButtons] = React.useState(false);
  const [showSuccessPopup, setShowSuccessPopup] = React.useState(false);
  //const [listCreationSuccess, setListCreationSuccess] = React.useState(false);
  //const [documentLibraryCreationSuccess, setDocumentLibraryCreationSuccess] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const [popupMessage, setPopupMessage] = React.useState("");

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
      const message = `Invalid data found in the following cells:\n${invalidCells
        .map((cell) => `Row ${cell.row}, Column ${cell.col}: ${cell.issue}`)
        .join("\n")}`;
      alert(message);
      return false; // Return false if data is invalid
    }

    return true; // Return true if all data is valid
  };
  const createSharePointList = async (): Promise<boolean> => {
    if (!listName || !uniqueId) {
      alert("Please provide a list name and select a unique ID.");
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
      alert("Error creating the SharePoint list or adding data.");
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
      alert("Error creating the document library.");
      return false;
    }
  };
  const addDatatoList = async (): Promise<boolean> => {
    const requestDigest = await getRequestDigest(); // Fetch digest dynamically
    let allDataAddedSuccessfully = true; // Flag to track overall success

    const parseCurrency = (value: string | number): number => {
      if (typeof value === "string") {
        return Number(value.replace(/[^0-9.]/g, "")); // Remove '$' and other non-numeric characters
      }
      return Number(value);
    };

    // Now add data to the list
    for (const row of tableData) {
      const itemPayload: Record<string, any> = {};

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
        return true;
      } catch (error) {
        console.error("Failed to add data to list:", error);
        allDataAddedSuccessfully = false; // Mark failure
        return false;
      }
    }

    // Return true if all items were added successfully, otherwise return false
    return allDataAddedSuccessfully;
  };
  return (
    <div className={styles.mainBox}>
      {/* Success Popup */}
      {showSuccessPopup && (
        <div className={`${styles.successPopup}`}>
          <div className={`${styles.popupContent}`}>
            <div className={`${styles.progressBar}`}>
              <div
                className={`${styles.progress}`}
                style={{ width: `${progress}%` }}
              ></div>
            </div>
            <p>{popupMessage}</p>
          </div>
        </div>
      )}

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

      {/* Form Group for list name */}
      <div className={styles["form-group"]}>
        <label htmlFor="listName">List Name:</label>
        <input
          type="text"
          id="listName"
          placeholder="Enter list name"
          value={listName}
          onChange={(e) => setListName(e.target.value)}
        />
        <i className={styles.listinfo} data-tooltip="Enter a valid list name (letters, numbers, and no special characters).">i</i>
      </div>

      {/* Form Group for file upload */}
      {showTable && (
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
                    {tableHeaders.concat("Attachment").map((header, index) => (
                      <th key={index}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {tableData.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, cellIndex) => (
                        <td key={cellIndex}>{cell}</td>
                      ))}
                      <td key="attachment">
                        <input type="file" />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* Back and Submit Buttons */}
      {showButtons && (
        <div className={`${styles.backSubmitbtn}`}>
          <div className={`${styles.homeBtn}`}>
            <button>
              <Link to="/selectlisttype">Back</Link>
            </button>
          </div>
          <div className={`${styles.homeBtn}`}>
            {showTable ? (
              <button
                onClick={async () => {
                  const isValid = await validateColumns();
                  if (isValid) {
                    setShowTable(!showTable);
                  } else {
                    alert("Validation failed. Please correct the data.");
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
                    setPopupMessage("Creating SharePoint list...");
                    setProgress(33.33);
                    setShowSuccessPopup(true);

                    const isListCreationSuccess = await createSharePointList();
                    if (!isListCreationSuccess) {
                      setPopupMessage("Failed to create SharePoint list.");
                      await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds
                      setShowSuccessPopup(false);
                      return; // Stop the process
                    }
                    await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds

                    // Step 2: Create Document Library
                    setPopupMessage("Creating document library...");
                    setProgress(66.66);

                    const isLibraryCreationSuccess =
                      await createDocumentLibrary();
                    if (!isLibraryCreationSuccess) {
                      setPopupMessage("Failed to create document library.");
                      await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds
                      setShowSuccessPopup(false);
                      return; // Stop the process
                    }
                    await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds

                    // Step 3: Add Data to List
                    setPopupMessage("Submitting data...");
                    setProgress(100);

                    const isDataSubmissionSuccess = await addDatatoList();
                    if (!isDataSubmissionSuccess) {
                      setPopupMessage("Failed to submit data.");
                      await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds
                      setShowSuccessPopup(false);
                      return; // Stop the process
                    }
                    await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds

                    setPopupMessage("Data successfully submitted.");
                    await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds

                    // Hide popup after completion
                    setShowSuccessPopup(false);
                    setShowTable(false);
                  } catch (error) {
                    setPopupMessage(
                      error.message || "An unexpected error occurred."
                    );
                    await new Promise((resolve) => setTimeout(resolve, 3000)); // Show for 3 seconds
                    setShowSuccessPopup(false);
                  }
                }}
              >
                Submit
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
export default NewList;
