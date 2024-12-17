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

  const siteUrl =
    context.pageContext.web.absoluteUrl ||
    "https://realitycraftprivatelimited.sharepoint.com/sites/BulkUpload";

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {

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

          // Format the rows to handle Excel date conversion
          const formattedRows = rows.map((row) =>
            row.map((cell, index) => {
              // Check if the column type is 'Date' and the cell is a number
              if (
                typeof cell === "number" &&
                headers[index] === "Date"
              ) {
                const excelDate = new Date((cell - 25569) * 86400000); // Excel date to JavaScript date
                return excelDate.toLocaleDateString("en-US"); // Format to MM/DD/YYYY
              }
              return cell; // Keep other values as-it is
            })
          );
          setTableHeaders(headers as string[]); // Set headers

          setTableData(formattedRows); // Set rows

          setColumnTypes(Array(headers.length).fill("Single line of text"));
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

  //Data Validation 
  const validateColumns = async () => {
    const invalidCells: { row: number; col: string; issue: string }[] = [];

    tableData.forEach((row, rowIndex) => {
      columnTypes.forEach((type, colIndex) => {
        const cellValue = row[colIndex];

        if (type === "Number" && isNaN(Number(cellValue))) {
          // Collect invalid cells for Number columns
          invalidCells.push({
            row: rowIndex + 1,
            col: tableHeaders[colIndex],
            issue: "Expected a number",
          });
        } else if (type === "Single line of text" && typeof cellValue === "number") {
          // Convert numbers to strings for text columns
          row[colIndex] = String(cellValue);
        }
      });
    });

    if (invalidCells.length > 0) {
      const message = `Invalid data found in the following cells:\n${invalidCells
        .map(
          (cell) =>
            `Row ${cell.row}, Column ${cell.col}: ${cell.issue}`
        )
        .join("\n")}`;
      alert(message);
      return false; // Return false if data is invalid
    }

    return true; // Return true if all data is valid
  };



  //new createsharepoint list
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
            fieldTypeKind = 8; // Currency Field
            break;
          case "Single line of text":
          default:
            fieldTypeKind = 2; // Single Line of Text
        }
        await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
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
        });

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

      alert("List created successfully!");
      return true; // Return true to indicate success
    } catch (error) {
      console.error(error);
      alert("Error creating the SharePoint list or adding data.");
      return false; // Return false if an error occurred
    }
  };



  const createDocumentLibrary = async () => {
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

      alert("Document library created successfully!");
    } catch (error) {
      console.error(error);
      alert("Error creating the document library.");
    }
  };


  //Date Format
  /*const formatToISO = (dateValue: string): string => {
    let month, day, year;

    if (dateValue.includes("/")) {
      // Split using "/"
      [month, day, year] = dateValue.split("/");

      // Check if the format is indeed MM/DD/YYYY
      if (month.length !== 2 || day.length !== 2 || year.length !== 4) {
        throw new Error("Unsupported date format");
      }
    } else if (dateValue.includes("-")) {
      // Split using "-"
      [month, day, year] = dateValue.split("-");

      // Check if the format is indeed MM-DD-YYYY
      if (month.length !== 2 || day.length !== 2 || year.length !== 4) {
        throw new Error("Unsupported date format");
      }
    } else {
      throw new Error("Unsupported date format");
    }

    // Convert to ISO format
    const isoDate = new Date(`${year}-${month}-${day}T00:00:00Z`);
    return isoDate.toISOString();
  };*/

  //Add Data to list
  const addDatatoList = async () => {

    const requestDigest = await getRequestDigest(); // Fetch digest dynamically
    let allDataAddedSuccessfully = true; // Flag to track overall success

    // Now add data to the list
    for (const row of tableData) {
      const itemPayload: Record<string, any> = {};

      tableHeaders.forEach((header, index) => {

        const internalColumnName = header
          .replace(/\s+/g, "_x0020_")  // Replace spaces with _x0020_
          .replace(/\//g, "_x002f_");  // Replace slashes with _x002f_
        //let cellValue = row[index];

        // Convert date if necessary
        //if (header.includes("Date") && typeof cellValue === "string" && (cellValue.includes("/") || cellValue.includes("-"))) {
        //cellValue = formatToISO(cellValue); // Use your date format conversion function here
        //}
        itemPayload[internalColumnName] = row[index]; // Map each header to its corresponding cell value
      });

      try {
        const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": requestDigest,
          },
          body: JSON.stringify({
            __metadata: { type: "SP.ListItem" },
            ...itemPayload,
          }),
        });

        if (!response.ok) {
          throw new Error(`Error adding item to list: ${response.statusText}`);
        }
        else {

        }
      } catch (error) {
        console.error("Failed to add data to list:", error);
        allDataAddedSuccessfully = false; // Mark failure
      }
    }

    // Show success alert only if all items were added successfully
    if (allDataAddedSuccessfully) {
      alert("Data added successfully");
    } else {
      alert("Error adding data to the SharePoint list.");
    }
  };
  return (
    <div className={styles.mainBox}>
      <div className={`${styles.homeBtn}`}>
        {" "}
        <button>
          <Link to="/">Home</Link>
        </button>
      </div>
      <div className={`${styles.InnerBox}`}>
        <h1>Bulk Upload</h1>
      </div>
      <div className={styles["form-group"]}>
        <label htmlFor="listName">List Name:</label>
        <input
          type="text"
          id="listName"
          placeholder="Enter list name"
          value={listName}
          onChange={(e) => setListName(e.target.value)}
        />
      </div>
      {showTable && ( // Conditionally render the file upload section
        <div className={styles["form-group"]}>
          <label htmlFor="fileUpload">Upload File:</label>
          <input
            type="file"
            id="fileUpload"
            accept=".xlsx, .xls, .csv"
            onChange={handleFileUpload}
          />
        </div>
      )}
      {tableData.length > 0 && (
        <div className={styles.tableContainer}>
          <div className={styles.verticalTableWrapper}>
            {showTable ? (
              <table className={styles.verticalTable}>
                <thead>
                  <tr>
                    <th>Column Names</th>
                    <th>Column Type</th>
                    <th>Unique ID</th>
                  </tr>
                </thead>
                <tbody>
                  {tableHeaders.map((header, index) => (
                    <tr key={index}>
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
                      </td>
                      <td>
                        <input
                          type="radio"
                          name="uniqueId"
                          checked={uniqueId === header}
                          onChange={() => handleUniqueIdChange(index)}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>) :
              (<table className={styles.dataTable}>
                <thead>
                  <tr>
                    {tableHeaders.map((header, index) => (
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
                    </tr>
                  ))}
                </tbody>
              </table>)}
          </div>
          <div className={`${styles.homeBtn}`}>
            <button>
              <Link to="/selectlisttype">Back</Link>
            </button>
          </div>
          <div className={`${styles.homeBtn} ${styles.validateBtn}`}>
            {showTable ? (
              <button
                onClick={async () => {
                  const isValid = await validateColumns();
                  if (isValid) {
                    setShowTable(!showTable);
                  }
                  else {
                    alert("Validation failed. Please correct the data.");
                  }
                }}
              >
                Validate
              </button>) :

              (<button
                onClick={async () => {
                  const islistCreationSuccess = await createSharePointList();
                  if (islistCreationSuccess) {
                    await createDocumentLibrary(); // Capture if list creation is successful
                    await addDatatoList();
                  }
                  else {
                    alert("Error in creating the sharepoint list.");
                  }
                }}>
                Submit
              </button>)}

          </div>
        </div>
      )}
    </div>
  );
};

export default NewList;
