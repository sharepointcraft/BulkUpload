import * as React from "react";
import * as XLSX from "xlsx";
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

  const siteUrl =
    context.pageContext.web.absoluteUrl ||
    "https://realitycraftprivatelimited.sharepoint.com/sites/DevJay";

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
          setTableHeaders(headers as string[]); // Set headers
          setTableData(rows); // Set rows

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
  // new createhsharepoint list with the number validation of the data
  const createSharePointList = async (): Promise<boolean> => {
    if (!listName || !uniqueId) {
      alert("Please provide a list name and select a unique ID.");
      return false; // Return false to indicate failure
    }
  
    // Validate data for Number column types
    const invalidCells: { row: number; col: string }[] = [];
    tableData.forEach((row, rowIndex) => {
      columnTypes.forEach((type, colIndex) => {
        if (type === "Number" && isNaN(Number(row[colIndex]))) {
          invalidCells.push({ row: rowIndex + 1, col: tableHeaders[colIndex] });
        }
      });
    });
  
    if (invalidCells.length > 0) {
      alert(
        `Invalid data found in the following cells:\n${invalidCells
          .map((cell) => `Row ${cell.row}, Column ${cell.col}`)
          .join("\n")}`
      );
      return false; // Return false if data is invalid
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
        await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": requestDigest,
          },
          body: JSON.stringify({
            __metadata: { type: "SP.Field" },
            Title: tableHeaders[i],
            FieldTypeKind: columnTypes[i] === "Number" ? 9 : 2, // Adjust field types
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
      alert("Error creating the SharePoint list.");
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
  return (
    <div className={styles.mainBox}>
      <h2>Create SharePoint List</h2>
      <input
        type="text"
        placeholder="Enter list name"
        value={listName}
        onChange={(e) => setListName(e.target.value)}
      />
      <input
        type="file"
        accept=".xlsx, .xls, .csv"
        onChange={handleFileUpload}
      />
      {tableData.length > 0 && (
        <div className={styles.tableContainer}>
          <div className={styles.verticalTableWrapper}>
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
            </table>
            <table className={styles.dataTable}>
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
            </table>
          </div>
          <button
            onClick={async () => {
              const listCreationSuccess = await createSharePointList(); // Capture if list creation is successful
              if (listCreationSuccess) {
                await createDocumentLibrary(); // Create library only if list creation is successful
              }
            }}
          >
            Create List and Library
          </button>
        </div>
      )}
    </div>
  );
};

export default NewList;
