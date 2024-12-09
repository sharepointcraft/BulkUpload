  import * as React from "react";
  import * as XLSX from "xlsx";
  import styles from "../../Components/BulkUpload/NewList.module.scss";

  const NewList: React.FC = () => {
    const [tableData, setTableData] = React.useState<string[][]>([]);
    const [tableHeaders, setTableHeaders] = React.useState<string[]>([]);
    const [columnTypes, setColumnTypes] = React.useState<string[]>([]);
    const [uniqueId, setUniqueId] = React.useState<string | null>(null);

    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
          const data = e.target?.result;
          if (typeof data === "string" || data instanceof ArrayBuffer) {
            const workbook = XLSX.read(data, { type: "binary" });
            const sheetName = workbook.SheetNames[0];
            const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
              header: 1,
            }) as string[][];

            const [headers, ...rows] = sheetData;
            setTableHeaders(headers as string[]);
            setTableData(rows);
            
            // Initialize column types (set to 'Single line of text' by default)
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
      setUniqueId(tableHeaders[index]);  // Set the selected unique column
    };

    return (
      <div className={styles.mainBox}>
        <h2>Upload File</h2>
        <input
          type="file"
          accept=".xlsx, .xls, .csv"
          onChange={handleFileUpload}
        />
        {tableData.length > 0 && (
          <div className={styles.tableContainer}>
            {/* Table for displaying column names and types */}
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
                          onChange={(e) => handleColumnTypeChange(index, e.target.value)}
                        >
                          <option value="Single line of text">Single line of text</option>
                          <option value="Multiple Line of text">Multiple Line of text</option>
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
            </div>

            {/* Main table */}
            <table className={styles.tableHeader}>
              <thead>
                <tr>
                  {tableHeaders.map((header, index) => (
                    <th key={index}>{header}</th>
                  ))}
                </tr>
              </thead>
            </table>
            <div className={styles.tableBodyWrapper}>
              <table className={styles.table}>
                <tbody>
                  {tableData.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell: any, cellIndex: number) => (
                        <td key={cellIndex}>{cell}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    );
  };

  export default NewList;
