// TableSection.tsx
import * as React from 'react';
import styles from './TableSection.module.scss';

interface TableSectionProps {
  tableData: string[][];
  tableHeaders: string[];
  columnTypes: string[];
  uniqueId: string | null;
  createDocLib: string;
  showTable: boolean;
  handleUniqueIdChange: (index: number) => void;
  handleColumnTypeChange: (index: number, value: string) => void;
}

const TableSection: React.FC<TableSectionProps> = ({
  tableData,
  tableHeaders,
  columnTypes,
  uniqueId,
  createDocLib,
  showTable,
  handleUniqueIdChange,
  handleColumnTypeChange,
}) => {
  // If tableData is empty, return null (no content)
  if (tableData.length === 0) return null;

  return (
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
                      checked={uniqueId === header}
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
  );
};

export default TableSection;
