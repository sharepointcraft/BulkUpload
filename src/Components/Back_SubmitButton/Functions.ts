import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs"
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/content-types/types";


//////// METHOD TO VALIDATE THE COLUMNS ////////

export const validateColumns = async (
  tableHeaders: string[],
  tableData: any[][],
  columnTypes: string[],
  selectedColumnIndex: number | null,
  setErrorPopupMessage: (message: string) => void,
  setIsPopupOpen: (isOpen: boolean) => void
): Promise<boolean> => {
  const invalidCells: { row: number; col: string; issue: string }[] = [];

  if (selectedColumnIndex === null) {
    setErrorPopupMessage("Please select a unique ID.");
    setIsPopupOpen(true);
    return false; // Return false to indicate failure
  }

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
      ${invalidCells
        .slice(0, 5)
        .map(
          (cell) =>
            `<li>Row ${cell.row}, Column ${cell.col}: ${cell.issue}</li>`
        )
        .join("")}
      ${invalidCells.length > 5
        ? '<button id="moreErrorsButton">More Errors</button>'
        : ""
      }
    `;

    setErrorPopupMessage(message);
    setIsPopupOpen(true);

    const allErrorsMessage = `
      <strong>All Errors:</strong><br/>
      <ul>
        ${invalidCells
        .map(
          (cell) =>
            `<li>Row ${cell.row}, Column ${cell.col}: ${cell.issue}</li>`
        )
        .join("")}
      </ul>
    `;

    setTimeout(() => {
      const moreErrorsButton = document.getElementById("moreErrorsButton");
      if (moreErrorsButton) {
        moreErrorsButton.addEventListener("click", () => {
          const newTab = window.open("", "_blank");
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

//////// METHOD TO CREATE LIST IN THE SHAREPOINT ////////

export const createSharePointList = async (
  siteUrl: string,
  listName: string,
  tableHeaders: string[],
  columnTypes: string[],
  setErrorPopupMessage: (message: string) => void,
  setIsPopupOpen: (isOpen: boolean) => void,
  getRequestDigest: () => Promise<string>
): Promise<boolean> => {
  
  if (!listName) {
    setErrorPopupMessage("Please provide a list name.");
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

    return true; // Return true to indicate success
  } catch (error) {
    console.error(error);
    setErrorPopupMessage("Error creating the SharePoint list or adding data.");
    setIsPopupOpen(true);
    return false; // Return false if an error occurred
  }
};

//////// METHOD TO CREATE LIBRARY IN THE SHAREPOINT ////////


export const createDocumentLibrary = async (
  getRequestDigest: () => Promise<string>,
  listName: string,
  siteUrl: string,
  selectedColumnValues: string[], // Pass selected column values
  recordFiles: Record<string, File>
): Promise<boolean> => {
  try {
    const requestDigest = await getRequestDigest(); // Fetch digest dynamically

    // Step 1: Create the document library if it doesn't exist
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
      const errorDetails = await response.text();
      throw new Error(`Error creating library: ${response.statusText}, ${errorDetails}`);
    }

    console.log(`Library '${listName}_Documents' created successfully.`);

    // Step 2: Enable content type management
    const list = await sp.web.lists.getByTitle(`${listName}_Documents`);
    await list.update({ ContentTypesEnabled: true });

    console.log(`Content Types Management enabled for library '${listName}_Documents'.`);

    try {
      // Step 3: Activate Document Set Feature
      await sp.site.features.add("3bae86a2-776d-499d-9db8-fa4cdc7884f8", true);

      // Step 4: Add Document Set Content Type
      const lib = sp.web.lists.getByTitle(`${listName}_Documents`);
      const documentSetContentTypeId = "0x0120D520";
      console.log(` ${documentSetContentTypeId}`);

      await lib.contentTypes.addAvailableContentType(documentSetContentTypeId);

      selectedColumnValues.sort((a, b) => Number(a) - Number(b));

      // Step 5: Create Document Sets for records with attachments
      for (const record of selectedColumnValues) {
        console.log(`Processing record:`, record);

        const docSetName = String(record).trim();
        const selectedFile = recordFiles[docSetName]; // Get the file for the current record

        // Ensure a file exists for this record
        if (!selectedFile) {
          console.warn(`No file found for record: '${docSetName}', skipping Document Set creation.`);
          continue;
        }
        try {
          console.log(`Creating Document Set with name: '${docSetName}'`);

          // Create a Document Set
          await lib.rootFolder.folders.addUsingPath(docSetName);

          // Retrieve the folder's list item to update its content type
          const docSetFolder = await lib.rootFolder.folders.getByName(docSetName).listItemAllFields.get();
          console.log(`Created Document Set:`, docSetFolder);

          // Update Content Type
          await lib.items.getById(docSetFolder.Id).update({
            ContentTypeId: documentSetContentTypeId,
          });

          console.log(`Document Set '${docSetName}' created successfully.`);

          // Upload document inside the created Document Set
          const folderUrl = `/sites/BulkUpload/${listName}_Documents/${docSetName}`;
          console.log(`Uploading document to Document Set: ${folderUrl}`);

          const fileBuffer = await selectedFile.arrayBuffer(); // Convert file to buffer
          const parameters = { Overwrite: true };

          await sp.web.getFolderByServerRelativeUrl(folderUrl).files.addUsingPath(selectedFile.name, fileBuffer, parameters);
          console.log(`File '${selectedFile.name}' uploaded successfully to '${folderUrl}'.`);

        } catch (error) {
          console.error(`Error processing record '${docSetName}': ${error.message}`);
        }
      }
    } catch (error) {
      console.error(`Error uploading document to folderUrl':`, error);
    }
    return true;
  } catch (error) {
    console.error("Error creating the document library or document set:", error);
    return false;
  }
};


//////// METHOD TO ADD DATA IN THE SHAREPOINT LIST ////////

export const addDataToList = async (
  tableData: any[],
  tableHeaders: string[],
  columnTypes: string[],
  siteUrl: string,
  listName: string,
  getRequestDigest: () => Promise<string>
): Promise<boolean> => {
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
