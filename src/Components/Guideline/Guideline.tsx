import * as React from "react";
import styles from "./Guideline.module.scss";

interface GuidelineProps {
  type: "new" | "existing"; // Define prop for choosing guideline type
}

const Guideline: React.FC<GuidelineProps> = ({ type }) => {
  return (
    <div className={styles.guidelineContainer}>
      {type === "new" ? (
        <>
          <h2>Guidelines for Creating a New List</h2>
          <p>
            - Choose a unique <b>List Name</b> and provide all necessary fields.
          </p>
          <p>
            - Upload an Excel or CSV file with a proper structure matching the list format.
          </p>
          <p>
            - Ensure that column names in your file match the SharePoint list headers exactly.
          </p>
          <p className={styles.note}>
            <b>Note:</b> Once created, you can add or modify list items but cannot rename the list.
          </p>
        </>
      ) : (
        <>
          <h2>Guidelines for Uploading to an Existing List</h2>
          <p>
            - Select an existing list from the dropdown before uploading a file.
          </p>
          <p>
            - The uploaded file should have the same column structure as the selected list.
          </p>
          <p>
            - If an <b>ID</b> column is present, the system will update existing records instead of creating new ones.
          </p>
          <p className={styles.note}>
            <b>Note:</b> Ensure data consistency to avoid mismatches or validation errors.
          </p>
        </>
      )}
    </div>
  );
};

export default Guideline;
