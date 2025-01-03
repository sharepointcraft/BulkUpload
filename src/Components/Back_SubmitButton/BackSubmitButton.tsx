import * as React from 'react';
import { Link } from 'react-router-dom';
import styles from './BackSubmitButtons.module.scss'; // Assuming you have a separate CSS module for styling
 
interface BackSubmitButtonsProps {
  showButtons: boolean;
  showTable: boolean;
  validateColumns: () => Promise<boolean>;
  createSharePointList: () => Promise<boolean>;
  createDocumentLibrary: () => Promise<boolean>;
  addDataToList: () => Promise<boolean>;
  createDocLib: string;
  setIsDialogVisible: (value: boolean) => void;
  setPopupMessage: (message: string) => void;
  setShowSuccessPopup: (value: boolean) => void;
  setErrorPopupMessage: (message: string) => void;
  setIsPopupOpen: (value: boolean) => void;
  setShowTable: (value: boolean) => void;
  setShowSuccessIcon: (value: boolean) => void;
}
 
const BackSubmitButtons: React.FC<BackSubmitButtonsProps> = ({
  showButtons,
  showTable,
  validateColumns,
  createSharePointList,
  createDocumentLibrary,
  addDataToList,
  createDocLib,
  setIsDialogVisible,
  setPopupMessage,
  setShowSuccessPopup,
  setErrorPopupMessage,
  setIsPopupOpen,
  setShowTable,
  setShowSuccessIcon,
}) => {
  const handleValidateClick = async () => {
    const isValid = await validateColumns();
    if (isValid) {
      setIsDialogVisible(true);
    }
  };
 
  const handleSubmitClick = async () => {
    try {
      setPopupMessage('Creating SharePoint list...');
      setShowSuccessPopup(true);
 
      const isListCreated = await createSharePointList();
      if (!isListCreated) {
        setShowSuccessPopup(false);
        setErrorPopupMessage('List Name: List name already exists, Please use different name.');
        setIsPopupOpen(true);
        return;
      }
 
      if (createDocLib === 'yes') {
        setPopupMessage('Creating document library...');
        const isLibraryCreated = await createDocumentLibrary();
        if (!isLibraryCreated) {
          setShowSuccessPopup(false);
          setErrorPopupMessage('Failed to create document library.');
          setIsPopupOpen(true);
          return;
        }
      }
 
      setPopupMessage('Submitting data...');
      const isDataSubmitted = await addDataToList();
      if (!isDataSubmitted) {
        setShowSuccessPopup(false);
        setErrorPopupMessage('Failed to submit data.');
        setIsPopupOpen(true);
        return;
      }
 
      setPopupMessage('Data successfully submitted.');
      setShowSuccessIcon(false);
      setShowTable(false);
    } catch (error: any) {
      setErrorPopupMessage(`An unexpected error occurred. ${error.message}`);
      setIsPopupOpen(true);
    }
  };
 
  if (!showButtons) return null;
 
  return (
<div className={`${styles.backSubmitbtn}`}>
<div className={`${styles.backBtn}`}>
<button>
<Link to="/selectlisttype">Back</Link>
</button>
</div>
<div className={`${styles.validateBtn}`}>
        {showTable ? (
<button onClick={handleValidateClick}>Validate</button>
        ) : (
<button onClick={handleSubmitClick}>Submit</button>
        )}
</div>
</div>
  );
};
 
export default BackSubmitButtons;