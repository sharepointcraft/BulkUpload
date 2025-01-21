// SuccessPopUp.tsx
import * as React from 'react';
import styles from './SuccessPopUp.module.scss';

interface SuccessPopUpProps {
  showSuccessPopup: boolean;
  showSuccessIcon: boolean;
  popupMessage: string;
  setShowSuccessPopup: (show: boolean) => void;
  setShowSuccessIcon: (show: boolean) => void;
  resetForm: () => void;
}

const SuccessPopUp: React.FC<SuccessPopUpProps> = ({
    showSuccessPopup,
    showSuccessIcon,
    popupMessage,
    setShowSuccessPopup,
    setShowSuccessIcon,
    resetForm,
  }) => {
    // If showSuccessPopup is false, return an empty div instead of false.
    return showSuccessPopup ? (
      <div className={styles.successPopup}>
        <div
          className={styles.popupContent}
          style={{
            borderColor: showSuccessPopup ? 'yellow' : 'green',
            borderWidth: '2px',
            borderStyle: 'solid',
          }}
        >
          {showSuccessIcon ? (
            <div className={styles.circularProgress}>
              <div className={styles.loadingSpinner}></div>
            </div>
          ) : (
            <span className={styles['success-icon']}>✔</span>
          )}
          <p>{popupMessage}</p>
          {!showSuccessIcon && (
            <button
              className={styles.okButton}
              onClick={() => {
                setShowSuccessIcon(true);
                setShowSuccessPopup(false);
                resetForm();
              }}
            >
              OK
            </button>
          )}
        </div>
      </div>
    ) : (
      <div></div> // Empty div or any element you want to render when false
    );
  };
  

export default SuccessPopUp;
