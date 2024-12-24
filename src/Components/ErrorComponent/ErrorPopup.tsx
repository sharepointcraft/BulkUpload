import * as React from 'react';
import styles from './ErrorPopup.module.scss'; // Import custom styles
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faExclamationCircle } from '@fortawesome/free-solid-svg-icons';

interface ErrorPopupProps {
  isOpen: boolean;
  message: string;
  onClose: () => void;
}

const ErrorPopup: React.FC<ErrorPopupProps> = ({ isOpen, message, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className={`${styles['popup-overlay']}`}>
      <div className={`${styles['popup-container']}`}>
        <button className={`${styles['popup-close']}`} onClick={onClose}>
          X
        </button>
        <div className={`${styles['popup-icon']}`}>
          <FontAwesomeIcon icon={faExclamationCircle} size="6x" color="red" />
        </div>
        <h3>ALERT!</h3>
        <div
          className={`${styles['popup-message']}`}
          dangerouslySetInnerHTML={{ __html: message }} // Inject HTML safely
        />
      </div>
    </div>
  );
};

export default ErrorPopup;
