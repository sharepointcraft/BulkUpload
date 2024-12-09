import * as React from "react";
import styles from "../../webparts/bulkUpload/components/BulkUpload.module.scss";
import { Link } from "react-router-dom";

const SelectList: React.FC = () => {
  return (
    <div className={`${styles.OuterBox}`}>
      <div className={`${styles.InnerBox}`}>
        <h1>Welcome! Select Your List</h1>
      </div>
      <div className={`${styles.ButtonBox}`}>
        <div className={`${styles.ExistListPage}`}>
          {/* <div className={`${styles.ExistListTitle}`}>
            <h3>Existing List</h3>
          </div> */}
          <div className={`${styles.ExistListImg}`}>
            <img
              src={require("../../../src/webparts/bulkUpload/assets/ExistingList.png")}
              alt="Existing List Image"
            />
          </div>
          <div className={`${styles.ExistListBtn}`}>
            <button>
              <Link to="/newlist">Existing List</Link>
            </button>
          </div>
        </div>
        <div className={`${styles.NewListPage}`}>
          {/* <div className={`${styles.NewListTitle}`}>
          {" "}
          <h3>New List</h3>
        </div> */}
          <div className={`${styles.NewListImg}`}>
            <img
              src={require("../../../src/webparts/bulkUpload/assets/NewList.png")}
              alt="New List Image"
            />
          </div>
          <div className={`${styles.NewListBtn}`}>
            <button>
              <Link to="/existlist">Create New List</Link>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SelectList;
