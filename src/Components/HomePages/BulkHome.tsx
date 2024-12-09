import * as React from "react";
import { Link } from "react-router-dom";
import styles from "../../webparts/bulkUpload/components/BulkUpload.module.scss";

const Home: React.FC = () => {
  return (
    <div className={`${styles.OuterBox}`}>
      <div className={`${styles.InnerBox}`}>
        <h1>Bulk Operations</h1>
      </div>
      <div className={`${styles.ButtonBox}`}>
        <div className={`${styles.HomePage}`}>
          <div className={`${styles.bulkTitle}`}>
            <h3>Bulk Upload</h3>
          </div>
          <div className={`${styles.bulkImage}`}>
            {" "}
            <img
              src={require("../../../src/webparts/bulkUpload/assets/Home1.png")}
              alt="Bulk-Upload Image"
            />
          </div>
          <div className={`${styles.bulkBtn}`}>
            {" "}
            <button>
              <Link to="/selectlisttype">Upload File</Link>
            </button>
          </div>

          <br />
        </div>
        <div className={`${styles.DashPage}`}>
          <div className={`${styles.dashTitle}`}>
            <h3>Dash Board</h3>
          </div>
          <div className={`${styles.dashImage}`}>
            <img
              src={require("../../../src/webparts/bulkUpload/assets/DashBoard1.png")}
              alt="Dashboard Image"
            />
          </div>
          <div className={`${styles.dashBtn}`}>
            {" "}
            <button>
              <Link to="/">DashBoard View</Link>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Home;
