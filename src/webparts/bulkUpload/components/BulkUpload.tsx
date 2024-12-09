import * as React from 'react';
// import styles from './BulkUpload.module.scss';
import type { IBulkUploadProps } from './IBulkUploadProps';
import { HashRouter as Router, Routes, Route } from 'react-router-dom';
import BulkHome from "../../../Components/HomePages/BulkHome"
import ListHome from "../../../Components/HomePages/ListHome"
import NewList from "../../../Components/BulkUpload/NewList";
import ExistingList from "../../../Components/BulkUpload/ExistingList";
export default class BulkUpload extends React.Component<IBulkUploadProps> {
  public render(): React.ReactElement<IBulkUploadProps> {
    const {
      
    } = this.props;

    return (
      <Router>
          {console.log("Router is rendering")}
          <Routes>
            <Route path='/' element={<BulkHome/>}></Route>
            <Route path='/selectlisttype' element={<ListHome/>}></Route>
            <Route path='/newlist' element={<NewList/>}></Route>
            <Route path='/existlist' element={<ExistingList/>}></Route>
          </Routes>
        </Router>
    );
  }
}
