import * as React from 'react';
import { HashRouter as Router, Routes, Route } from 'react-router-dom';
import BulkHome from "../../../Components/HomePages/BulkHome";
import ListHome from "../../../Components/HomePages/ListHome";
import NewList from "../../../Components/BulkUpload/NewList";
import ExistingList from "../../../Components/BulkUpload/ExistingList";
// import { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IBulkUploadProps } from './IBulkUploadProps';

export default class BulkUpload extends React.Component<IBulkUploadProps> {
  public render(): React.ReactElement<IBulkUploadProps> {
    const { context } = this.props;

    return (
      <Router>
        {console.log("Router is rendering")}
        <Routes>
          <Route path='/' element={<BulkHome/>}></Route>
          <Route path='/selectlisttype' element={<ListHome/>}></Route>
          <Route path='/newlist' element={<NewList context={context}/>}></Route>
          <Route path='/existlist' element={<ExistingList/>}></Route>
        </Routes>
      </Router>
    );
  }
}
