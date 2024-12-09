import * as React from "react";
import * as ReactDom from "react-dom";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// import * as strings from "BulkUploadWebPartStrings";
import BulkUpload from "./components/BulkUpload";
import { IBulkUploadProps } from "./components/IBulkUploadProps";

export interface IBulkUploadWebPartProps {
  description: string;
}

export default class BulkUploadWebPart extends BaseClientSideWebPart<IBulkUploadWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IBulkUploadProps> =
      React.createElement(BulkUpload,{});

    ReactDom.render(element, this.domElement);
  }
}


