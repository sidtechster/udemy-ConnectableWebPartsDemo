import * as React from 'react';
import styles from './ConsumerWebPartDemo.module.scss';
import { IConsumerWebPartDemoProps } from './IConsumerWebPartDemoProps';
import { escape, truncate } from '@microsoft/sp-lodash-subset';
import { IConsumerWebPartDemoState } from './IConsumerWebPartDemoState';

import {
  autobind,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  DetailsRowCheck,
  Selection
} from 'office-ui-fabric-react';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IWebPartPropertiesMetadata } from "@microsoft/sp-webpart-base";
import { IEmployee } from './IEmployee';

let _employeeListColumns = [
  {
    key: 'ID',
    name: 'ID',
    fieldName: 'ID',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'Title',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'DeptTitle',
    name: 'DeptTitle',
    fieldName: 'DeptTitleId',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'Designation',
    name: 'Designation',
    fieldName: 'Designation',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  }
];

export default class ConsumerWebPartDemo extends React.Component<IConsumerWebPartDemoProps, IConsumerWebPartDemoState> {
  
  constructor(props: IConsumerWebPartDemoProps, state: IConsumerWebPartDemoState) {
    super(props);

    this.state = {
      status: 'Ready',
      EmployeeListItems: [],
      EmployeeListItem: {
        Id: 0,
        Title: "",
        DeptTitle: "",
        Designation: ""
      },
      DeptTitleId: ""
    };
  }

  private _getListItems(): Promise<IEmployee[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('Employees')/items?filter=DeptTitleId eq " + this.props.DeptTitleId.tryGetValue();
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    })
    .then(json => {
      return json.value;
    }) as Promise<IEmployee[]>;
  }

  public bindDetailsList(message: string) : void {
    this._getListItems().then(listItems => {
      this.setState({ EmployeeListItems: listItems, status: message,
        DeptTitleId: this.props.DeptTitleId.tryGetValue().toString() });
    });
  }
  
  public render(): React.ReactElement<IConsumerWebPartDemoProps> {

    if(this.state.DeptTitleId != this.props.DeptTitleId.tryGetValue())
    {
      this.bindDetailsList("All records has been loaded successfully");
    }

    return (
      <div className={ styles.consumerWebPartDemo }>
        <div>
          {/* <h1>Selected department is : {this.props.DeptTitleId.tryGetValue()}</h1> */}
          <h1>Selected department is : {this.state.DeptTitleId}</h1>
        </div>
        <DetailsList 
          items = { this.state.EmployeeListItems }
          columns = { _employeeListColumns }
          setKey = 'Id'
          checkboxVisibility = { CheckboxVisibility.always }
          selectionMode = { SelectionMode.single }
          layoutMode = { DetailsListLayoutMode.fixedColumns }
          compact = { true  }
        />
      </div>
    );
  }
}
