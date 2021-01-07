import * as React from 'react';
import styles from './ProviderWebPartDemo.module.scss';
import { IProviderWebPartDemoProps } from './IProviderWebPartDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IProviderWebPartDemoState } from './IDepartmentState';

import {
  autobind,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  DetailsRowCheck,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IDepartment } from './IDepartment';

let _departmentListColumns = [
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
  }
];

export default class ProviderWebPartDemo extends React.Component<IProviderWebPartDemoProps, IProviderWebPartDemoState> {

  private _selection: Selection;

  private _onItemsSelectionChanged = () => {
    this.props.onDepartmentSelected(this._selection.getSelection()[0] as IDepartment);

    this.setState({
      DepartmentListItem: (this._selection.getSelection()[0] as IDepartment)
    });
  }

  constructor(props: IProviderWebPartDemoProps, state: IProviderWebPartDemoState) {
    super(props);

    this.state = {
      status: 'Ready',
      DepatmentListItems: [],
      DepartmentListItem: {
        Id: 0,
        Title: ""
      }
    };

    this._selection = new Selection({
      onSelectionChanged: this._onItemsSelectionChanged
    });
  }

  private _getListItems(): Promise<IDepartment[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('Department')/items";
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    })
    .then(json => {
      return json.value;
    }) as Promise<IDepartment[]>;
  }

  public bindDetailsList(message: string) : void {
    this._getListItems().then(listItems => {
      this.setState({ DepatmentListItems: listItems, status: message });
    });
  }

  public componentDidMount(): void {
    this.bindDetailsList("All records have been loaded successfully");
  }

  public render(): React.ReactElement<IProviderWebPartDemoProps> {
    return (
      <div className={ styles.providerWebPartDemo }>
        <DetailsList 
          items = { this.state.DepatmentListItems }
          columns = { _departmentListColumns }
          setKey = 'Id'
          checkboxVisibility = { CheckboxVisibility.always }
          selectionMode = { SelectionMode.single }
          layoutMode = { DetailsListLayoutMode.fixedColumns }
          compact = { true }
          selection = { this._selection }
        />
      </div>
    );
  }
}
