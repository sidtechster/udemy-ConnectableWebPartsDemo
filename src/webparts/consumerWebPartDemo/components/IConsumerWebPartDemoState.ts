import { IEmployee } from "./IEmployee";
import { DynamicProperty } from "@microsoft/sp-component-base";

export interface IConsumerWebPartDemoState {
    status: string;
    EmployeeListItems: IEmployee[];
    EmployeeListItem: IEmployee;
    DeptTitleId: string;
}