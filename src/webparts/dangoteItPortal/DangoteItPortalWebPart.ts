import * as React from 'react';
import * as ReactDom from 'react-dom';
import {Version} from '@microsoft/sp-core-library';
import {IPropertyPaneConfiguration,PropertyPaneTextField} from '@microsoft/sp-property-pane';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import * as strings from 'DangoteItPortalWebPartStrings';
import DangoteItPortal from './components/DangoteItPortal';
import {IDangoteItPortalProps} from './components/IDangoteItPortalProps';
export interface IDangoteItPortalWebPartProps {description:string;}
export default class DangoteItPortalWebPart extends BaseClientSideWebPart<IDangoteItPortalWebPartProps>{
 public render():void {const element:React.ReactElement<IDangoteItPortalProps>=React.createElement(DangoteItPortal,{description:this.properties.description,isDarkTheme:false,environmentMessage:this.context.isServedFromLocalhost?'Running locally':'Running in SharePoint'});ReactDom.render(element,this.domElement);}
 protected onDispose():void{ReactDom.unmountComponentAtNode(this.domElement);}
 protected get dataVersion():Version{return Version.parse('1.0');}
 protected getPropertyPaneConfiguration():IPropertyPaneConfiguration{return {pages:[{header:{description:strings.PropertyPaneDescription},groups:[{groupName:strings.BasicGroupName,groupFields:[PropertyPaneTextField('description',{label:strings.DescriptionFieldLabel})]}]}]};}
}
