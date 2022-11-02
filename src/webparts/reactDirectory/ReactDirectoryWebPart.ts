import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  IPropertyPaneToggleProps,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactDirectoryWebPartStrings';
import DirectoryHook from "./components/DirectoryHook";
import { IReactDirectoryProps } from './components/IReactDirectoryProps';

export interface IReactDirectoryWebPartProps {
  title: string;
  searchFirstName: boolean;
  searchProps: string;
  clearTextSearchProps: string;
  pageSize: number;
  prefLang: string;
}

export default class ReactDirectoryWebPart extends BaseClientSideWebPart<IReactDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactDirectoryProps> =
      React.createElement(
        // ReactDirectory,
        // {
        //   title: this.properties.title,
        //   context: this.context,
        //   searchFirstName: this.properties.searchFirstName,
        //   displayMode: this.displayMode,
        //   updateProperty: (value: string) => {
        //     this.properties.title = value;
        //   }
        // },
        DirectoryHook,
        {
          title: this.properties.title,
          context: this.context,
          searchFirstName: this.properties.searchFirstName,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          searchProps: this.properties.searchProps,
          clearTextSearchProps: this.properties.clearTextSearchProps,
          pageSize: this.properties.pageSize,
          prefLang: this.properties.prefLang,
        }
      );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneDropdown('prefLang', {
                  label: 'Preferred Language',
                  options: [
                    { key: 'account', text: 'Account' },
                    { key: 'en-us', text: 'English' },
                    { key: 'fr-fr', text: 'Français' }
                  ]}),
                PropertyPaneToggle("searchFirstName", {
                  checked: false,
                  label: "Search on First Name Only?"
                }),
                // PropertyPaneTextField('searchProps', {
                //   label: strings.SearchPropsLabel,
                //   description: strings.SearchPropsDesc,
                //   value: this.properties.searchProps,
                //   multiline: false,
                //   resizable: false
                // }),
                // PropertyPaneTextField('clearTextSearchProps', {
                //   label: strings.ClearTextSearchPropsLabel,
                //   description: strings.ClearTextSearchPropsDesc,
                //   value: this.properties.clearTextSearchProps,
                //   multiline: false,
                //   resizable: false
                // }),
                PropertyPaneSlider('pageSize', {
                  label: 'Results per page',
                  showValue: true,
                  max: 20,
                  min: 2,
                  step: 2,
                  value: this.properties.pageSize
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
