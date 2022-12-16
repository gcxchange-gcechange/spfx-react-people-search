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
  hidingUsers: string;
  
}

export default class ReactDirectoryWebPart extends BaseClientSideWebPart<IReactDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactDirectoryProps> =
      React.createElement(DirectoryHook, {
        title: this.properties.title,
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        pageSize: this.properties.pageSize,
        prefLang: this.properties.prefLang,
        hidingUsers: this.properties.hidingUsers,
        searchFirstName: this.properties.searchFirstName,
      });

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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel,
                }),
                PropertyPaneDropdown("prefLang", {
                  label: "Preferred Language",
                  options: [
                    { key: "account", text: "Account" },
                    { key: "en-us", text: "English" },
                    { key: "fr-fr", text: "Français" },
                  ],
                }),
                PropertyPaneTextField("hidingUsers", {
                  label: "Users not in serach",
                  description:"Enter the users' emails who don't need to be in search separated by comma "
                }),

                PropertyPaneSlider("pageSize", {
                  label: "Results per page",
                  showValue: true,
                  max: 20,
                  min: 2,
                  step: 2,
                  value: this.properties.pageSize,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
