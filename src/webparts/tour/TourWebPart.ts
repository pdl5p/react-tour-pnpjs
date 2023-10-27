import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TourWebPartStrings';
import Tour from './components/Tour';
import { ITourProps } from './components/ITourProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { sp, ClientSidePage, ClientSideWebpart, IClientControlEmphasis } from '@pnp/sp';
import { PartSelector, IPartSelectorProps } from './components/PartSelector';
import { StepText, IStepTextProps } from './components/StepText';

export interface ITourWebPartProps {
  actionValue: string;
  description: string;
  collectionData: any[];
  tourVersion: string;
}


export default class TourWebPart extends BaseClientSideWebPart<ITourWebPartProps> {

  private loadIndicator: boolean = false;
  private webpartList: any[] = new Array<any[]>();
  private isFullWidthWebPart: boolean = false;

  public onInit(): Promise<void> {

    this.isFullWidthWebPart = this.domElement.closest("CanvasControl--fullWidth") !== null;

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  private getKey(): string {
    const key = `SiteTour_${this.properties.tourVersion}_${this.context.pageContext.site.serverRequestPath.toLowerCase()}`;
    console.log("KEY", key);
    return key;
  }

  private onClose(): void {
    //alert("closing");

    const key = this.getKey();
    window.localStorage.setItem(key, "closed");

    this.render();
  }

  public render(): void {

    const editMode = this.displayMode === DisplayMode.Edit;
    const key = this.getKey();
    if(window.localStorage.getItem(key) && !editMode){
      ReactDom.render(React.createElement("div", {}), this.domElement);
    }
    else
    {

      const element: React.ReactElement<ITourProps> = React.createElement(
        Tour,
        {
          actionValue: this.properties.actionValue,
          description: this.properties.description,
          collectionData: this.properties.collectionData,
          onClose: this.onClose.bind(this),
          editMode,
        }
      );
      ReactDom.render(element, this.domElement);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public async GetAllWebpart(): Promise<any[]> {
    console.log("GGG");
    // page file
    const file = sp.web.getFileByServerRelativePath(this.context.pageContext.site.serverRequestPath);

    const page = await ClientSidePage.fromFile(file);

    const wpData: any[] = [];

    page.sections.forEach(section => {
      section.columns.forEach(column => {
        column.controls.forEach(control => {
          var wpName = {};
          var wp = {};
          if (control.data.webPartData != undefined) {
            wpName = `sec[${section.order}] col[${column.order}] - ${control.data.webPartData.title}`;
            wp = { text: wpName, key: `[data-sp-feature-instance-id="${control.data.webPartData.instanceId}"]` };
            wpData.push(wp);
          } else {
            wpName = `sec[${section.order}] col[${column.order}] - "Webpart"`;
            wp = { text: wpName, key: `[data-sp-feature-instance-id="${control.data.id}"]` };
            wpData.push(wp);
          }
        });
      });
    });

    console.log("WPDATA", wpData);
    wpData.push({ text: "Custom CSS Selector", key: "custom" });

    return wpData;
  }

  protected onPropertyPaneConfigurationStart(): void {
    var self = this;
    this.GetAllWebpart().then(res => {
      self.webpartList = res;
      self.loadIndicator = false;
      self.context.propertyPane.refresh();
    });
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
                PropertyPaneTextField('actionValue', {
                  label: strings.ActionValueFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Tour steps",
                  panelHeader: "Tour steps configuration",
                  manageBtnLabel: "Configure tour steps",
                  value: this.properties.collectionData,
                  enableSorting: true,
                  fields: [
                    {
                        id: "CssSelector",
                        title: "WebPart or CSS Selector",
                        type: CustomCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                          const props: IPartSelectorProps = {
                            options: this.webpartList,
                            value,
                            onUpdate,
                            onError,
                            field
                          };
                          return (
                            React.createElement(PartSelector, props)
                          );
                        },
                    },
                    // {
                    //   id: "WebPart",
                    //   title: "section[x] column[y] - WebPart Title",
                    //   type: CustomCollectionFieldType.dropdown,
                    //   options: this.webpartList,
                    //   required: true,
                    // },
                    // {
                    //   id: "CssSelector",
                    //   title: "CSS Selector",
                    //   type: CustomCollectionFieldType.string,
                    //   required: true,
                    // },
                    {
                      id: "StepDescription",
                      title: "Step Description",
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {

                        const props: IStepTextProps = {
                          value,
                          onUpdate,
                          onError,
                          field
                        };

                        return (
                          React.createElement(StepText, props)
                        );
                        // return (
                        //   React.createElement("div", null,
                        //     React.createElement("textarea",
                        //       {
                        //         style: { width: "400px", height: "100px" },
                        //         placeholder: "Step description",
                        //         key: itemId,
                        //         value: value,
                        //         onChange: (event: React.FormEvent<HTMLTextAreaElement>) => {
                        //           console.log(event);
                        //           onUpdate(field.id, event.currentTarget.value);
                        //         }
                        //       })
                        //   )
                        // );
                      }
                    },
                    // {
                    //   id: "Position",
                    //   title: "Position",
                    //   type: CustomCollectionFieldType.number,
                    //   required: true,
                      
                    // },
                    {
                      id: "Enabled",
                      title: "Enabled",
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: true
                    }
                  ],
                  disabled: false
                }),
                PropertyPaneTextField('tourVersion', {
                  label: "Tour version",
                  description: "Update this to reset the tour for users who have closed the web part",
                  value: "v1",
                })
              ]
            }
          ]
        }
      ],
      loadingIndicatorDelayTime: 5,
      showLoadingIndicator: this.loadIndicator
    };
  }
}
