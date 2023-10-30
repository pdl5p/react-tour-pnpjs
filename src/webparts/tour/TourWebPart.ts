import * as React from "react";
import * as ReactDom from "react-dom";
import { Version, DisplayMode } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";

import * as strings from "TourWebPartStrings";
import Tour from "./components/Tour";
import { ITourProps } from "./components/Tour";
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType,
} from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import {
  sp,
  ClientSidePage,
  ClientSideWebpart,
  IClientControlEmphasis,
} from "@pnp/sp";
import { PartSelector, IPartSelectorProps } from "./components/PartSelector";
import { StepText, IStepTextProps } from "./components/StepText";

export interface ITourWebPartProps {
  icon: string;
  actionValue: string;
  description: string;
  collectionData: any[];
  tourVersion: string;
}

export default class TourWebPart extends BaseClientSideWebPart<ITourWebPartProps> {
  //private loadingWebPartData: boolean = true;
  private webPartDataReady: boolean = false;
  private webpartList: any[] = new Array<any[]>();
  private isFullWidthWebPart: boolean = false;

  public onInit(): Promise<void> {
    this.isFullWidthWebPart =
      this.domElement.closest(".CanvasZone--fullWidth") !== null;

    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }

  private getKey(): string {
    const key = `SiteTour_${this.context.pageContext.site.serverRequestPath.toLowerCase()}`;
    console.log("KEY", key);
    return key;
  }

  private isUserClosed(): boolean {
    const key = this.getKey();
    return window.localStorage.getItem(key) !== null;
  }

  private onClose(): void {
    const key = this.getKey();
    window.localStorage.setItem(key, "closed");

    this.render();
  }

  public render(): void {
    const editMode = this.displayMode === DisplayMode.Edit;
    const key = this.getKey();
    if (this.isUserClosed() && !editMode) {
      ReactDom.render(React.createElement("div", {}), this.domElement);
    } else {
      const element: React.ReactElement<ITourProps> = React.createElement(
        Tour,
        {
          actionValue: this.properties.actionValue,
          description: this.properties.description,
          icon: this.properties.icon,
          collectionData: this.properties.collectionData,
          onClose: this.onClose.bind(this),
          editMode,
          isFullWidth: this.isFullWidthWebPart,
        }
      );
      ReactDom.render(element, this.domElement);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  public async GetAllWebpart(): Promise<any[]> {
    // page file
    const file = sp.web.getFileByServerRelativePath(
      this.context.pageContext.site.serverRequestPath
    );

    const page = await ClientSidePage.fromFile(file);

    const wpData: any[] = [];

    page.sections.forEach((section) => {
      section.columns.forEach((column) => {
        column.controls.forEach((control) => {
          if (control.data.webPartData != undefined) {
            wpData.push({
              text: `sec[${section.order}] col[${column.order}] - ${control.data.webPartData.title}`,
              key: `${control.data.webPartData.instanceId}`,
            });
          } else {
            wpData.push({
              text: `sec[${section.order}] col[${column.order}] - "Webpart"`,
              key: `${control.data.id}`,
            });
          }
        });
      });
    });

    wpData.push({ text: "Custom CSS Selector", key: "custom" });

    return wpData;
  }

  protected onPropertyPaneConfigurationStart(): void {
    console.log("onPropertyPaneConfigurationStart");
    var self = this;
    self.webPartDataReady = false;
    this.GetAllWebpart().then((res) => {
      self.webpartList = res;
      self.webPartDataReady = true;
      setTimeout(() => {
        self.context.propertyPane.refresh();
      }, 10);
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log("getPropertyPaneConfiguration");
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
                PropertyPaneTextField("actionValue", {
                  label: strings.ActionValueFieldLabel,
                }),
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneTextField("icon", {
                  label: strings.IconNameFieldLabel,
                }),
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Tour steps",
                  panelHeader: "Tour steps configuration",
                  manageBtnLabel: "Configure tour steps",
                  value: this.properties.collectionData,
                  //manageBtnEnabled: !this.loadIndicator,
                  enableSorting: true,
                  fields: [
                    {
                      id: "CssSelector",
                      title: "WebPart or CSS Selector",
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (
                        field,
                        value,
                        onUpdate,
                        item,
                        itemId,
                        onError
                      ) => {
                        const props: IPartSelectorProps = {
                          options: this.webpartList,
                          value,
                          onUpdate,
                          onError,
                          field,
                        };
                        return React.createElement(PartSelector, props);
                      },
                    },
                    {
                      id: "StepDescription",
                      title: "Step Description",
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (
                        field,
                        value,
                        onUpdate,
                        item,
                        itemId,
                        onError
                      ) => {
                        const props: IStepTextProps = {
                          value,
                          onUpdate,
                          onError,
                          field,
                        };
                        return React.createElement(StepText, props);
                      },
                    },
                    {
                      id: "Enabled",
                      title: "Enabled",
                      type: CustomCollectionFieldType.boolean,
                      defaultValue: true,
                    },
                  ],
                  disabled: !this.webPartDataReady,
                }),
                PropertyPaneTextField("tourVersion", {
                  label: "Tour version",
                  description:
                    "Update this to reset the tour for users who have closed the web part",
                  value: "v1",
                }),
              ],
            },
          ],
        },
      ],
      loadingIndicatorDelayTime: 5,
      showLoadingIndicator: this.webPartDataReady ? false : true,
    };
  }
}
