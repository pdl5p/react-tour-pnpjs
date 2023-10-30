import { find } from "@microsoft/sp-lodash-subset";
export class TourHelper {
  private static reGuid = new RegExp(
    "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
  );

  public static isGuid(value: string): boolean {
    return this.reGuid.test(value);
  }

  public static extractGuid(value: string): string {
    return this.reGuid.exec(value)[0] || "";
  }

  private static matchViewPortId(
    viewPortIds: string[],
    selector: string
  ): string {
    // viewPortIds.forEach((id) => {
    //   console.log("id", id, selector, id.indexOf(selector));

    //   if (id.indexOf(selector) > -1) {
    //     match = `[data-viewport-id="${id}"]`;
    //   }
    // });

    const match = find(viewPortIds, (id) => id.indexOf(selector) > -1);
    return match ? `[data-viewport-id="${match}"]` : "";

    //viewPortIds.find((id) => { return id.indexOf(selector) > -1; });
    //return match;
  }

  public static getTourSteps(settings: any[]): any[] {
    var result: any[] = new Array<any>();

    const viewPorts = document.querySelectorAll("div[data-viewport-id]");
    const viewPortIds: string[] = Array.prototype.map
      .call(viewPorts, (p: Element) => {
        const dvp = p.getAttribute("data-viewport-id") || null;
        return dvp;
      })
      .filter((p: string | null) => {
        return p !== null;
      });

    //console.log("VPID", viewPortIds);
    if (settings != undefined) {
      settings.forEach((ele) => {
        if (ele.Enabled) {
          let selector = ele.CssSelector;
          if (TourHelper.isGuid(selector)) {
            //selector = `[data-sp-feature-instance-id="${selector}"]`;
            //selector = `#${selector}`;
            selector = TourHelper.matchViewPortId(viewPortIds, selector);
          }
          console.log("selector", ele.CssSelector, selector);
          if (selector) {
            result.push({
              selector: selector,
              content: ele.StepDescription,
            });
          }
        }
      });
    }
    //console.log("result", result);
    return result;
  }
}
