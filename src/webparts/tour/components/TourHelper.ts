import {
  sp
} from "@pnp/sp";
import { graph } from "@pnp/graph";


export class TourHelper {

  public static getTourSteps(settings: any[]): any[] {

    var result: any[] = new Array<any>();

    if (settings != undefined) {
      settings.forEach(ele => {
        if (ele.Enabled) {
          result.push(
            {
              selector: ele.CssSelector,
              content: ele.StepDescription
            });
        }
      });

    }
    return result;
  }

}
