import { UserProfile, ExpressResponse } from "../models/ExpressRequest";
import { IDriveItem } from "../dal/OneDriveAPIDao";


export class Render {
  public static sendMail(profile: UserProfile, actual_recipient: string | undefined, res: ExpressResponse) {
    const vm = {
      display_name: profile.displayName,
      email_address: profile.emails[0].address,
      actual_recipient: actual_recipient
    }

    res.render('sendMail', vm);
  }

  public static folder({ driveId, items }: { driveId: string, items: IDriveItem[] }, res: ExpressResponse) {

    const vm = {
      driveId,
      items: items.map(it => {
        return {
          ...it,
          downloadUrl: it.file && it['@microsoft.graph.downloadUrl'] ? it['@microsoft.graph.downloadUrl'] : null
        };
      })
    }

    res.render('folder', vm);
  }


  public static error(e, res: ExpressResponse) {
    e.innerError = (e.response) ? e.response.text : '';
    res.render('error', {
      error: e
    });
  }
}
