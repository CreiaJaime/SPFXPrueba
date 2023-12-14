import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import styles from './CompanyAnnouncementsWebPart.module.scss';

export interface ICompanyAnnouncementsWebPartProps {
}

export default class CompanyAnnouncementsWebPart extends BaseClientSideWebPart<ICompanyAnnouncementsWebPartProps> {
  public render(): void {
    this.context.spHttpClient
  .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Announcements')/items?$select=Title,Description,Important`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'accept': 'application/json;odata.metadata=none'
      }
    })
  .then(response => response.json())
  .then(announcements => {
    const announcementsHtml = announcements.value.map((announcement: any) =>
      `<dt${announcement.Important ?
        ` class="${styles.important}"` : ''}>${announcement.Title}</dt>
      <dd>${announcement.Description}</dd>`); 
    
    this.domElement.innerHTML = `
    <div class="${styles.companyAnnouncements}">
      <div class="${styles.container}">
        <div class="${styles.title}">Announcements</div>
        <dl>
          ${announcementsHtml.join('')}
        </dl>
      </div>
    </div>
    `;
  })
  .catch(error => this.context.statusRenderer.renderError(this.domElement, error));
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
