import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneButton,
  PropertyPaneButtonType,  
  IPropertyPaneGroup,
  IPropertyPanePage,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ImportantLinksWebPart.module.scss';
import * as strings from 'ImportantLinksWebPartStrings';
//import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';

export interface IImportantLinksWebPartProps {
  description: string;
  numGroups : number;
  useList : boolean;
  
  groupTitle1 : string;
  numLinks1: number;
  linkgroup1_Title1: string;
  linkgroup1_Title2: string;
  linkgroup1_Title3: string;
  linkgroup1_Title4: string;
  linkgroup1_Title5: string;
  linkgroup1_Title6: string;
  linkgroup1_Title7: string;
  linkgroup1_Title8: string;
  linkgroup1_Title9: string;
  linkgroup1_Title10: string;
  linkgroup1_Title11: string;
  linkgroup1_Title12: string;
  linkgroup1_Title13: string;
  linkgroup1_Title14: string;
  linkgroup1_Title15: string;
  linkgroup1_URL1: string;
  linkgroup1_URL2: string;
  linkgroup1_URL3: string;
  linkgroup1_URL4: string;
  linkgroup1_URL5: string;
  linkgroup1_URL6: string;
  linkgroup1_URL7: string;
  linkgroup1_URL8: string;
  linkgroup1_URL9: string;
  linkgroup1_URL10: string;
  linkgroup1_URL11: string;
  linkgroup1_URL12: string;
  linkgroup1_URL13: string;
  linkgroup1_URL14: string;
  linkgroup1_URL15: string;
  linkgroup1_Browse1: string;
  linkgroup1_Browse2: string;
  linkgroup1_Browse3: string;
  linkgroup1_Browse4: string;
  linkgroup1_Browse5: string;
  linkgroup1_Browse6: string;
  linkgroup1_Browse7: string;
  linkgroup1_Browse8: string;
  linkgroup1_Browse9: string;
  linkgroup1_Browse10: string;
  linkgroup1_Browse11: string;
  linkgroup1_Browse12: string;
  linkgroup1_Browse13: string;
  linkgroup1_Browse14: string;
  linkgroup1_Browse15: string;

  groupTitle2 : string;
  numLinks2: number;
  groupTitle3 : string;
  numLinks3: number;
  groupTitle4 : string;
  numLinks4: number;
  groupTitle5 : string;
  numLinks5: number;
  groupTitle6 : string;
  numLinks6: number;
  groupTitle7 : string;
  numLinks7: number;
  groupTitle8 : string;
  numLinks8: number;
  groupTitle9 : string;
  numLinks9: number;
  groupTitle10 : string;
  numLinks10: number;
}

export default class ImportantLinksWebPart extends BaseClientSideWebPart<IImportantLinksWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.importantLinks} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>

      <div>
        <h5>${this.properties.groupTitle1} links=${this.properties.numLinks1}</h5>
        <h5>${this.properties.groupTitle2} links=${this.properties.numLinks2}</h5>
        <h5>${this.properties.groupTitle3} links=${this.properties.numLinks3}</h5>
        <h5>${this.properties.groupTitle4} links=${this.properties.numLinks4}</h5>
      </div>
    </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected textBoxValidationMethod(value: string): string {
    // check for image extension
    //if (value.length < 10) 
    //{
    //  return "Name should be at least 10 characters!"; 
    //}
    //else 
    //{ 
      return ""; 
    //}
  } 
  
  private buttonClick(): void {  
    const currentWebUrl = this.context.pageContext.web.absoluteUrl; 
    window.open(currentWebUrl+'/Lists/ImportantLinks/AllItems.aspx','_blank');  
    //return "test"  
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let linkList : any={};
    const linkGroups : IPropertyPaneGroup["groupFields"]=[];
    const group1Links : IPropertyPaneGroup["groupFields"]=[];
    const listPanels: IPropertyPaneGroup[] =[];
    //const page2 : IPropertyPanePage["groups"]=[];
    //const page3 : IPropertyPanePage["groups"]=[];

    for(let x=1; x<=this.properties.numGroups;x++){
      if(this.properties.useList){
        linkGroups.push(PropertyPaneTextField('groupTitle'+x, {
            label: `Link Group ${x} Name`,
            value: `Link Group ${x}`,
            placeholder: "Please enter the link group name"  //,"description": "Name property field"
          })
        )
      }else{
        linkGroups.push(PropertyPaneTextField('groupTitle'+x, {
            label: `Link Group ${x} Name`,
            value: `Link Group ${x}`,
            placeholder: "Please enter the link group name"  //,"description": "Name property field"
          }),        
          PropertyPaneSlider('numLinks'+x, {
            label:'How Many Links for this group? (max 15)',
            min:1,
            max:15,
            value:1
          }) 
        )
      }

      switch(x){
        case 1:
          this.properties.groupTitle1 = `Link Group ${x}`;
          break;
        case 2:
          this.properties.groupTitle2 = `Link Group ${x}`;
          break;
        case 3:
          this.properties.groupTitle3 = `Link Group ${x}`;
          break;
        case 4:
          this.properties.groupTitle4 = `Link Group ${x}`;
          break;
        case 5:
          this.properties.groupTitle5 = `Link Group ${x}`;
          break;
        case 6:
          this.properties.groupTitle6 = `Link Group ${x}`;
          break;
        case 7:
          this.properties.groupTitle7 = `Link Group ${x}`;
          break;
        case 8:
          this.properties.groupTitle8 = `Link Group ${x}`;
          break;
        case 9:
          this.properties.groupTitle9 = `Link Group ${x}`;
          break;
        case 10:
          this.properties.groupTitle10 = `Link Group ${x}`;
          break;            
      }
    }

    if(this.properties.useList){
      linkList=PropertyPaneButton('', {
        text: "Open List",
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.buttonClick.bind(this) 
      })
    }else{
      linkList={};
      for(let lnk=1;lnk<=this.properties.numLinks1;lnk++){
        group1Links.push(PropertyPaneTextField('linkGroup1_Title'+lnk, {
            label: `Link Title ${lnk}`,
            value: `Link Title ${lnk}`,
            placeholder: "Please enter the link title"  //,"description": "Name property field"
          }),
          PropertyPaneTextField('linkGroup1_URL'+lnk, {
            label: `Link URL ${lnk}`,
            value: `Link URL ${lnk}`,
            placeholder: "Please enter the link URL"  //,"description": "Name property field"
          }),
          PropertyPaneDropdown('linkGroup1_Browse'+lnk, {
            label:'Please choose link browse option',
            options: [
              { key: '_self', text: 'Current Tab' },
              { key: '_blank', text: 'New Tab' },
              { key: '_parent', text: 'Current Browser - New Window' },
              { key: '_top', text: 'New Browser' },
            ]
          }),
          PropertyPaneHorizontalRule()      
        )
      }
    }
    
    //page2.push({groupName:"Link Groups",groupFields : linkGroups});
  
    for (var x = 1; x <= this.properties.numGroups; x++) {

      var singlePanel: IPropertyPaneGroup = {
        groupName: "Panel"+x,
        isCollapsed: true,
        groupFields: [
          PropertyPaneHorizontalRule(),
          PropertyPaneTextField('Title', {
            label: 'Title Panel',
            placeholder: 'Insert Title',
          }),
          PropertyPaneTextField('ImgPath', {
            label: 'Image Path',
            placeholder: 'Insert Path Image URL',
          }),
          PropertyPaneTextField('Link', {
            label: 'Link Image',
            placeholder: 'Insert Link',
          }),
          PropertyPaneDropdown('WidthValue', {
            label: 'Width Value',
            disabled: false,
            options: [{key: 'One', text: 'One'}, {key: 'Two', text: 'Two'}, {key: 'Three', text: 'Three'}]
          }),
        ]
      };
      listPanels.push(singlePanel);
      console.log(listPanels.length);
  }
/*    
      groups: [
        {
          groupName: "Link Groups",
          groupFields: linkGroups
        }
      ]
    )
*/

    // setup pages based on number of groups selected.
    
    return {
      pages: [
        {
          header: {
            description: "Page 1 - App Setup"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('App title for page', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider('numGroups', {
                  label:'How Many Link Groups? (max 10)',
                  min:1,
                  max:10,
                  value:1
                }),
                PropertyPaneCheckbox('useList', {
                  text: 'Use SharePoint List as link data?'
                }),
                linkList  
              ]
            }
          ]
        },
        { //Page 2
          displayGroupsAsAccordion: true,
          header : {
            description : "Page 2 = Groups Setup"
          },
          groups : listPanels
        },
        { //Page 3
          header: {
            description: "Page 3 – Group 1 Setup"
          },
          groups: [
            {
              groupName: "Group Links",
              groupFields: group1Links
            }
          ]
        }

      ]
    };
  }
}