import * as React from 'react';
import styles from './DevScopeGroupsSearch.module.scss';
import { IDevScopeGroupsSearchProps } from './IDevScopeGroupsSearchProps';
import { IGraphConsumerState } from './IGraphConsumerState';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ServiceScope } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import { IGraphGroup } from './IGraphGroup';
import * as $ from "jquery";
import { PopupWindowPosition } from '@microsoft/sp-property-pane/lib/propertyPaneFields/propertyPaneLink/IPropertyPaneLink';
import 'datatables.net';
require('datatables.net-buttons');
require('datatables.net-responsive');
require('datatables.net-select');

const logo: any = require('../assets/groupssearchlogo.png');
const loading: any = require('../assets/loading.gif');
const powerBIUrl: string = "https://app.powerbi.com/";
const powerBiImg: any = require("../assets/powerBI.png");
const oneNoteImg: any = require("../assets/onenote.png");
const teamsImg: any = require("../assets/teams.png");
const mailImg: any = require("../assets/mail.png");

export default class DevScopeGroupsSearch extends React.Component<IDevScopeGroupsSearchProps, IGraphConsumerState, {}> {

  constructor(props: IDevScopeGroupsSearchProps, state: IGraphConsumerState) {
    super(props);
    
    // Initialize the state of the component
    this.state = { 
      groups: [],
      searchFor: ""
    };
  }

  public render(): React.ReactElement<IDevScopeGroupsSearchProps> {
    let dtCssUrl = "//cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css";
    SPComponentLoader.loadCss(dtCssUrl);
  
    return (
      <div>
        <div id="content-header">
            <div className={"padding"}>
                <p>
                    <img src={logo} style={{verticalAlign: "middle"}} />
                    <span style={{marginBottom:"50px"}}>Search Office 365 Groups</span>
                </p>
            </div>
        </div>
        <div id="content-main">
            <table id="groupsTable" className={"display no-wrap"} style={{width:"100%"}}>
                <thead>
                    <tr>
                        <th>Group</th>
                        <th>Email</th>
                        <th>Team</th>
                        <th>PowerBI</th>
                        <th>OneNote</th>
                    </tr>
                </thead>
                <tfoot>
                    <tr>
                        <th>Group</th>
                        <th>Email</th>
                        <th>Team</th>
                        <th>PowerBI</th>
                        <th>OneNote</th>
                    </tr>
                </tfoot>
            </table>
        </div>
        <div id="loading" style={{display:"none"}}>
            <img src={loading} alt="Loading" /><br />
            Loading...
        </div>
      </div>
    );
  }

  public componentDidMount(): void{
    this._searchGroups();
  } 

  private _getGroups(): Promise<any[]> {
    return new Promise((resolve, reject) => {
    var self = this;
    // Log the current operation
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("groups")
          .version("v1.0")
          .filter(`groupTypes/any(c:c+eq+'Unified')&$top=` +  this.props.itemnumberproperty)
          .get((err, res) => {  

            if (err) {
              console.error(err);
              reject(err);
            }

            // Prepare the output array
            var groups: Array<IGraphGroup> = new Array<IGraphGroup>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              groups.push( { 
                groupTitle: item.displayName,
                groupMail: item.mail,
                groupId: item.id
              });
            });

            resolve(groups);
          });
        });
      });
  }

  private _searchGroups(): void {
    var self = this;

    this._getGroups().then((res) => {
      var table = $('#groupsTable').DataTable({
        data: res,
        responsive: true,
        "columns": [
            {
                data: null, render: function (data, type, row) {
                    return '<a href="#" id="groupLink" groupId="' + data.groupId + '" >' + data.groupTitle + '</a>'
                }
            },
            {
                data: null, render: function (data, type, row) {
                    return "<a href=mailto:'" + data.groupMail + "'><img src='" + mailImg + "' /></a>"
                },
                className: 'dt-body-center'
            },
            {
              data: null, render: function (data, type, row) {
                  return "<a href='#' id='teamsLink' groupId='" + data.groupId + "' alt='_blank'><img src='" + teamsImg + "' /></a>"
              },
              className: 'dt-body-center'
            },
            {
                data: null, render: function (data, type, row) {
                    return "<a href='" + powerBIUrl + "groups/" +  data.groupId + "' alt='_blank'><img src='" + powerBiImg + "' /></a>"
                },
                className: 'dt-body-center'
            },
            {
                data: null, render: function (data, type, row) {
                    return "<a href='#' id='oneNoteLink' groupId='" + data.groupId + "' alt='_blank'><img src='" + oneNoteImg + "' /></a>"
                },
                className: 'dt-body-center'
            }
        ],
        "order": [[ self.props.orderbyfieldproperty, self.props.ordermodefieldproperty ]]
      });
      
      $('#groupsTable').delegate('#groupLink', 'click', function (el) {
        var groupId = el.target.attributes["groupId"].value;
        self.openUrl(groupId, "Group");
      });

      $('#groupsTable').delegate('#oneNoteLink', 'click', function (el) {
          var groupId = el.target.parentElement.attributes["groupId"].value;
          self.openUrl(groupId, "OneNote");
      });

       $('#groupsTable').delegate('#teamsLink', 'click', function (el) {
          var groupId = el.target.parentElement.attributes["groupId"].value;
          self.openTeamsUrl(groupId);
      });
    });
  }

  private openUrl(groupId: string,  linkType: string): void {

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("groups/" + groupId + "/drive/root/webUrl")
          .version("v1.0")
          .get((err, res) => {  

            if (err) {
              alert("You do not have access to this resource or does not exist.");
              return;
            }
            
            var spUrl = res.value;
            if(linkType == "OneNote")
            {
              spUrl = spUrl.replace("Shared%20Documents", "_layouts/15/groupstatus.aspx?Target=NOTEBOOK");
            }
            window.open(spUrl, '_blank');
          })
        });
  }

  private openTeamsUrl(groupId: string): void {

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("groups/" + groupId + "/team")
          .version("v1.0")
          .get((err, res) => {  

            if (err) {
              alert("You do not have access to this Teams or does not exist.");
              return;
            }
            
            var spUrl = res.webUrl;
            window.open(spUrl, '_blank');
          })
        });
  }
}
