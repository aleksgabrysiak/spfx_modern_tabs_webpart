
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ModernTabsWebPartStrings';
import * as $ from 'jquery';
import * as jQuery from 'jquery';
import PnPTelemetry from '@pnp/telemetry-js'; //Get PnPTelemetry to disable it below

import {  Web } from "@pnp/sp";

const telemetry = PnPTelemetry.getInstance();
telemetry.optOut();

export interface IModernTabsWebPartProps {
  description: string;
  sectionClass: string;
  webpartClass: string;
  tabData: any[];
  tabNames:any[];
  tabinit:any[];
}

export default class ModernTabsWebPart extends BaseClientSideWebPart<IModernTabsWebPartProps> {
  private usergroups:any[string]=[];
  public tabinit:any[]=[];

  public render(): void {
    require('./AddTabs.css');
    //require('./jquery.min.js');

    
    this._checkGroup().then(results=>{results.forEach(result=>{this.usergroups.push(result.Title)})}).then(()=>{
    
      
    if (this.displayMode === DisplayMode.Read)
    {
     
    var tabWebPartID = $(this.domElement).closest("div." + this.properties.webpartClass).attr("id");  
    var tabsDiv = tabWebPartID + "tabs";
    var contentsDiv = tabsDiv +"Contents";
      
    this.domElement.innerHTML = "<div data-addui='tabs'><div role='tabs' id='"+tabsDiv+"'></div><div role='contents' id='"+contentsDiv+"'></div></div>";

      var thisTabData = this.properties.tabData;
      const groupped =  this.groupBy(thisTabData,"TabLabel");

      Object.keys(groupped).forEach((key)=>{ 
        var tabName = this.properties.tabNames[Number(key)].TabLabel;
        var usergroupsstring = this.usergroups.toString();
        var shouldbehidden=false;

        this.properties.tabNames.every((tab)=>{ 
        if(tab.UserGroup && tab.TabLabel === tabName  && usergroupsstring.indexOf(tab.UserGroup)===-1) //decide on visibility of the tab by checking if user is in specific group (dirty way:())
        {
         shouldbehidden = true;
         return !shouldbehidden;
        }
        else{
          shouldbehidden=false;
          return !shouldbehidden;
        }
        }); 


        shouldbehidden?()=>{return false}: $("#"+tabsDiv).append("<div>"+tabName+"</div>");

        let firstWebPartId="";
        groupped[key].forEach((element:any,index:number) => {
            if(index==0){
            $("#"+contentsDiv).append($("#"+element.WebPartID));
            firstWebPartId = element.WebPartID.toString();
            }
            else{
             $("#"+firstWebPartId).children().append($("#"+element.WebPartID));
             $("#"+element.WebPartID).css('padding-bottom','10px');
            }
          });
          
      });
      {this.RenderTabs(0)};
     
    }else{
      this.domElement.innerHTML = `
      <div>
       
                <ul>
                  <li>Place this web part in the <b>same section</b> of the page as the web parts you would like to put into tabs.</li> 
                  <li>Add the web parts to the section and then edit the properties of this web part.</li>
                  <li>Click on the button  'Manage Tabs' to specify the labels for each tab and whether each tab should be visible for particular user's group (leave empty if should be visible for everyone).</li>
                  <li>Once Tabs are defined, click on 'Manage WebParts' and assign each WebPart to your Tabs. If visibility to specific group of the tab has been defined in previous step, you will see this information next to selected Tab name.</li>  
                  <li>You can place many webparts in the same tab.</li>              
                  </ul> 
                  <br/>Credits to <a href="https://github.com/mrackley/">orignal author of the Hillbilly Tabs.</a>
                  <br/>The script for rendering tabs as well as CSS has been developed by <a href="https://www.jqueryscript.net/other/Minimal-Handy-jQuery-Tabs-Plugin-AddTabs.html">dustinpoissant.</a>
                 
      </div>`;
      var zone=  $(this.domElement).closest("div." + this.properties.sectionClass)[0];
      zone.querySelectorAll("h4").forEach((header)=>{header.remove()});
      
      var initwebparts = this.getZones();
        initwebparts.forEach((webpart, index)=>{
        $("#"+webpart[0]).parent().prepend("<h4 style=\"color:red\">WebPart #"+(index+1)+" ↓</h4>");
      });

      if(zone){
        var observer = new MutationObserver(()=>{
        var webparts = this.getZones();
        zone.querySelectorAll("h4").forEach((header)=>{header.remove()});
        webparts.forEach((webpart, index)=>{
        $("#"+webpart[0]).parent().prepend("<h4 style=\"color:red\">WebPart #"+(index+1)+" ↓</h4>");  
      });
        });
        observer.observe(zone,{ childList: true });
      }
    }
  });
  }


  private async _checkGroup():Promise<any[]> {
    const web = new Web(this.context.pageContext.web.absoluteUrl);
    const groups:any[]= await web.currentUser.groups.get();
    return groups;
  }

  protected groupBy(arr:any, property:string) {
    return arr.reduce(function (memo:any, x:any) {
        if (!memo[x[property]]) { memo[x[property]] = []; }
        memo[x[property]].push(x);
        return memo;
    }, {});
  }

  protected async RenderTabs(firstindex:number)
  {
      //The script in this funciton and the css AddTabs was developed by dustinpoissant 
      //https://www.jqueryscript.net/other/Minimal-Handy-jQuery-Tabs-Plugin-AddTabs.html
          
     var active = firstindex;
     //@ts-ignore
      if(typeof($add)=="undefined")var $add={version:{Tabs:"1.1.0"},auto:{disabled:false}};
      (function($){
        //@ts-ignore
        $add.Tabs = function(selector, settings){
          $(selector).each(function(i, el){
            var $el = $(el);
            var S = $.extend({
              change: "click"
            }, $el.data(), settings);
            var $tabHolder = $el.find("[role=tabs]");
            $tabHolder.addClass("addui-Tabs-tabHolder");
            var $tabs = $tabHolder.children();
            var $contentHolder = $el.find("[role=contents]");
            $contentHolder.addClass("addui-Tabs-contentHolder");
            var $contents = $contentHolder.children();
            $el.addClass("addui-Tabs").attr("role", "").removeAttr("role");
            $tabs.addClass("addui-Tabs-tab");
            $contents.addClass("addui-Tabs-content").each(function(i, c){
              if($(c).hasClass("active")){
                $(c).removeClass("active");
                active = i;
              }
            });
            $contents.removeClass("addui-Tabs-active").eq(active).addClass("addui-Tabs-active");
            $tabs.removeClass("addui-Tabs-active").eq(active).addClass("addui-Tabs-active");
            var event = "click";
            if(S.change === "hover") event = "mouseenter";
            $tabs.on(event, function(e){
              $tabs.each(function(i, t){
                if(t === e.target){
                  active = i;
                  $contents.removeClass("addui-Tabs-active").eq(active).addClass("addui-Tabs-active");
                  $tabs.removeClass("addui-Tabs-active").eq(active).addClass("addui-Tabs-active");
                }
              });
            })
          });
          return this;
        };
        //@ts-ignore
        $.fn.addTabs = function(settings){$add.Tabs(this, settings);};
        //@ts-ignore
        $add.auto.Tabs = function(){
          if(!$add.auto.disabled){
            //@ts-ignore
            $("[data-addui=tabs]").addTabs();
          }
        }
      })(jQuery);
      $(function(){for(var k in $add.auto){if(typeof(
        //@ts-ignore
        $add.auto[k]
        )=="function"){
        //@ts-ignore
        $add.auto[k]();}}});
        
  }


 
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
    
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  

   private getTabs():Array<[number,string,string]>{
    const tabNamesR = new Array<[number,string,string]>();
    this.properties.tabNames?
    this.properties.tabNames.forEach((tab,index)=>{
      tabNamesR.push(
        [
         index,
         tab.TabLabel,
         tab.UserGroup
        ])
    }): tabNamesR.push([0,'','']) ; 
    return tabNamesR;
  }

  private getZones(): Array<[string,string]> {
    const zones = new Array<[string,string]>();

    let tabWebPartID = $(this.domElement).closest("div." + this.properties.webpartClass).attr("id");       
    let zoneDIV = $(this.domElement).closest("div." + this.properties.sectionClass);
    let count = 1;
    $(zoneDIV).find("."+this.properties.webpartClass).each(function(){
      let thisWPID = $(this).attr("id");
      if (thisWPID !== tabWebPartID)
      {
        let zoneId = $(this).attr("id");
        let zoneName:string = "Webpart #" + count + " (Id: "+thisWPID+")";
        count++;
        //@ts-ignore
        zones.push([zoneId, zoneName]);
      }
    });

    return zones;
  } 



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ""
          },
          groups: [
            {
              groupName: "Settings:",
              groupFields: [
                PropertyPaneTextField('sectionClass', {
                  label: "SectionClass:",
                  description: "Class identifier for Page Section.",
                  disabled: true
                }),
                PropertyPaneTextField('webpartClass', {
                  label: "WebPartClass:",
                  description: "Class identifier for Web Part.", 
                  disabled: true
                }),
                PropertyFieldCollectionData("tabNames", {
                  key: "tabNames",
                  label: "Manage Tabs & WebParts\n",
                  panelHeader: "Specify Labels for Tabs",
                  manageBtnLabel: "Manage Tabs",
                  value: this.properties.tabNames,
                  fields: [
                    {
                      id: "TabLabel",
                      title: "Tab Label",
                      type: CustomCollectionFieldType.string
                      
                    },
                    {
                      id: "UserGroup",
                      title: "Show Only for user in group",
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                  
                }),
                PropertyFieldCollectionData("tabData", {
                  key: "tabData",
                  label: strings.TabLabels,
                  panelHeader: "Assign WebParts to Tabs",
                  manageBtnLabel: "Manage WebParts",
                  value: this.properties.tabData,
                  fields: [
                    {
                      id: "WebPartID",
                      title: "Web Part:",
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this.getZones().map((zone:[string,string]) => {
                        return {
                          key: zone["0"],
                          text: zone["1"],
                        };
                      })

                    },
                    {
                      id: "TabLabel",
                      title: "Tab Label:",
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this.getTabs().map((tab:[number,string,string])=>{
                        var tabtext = tab['1'] + (tab['2'] && tab['2'].length>0 ? " (visible for group "+tab['2']+")":"");
                        return {
                          key: tab['0'],
                          text: tabtext
                        };})
                    }
                  ],
                  disabled: !(this.properties.tabNames && this.properties.tabNames.length)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
