import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TasksCalenderWebPart.module.scss';
import * as strings from 'TasksCalenderWebPartStrings';

var $: any = require('jquery');
var moment: any = require('moment');

import 'fullcalendar';

var COLORS = ['#466365', '#B49A67', '#93B7BE', '#E07A5F', '#849483', '#084C61', '#DB3A34'];

/* export interface ITasksCalenderWebPartProps {
  description: string;
} */

import {ITasksCalendarWebPartProps} from './ITasksCalendarWebPartProps';

export default class TasksCalenderWebPart extends BaseClientSideWebPart<ITasksCalendarWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    
    <link type="text/css" rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.4.0/fullcalendar.min.css" />
    <div id="calendar"></div>
  </div>`;
  
  this.displayTasks();
  }
     
    private displayTasks() { 
      debugger
      $('#calendar').fullCalendar('destroy');
      $('#calendar').fullCalendar({
        weekends: false,
        header: {
          left: 'prev,next today',
          center: 'title',
          right: 'month,basicWeek,basicDay'
        },
        displayEventTime: false,
        // open up the display form when a user clicks on an event
        eventClick: (calEvent, jsEvent, view) => {
          (window as any).location = this.context.pageContext.web.absoluteUrl +
            "/Lists/" + escape('Calender') + "/DispForm.aspx?ID=" + calEvent.id;
        },
        editable: true,
        timezone: "UTC",
        droppable: true, // this allows things to be dropped onto the calendar
        // update the end date when a user drags and drops an event 
        eventDrop: (event, delta, revertFunc) => {
          this.updateTask(event.id, event.start, event.end);
        },
        // put the events on the calendar 
        events: (start, end, timezone, callback) => {
          var startDate = start.format('YYYY-MM-DD');
          var endDate = end.format('YYYY-MM-DD');
  
          var restQuery: string = `/_api/Web/Lists/GetByTitle('${escape('Calender')}')/items?$filter=((EventDate ge '${startDate}' and EventDate le '${endDate}')or(EndDate ge '${startDate}' and EndDate le '${endDate}'))`;
  
          $.ajax({
            url: this.context.pageContext.web.absoluteUrl + restQuery,
            type: "GET",
            dataType: "json",
            headers: {
              Accept: "application/json;odata=nometadata"
            }
          })
            .done((data, textStatus, jqXHR) => {
              var personColors = {};
              var colorNo = 0;
  
              var events = data.value.map((task) => {
                var color ;
                if (!color) {
                  color = COLORS[colorNo++];
                }
                if (colorNo >= COLORS.length) {
                  colorNo = 0;
                }
  
                return {
                  title: task.Title,
                  id: task.ID,
                  color: color, // specify the background color and border color can also create a class and use className parameter. 
                  start: moment.utc(task.EventDate).add("1", "days"),
                  end: moment.utc(task.EndDate).add("1", "days") // add one day to end date so that calendar properly shows event ending on that day
                };
              });
  
              callback(events);
            });
        }
      });
    }
  
    private updateTask(id, startDate, dueDate) {
      // subtract the previously added day to the date to store correct date
      var sDate = moment.utc(startDate).add("-1", "days").format('YYYY-MM-DD') + "T" +
        startDate.format("hh:mm") + ":00Z";
      if (!dueDate) {
        dueDate = startDate;
      }
      var dDate = moment.utc(dueDate).add("-1", "days").format('YYYY-MM-DD') + "T" +
        dueDate.format("hh:mm") + ":00Z";
  
      $.ajax({
        url: this.context.pageContext.web.absoluteUrl + '/_api/contextinfo',
        type: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata'
        }
      })
        .then((data, textStatus, jqXHR) => {
          return $.ajax({
            url: this.context.pageContext.web.absoluteUrl +
            "/_api/Web/Lists/getByTitle('" + escape('Calender') + "')/Items(" + id + ")",
            type: 'POST',
            data: JSON.stringify({
              EventDate: sDate,
              EndDate: dDate,
            }),
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "X-RequestDigest": data.FormDigestValue,
              "IF-MATCH": "*",
              "X-Http-Method": "PATCH"
            }
          });
        })
        .done((data, textStatus, jqXHR) => {
          alert("Update Successful");
        })
        .fail((jqXHR, textStatus, errorThrown) => {
          alert("Update Failed");
        })
        .always(() => {
          this.displayTasks();
        });
    }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('listName', {
                  label: "List Name",
                  value:"Calender"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
