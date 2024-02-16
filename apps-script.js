/**
 * This code is meant to enrich Calendar events right after they get created/updated.
 * 
 * Installation
 *  - create a new Google Apps Script project;
 *  - copy/paste the code from this file;
 *  - in the Apps Script editor, add: Services > Google Calendar API;
 *  - create a Google Calendar trigger.
 * 
 * The source of this file is at: https://github.com/puntonim/google-calendar-rich/blob/main/apps-script.js
 */

class BaseError extends Error {
    constructor(message) {
      super(message);
      this.name = this.constructor.name;
    }
  }
  
  class NoEventFound extends BaseError { }
  class EventNull extends BaseError { }
  class EventTypeNull extends BaseError { }
  
  
  const onCalendarUpdate = (triggerData) => {
    /**
     * Executed via a trigger defined on every Google Calendar update.
     * 
     * Args:
     *    triggerData (object): eg. {calendarId=pun...@gmail.com, triggerUid=21985234, authMode=FULL}
     */
    Logger.log("START onCalendarUpdate");
    
    // Notice that unfortunately the params passed to this function by the trigger 
    //  are {calendarId=pun...@gmail.com, triggerUid=21985234, authMode=FULL} and 
    //  do NOT include the eventId. So we need to get the latest eventId by ourselves.
    let event;
    try {
      event = _getLastEditedEvent(triggerData.calendarId);
    } catch (err) {
      if (err instanceof NoEventFound) {
        Logger.log("No events found, probably the update was a delete event");
        return;
      } else throw err;
    }
  
    let title = event.getTitle();
    const description = event.getDescription();
    Logger.log(title, description);
  
    // Do not edit anything if the description includes ":skip:".
    if (description.includes(":skip:")) return;
  
    const eventType = _updateEventTitleAndGetType(event);
    if (eventType) _processEventByType({event, eventType});
  
    Logger.log("END onCalendarUpdate");
  };
  
  
  const titleEnrichmentsAndTypes = {
    ":run:": {"emoji": "🏃‍♂️", "eventType": "RUN"},
  
    ":bike:": {"emoji": "🚴‍♂️", "eventType": "BIKE"},
  
    ":birthday:": {"emoji": "🎉", "eventType": "BIRTHDAY"},
    ":birth:": {"emoji": "🎉", "eventType": "BIRTHDAY"},
    ":compleanno:": {"emoji": "🎉", "eventType": "BIRTHDAY"},
    ":comple:": {"emoji": "🎉", "eventType": "BIRTHDAY"},
  
    ":dinner:": {"emoji": "🍗", "eventType": "DINNER"},
    ":lunch:": {"emoji": "🍗", "eventType": "DINNER"},
    ":cena:": {"emoji": "🍗", "eventType": "DINNER"},
    ":pranzo:": {"emoji": "🍗", "eventType": "DINNER"},
  
    ":hospital:": {"emoji": "🏥", "eventType": "HEALTH"},
    ":hosp:": {"emoji": "🏥", "eventType": "HEALTH"},
    ":ospedale:": {"emoji": "🏥", "eventType": "HEALTH"},
    ":osp:": {"emoji": "🏥", "eventType": "HEALTH"},
  
    ":workout:": {"emoji": "💪", "eventType": "GYM"},
    ":gym:": {"emoji": "💪", "eventType": "GYM"},
    ":muscle:": {"emoji": "💪", "eventType": "GYM"},
    ":bicep:": {"emoji": "💪", "eventType": "GYM"},
    ":biceps:": {"emoji": "💪", "eventType": "GYM"},
  
    ":star:": {"emoji": "⭐", "eventType": null},
    ":check:": {"emoji": "✅", "eventType": null},
    ":ita:": {"emoji": "🇮🇹", "eventType": null},
    ":party:": {"emoji": "🎉", "eventType": null},
    ":chicken:": {"emoji": "🍗", "eventType": null},
    ":cross:": {"emoji": "✝️", "eventType": null},
    ":death:": {"emoji": "✝️", "eventType": null},
    ":$:": {"emoji": "💰", "eventType": null},
    ":money:": {"emoji": "💰", "eventType": null},
    ":dollar:": {"emoji": "💰", "eventType": null},
  }
  
  
  const _updateEventTitleAndGetType = (event) => {
    /**
     * Update a CalendarEvent title and identify its known type.
     * 
     * Args:
     *    event (CalendarEvent): see https://developers.google.com/apps-script/reference/calendar/calendar-event.
     */
    let origTitle = event.getTitle()
    let newTitle = origTitle;
    let eventType = null;
  
    // Search for all the :xxx: patterns in the title (eg. ":run:").
    [...origTitle.matchAll(/:[a-z$]+:/g)].forEach((token) => {
      const data = titleEnrichmentsAndTypes[token];
      if (!(data)) return;  // `return` in `forEach` is like a `continue` in a `for` loop.
  
      // Replace the pattern eg. ":run:" -> "🏃‍♂️".
      newTitle = newTitle.replaceAll(token, data.emoji);
      if (!(eventType)) eventType = data.eventType;
    });
    
    Logger.log(`Event type: ${eventType}`);
    if (newTitle !== origTitle) {
      Logger.log(`Setting title to: ${newTitle}`);
      event.setTitle(newTitle);
    }
  
    return eventType;
  };
  
  
  const _processEventByType = ({event, eventType = null}) => {
    /**
     * Process a CalendarEvent by its known type and perform actions like changing its color.
     * 
     * Args:
     *    event (CalendarEvent): see https://developers.google.com/apps-script/reference/calendar/calendar-event.
     */
    if (!(event)) throw new EventNull();
    if (!(eventType)) throw new EventTypeNull();
  
    // Colors: https://developers.google.com/apps-script/reference/calendar/event-color.
    let color = null;
  
    switch (eventType) {
      case "RUN":
        color = CalendarApp.EventColor.GRAY;
        break;
      case "BIKE":
        color = CalendarApp.EventColor.GRAY;
        break;
      case "BIRTHDAY":
        color = CalendarApp.EventColor.YELLOW;
        break;
      case "DINNER":
        color = CalendarApp.EventColor.BLUE;
        break;
      case "HEALTH":
        color = CalendarApp.EventColor.RED;
        break;
      case "GYM":
        color = CalendarApp.EventColor.PALE_RED;
        break;
    } 
  
    if (color) {
      Logger.log(`Setting color to ${color}...`);
      event.setColor(color);
    }
  };
  
  
  const _getLastEditedEvent = (calendarId) => {
    /**
     * Get the most recently updated event in Google Calendar.
     * 
     * Args:
     *    calendarId (str): eg. "pun...@gmail.com".
     * 
     * Return: CalendarEvent, docs: https://developers.google.com/apps-script/reference/calendar/calendar-event
     * 
     * Inspired by: https://stackoverflow.com/a/74266545/1969672
     */
    // List all events filtering by updatedMin = 1 min ago and sorting by "updated";
    const now = new Date();
    const oneMinAgo = new Date(now.getTime() - (60 * 1000));
    const options = {
      updatedMin: oneMinAgo.toISOString(),
      maxResults: 50,
      orderBy: "updated",
      singleEvents: true,
      showDeleted: false,
      fields: "items(id)",
    }
    const events = Calendar.Events.list(calendarId, options);
    if(!events.items) throw new NoEventFound("Calendar.Events.list() returned zero events");
  
    // Get the id of the first event (the most recently updated).
    const event = events.items[events.items.length-1];
    // event is null when the update was a delete of ann evennt from the calendar.
    if(!event) throw new NoEventFound("The last event in events.items is null");

    // Get and return the details of the event.
    return CalendarApp.getEventById(event.id);
  };
