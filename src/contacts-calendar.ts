import { JmapClient } from './jmap-client.js';

export class ContactsCalendarClient extends JmapClient {
  
  async getContacts(limit: number = 50): Promise<any[]> {
    const session = await this.getSession();
    
    const request = {
      using: ['urn:ietf:params:jmap:core', 'com:fastmail:contacts'],
      methodCalls: [
        ['Contact/query', {
          accountId: session.accountId,
          limit
        }, 'query'],
        ['Contact/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Contact/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[1][1].list;
  }

  async getContactById(id: string): Promise<any> {
    const session = await this.getSession();
    
    const request = {
      using: ['urn:ietf:params:jmap:core', 'com:fastmail:contacts'],
      methodCalls: [
        ['Contact/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].list[0];
  }

  async searchContacts(query: string, limit: number = 20): Promise<any[]> {
    const session = await this.getSession();
    
    const request = {
      using: ['urn:ietf:params:jmap:core', 'com:fastmail:contacts'],
      methodCalls: [
        ['Contact/query', {
          accountId: session.accountId,
          filter: { text: query },
          limit
        }, 'query'],
        ['Contact/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Contact/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[1][1].list;
  }

  async getCalendars(): Promise<any[]> {
    const session = await this.getSession();
    
    const request = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['Calendar/get', {
          accountId: session.accountId
        }, 'calendars']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].list;
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50): Promise<any[]> {
    const session = await this.getSession();
    
    const filter = calendarId ? { inCalendar: calendarId } : {};
    
    const request = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'start', isAscending: true }],
          limit
        }, 'query'],
        ['CalendarEvent/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'CalendarEvent/query', path: '/ids' },
          properties: ['id', 'title', 'description', 'start', 'end', 'location', 'participants']
        }, 'events']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[1][1].list;
  }

  async getCalendarEventById(id: string): Promise<any> {
    const session = await this.getSession();
    
    const request = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'event']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].list[0];
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string; // ISO 8601 format
    end: string;   // ISO 8601 format
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
  }): Promise<string> {
    const session = await this.getSession();

    const eventObject = {
      calendarId: event.calendarId,
      title: event.title,
      description: event.description || '',
      start: event.start,
      end: event.end,
      location: event.location || '',
      participants: event.participants || []
    };

    const request = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          create: { newEvent: eventObject }
        }, 'createEvent']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].created.newEvent.id;
  }
}