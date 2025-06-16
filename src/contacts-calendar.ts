import { JmapClient, JmapRequest } from './jmap-client.js';

export class ContactsCalendarClient extends JmapClient {
  
  async getContacts(limit: number = 50): Promise<any[]> {
    const session = await this.getSession();
    
    // Try CardDAV namespace first, then Fastmail specific
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
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

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[1][1].list;
    } catch (error) {
      // Fallback: try to get contacts using AddressBook methods
      const fallbackRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
        methodCalls: [
          ['AddressBook/get', {
            accountId: session.accountId
          }, 'addressbooks']
        ]
      };
      
      try {
        const fallbackResponse = await this.makeRequest(fallbackRequest);
        return fallbackResponse.methodResponses[0][1].list || [];
      } catch (fallbackError) {
        throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  }

  async getContactById(id: string): Promise<any> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['Contact/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].list[0];
    } catch (error) {
      throw new Error(`Contact access not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async searchContacts(query: string, limit: number = 20): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
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

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[1][1].list;
    } catch (error) {
      throw new Error(`Contact search not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async getCalendars(): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['Calendar/get', {
          accountId: session.accountId
        }, 'calendars']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].list;
    } catch (error) {
      // Calendar access might require special permissions
      throw new Error(`Calendar access not supported or requires additional permissions. This may be due to account settings or JMAP scope limitations: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50): Promise<any[]> {
    const session = await this.getSession();
    
    const filter = calendarId ? { inCalendar: calendarId } : {};
    
    const request: JmapRequest = {
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

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[1][1].list;
    } catch (error) {
      throw new Error(`Calendar events access not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  async getCalendarEventById(id: string): Promise<any> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'event']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].list[0];
    } catch (error) {
      throw new Error(`Calendar event access not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
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

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        ['CalendarEvent/set', {
          accountId: session.accountId,
          create: { newEvent: eventObject }
        }, 'createEvent']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return response.methodResponses[0][1].created.newEvent.id;
    } catch (error) {
      throw new Error(`Calendar event creation not supported: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}