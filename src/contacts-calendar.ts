import { JmapClient, type JmapRequest } from './jmap-client.js';

export class ContactsCalendarClient extends JmapClient {
  private async checkContactsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:contacts'];
  }

  private async checkCalendarsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:calendars'];
  }

  async getContacts(limit = 50): Promise<Record<string, unknown>[]> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error(
        'Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.',
      );
    }

    const session = await this.getSession();

    // Try CardDAV namespace first, then Fastmail specific
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        [
          'ContactCard/query',
          {
            accountId: session.accountId,
            limit,
          },
          'query',
        ],
        [
          'ContactCard/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
            properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes'],
          },
          'contacts',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      // Fallback: try to get contacts using AddressBook methods
      const fallbackRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
        methodCalls: [
          [
            'AddressBook/get',
            {
              accountId: session.accountId,
            },
            'addressbooks',
          ],
        ],
      };

      try {
        const fallbackResponse = await this.makeRequest(fallbackRequest);
        return this.getListResult(fallbackResponse, 0);
      } catch (fallbackError) {
        throw new Error(
          `Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`,
        );
      }
    }
  }

  async getContactById(id: string): Promise<Record<string, unknown>> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error(
        'Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.',
      );
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        [
          'ContactCard/get',
          {
            accountId: session.accountId,
            ids: [id],
          },
          'contact',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      const contact = this.getListResult(response, 0)[0];
      if (!contact) {
        throw new Error(`Contact with ID '${id}' not found`);
      }
      return contact;
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('Contact with ID')) {
        throw error;
      }
      throw new Error(
        `Contact access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`,
      );
    }
  }

  async searchContacts(query: string, limit = 20): Promise<Record<string, unknown>[]> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error(
        'Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.',
      );
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        [
          'ContactCard/query',
          {
            accountId: session.accountId,
            filter: { text: query },
            limit,
          },
          'query',
        ],
        [
          'ContactCard/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
            properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes'],
          },
          'contacts',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(
        `Contact search not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`,
      );
    }
  }

  async getCalendars(): Promise<Record<string, unknown>[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error(
        'Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.',
      );
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        [
          'Calendar/get',
          {
            accountId: session.accountId,
          },
          'calendars',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 0);
    } catch (error) {
      // Calendar access might require special permissions
      throw new Error(
        `Calendar access not supported or requires additional permissions. This may be due to account settings or JMAP scope limitations: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`,
      );
    }
  }

  async getCalendarEvents(calendarId?: string, limit = 50): Promise<Record<string, unknown>[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error(
        'Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.',
      );
    }

    const session = await this.getSession();

    const filter = calendarId ? { inCalendar: calendarId } : {};

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        [
          'CalendarEvent/query',
          {
            accountId: session.accountId,
            filter,
            sort: [{ property: 'start', isAscending: true }],
            limit,
          },
          'query',
        ],
        [
          'CalendarEvent/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'CalendarEvent/query', path: '/ids' },
            properties: ['id', 'title', 'description', 'start', 'end', 'location', 'participants'],
          },
          'events',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(
        `Calendar events access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`,
      );
    }
  }

  async getCalendarEventById(id: string): Promise<Record<string, unknown>> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error(
        'Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.',
      );
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        [
          'CalendarEvent/get',
          {
            accountId: session.accountId,
            ids: [id],
          },
          'event',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      const event = this.getListResult(response, 0)[0];
      if (!event) {
        throw new Error(`Calendar event with ID '${id}' not found`);
      }
      return event;
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('Calendar event with ID')) {
        throw error;
      }
      throw new Error(
        `Calendar event access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`,
      );
    }
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string; // ISO 8601 format
    end: string; // ISO 8601 format
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
  }): Promise<string> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error(
        'Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.',
      );
    }

    const session = await this.getSession();

    const eventObject = {
      calendarId: event.calendarId,
      title: event.title,
      description: event.description || '',
      start: event.start,
      end: event.end,
      location: event.location || '',
      participants: event.participants || [],
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:calendars'],
      methodCalls: [
        [
          'CalendarEvent/set',
          {
            accountId: session.accountId,
            create: { newEvent: eventObject },
          },
          'createEvent',
        ],
      ],
    };

    try {
      const response = await this.makeRequest(request);
      const result = this.getMethodResult(response, 0);
      const eventId = result.created?.newEvent?.id;
      if (!eventId) {
        throw new Error('Calendar event creation returned no event ID');
      }
      return eventId;
    } catch (error) {
      throw new Error(
        `Calendar event creation not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`,
      );
    }
  }
}
