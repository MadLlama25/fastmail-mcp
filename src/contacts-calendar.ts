import { JmapClient, JmapRequest, QueryResult } from './jmap-client.js';

export class ContactsCalendarClient extends JmapClient {
  
  private async checkContactsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:contacts'];
  }
  
  private async checkCalendarsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:calendars'];
  }

  /** Contacts may live on a different primary account than mail. */
  private async contactsAccountId(): Promise<string> {
    const session = await this.getSession();
    return session.primaryAccounts?.['urn:ietf:params:jmap:contacts'] ?? session.accountId;
  }
  
  async getContacts(limit: number = 50): Promise<QueryResult> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();

    // Try CardDAV namespace first, then Fastmail specific
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/query', {
          accountId: session.accountId,
          limit,
          calculateTotal: true
        }, 'query'],
        ['ContactCard/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getQueryResult(response, 0, 1);
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
        const items = this.getListResult(fallbackResponse, 0);
        return { items };
      } catch (fallbackError) {
        throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
      }
    }
  }

  async getContactById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/get', {
          accountId: session.accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    let contact;
    try {
      const response = await this.makeRequest(request);
      contact = this.getListResult(response, 0)[0];
    } catch (error) {
      throw new Error(`Contact access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
    // ContactCard/get reports unknown ids via notFound, leaving list empty — a
    // bare undefined here used to serialize as a successful empty tool response.
    if (!contact) {
      throw new Error(`Contact not found: ${id}`);
    }
    return contact;
  }

  async searchContacts(query: string, limit: number = 20): Promise<QueryResult> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/query', {
          accountId: session.accountId,
          filter: { text: query },
          limit,
          calculateTotal: true
        }, 'query'],
        ['ContactCard/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
          properties: ['id', 'name', 'emails', 'phones', 'addresses', 'notes']
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getQueryResult(response, 0, 1);
    } catch (error) {
      throw new Error(`Contact search not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  async getCalendars(): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

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
      return this.getListResult(response, 0);
    } catch (error) {
      // Calendar access might require special permissions
      throw new Error(`Calendar access not supported or requires additional permissions. This may be due to account settings or JMAP scope limitations: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50): Promise<any[]> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

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
      return this.getListResult(response, 1);
    } catch (error) {
      throw new Error(`Calendar events access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  async getCalendarEventById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

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
      return this.getListResult(response, 0)[0];
    } catch (error) {
      throw new Error(`Calendar event access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
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
    // Check permissions first
    const hasPermission = await this.checkCalendarsPermission();
    if (!hasPermission) {
      throw new Error('Calendar access not available. This account may not have JMAP calendar permissions enabled. Please check your Fastmail account settings or contact support to enable calendar API access.');
    }

    const session = await this.getSession();

    const eventObject: Record<string, unknown> = {
      calendarId: event.calendarId,
      title: event.title,
      description: event.description || '',
      start: event.start,
      end: event.end,
      location: event.location || '',
    };
    // TODO: participants should be an RFC 8984 object/map, not an array.
    // TODO: startDate/endDate not passed through to JMAP.
    if (event.participants?.length) {
      eventObject.participants = event.participants;
    }

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
      const result = this.getMethodResult(response, 0);
      const eventId = result.created?.newEvent?.id;
      if (!eventId) {
        throw new Error('Calendar event creation returned no event ID');
      }
      return eventId;
    } catch (error) {
      throw new Error(`Calendar event creation not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling calendar API access in Fastmail settings.`);
    }
  }

  // ---------- contacts write (JMAP ContactCard/set, RFC 9610) ----------
  //
  // A live probe confirmed Fastmail accepts ContactCard/set with an RFC 9610
  // Card shape; the server assigns the default address book, uid, and prodId.
  // Note: creation-id references ("#id") are NOT resolved in destroy arrays by
  // Fastmail's backend — always destroy by real id.

  /** Map the flat tool-facing input onto an RFC 9610 Card (arrays -> Id-maps). */
  private buildCardProperties(input: {
    name?: { given?: string; surname?: string; full?: string };
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    addresses?: Array<{ full: string; label?: string }>;
    notes?: string;
  }): Record<string, any> {
    const card: Record<string, any> = {};

    if (input.name) {
      const components: Array<{ kind: string; value: string }> = [];
      if (input.name.given) components.push({ kind: 'given', value: input.name.given });
      if (input.name.surname) components.push({ kind: 'surname', value: input.name.surname });
      card.name = {
        ...(components.length && { components }),
        ...(input.name.full && { full: input.name.full }),
      };
    }
    const toIdMap = (items: any[] | undefined, prefix: string) => {
      if (!items?.length) return undefined;
      const map: Record<string, any> = {};
      items.forEach((item, i) => { map[`${prefix}${i}`] = item; });
      return map;
    };
    const emails = toIdMap(input.emails, 'e');
    const phones = toIdMap(input.phones, 'p');
    const addresses = toIdMap(input.addresses, 'a');
    if (emails) card.emails = emails;
    if (phones) card.phones = phones;
    if (addresses) card.addresses = addresses;
    if (input.notes) card.notes = { n0: { note: input.notes } };

    return card;
  }

  async createContact(input: {
    name?: { given?: string; surname?: string; full?: string };
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    addresses?: Array<{ full: string; label?: string }>;
    notes?: string;
    addressBookId?: string;
  }): Promise<string> {
    const hasName = !!(input.name?.full || input.name?.given || input.name?.surname);
    if (!hasName && !input.emails?.length) {
      throw new Error('A contact needs a name or at least one email address');
    }

    const accountId = await this.contactsAccountId();
    const card: Record<string, any> = {
      '@type': 'Card',
      version: '1.0',
      ...this.buildCardProperties(input),
      ...(input.addressBookId && { addressBookIds: { [input.addressBookId]: true } }),
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', { accountId, create: { newContact: card } }, 'createContact'],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    if (result.notCreated?.newContact) {
      const err = result.notCreated.newContact;
      throw new Error(`Failed to create contact: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }
    const id = result.created?.newContact?.id;
    if (!id) {
      throw new Error('Contact creation returned no id');
    }
    return id;
  }

  async updateContact(id: string, patch: {
    name?: { given?: string; surname?: string; full?: string };
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    addresses?: Array<{ full: string; label?: string }>;
    notes?: string;
    expectState?: string;
  }): Promise<void> {
    const { expectState, ...fields } = patch;
    const patchObject = this.buildCardProperties(fields);
    if (Object.keys(patchObject).length === 0) {
      throw new Error('At least one field to update must be provided (name, emails, phones, addresses, or notes)');
    }

    const accountId = await this.contactsAccountId();

    // Existence check first, for a clean not-found error (repo convention).
    const getResponse = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [['ContactCard/get', { accountId, ids: [id], properties: ['id'] }, 'g']],
    });
    if (!this.getListResult(getResponse, 0)[0]) {
      throw new Error(`Contact not found: ${id}`);
    }

    // JMAP PatchObject semantics: each provided top-level field wholly
    // replaces the stored value (e.g. emails: [] clears all emails).
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId,
          update: { [id]: patchObject },
          ...(expectState && { ifInState: expectState }),
        }, 'updateContact'],
      ],
    });
    const result = this.getMethodResult(response, 0);
    if (result.notUpdated?.[id]) {
      const err = result.notUpdated[id];
      throw new Error(`Failed to update contact: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }
  }

  async deleteContact(id: string, expectState?: string): Promise<void> {
    const accountId = await this.contactsAccountId();
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId,
          destroy: [id],
          ...(expectState && { ifInState: expectState }),
        }, 'deleteContact'],
      ],
    });
    const result = this.getMethodResult(response, 0);
    if (result.notDestroyed?.[id]) {
      const err = result.notDestroyed[id];
      if (err.type === 'notFound') {
        throw new Error(`Contact not found: ${id}`);
      }
      throw new Error(`Failed to delete contact: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }
  }
}
