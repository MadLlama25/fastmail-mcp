import assert from 'node:assert/strict';
import { beforeEach, describe, it, mock } from 'node:test';
import {
  CalDAVCalendarClient,
  escapeICalText,
  extractVEvent,
  formatICalDate,
  parseCalendarObject,
  parseICalValue,
  unescapeICalText,
} from './caldav-client.js';

describe('extractVEvent', () => {
  it('extracts VEVENT block from iCalendar data', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'DTSTART:19700101T000000',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'SUMMARY:Test Event',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    assert.ok(vevent.includes('SUMMARY:Test Event'));
    assert.ok(vevent.includes('DTSTART;TZID=Europe/Rome:20260320T083000'));
    assert.ok(!vevent.includes('VTIMEZONE'));
    assert.ok(!vevent.includes('TZID:Europe/Rome'));
  });

  it('returns original data when no VEVENT block found', () => {
    const data = 'no vevent here';
    assert.equal(extractVEvent(data), data);
  });

  it('ignores VTIMEZONE DTSTART when extracting VEVENT', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'END:STANDARD',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'SUMMARY:Meeting',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    // Should only have the VEVENT DTSTART, not the VTIMEZONE one
    const dtstartMatches = vevent.match(/DTSTART/g);
    assert.equal(dtstartMatches?.length, 1);
    assert.ok(vevent.includes('20260320T083000'));
  });
});

describe('parseICalValue', () => {
  it('handles simple KEY:value format', () => {
    const vevent = 'SUMMARY:Test Event\nDTSTART:20260320T083000Z';
    assert.equal(parseICalValue(vevent, 'SUMMARY'), 'Test Event');
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000Z');
  });

  it('handles parameterized KEY;TZID=...:value format', () => {
    const vevent = 'DTSTART;TZID=Europe/Rome:20260320T083000\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000');
  });

  it('handles VALUE=DATE format', () => {
    const vevent = 'DTSTART;VALUE=DATE:20260324\nDTEND;VALUE=DATE:20260325';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260324');
    assert.equal(parseICalValue(vevent, 'DTEND'), '20260325');
  });

  it('returns undefined for missing keys', () => {
    const vevent = 'SUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'LOCATION'), undefined);
  });

  it('handles line folding (continuation lines)', () => {
    const vevent = 'DESCRIPTION:This is a long\n description that wraps\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), 'This is a longdescription that wraps');
  });
});

describe('formatICalDate', () => {
  it('formats datetime without timezone', () => {
    assert.equal(formatICalDate('20260320T083000'), '2026-03-20T08:30:00');
  });

  it('formats datetime with Z suffix', () => {
    assert.equal(formatICalDate('20260320T083000Z'), '2026-03-20T08:30:00Z');
  });

  it('formats all-day date', () => {
    assert.equal(formatICalDate('20260324'), '2026-03-24');
  });

  it('returns undefined for undefined input', () => {
    assert.equal(formatICalDate(undefined), undefined);
  });

  it('returns cleaned string for unrecognized formats', () => {
    assert.equal(formatICalDate('something-else'), 'something-else');
  });

  it('strips carriage returns', () => {
    assert.equal(formatICalDate('20260320T083000\r'), '2026-03-20T08:30:00');
  });
});

describe('parseCalendarObject', () => {
  it('parses a full calendar object with VTIMEZONE + VEVENT', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'TZOFFSETFROM:+0200',
      'TZOFFSETTO:+0100',
      'END:STANDARD',
      'BEGIN:DAYLIGHT',
      'DTSTART:19700329T020000',
      'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
      'TZOFFSETFROM:+0100',
      'TZOFFSETTO:+0200',
      'END:DAYLIGHT',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:abc123@fastmail',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'DTEND;TZID=Europe/Rome:20260320T093000',
      'SUMMARY:Morning Meeting',
      'DESCRIPTION:Discuss project\\nSecond line',
      'LOCATION:Room A\\, Building 1',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: 'https://caldav.example.com/cal/abc.ics' });

    assert.equal(event.id, 'abc123@fastmail');
    assert.equal(event.url, 'https://caldav.example.com/cal/abc.ics');
    assert.equal(event.title, 'Morning Meeting');
    assert.equal(event.description, 'Discuss project\nSecond line');
    assert.equal(event.location, 'Room A, Building 1');
    // Should get the VEVENT DTSTART, not the VTIMEZONE one
    assert.equal(event.start, '2026-03-20T08:30:00');
    assert.equal(event.end, '2026-03-20T09:30:00');
  });

  it('parses an all-day event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday1@fastmail',
      'DTSTART;VALUE=DATE:20260324',
      'DTEND;VALUE=DATE:20260325',
      'SUMMARY:All Day Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-24');
    assert.equal(event.end, '2026-03-25');
    assert.equal(event.title, 'All Day Event');
  });

  it('parses a UTC event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:utc1@fastmail',
      'DTSTART:20260320T083000Z',
      'DTEND:20260320T093000Z',
      'SUMMARY:UTC Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-20T08:30:00Z');
    assert.equal(event.end, '2026-03-20T09:30:00Z');
  });

  it('defaults title to Untitled when SUMMARY is missing', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:notitle@fastmail',
      'DTSTART:20260320T083000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.title, 'Untitled');
  });

  it('handles missing optional fields', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:minimal@fastmail',
      'DTSTART:20260320T083000Z',
      'SUMMARY:Minimal',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.description, undefined);
    assert.equal(event.location, undefined);
    assert.equal(event.end, undefined);
  });

  it('uses empty string fallback when data is undefined', () => {
    const event = parseCalendarObject({ url: 'https://cal.example.com/e.ics' } as {
      data: string;
      url: string;
    });
    assert.equal(event.url, 'https://cal.example.com/e.ics');
    assert.equal(event.title, 'Untitled');
  });

  it('falls back to url for id when UID is missing', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'DTSTART:20260320T083000Z',
      'SUMMARY:No UID',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: 'https://cal.example.com/nouid.ics' });
    assert.equal(event.id, 'https://cal.example.com/nouid.ics');
  });
});

// ---------- CalDAVCalendarClient ----------

// Helper to create a client with a mocked internal DAV client
function makeCaldavClient() {
  const caldavClient = new CalDAVCalendarClient({
    username: 'user@example.com',
    password: 'password123',
  });

  const mockDavClient = {
    login: async () => {},
    fetchCalendars: async () => [
      {
        displayName: 'Personal',
        url: '/dav/cal/personal/',
        description: 'My personal calendar',
        calendarColor: '#0000FF',
      },
      {
        displayName: 'Work',
        url: '/dav/cal/work/',
        description: undefined,
        calendarColor: undefined,
      },
      {
        displayName: 'DEFAULT_TASK_CALENDAR_NAME',
        url: '/dav/cal/tasks/',
      },
    ],
    fetchCalendarObjects: async () => [] as { data: string; url: string }[],
    createCalendarObject: async () => {},
  };

  // Inject the mock DAV client by setting the private field
  // biome-ignore lint/suspicious/noExplicitAny: test mock injection
  (caldavClient as any).client = mockDavClient;

  return { caldavClient, mockDavClient };
}

describe('CalDAVCalendarClient.getCalendars', () => {
  it('returns calendars excluding task calendar', async () => {
    const { caldavClient } = makeCaldavClient();
    const calendars = await caldavClient.getCalendars();

    assert.equal(calendars.length, 2);
    assert.equal(calendars[0].displayName, 'Personal');
    assert.equal(calendars[0].url, '/dav/cal/personal/');
    assert.equal(calendars[0].description, 'My personal calendar');
    assert.equal(calendars[0].color, '#0000FF');
    assert.equal(calendars[1].displayName, 'Work');
    assert.equal(calendars[1].description, undefined);
    assert.equal(calendars[1].color, undefined);
  });

  it('handles calendars with missing displayName', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    mockDavClient.fetchCalendars = async () => [
      { url: '/dav/cal/unnamed/', displayName: undefined },
    ];

    const calendars = await caldavClient.getCalendars();
    assert.equal(calendars.length, 1);
    assert.equal(calendars[0].displayName, 'Unnamed');
  });
});

describe('CalDAVCalendarClient.getCalendarEvents', () => {
  it('returns events from all calendars', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    mockDavClient.fetchCalendarObjects = async () => [
      {
        data: [
          'BEGIN:VCALENDAR',
          'BEGIN:VEVENT',
          'UID:evt1@fastmail',
          'DTSTART:20260320T083000Z',
          'SUMMARY:Event 1',
          'END:VEVENT',
          'END:VCALENDAR',
        ].join('\r\n'),
        url: '/dav/cal/personal/evt1.ics',
      },
    ];

    const events = await caldavClient.getCalendarEvents();
    assert.equal(events.length, 2); // 1 from each non-task calendar (same mock returns for both)
    assert.equal(events[0].title, 'Event 1');
  });

  it('filters events by calendar ID', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    // Pre-populate calendars cache
    await caldavClient.getCalendars();

    mockDavClient.fetchCalendarObjects = async () => [
      {
        data: [
          'BEGIN:VCALENDAR',
          'BEGIN:VEVENT',
          'UID:evt1@fastmail',
          'DTSTART:20260320T083000Z',
          'SUMMARY:Personal Event',
          'END:VEVENT',
          'END:VCALENDAR',
        ].join('\r\n'),
        url: '/dav/cal/personal/evt1.ics',
      },
    ];

    const events = await caldavClient.getCalendarEvents('/dav/cal/personal/');
    assert.equal(events.length, 1);
    assert.equal(events[0].title, 'Personal Event');
  });

  it('respects the limit parameter', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    mockDavClient.fetchCalendarObjects = async () => [
      {
        data: 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nUID:e1\nDTSTART:20260320T083000Z\nSUMMARY:E1\nEND:VEVENT\nEND:VCALENDAR',
        url: '/e1.ics',
      },
      {
        data: 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nUID:e2\nDTSTART:20260321T083000Z\nSUMMARY:E2\nEND:VEVENT\nEND:VCALENDAR',
        url: '/e2.ics',
      },
    ];

    const events = await caldavClient.getCalendarEvents(undefined, 1);
    assert.equal(events.length, 1);
  });

  it('fetches calendars when cache is empty', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    // Reset the calendars cache
    // biome-ignore lint/suspicious/noExplicitAny: test cache reset
    (caldavClient as any).calendars = null;

    mockDavClient.fetchCalendarObjects = async () => [];

    const events = await caldavClient.getCalendarEvents();
    assert.deepEqual(events, []);
  });
});

describe('CalDAVCalendarClient.getCalendarEventById', () => {
  it('finds event by UID', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    mockDavClient.fetchCalendarObjects = async () => [
      {
        data: [
          'BEGIN:VCALENDAR',
          'BEGIN:VEVENT',
          'UID:target-uid@fastmail',
          'DTSTART:20260320T083000Z',
          'SUMMARY:Found It',
          'END:VEVENT',
          'END:VCALENDAR',
        ].join('\r\n'),
        url: '/dav/cal/personal/target.ics',
      },
    ];

    const event = await caldavClient.getCalendarEventById('target-uid@fastmail');
    assert.ok(event);
    assert.equal(event.title, 'Found It');
  });

  it('finds event by URL', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    mockDavClient.fetchCalendarObjects = async () => [
      {
        data: [
          'BEGIN:VCALENDAR',
          'BEGIN:VEVENT',
          'UID:other-uid',
          'DTSTART:20260320T083000Z',
          'SUMMARY:By URL',
          'END:VEVENT',
          'END:VCALENDAR',
        ].join('\r\n'),
        url: '/dav/cal/personal/target.ics',
      },
    ];

    const event = await caldavClient.getCalendarEventById('/dav/cal/personal/target.ics');
    assert.ok(event);
    assert.equal(event.title, 'By URL');
  });

  it('returns null when event not found', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    mockDavClient.fetchCalendarObjects = async () => [];

    const event = await caldavClient.getCalendarEventById('nonexistent');
    assert.equal(event, null);
  });

  it('fetches calendars when cache is empty', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    // biome-ignore lint/suspicious/noExplicitAny: test cache reset
    (caldavClient as any).calendars = null;

    mockDavClient.fetchCalendarObjects = async () => [];
    const event = await caldavClient.getCalendarEventById('nonexistent');
    assert.equal(event, null);
  });
});

describe('CalDAVCalendarClient.createCalendarEvent', () => {
  it('creates an event and returns UID', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    // Pre-populate calendars cache
    await caldavClient.getCalendars();

    let capturedArgs: { calendar: unknown; filename: string; iCalString: string } | null = null;
    mockDavClient.createCalendarObject = async (args: {
      calendar: unknown;
      filename: string;
      iCalString: string;
    }) => {
      capturedArgs = args;
    };

    const uid = await caldavClient.createCalendarEvent({
      calendarId: '/dav/cal/personal/',
      title: 'New Event',
      description: 'A description',
      start: '2026-03-20T08:30:00',
      end: '2026-03-20T09:30:00',
      location: 'Room A',
    });

    assert.ok(uid.includes('@fastmail-mcp'));
    assert.ok(capturedArgs);
    assert.ok(capturedArgs.iCalString.includes('SUMMARY:New Event'));
    assert.ok(capturedArgs.iCalString.includes('DESCRIPTION:A description'));
    assert.ok(capturedArgs.iCalString.includes('LOCATION:Room A'));
  });

  it('creates event without optional description and location', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    await caldavClient.getCalendars();

    let capturedIcal = '';
    mockDavClient.createCalendarObject = async (args: {
      calendar: unknown;
      filename: string;
      iCalString: string;
    }) => {
      capturedIcal = args.iCalString;
    };

    await caldavClient.createCalendarEvent({
      calendarId: '/dav/cal/personal/',
      title: 'Minimal Event',
      start: '2026-03-20T08:30:00',
      end: '2026-03-20T09:30:00',
    });

    assert.ok(!capturedIcal.includes('DESCRIPTION'));
    assert.ok(!capturedIcal.includes('LOCATION'));
  });

  it('throws when calendar not found', async () => {
    const { caldavClient } = makeCaldavClient();
    await caldavClient.getCalendars();

    await assert.rejects(
      () =>
        caldavClient.createCalendarEvent({
          calendarId: '/dav/cal/nonexistent/',
          title: 'Event',
          start: '2026-03-20T08:30:00',
          end: '2026-03-20T09:30:00',
        }),
      (err: Error) => {
        assert.match(err.message, /Calendar not found/);
        return true;
      },
    );
  });

  it('fetches calendars when cache is empty', async () => {
    const { caldavClient, mockDavClient } = makeCaldavClient();
    // biome-ignore lint/suspicious/noExplicitAny: test cache reset
    (caldavClient as any).calendars = null;

    mockDavClient.createCalendarObject = async () => {};

    const uid = await caldavClient.createCalendarEvent({
      calendarId: '/dav/cal/personal/',
      title: 'New Event',
      start: '2026-03-20T08:30:00',
      end: '2026-03-20T09:30:00',
    });

    assert.ok(uid.includes('@fastmail-mcp'));
  });
});

describe('CalDAVCalendarClient constructor', () => {
  it('uses default server URL when not provided', () => {
    const caldavClient = new CalDAVCalendarClient({
      username: 'user@example.com',
      password: 'password123',
    });
    // biome-ignore lint/suspicious/noExplicitAny: inspect private field
    assert.equal((caldavClient as any).config.username, 'user@example.com');
  });

  it('accepts custom server URL', () => {
    const caldavClient = new CalDAVCalendarClient({
      username: 'user@example.com',
      password: 'password123',
      serverUrl: 'https://custom-caldav.example.com',
    });
    // biome-ignore lint/suspicious/noExplicitAny: inspect private field
    assert.equal((caldavClient as any).config.serverUrl, 'https://custom-caldav.example.com');
  });
});

describe('CalDAVCalendarClient HTTPS enforcement', () => {
  it('rejects http:// server URLs', () => {
    assert.throws(
      () =>
        new CalDAVCalendarClient({ username: 'u', password: 'p', serverUrl: 'http://evil.com' }),
      /HTTPS is required/,
    );
  });

  it('rejects ftp:// server URLs', () => {
    assert.throws(
      () => new CalDAVCalendarClient({ username: 'u', password: 'p', serverUrl: 'ftp://evil.com' }),
      /HTTPS is required/,
    );
  });

  it('accepts https:// server URLs', () => {
    assert.doesNotThrow(
      () =>
        new CalDAVCalendarClient({
          username: 'u',
          password: 'p',
          serverUrl: 'https://caldav.fastmail.com',
        }),
    );
  });

  it('accepts default (no serverUrl)', () => {
    assert.doesNotThrow(() => new CalDAVCalendarClient({ username: 'u', password: 'p' }));
  });
});

describe('escapeICalText', () => {
  it('escapes backslashes', () => {
    assert.equal(escapeICalText('a\\b'), 'a\\\\b');
  });

  it('escapes semicolons', () => {
    assert.equal(escapeICalText('a;b'), 'a\\;b');
  });

  it('escapes commas', () => {
    assert.equal(escapeICalText('a,b'), 'a\\,b');
  });

  it('escapes newlines', () => {
    assert.equal(escapeICalText('a\nb'), 'a\\nb');
  });

  it('escapes carriage return + newline', () => {
    assert.equal(escapeICalText('a\r\nb'), 'a\\nb');
  });

  it('prevents iCal injection via title', () => {
    const malicious = 'Meeting\r\nATTENDEE:mailto:evil@attacker.com';
    const escaped = escapeICalText(malicious);
    assert.ok(!escaped.includes('\n'));
    assert.ok(!escaped.includes('\r'));
  });
});

describe('unescapeICalText', () => {
  it('unescapes \\n to newline', () => {
    assert.equal(unescapeICalText('line1\\nline2'), 'line1\nline2');
  });

  it('unescapes \\, to comma', () => {
    assert.equal(unescapeICalText('Room A\\, Building 1'), 'Room A, Building 1');
  });

  it('unescapes \\; to semicolon', () => {
    assert.equal(unescapeICalText('a\\;b'), 'a;b');
  });

  it('unescapes \\\\ to backslash', () => {
    assert.equal(unescapeICalText('a\\\\b'), 'a\\b');
  });

  it('handles multiple escape sequences in one string', () => {
    assert.equal(unescapeICalText('line1\\nRoom A\\, Bldg\\;1'), 'line1\nRoom A, Bldg;1');
  });

  it('is inverse of escapeICalText for round-trip', () => {
    const original = 'Meeting, Room 1; Floor 2\nSee you there\\';
    assert.equal(unescapeICalText(escapeICalText(original)), original);
  });

  it('leaves plain strings unchanged', () => {
    assert.equal(unescapeICalText('No special chars'), 'No special chars');
  });
});

describe('CalDAVCalendarClient.getCalendarEvents', () => {
  function makeIcal(uid: string, summary: string, dtstart: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTART:${dtstart}`,
      `SUMMARY:${summary}`,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedClient(calendarObjects: Array<{ data: string; url: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    // Override the private getClient method to return a mock DAVClient
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
    };
    // biome-ignore lint/suspicious/noExplicitAny: override private field in test
    (client as unknown as Record<string, unknown>).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('sorts events by start date ascending', async () => {
    const objects = [
      { data: makeIcal('c@fm', 'Evening', '20260325T200000Z'), url: '/c.ics' },
      { data: makeIcal('a@fm', 'Morning', '20260325T080000Z'), url: '/a.ics' },
      { data: makeIcal('b@fm', 'Afternoon', '20260325T140000Z'), url: '/b.ics' },
    ];
    const { client } = createMockedClient(objects);
    const events = await client.getCalendarEvents(undefined, 50);

    assert.equal(events.length, 3);
    assert.equal(events[0].title, 'Morning');
    assert.equal(events[1].title, 'Afternoon');
    assert.equal(events[2].title, 'Evening');
  });

  it('passes timeRange to fetchCalendarObjects when startDate/endDate provided', async () => {
    const objects = [{ data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' }];
    const { client, mockDAVClient } = createMockedClient(objects);
    await client.getCalendarEvents(undefined, 50, '2026-03-25T00:00:00Z', '2026-03-26T00:00:00Z');

    const callArgs = mockDAVClient.fetchCalendarObjects.mock.calls[0].arguments[0];
    assert.deepEqual(callArgs.timeRange, {
      start: '2026-03-25T00:00:00Z',
      end: '2026-03-26T00:00:00Z',
    });
  });

  it('does not pass timeRange when no dates provided', async () => {
    const objects = [{ data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' }];
    const { client, mockDAVClient } = createMockedClient(objects);
    await client.getCalendarEvents(undefined, 50);

    const callArgs = mockDAVClient.fetchCalendarObjects.mock.calls[0].arguments[0];
    assert.equal(callArgs.timeRange, undefined);
  });
});
