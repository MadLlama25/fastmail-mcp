import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import {
  extractVEvent,
  parseICalValue,
  formatICalDate,
  parseCalendarObject,
  escapeICalText,
  validateAndFormatICalDate,
  parseRecurrence,
  CalDAVCalendarClient,
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
});

describe('escapeICalText', () => {
  it('escapes backslashes', () => {
    assert.equal(escapeICalText('path\\to\\file'), 'path\\\\to\\\\file');
  });

  it('escapes semicolons', () => {
    assert.equal(escapeICalText('a;b;c'), 'a\\;b\\;c');
  });

  it('escapes commas', () => {
    assert.equal(escapeICalText('Room A, Building 1'), 'Room A\\, Building 1');
  });

  it('escapes newlines', () => {
    assert.equal(escapeICalText('line1\nline2'), 'line1\\nline2');
    assert.equal(escapeICalText('line1\r\nline2'), 'line1\\nline2');
  });

  it('leaves plain text unchanged', () => {
    assert.equal(escapeICalText('Team Standup'), 'Team Standup');
  });

  it('prevents ICS property injection via CRLF', () => {
    const malicious = 'Meeting\r\nATTENDEE:mailto:attacker@evil.com';
    const escaped = escapeICalText(malicious);
    // No literal newlines means the injected ATTENDEE stays inside the text value,
    // not on its own ICS property line
    assert.ok(!escaped.includes('\n'), 'escaped text must not contain literal newlines');
    assert.ok(!escaped.includes('\r'), 'escaped text must not contain literal carriage returns');
    assert.equal(escaped, 'Meeting\\nATTENDEE:mailto:attacker@evil.com');
  });

  it('prevents injection of extra VEVENT components', () => {
    const malicious = 'Meeting\r\nEND:VEVENT\r\nBEGIN:VEVENT\r\nSUMMARY:Injected';
    const escaped = escapeICalText(malicious);
    // No literal newlines means the injected properties stay inside the SUMMARY value
    assert.ok(!escaped.includes('\n'), 'escaped text must not contain literal newlines');
    assert.ok(!escaped.includes('\r'), 'escaped text must not contain literal carriage returns');
    // The entire payload is on one logical ICS line, so END:VEVENT can't terminate the block
    assert.ok(escaped.startsWith('Meeting\\n'), 'newlines should be escaped, not literal');
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
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
    };
    (client as any).client = mockDAVClient;
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
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);
    await client.getCalendarEvents(undefined, 50, '2026-03-25T00:00:00Z', '2026-03-26T00:00:00Z');

    const callArgs = mockDAVClient.fetchCalendarObjects.mock.calls[0].arguments[0];
    assert.deepEqual(callArgs.timeRange, {
      start: '2026-03-25T00:00:00Z',
      end: '2026-03-26T00:00:00Z',
    });
  });

  it('does not pass timeRange when no dates provided', async () => {
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);
    await client.getCalendarEvents(undefined, 50);

    const callArgs = mockDAVClient.fetchCalendarObjects.mock.calls[0].arguments[0];
    assert.equal(callArgs.timeRange, undefined);
  });
});

describe('validateAndFormatICalDate', () => {
  it('accepts and formats date-only', () => {
    assert.equal(validateAndFormatICalDate('2026-04-18', 'start'), '20260418');
  });

  it('accepts and formats UTC datetime', () => {
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00Z', 'start'), '20260418T100000Z');
  });

  it('accepts and normalizes positive offset to UTC', () => {
    // 2026-04-18T10:00:00+02:00 = 2026-04-18T08:00:00Z
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00+02:00', 'start'), '20260418T080000Z');
  });

  it('accepts and normalizes negative offset to UTC', () => {
    // 2026-04-18T10:00:00-05:00 = 2026-04-18T15:00:00Z
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00-05:00', 'start'), '20260418T150000Z');
  });

  it('accepts floating datetime (no zone)', () => {
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00', 'start'), '20260418T100000');
  });

  it('rejects CRLF injection attempt', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z\r\nATTENDEE:mailto:attacker@example.com', 'start'),
      /control characters/,
    );
  });

  it('rejects bare LF injection attempt', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z\nATTENDEE:mailto:attacker@example.com', 'start'),
      /control characters/,
    );
  });

  it('rejects null byte', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z\0', 'start'),
      /control characters/,
    );
  });

  it('rejects malformed date', () => {
    assert.throws(
      () => validateAndFormatICalDate('not-a-date', 'start'),
      /must be ISO-8601/,
    );
  });

  it('rejects extra trailing content', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z bonus', 'start'),
      /must be ISO-8601/,
    );
  });

  it('rejects non-string input', () => {
    assert.throws(
      () => validateAndFormatICalDate(undefined as any, 'start'),
      /must be a string/,
    );
  });

  it('throws with field name in error', () => {
    try {
      validateAndFormatICalDate('garbage', 'event.end');
      assert.fail('should have thrown');
    } catch (e) {
      assert.match((e as Error).message, /event\.end/);
    }
  });
});

describe('parseRecurrence', () => {
  it('passes through a raw RRULE string unchanged', () => {
    assert.equal(parseRecurrence('FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=29'), 'FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=29');
  });

  it('builds RRULE from structured object — yearly by month+day', () => {
    const result = parseRecurrence({ frequency: 'YEARLY', byMonth: 4, byMonthDay: 29 });
    assert.ok(result.startsWith('FREQ=YEARLY'), 'must start with FREQ');
    assert.ok(result.includes('BYMONTH=4'), 'must include BYMONTH');
    assert.ok(result.includes('BYMONTHDAY=29'), 'must include BYMONTHDAY');
  });

  it('builds RRULE from structured object — weekly with interval and count', () => {
    const result = parseRecurrence({ frequency: 'WEEKLY', interval: 2, count: 10, byDay: 'MO,WE,FR' });
    assert.ok(result.includes('FREQ=WEEKLY'));
    assert.ok(result.includes('INTERVAL=2'));
    assert.ok(result.includes('COUNT=10'));
    assert.ok(result.includes('BYDAY=MO,WE,FR'));
  });

  it('normalizes until date via validateAndFormatICalDate', () => {
    const result = parseRecurrence({ frequency: 'DAILY', until: '2027-12-31' });
    assert.ok(result.includes('UNTIL=20271231'), `got: ${result}`);
  });
});

describe('parseCalendarObject — rrule field', () => {
  it('extracts RRULE from a recurring event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:recur1@fastmail',
      'DTSTART:20270429T213300Z',
      'DTEND:20270429T223300Z',
      'SUMMARY:Free frosty keychain',
      'RRULE:FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=29',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: 'https://caldav.example.com/cal/recur1.ics' });
    assert.equal(event.rrule, 'FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=29');
    assert.equal(event.title, 'Free frosty keychain');
  });

  it('returns undefined rrule for non-recurring events', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:norule@fastmail',
      'DTSTART:20270429T213300Z',
      'SUMMARY:One-off',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.rrule, undefined);
  });
});

describe('CalDAVCalendarClient — createCalendarEvent with recurrence', () => {
  function createMockedClientForCreate() {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const createdIcals: string[] = [];
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      createCalendarObject: mock.fn(async (params: any) => {
        createdIcals.push(params.iCalString);
        return {};
      }),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient, createdIcals };
  }

  it('embeds RRULE inside VEVENT, not at top level', async () => {
    const { client, createdIcals } = createMockedClientForCreate();
    await client.createCalendarEvent({
      calendarId: '/cal/personal/',
      title: 'Free frosty keychain',
      start: '2027-04-29T21:33:00-07:00',
      end: '2027-04-29T22:33:00-07:00',
      recurrence: { frequency: 'YEARLY', byMonth: 4, byMonthDay: 29 },
    });

    assert.equal(createdIcals.length, 1);
    const ical = createdIcals[0];

    // RRULE must be inside VEVENT
    const veventMatch = ical.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/);
    assert.ok(veventMatch, 'must have VEVENT block');
    const vevent = veventMatch![0];
    assert.ok(vevent.includes('RRULE:'), 'RRULE must be inside VEVENT');

    // RRULE must NOT appear between VCALENDAR and VEVENT
    const beforeVevent = ical.split('BEGIN:VEVENT')[0];
    assert.ok(!beforeVevent.includes('RRULE:'), 'RRULE must not appear before VEVENT');

    // RRULE must NOT appear between END:VEVENT and END:VCALENDAR
    const afterVevent = ical.split('END:VEVENT')[1] || '';
    assert.ok(!afterVevent.includes('RRULE:'), 'RRULE must not appear after VEVENT');

    assert.ok(ical.includes('FREQ=YEARLY'), 'RRULE must contain FREQ=YEARLY');
  });

  it('creates event without RRULE when recurrence is not provided', async () => {
    const { client, createdIcals } = createMockedClientForCreate();
    await client.createCalendarEvent({
      calendarId: '/cal/personal/',
      title: 'One-off',
      start: '2027-04-29T21:33:00Z',
      end: '2027-04-29T22:33:00Z',
    });

    const ical = createdIcals[0];
    assert.ok(!ical.includes('RRULE:'), 'should not have RRULE for non-recurring event');
  });
});

describe('CalDAVCalendarClient — updateCalendarEvent', () => {
  function makeIcalWithRrule(uid: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTART:20270429T213300Z',
      'DTEND:20270429T223300Z',
      'SUMMARY:Free frosty keychain',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedClientForUpdate(uid: string) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const updatedObjects: any[] = [];
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => [
        { data: makeIcalWithRrule(uid), url: `/cal/personal/${uid}.ics`, etag: '"etag123"' },
      ]),
      updateCalendarObject: mock.fn(async (params: any) => {
        updatedObjects.push(params.calendarObject);
        return {};
      }),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient, updatedObjects };
  }

  it('calls updateCalendarObject with correct url and new iCalString', async () => {
    const uid = '1777523653761-r2ohjha07qq@fastmail-mcp';
    const { client, updatedObjects } = createMockedClientForUpdate(uid);

    await client.updateCalendarEvent(uid, {
      recurrence: 'FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=29',
    });

    assert.equal(updatedObjects.length, 1);
    const co = updatedObjects[0];
    assert.equal(co.url, `/cal/personal/${uid}.ics`);
    assert.equal(co.etag, '"etag123"');
    assert.ok(typeof co.data === 'string', 'data must be a string (iCalString)');
    assert.ok(co.data.includes('RRULE:FREQ=YEARLY;BYMONTH=4;BYMONTHDAY=29'), 'must embed RRULE');
    // RRULE must be inside VEVENT
    const vevent = co.data.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/)?.[0] ?? '';
    assert.ok(vevent.includes('RRULE:'), 'RRULE must be inside VEVENT block');
    assert.ok(co.data.includes('SUMMARY:Free frosty keychain'), 'must preserve existing title');
  });

  it('preserves existing fields when only recurrence is updated', async () => {
    const uid = 'test-preserve@fastmail-mcp';
    const { client, updatedObjects } = createMockedClientForUpdate(uid);

    await client.updateCalendarEvent(uid, { recurrence: { frequency: 'YEARLY' } });

    const data = updatedObjects[0].data;
    assert.ok(data.includes('DTSTART:20270429T213300Z'), 'must preserve original start');
    assert.ok(data.includes('DTEND:20270429T223300Z'), 'must preserve original end');
  });

  it('throws when event not found', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.updateCalendarEvent('nonexistent@fastmail-mcp', { title: 'New Title' }),
      /not found/,
    );
  });
});

describe('CalDAVCalendarClient — deleteCalendarEvent', () => {
  it('calls deleteCalendarObject with correct url and etag', async () => {
    const uid = 'delete-me@fastmail-mcp';
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTART:20270101T100000Z',
      'SUMMARY:To Delete',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const deletedObjects: any[] = [];
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => [
        { data: ical, url: `/cal/personal/${uid}.ics`, etag: '"etag456"' },
      ]),
      deleteCalendarObject: mock.fn(async (params: any) => {
        deletedObjects.push(params.calendarObject);
        return {};
      }),
    };
    (client as any).client = mockDAVClient;

    await client.deleteCalendarEvent(uid);

    assert.equal(deletedObjects.length, 1);
    assert.equal(deletedObjects[0].url, `/cal/personal/${uid}.ics`);
    assert.equal(deletedObjects[0].etag, '"etag456"');
  });

  it('throws when event not found', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => []),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.deleteCalendarEvent('ghost@fastmail-mcp'),
      /not found/,
    );
  });
});

// Integration smoke test — requires FASTMAIL_INTEGRATION=1 and CalDAV credentials
// Skips silently when env var is not set so CI stays clean
describe('integration: update_calendar_event round-trip (FASTMAIL_INTEGRATION=1)', { skip: !process.env.FASTMAIL_INTEGRATION }, () => {
  it('creates a recurring event, fetches it, confirms RRULE survived, then deletes it', async () => {
    const username = process.env.FASTMAIL_CALDAV_USERNAME;
    const password = process.env.FASTMAIL_CALDAV_PASSWORD;
    assert.ok(username, 'FASTMAIL_CALDAV_USERNAME required for integration test');
    assert.ok(password, 'FASTMAIL_CALDAV_PASSWORD required for integration test');

    const client = new CalDAVCalendarClient({ username, password });

    // List calendars to find the default one
    const calendars = await client.getCalendars();
    assert.ok(calendars.length > 0, 'must have at least one calendar');
    const calendarId = calendars[0].url;

    // Create a recurring event
    const uid = await client.createCalendarEvent({
      calendarId,
      title: 'Integration Test Recurring Event',
      start: '2027-06-15T10:00:00Z',
      end: '2027-06-15T11:00:00Z',
      recurrence: { frequency: 'YEARLY', byMonth: 6, byMonthDay: 15 },
    });

    // Fetch it back
    const fetched = await client.getCalendarEventById(uid);
    assert.ok(fetched, 'event must be fetchable after creation');
    assert.ok(fetched!.rrule, 'RRULE must survive CalDAV round-trip');
    assert.ok(fetched!.rrule!.includes('FREQ=YEARLY'), `expected YEARLY, got: ${fetched!.rrule}`);

    // Update it to add a different recurrence
    await client.updateCalendarEvent(uid, {
      recurrence: 'FREQ=YEARLY;BYMONTH=6;BYMONTHDAY=15;COUNT=5',
    });

    const updated = await client.getCalendarEventById(uid);
    assert.ok(updated?.rrule?.includes('COUNT=5'), `updated RRULE must include COUNT=5, got: ${updated?.rrule}`);

    // Clean up
    await client.deleteCalendarEvent(uid);
    const gone = await client.getCalendarEventById(uid);
    assert.equal(gone, null, 'event must be gone after deletion');
  });
});
