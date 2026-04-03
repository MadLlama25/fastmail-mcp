import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import {
  extractVEvent,
  parseICalValue,
  findValueBoundary,
  parseAllICalProperties,
  parseAttendee,
  parseICalDuration,
  formatICalDate,
  parseCalendarObject,
  escapeICalText,
  unescapeICalText,
  toICalUTC,
  foldICalLine,
  detectLineEnding,
  replaceICalProperty,
  removeAllICalProperties,
  removeOrphanedVTimezones,
  removeExceptionVEvents,
  insertBeforeEndVEvent,
  validateAttendeeEmail,
  quoteParamValue,
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

  it('returns null when no VEVENT block found', () => {
    const data = 'no vevent here';
    assert.equal(extractVEvent(data), null);
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

describe('CalDAVCalendarClient.updateCalendarEvent', () => {
  function makeFullIcal(uid: string, summary: string, dtstart: string, dtend: string, description?: string, location?: string): string {
    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:20260401T000000Z`,
      `DTSTART:${dtstart}`,
      `DTEND:${dtend}`,
      `SUMMARY:${escapeICalText(summary)}`,
    ];
    if (description) lines.push(`DESCRIPTION:${escapeICalText(description)}`);
    if (location) lines.push(`LOCATION:${escapeICalText(location)}`);
    lines.push('END:VEVENT', 'END:VCALENDAR');
    return lines.join('\r\n');
  }

  function createMockedClientWithUpdateDelete(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
      updateCalendarObject: mock.fn(async () => ({})),
      deleteCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('updates only the title, preserving other fields', async () => {
    const ical = makeFullIcal('evt1@fm', 'Original Title', '20260401T100000Z', '20260401T110000Z', 'My description', 'Room A');
    const objects = [{ data: ical, url: '/cal/evt1.ics', etag: '"etag1"' }];
    const { client, mockDAVClient } = createMockedClientWithUpdateDelete(objects);

    const result = await client.updateCalendarEvent('evt1@fm', { title: 'New Title' });

    assert.equal(result, 'evt1@fm');
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedObj = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject;
    assert.ok(updatedObj.data.includes('SUMMARY:New Title'));
    assert.ok(updatedObj.data.includes('DESCRIPTION:My description'));
    assert.ok(updatedObj.data.includes('LOCATION:Room A'));
    assert.ok(updatedObj.data.includes('DTSTART:20260401T100000Z'));
    assert.ok(updatedObj.data.includes('UID:evt1@fm'));
  });

  it('updates start and end times', async () => {
    const ical = makeFullIcal('evt2@fm', 'Meeting', '20260401T100000Z', '20260401T110000Z');
    const objects = [{ data: ical, url: '/cal/evt2.ics' }];
    const { client, mockDAVClient } = createMockedClientWithUpdateDelete(objects);

    await client.updateCalendarEvent('evt2@fm', {
      start: '2026-04-02T14:00:00Z',
      end: '2026-04-02T15:00:00Z',
    });

    const updatedObj = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject;
    assert.ok(updatedObj.data.includes('DTSTART:20260402T140000Z'));
    assert.ok(updatedObj.data.includes('DTEND:20260402T150000Z'));
    assert.ok(updatedObj.data.includes('SUMMARY:Meeting'));
  });

  it('throws when event not found', async () => {
    const { client } = createMockedClientWithUpdateDelete([]);
    await assert.rejects(
      () => client.updateCalendarEvent('nonexistent@fm', { title: 'X' }),
      /Calendar event not found: nonexistent@fm/
    );
  });
});

describe('CalDAVCalendarClient.deleteCalendarEvent', () => {
  function createMockedClientWithDelete(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
      deleteCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('deletes an event by UID', async () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:del1@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:To Delete',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: ical, url: '/cal/del1.ics', etag: '"etag1"' }];
    const { client, mockDAVClient } = createMockedClientWithDelete(objects);

    await client.deleteCalendarEvent('del1@fm');

    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 1);
    const deletedObj = mockDAVClient.deleteCalendarObject.mock.calls[0].arguments[0].calendarObject;
    assert.equal(deletedObj.url, '/cal/del1.ics');
  });

  it('throws when event not found', async () => {
    const { client } = createMockedClientWithDelete([]);
    await assert.rejects(
      () => client.deleteCalendarEvent('nonexistent@fm'),
      /Calendar event not found: nonexistent@fm/
    );
  });
});

// ============================================================
// New tests for calendar attendee support & non-destructive updates
// ============================================================

describe('findValueBoundary', () => {
  it('finds colon in simple property', () => {
    assert.equal(findValueBoundary('SUMMARY:Test'), 7);
  });

  it('skips colons inside quoted parameter values', () => {
    const line = 'ATTENDEE;DELEGATED-FROM="mailto:boss@example.com";CN="Smith, John":mailto:john@example.com';
    const idx = findValueBoundary(line);
    assert.equal(line.substring(idx + 1), 'mailto:john@example.com');
  });

  it('handles ALTREP with URL', () => {
    const line = 'DESCRIPTION;ALTREP="http://example.com/desc":Plain text';
    const idx = findValueBoundary(line);
    assert.equal(line.substring(idx + 1), 'Plain text');
  });

  it('returns -1 when no colon found', () => {
    assert.equal(findValueBoundary('NOCOLON'), -1);
  });
});

describe('parseAllICalProperties', () => {
  it('returns multiple ATTENDEEs', () => {
    const vevent = [
      'BEGIN:VEVENT',
      'ATTENDEE;CN=Alice:mailto:alice@example.com',
      'ATTENDEE;CN=Bob:mailto:bob@example.com',
      'SUMMARY:Test',
      'END:VEVENT',
    ].join('\n');
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 2);
    assert.ok(results[0].includes('alice@'));
    assert.ok(results[1].includes('bob@'));
  });

  it('returns empty array when none found', () => {
    const vevent = 'BEGIN:VEVENT\nSUMMARY:Test\nEND:VEVENT';
    assert.deepEqual(parseAllICalProperties(vevent, 'ATTENDEE'), []);
  });

  it('handles folded ATTENDEE lines', () => {
    const vevent = [
      'BEGIN:VEVENT',
      'ATTENDEE;CN=Very Long Name;PARTSTAT=ACCEPTED:mailto:long',
      ' name@example.com',
      'SUMMARY:Test',
      'END:VEVENT',
    ].join('\n');
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 1);
    assert.ok(results[0].includes('longname@example.com'));
  });

  it('handles CRLF input', () => {
    const vevent = 'BEGIN:VEVENT\r\nATTENDEE;CN=Alice:mailto:a@b.com\r\nEND:VEVENT';
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 1);
    assert.ok(!results[0].includes('\r'));
  });

  it('does not match partial property names', () => {
    const vevent = 'BEGIN:VEVENT\nATTENDEE-X:foo\nATTENDEE:bar\nEND:VEVENT';
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 1);
    assert.equal(results[0], 'ATTENDEE:bar');
  });
});

describe('parseAttendee', () => {
  it('parses simple ATTENDEE with CN', () => {
    const result = parseAttendee('ATTENDEE;CN=Alice:mailto:alice@example.com');
    assert.equal(result.email, 'alice@example.com');
    assert.equal(result.name, 'Alice');
  });

  it('parses quoted CN with comma', () => {
    const result = parseAttendee('ATTENDEE;CN="Doe, John":mailto:john@example.com');
    assert.equal(result.name, 'Doe, John');
    assert.equal(result.email, 'john@example.com');
  });

  it('parses PARTSTAT, ROLE, CUTYPE, RSVP', () => {
    const result = parseAttendee('ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED;ROLE=REQ-PARTICIPANT;CUTYPE=INDIVIDUAL;RSVP=TRUE:mailto:alice@example.com');
    assert.equal(result.status, 'ACCEPTED');
    assert.equal(result.role, 'REQ-PARTICIPANT');
    assert.equal(result.cutype, 'INDIVIDUAL');
    assert.equal(result.rsvp, true);
  });

  it('converts RSVP=FALSE to boolean false', () => {
    const result = parseAttendee('ATTENDEE;RSVP=FALSE:mailto:alice@example.com');
    assert.equal(result.rsvp, false);
  });

  it('parses ORGANIZER lines', () => {
    const result = parseAttendee('ORGANIZER;CN=Boss:mailto:boss@example.com');
    assert.equal(result.email, 'boss@example.com');
    assert.equal(result.name, 'Boss');
  });

  it('handles missing CN', () => {
    const result = parseAttendee('ATTENDEE:mailto:anon@example.com');
    assert.equal(result.email, 'anon@example.com');
    assert.equal(result.name, undefined);
  });

  it('handles non-mailto URI', () => {
    const result = parseAttendee('ATTENDEE:urn:uuid:550e8400-e29b-41d4-a716-446655440000');
    assert.equal(result.email, 'urn:uuid:550e8400-e29b-41d4-a716-446655440000');
  });

  it('handles bare email without mailto', () => {
    const result = parseAttendee('ATTENDEE:alice@example.com');
    assert.equal(result.email, 'alice@example.com');
  });

  it('handles DELEGATED-FROM with quoted mailto (colon inside quotes)', () => {
    const result = parseAttendee('ATTENDEE;DELEGATED-FROM="mailto:boss@example.com";CN=Alice:mailto:alice@example.com');
    assert.equal(result.email, 'alice@example.com');
    assert.equal(result.name, 'Alice');
  });

  it('omits empty CN', () => {
    const result = parseAttendee('ATTENDEE;CN=:mailto:alice@example.com');
    assert.equal(result.name, undefined);
  });

  it('handles CN with literal DQUOTE character', () => {
    // CN value with embedded quote — parseAttendee should store the raw value
    const result = parseAttendee('ATTENDEE;CN="John \'Doc\' Smith":mailto:john@example.com');
    assert.equal(result.name, "John 'Doc' Smith");
    assert.equal(result.email, 'john@example.com');
  });
});

describe('parseICalDuration', () => {
  it('parses PT2H', () => {
    assert.equal(parseICalDuration('PT2H', '2026-04-01T10:00:00Z'), '2026-04-01T12:00:00Z');
  });

  it('parses P1D', () => {
    assert.equal(parseICalDuration('P1D', '2026-04-01T10:00:00Z'), '2026-04-02T10:00:00Z');
  });

  it('parses P1W', () => {
    assert.equal(parseICalDuration('P1W', '2026-04-01T10:00:00Z'), '2026-04-08T10:00:00Z');
  });

  it('parses P1DT2H30M', () => {
    assert.equal(parseICalDuration('P1DT2H30M', '2026-04-01T10:00:00Z'), '2026-04-02T12:30:00Z');
  });

  it('parses PT90M', () => {
    assert.equal(parseICalDuration('PT90M', '2026-04-01T10:00:00Z'), '2026-04-01T11:30:00Z');
  });

  it('parses PT0S (zero duration)', () => {
    assert.equal(parseICalDuration('PT0S', '2026-04-01T10:00:00Z'), '2026-04-01T10:00:00Z');
  });

  it('parses P1DT0H0M0S (verbose)', () => {
    assert.equal(parseICalDuration('P1DT0H0M0S', '2026-04-01T10:00:00Z'), '2026-04-02T10:00:00Z');
  });

  it('returns undefined for malformed input', () => {
    assert.equal(parseICalDuration('2H', '2026-04-01T10:00:00Z'), undefined);
    assert.equal(parseICalDuration('PTXYZ', '2026-04-01T10:00:00Z'), undefined);
  });

  it('rejects bare P', () => {
    assert.equal(parseICalDuration('P', '2026-04-01T10:00:00Z'), undefined);
  });

  it('rejects P1DT (T with no time components)', () => {
    assert.equal(parseICalDuration('P1DT', '2026-04-01T10:00:00Z'), undefined);
  });

  it('handles date-only start', () => {
    assert.equal(parseICalDuration('P1D', '2026-04-01'), '2026-04-02');
  });

  it('returns floating time for floating start (no Z)', () => {
    const result = parseICalDuration('PT2H', '2026-04-01T10:00:00');
    assert.equal(result, '2026-04-01T12:00:00');
    assert.ok(!result!.includes('Z'), 'floating start should produce floating end');
  });
});

describe('parseCalendarObject with participants', () => {
  it('parses ATTENDEE and ORGANIZER', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:evt@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:Meeting',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED;ROLE=REQ-PARTICIPANT:mailto:alice@example.com',
      'ATTENDEE;CN=Bob;PARTSTAT=TENTATIVE:mailto:bob@example.com',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' }, { includeParticipants: true });
    assert.equal(event.organizer?.email, 'boss@example.com');
    assert.equal(event.organizer?.name, 'Boss');
    assert.equal(event.participants?.length, 2);
    assert.equal(event.participants?.[0].email, 'alice@example.com');
    assert.equal(event.participants?.[0].status, 'ACCEPTED');
    assert.equal(event.participants?.[1].email, 'bob@example.com');
  });

  it('omits participants when includeParticipants is false', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:evt@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:Meeting',
      'ATTENDEE;CN=Alice:mailto:alice@example.com',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.participants, undefined);
    assert.equal(event.organizer, undefined);
  });

  it('computes end from DURATION', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur@fm',
      'DTSTART:20260401T100000Z',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-04-01T10:00:00Z');
    assert.equal(event.end, '2026-04-01T12:00:00Z');
  });

  it('returns minimal event when no VEVENT found', () => {
    const data = 'BEGIN:VCALENDAR\r\nEND:VCALENDAR';
    const event = parseCalendarObject({ data, url: '/test.ics' });
    assert.equal(event.title, 'Untitled');
    assert.equal(event.url, '/test.ics');
  });
});

describe('replaceICalProperty', () => {
  const simpleEvent = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'UID:test@fm',
    'SUMMARY:Original',
    'DTSTART:20260401T100000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\n');

  it('replaces an existing property', () => {
    const result = replaceICalProperty(simpleEvent, 'SUMMARY', 'SUMMARY:Updated');
    assert.ok(result.includes('SUMMARY:Updated'));
    assert.ok(!result.includes('SUMMARY:Original'));
  });

  it('preserves other properties when replacing', () => {
    const result = replaceICalProperty(simpleEvent, 'SUMMARY', 'SUMMARY:Updated');
    assert.ok(result.includes('UID:test@fm'));
    assert.ok(result.includes('DTSTART:20260401T100000Z'));
  });

  it('patches parameterized DTSTART (TZID form)', () => {
    const event = simpleEvent.replace('DTSTART:20260401T100000Z', 'DTSTART;TZID=Europe/Rome:20260401T100000');
    const result = replaceICalProperty(event, 'DTSTART', 'DTSTART;TZID=Europe/Rome:20260402T090000');
    assert.ok(result.includes('DTSTART;TZID=Europe/Rome:20260402T090000'));
    assert.ok(!result.includes('20260401'));
  });

  it('inserts a new property before END:VEVENT', () => {
    const result = replaceICalProperty(simpleEvent, 'LOCATION', 'LOCATION:Room A');
    assert.ok(result.includes('LOCATION:Room A'));
    const lines = result.split('\n');
    const locIdx = lines.findIndex(l => l === 'LOCATION:Room A');
    const endIdx = lines.findIndex(l => l === 'END:VEVENT');
    assert.ok(locIdx < endIdx);
  });

  it('removes a property when newLine is null', () => {
    const result = replaceICalProperty(simpleEvent, 'SUMMARY', null);
    assert.ok(!result.includes('SUMMARY'));
  });

  it('only patches first VEVENT in multi-VEVENT iCal', () => {
    const multiVevent = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'SUMMARY:Master',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'RECURRENCE-ID:20260401T100000Z',
      'SUMMARY:Exception',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(multiVevent, 'SUMMARY', 'SUMMARY:Updated Master');
    assert.ok(result.includes('SUMMARY:Updated Master'));
    assert.ok(result.includes('SUMMARY:Exception'));
  });

  it('skips properties inside VALARM', () => {
    const eventWithValarm = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:test@fm',
      'DESCRIPTION:Event description',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'ACTION:DISPLAY',
      'DESCRIPTION:Reminder',
      'END:VALARM',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(eventWithValarm, 'DESCRIPTION', 'DESCRIPTION:Updated event');
    assert.ok(result.includes('DESCRIPTION:Updated event'));
    assert.ok(result.includes('DESCRIPTION:Reminder')); // VALARM DESCRIPTION preserved
  });

  it('does not touch VTIMEZONE DTSTART', () => {
    const eventWithTz = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'DTSTART:19700101T000000',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'SUMMARY:Test',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(eventWithTz, 'DTSTART', 'DTSTART;TZID=Europe/Rome:20260402T090000');
    assert.ok(result.includes('DTSTART:19700101T000000')); // VTIMEZONE preserved
    assert.ok(result.includes('DTSTART;TZID=Europe/Rome:20260402T090000'));
  });

  it('throws on missing BEGIN:VEVENT', () => {
    assert.throws(() => replaceICalProperty('no vevent', 'SUMMARY', 'SUMMARY:x'), /BEGIN:VEVENT not found/);
  });

  it('throws on empty input', () => {
    assert.throws(() => replaceICalProperty('', 'SUMMARY', 'SUMMARY:x'), /empty input/);
  });

  it('handles folded property lines', () => {
    const event = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'SUMMARY:Very long title that was',
      ' folded across two lines',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:Short');
    assert.ok(result.includes('SUMMARY:Short'));
    assert.ok(!result.includes('folded across'));
  });
});

describe('removeAllICalProperties', () => {
  it('removes multiple ATTENDEEs', () => {
    const event = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:test@fm',
      'SUMMARY:Test',
      'ATTENDEE;CN=Alice:mailto:alice@example.com',
      'ATTENDEE;CN=Bob:mailto:bob@example.com',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(!result.includes('ATTENDEE'));
    assert.ok(result.includes('ORGANIZER'));
    assert.ok(result.includes('SUMMARY:Test'));
  });

  it('handles folded ATTENDEE lines', () => {
    const event = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'ATTENDEE;CN=Very Long Name;PARTSTAT=ACCEPTED:mailto:long',
      ' name@example.com',
      'SUMMARY:Test',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(!result.includes('ATTENDEE'));
    assert.ok(!result.includes('longname'));
    assert.ok(result.includes('SUMMARY:Test'));
  });

  it('preserves CRLF line endings', () => {
    const event = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nATTENDEE:mailto:a@b.com\r\nSUMMARY:Test\r\nEND:VEVENT\r\nEND:VCALENDAR';
    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(result.includes('\r\n'));
    assert.ok(!result.includes('ATTENDEE'));
  });

  it('preserves LF-only line endings', () => {
    const event = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nATTENDEE:mailto:a@b.com\nSUMMARY:Test\nEND:VEVENT\nEND:VCALENDAR';
    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(!result.includes('\r\n'));
    assert.ok(!result.includes('ATTENDEE'));
  });
});

describe('foldICalLine with custom line ending', () => {
  it('uses LF when specified', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long, '\n');
    assert.ok(!folded.includes('\r'));
    assert.ok(folded.includes('\n'));
  });

  it('defaults to CRLF', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long);
    assert.ok(folded.includes('\r\n'));
  });
});

describe('removeOrphanedVTimezones', () => {
  it('removes VTIMEZONE with no references', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART:20260401T100000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeOrphanedVTimezones(data);
    assert.ok(!result.includes('VTIMEZONE'));
    assert.ok(!result.includes('Europe/Rome'));
  });

  it('preserves VTIMEZONE with references', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeOrphanedVTimezones(data);
    assert.ok(result.includes('VTIMEZONE'));
    assert.ok(result.includes('Europe/Rome'));
  });
});

describe('removeExceptionVEvents', () => {
  it('removes only orphaned exception VEVENTs', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'DTSTART:20260401T100000Z',
      'RRULE:FREQ=WEEKLY',
      'SUMMARY:Master',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'RECURRENCE-ID:20260408T100000Z',
      'SUMMARY:Exception 1',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'RECURRENCE-ID:20260415T100000Z',
      'SUMMARY:Exception 2',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    // Only orphan the April 8 exception
    const orphaned = [new Date('2026-04-08T10:00:00Z')];
    const result = removeExceptionVEvents(data, orphaned);
    assert.ok(result.includes('SUMMARY:Master'));
    assert.ok(!result.includes('Exception 1'));
    assert.ok(result.includes('Exception 2'));
  });

  it('never touches master VEVENT', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'SUMMARY:Master',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeExceptionVEvents(data, [new Date()]);
    assert.ok(result.includes('SUMMARY:Master'));
  });
});

describe('validateAttendeeEmail', () => {
  it('accepts valid email', () => {
    assert.doesNotThrow(() => validateAttendeeEmail('alice@example.com'));
  });

  it('rejects empty email', () => {
    assert.throws(() => validateAttendeeEmail(''), /required/);
  });

  it('rejects email without @', () => {
    assert.throws(() => validateAttendeeEmail('notanemail'), /Invalid/);
  });

  it('rejects email with newline', () => {
    assert.throws(() => validateAttendeeEmail('a@b.com\r\nX-INJECT:true'), /illegal/i);
  });

  it('rejects email with colon', () => {
    assert.throws(() => validateAttendeeEmail('a:b@example.com'), /illegal/i);
  });

  it('rejects email with semicolon', () => {
    assert.throws(() => validateAttendeeEmail('a;b@example.com'), /illegal/i);
  });

  it('rejects email with double quote', () => {
    assert.throws(() => validateAttendeeEmail('a"b@example.com'), /illegal/i);
  });

  it('rejects email with backslash', () => {
    assert.throws(() => validateAttendeeEmail('a\\b@example.com'), /illegal/i);
  });

  it('rejects email with whitespace', () => {
    assert.throws(() => validateAttendeeEmail('a b@example.com'), /illegal/i);
    assert.throws(() => validateAttendeeEmail('a\tb@example.com'), /illegal/i);
  });
});

describe('quoteParamValue', () => {
  it('returns unquoted for simple values', () => {
    assert.equal(quoteParamValue('Alice'), 'Alice');
  });

  it('quotes values with comma', () => {
    assert.equal(quoteParamValue('Doe, John'), '"Doe, John"');
  });

  it('quotes values with semicolon', () => {
    assert.equal(quoteParamValue('A;B'), '"A;B"');
  });

  it('replaces double quotes with single quotes', () => {
    assert.equal(quoteParamValue('He said "hi"'), '"He said \'hi\'"');
  });

  it('strips newlines to prevent injection', () => {
    // Colon in result triggers DQUOTE quoting — that's correct
    assert.equal(quoteParamValue('Alice\r\nX-EVIL:payload'), '"Alice X-EVIL:payload"');
  });
});

describe('parseICalValue with CRLF and folded lines', () => {
  it('handles CRLF input with folded lines', () => {
    const vevent = 'DESCRIPTION:This is a long\r\n description that wraps\r\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), 'This is a longdescription that wraps');
  });

  it('uses quote-aware colon detection', () => {
    const vevent = 'ATTENDEE;DELEGATED-FROM="mailto:boss@example.com":mailto:alice@example.com\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'ATTENDEE'), 'mailto:alice@example.com');
  });
});

describe('toICalUTC', () => {
  it('throws on date-only input', () => {
    assert.throws(() => toICalUTC('2026-04-01'), /date-only input must be handled by caller/);
  });
});

describe('CalDAVCalendarClient.updateCalendarEvent (patch-based)', () => {
  function makeRichIcal(uid: string, extra: string[] = []): string {
    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Google Inc//Google Calendar//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTAMP:20260401T000000Z',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'DTEND;TZID=Europe/Rome:20260401T110000',
      'SUMMARY:Original Title',
      'DESCRIPTION:Original description',
      'LOCATION:Room A',
      'ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED:mailto:alice@example.com',
      'ATTENDEE;CN=Bob;PARTSTAT=TENTATIVE:mailto:bob@example.com',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'ACTION:DISPLAY',
      'DESCRIPTION:Reminder',
      'END:VALARM',
      'X-GOOGLE-CONFERENCE:https://meet.google.com/abc-def-ghi',
      'SEQUENCE:2',
      ...extra,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedPatchClient(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
      updateCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('preserves unknown properties when updating title only', async () => {
    const ical = makeRichIcal('evt1@fm');
    const objects = [{ data: ical, url: '/cal/evt1.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt1@fm', { title: 'New Title' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SUMMARY:New Title'));
    // Preserved properties
    assert.ok(updatedData.includes('ATTENDEE;CN=Alice'));
    assert.ok(updatedData.includes('ATTENDEE;CN=Bob'));
    assert.ok(updatedData.includes('ORGANIZER;CN=Boss'));
    assert.ok(updatedData.includes('VALARM'));
    assert.ok(updatedData.includes('X-GOOGLE-CONFERENCE'));
    assert.ok(updatedData.includes('DESCRIPTION:Original description'));
    assert.ok(updatedData.includes('LOCATION:Room A'));
  });

  it('does NOT re-emit DTSTART when only title changes (timezone preservation)', async () => {
    const ical = makeRichIcal('evt2@fm');
    const objects = [{ data: ical, url: '/cal/evt2.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt2@fm', { title: 'New Title' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    // Original DTSTART with TZID should be preserved exactly
    assert.ok(updatedData.includes('DTSTART;TZID=Europe/Rome:20260401T100000'));
  });

  it('increments SEQUENCE for location change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('evt3@fm');
    const objects = [{ data: ical, url: '/cal/evt3.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt3@fm', { location: 'New Room' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3')); // Was 2, now 3
  });

  it('does NOT increment SEQUENCE for title-only changes', async () => {
    const ical = makeRichIcal('evt4@fm');
    const objects = [{ data: ical, url: '/cal/evt4.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt4@fm', { title: 'New Title' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:2')); // Unchanged
  });

  it('does NOT increment SEQUENCE when no ATTENDEEs exist', async () => {
    const noAttendeeIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:solo@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Solo Event',
      'SEQUENCE:0',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noAttendeeIcal, url: '/cal/solo.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('solo@fm', { start: '2026-04-02T10:00:00Z' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:0'));
  });

  it('removes DURATION when setting end', async () => {
    const durationIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur@fm',
      'DTSTART:20260401T100000Z',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: durationIcal, url: '/cal/dur.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('dur@fm', { end: '2026-04-01T13:00:00Z' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('DTEND:20260401T130000Z'));
    assert.ok(!updatedData.includes('DURATION'));
  });

  it('does NOT remove DURATION when only setting start', async () => {
    const durationIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur2@fm',
      'DTSTART:20260401T100000Z',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: durationIcal, url: '/cal/dur2.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('dur2@fm', { start: '2026-04-02T10:00:00Z' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('DURATION:PT2H'));
  });

  it('participants: [] removes all ATTENDEEs, preserves ORGANIZER', async () => {
    const ical = makeRichIcal('evt5@fm');
    const objects = [{ data: ical, url: '/cal/evt5.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt5@fm', { participants: [] });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(!updatedData.includes('ATTENDEE'));
    assert.ok(updatedData.includes('ORGANIZER'));
  });

  it('floating end time preserves DTEND TZID', async () => {
    const ical = makeRichIcal('evt6@fm');
    const objects = [{ data: ical, url: '/cal/evt6.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt6@fm', { end: '2026-04-01T12:00:00' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('DTEND;TZID=Europe/Rome:20260401T120000'));
  });

  it('date-only start emits VALUE=DATE', async () => {
    const allDayIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday@fm',
      'DTSTART;VALUE=DATE:20260401',
      'DTEND;VALUE=DATE:20260402',
      'SUMMARY:All Day',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: allDayIcal, url: '/cal/allday.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('allday@fm', { start: '2026-04-05' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('DTSTART;VALUE=DATE:20260405'));
  });

  it('throws on event with no VEVENT', async () => {
    // Simulate an object found by URL but with no VEVENT block
    const noVeventIcal = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR';
    const objects = [{ data: noVeventIcal, url: '/cal/bad.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('/cal/bad.ics', { title: 'X' }),
      /no iCal data|not found/
    );
  });

  it('throws when VEVENT block is malformed (BEGIN without END)', async () => {
    // Has BEGIN:VEVENT (passes string check) but no END:VEVENT (extractVEvent returns null)
    const brokenIcal = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:broken@fm\r\nEND:VCALENDAR';
    const objects = [{ data: brokenIcal, url: '/cal/broken.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('/cal/broken.ics', { title: 'X' }),
      /not found|no VEVENT/
    );
  });

  it('throws on malformed start date format', async () => {
    const ical = makeRichIcal('evtbad@fm');
    const objects = [{ data: ical, url: '/cal/evtbad.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('evtbad@fm', { start: 'not-a-date' }),
      /Invalid start date format/
    );
  });

  it('throws on malformed end date format', async () => {
    const ical = makeRichIcal('evtbad2@fm');
    const objects = [{ data: ical, url: '/cal/evtbad2.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('evtbad2@fm', { end: 'garbage' }),
      /Invalid end date format/
    );
  });

  it('throws when adding participants with non-email CalDAV username', async () => {
    const noOrganizerIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noorg2@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Simple Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noOrganizerIcal, url: '/cal/noorg2.ics' }];
    // Use non-email username
    const client = new CalDAVCalendarClient({ username: 'not-an-email', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async () => objects),
      updateCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.updateCalendarEvent('noorg2@fm', {
        participants: [{ email: 'alice@example.com' }],
      }),
      /CalDAV username is not an email/
    );
  });

  it('adds ORGANIZER when adding participants to event with no existing ORGANIZER', async () => {
    const noOrganizerIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noorg@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Simple Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noOrganizerIcal, url: '/cal/noorg.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('noorg@fm', {
      participants: [{ email: 'alice@example.com', name: 'Alice' }],
    });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('ORGANIZER'));
    assert.ok(updatedData.includes('mailto:test@fastmail.com'));
    assert.ok(updatedData.includes('ATTENDEE'));
  });

  it('updates description only', async () => {
    const ical = makeRichIcal('evtdesc@fm');
    const objects = [{ data: ical, url: '/cal/evtdesc.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evtdesc@fm', { description: 'New description' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('DESCRIPTION:New description'));
    assert.ok(updatedData.includes('SUMMARY:Original Title')); // Other fields preserved
  });

  it('updates DTSTAMP and LAST-MODIFIED', async () => {
    const ical = makeRichIcal('evt7@fm');
    const objects = [{ data: ical, url: '/cal/evt7.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt7@fm', { title: 'X' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('LAST-MODIFIED:'));
    // DTSTAMP should be updated (not the original)
    assert.ok(!updatedData.includes('DTSTAMP:20260401T000000Z'));
  });

  it('preserves Google PRODID (does not stamp ours)', async () => {
    const ical = makeRichIcal('evt8@fm');
    const objects = [{ data: ical, url: '/cal/evt8.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt8@fm', { title: 'X' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('PRODID:-//Google Inc//Google Calendar//EN'));
  });
});

describe('CalDAVCalendarClient.updateCalendarEvent recurring events', () => {
  function makeRecurringIcal(exceptions: Array<{ recurrenceId: string; summary: string }> = []): string {
    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VEVENT',
      'UID:recur@fm',
      'DTSTAMP:20260401T000000Z',
      'DTSTART:20260406T100000Z',
      'DTEND:20260406T110000Z',
      'RRULE:FREQ=WEEKLY;COUNT=10',
      'SUMMARY:Weekly Meeting',
      'SEQUENCE:0',
      'END:VEVENT',
    ];
    for (const exc of exceptions) {
      lines.push(
        'BEGIN:VEVENT',
        'UID:recur@fm',
        `RECURRENCE-ID:${exc.recurrenceId}`,
        'DTSTAMP:20260401T000000Z',
        `DTSTART:${exc.recurrenceId}`,
        `DTEND:${exc.recurrenceId.replace('T10', 'T11')}`,
        `SUMMARY:${exc.summary}`,
        'END:VEVENT',
      );
    }
    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
  }

  function createMockedRecurringClient(ical: string) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const objects = [{ data: ical, url: '/cal/recur.ics' }];
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => objects),
      updateCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('time change with no exceptions proceeds without confirmRecurring', async () => {
    const ical = makeRecurringIcal([]); // No exceptions
    const { client, mockDAVClient } = createMockedRecurringClient(ical);

    await client.updateCalendarEvent('recur@fm', { start: '2026-04-07T10:00:00Z' });

    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('DTSTART:20260407T100000Z'));
  });

  it('time change with exceptions throws without confirmRecurring', async () => {
    const ical = makeRecurringIcal([
      { recurrenceId: '20260413T100000Z', summary: 'Exception Week 2' },
      { recurrenceId: '20260420T100000Z', summary: 'Exception Week 3' },
    ]);
    const { client } = createMockedRecurringClient(ical);

    await assert.rejects(
      () => client.updateCalendarEvent('recur@fm', { start: '2026-04-07T10:00:00Z' }),
      /confirmRecurring/
    );
  });

  it('time change with confirmRecurring: true removes orphaned exceptions', async () => {
    const ical = makeRecurringIcal([
      { recurrenceId: '20260413T100000Z', summary: 'Exception Week 2' },
    ]);
    const { client, mockDAVClient } = createMockedRecurringClient(ical);

    // Shift start by one day — the exception on April 13 (Monday) won't match
    // the new Tuesday schedule
    await client.updateCalendarEvent('recur@fm', {
      start: '2026-04-07T10:00:00Z',
      confirmRecurring: true,
    });

    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SUMMARY:Weekly Meeting')); // Master preserved
    assert.ok(updatedData.includes('DTSTART:20260407T100000Z')); // Start updated
  });

  it('time change with offset-bearing timestamp works for orphan detection', async () => {
    const ical = makeRecurringIcal([
      { recurrenceId: '20260413T100000Z', summary: 'Exception Week 2' },
    ]);
    const { client } = createMockedRecurringClient(ical);

    // Offset timestamp should be converted to UTC for rrule parsing
    await assert.rejects(
      () => client.updateCalendarEvent('recur@fm', { start: '2026-04-07T20:00:00+10:00' }),
      /confirmRecurring/
    );
  });

  it('confirmRecurring: true keeps valid exceptions and removes orphaned ones', async () => {
    // Weekly on Monday (April 6). Exception on April 13 (week 2) and April 20 (week 3).
    // Shift start by exactly 1 week — April 13 becomes the new first occurrence,
    // so the April 13 exception is still valid, but April 20 should also still be valid
    // (it's still week 3 of the new schedule).
    // Actually, let's shift to Tuesday (April 7). Now occurrences are April 7, 14, 21, 28...
    // Exception on April 13 (Monday) is orphaned. Exception on April 14 would match but
    // we have April 20 which is also orphaned (not a Tuesday).
    const ical = makeRecurringIcal([
      { recurrenceId: '20260413T100000Z', summary: 'Exception Mon Week 2' },
      { recurrenceId: '20260414T100000Z', summary: 'Exception Tue Week 2' },
    ]);
    const { client, mockDAVClient } = createMockedRecurringClient(ical);

    await client.updateCalendarEvent('recur@fm', {
      start: '2026-04-07T10:00:00Z',
      confirmRecurring: true,
    });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SUMMARY:Weekly Meeting')); // Master preserved
    // April 14 (Tuesday) should be kept — matches new weekly Tuesday schedule
    assert.ok(updatedData.includes('Exception Tue Week 2'));
    // April 13 (Monday) should be removed — doesn't match new Tuesday schedule
    assert.ok(!updatedData.includes('Exception Mon Week 2'));
  });

  it('non-time change on recurring event does not require confirmRecurring', async () => {
    const ical = makeRecurringIcal([
      { recurrenceId: '20260413T100000Z', summary: 'Exception Week 2' },
    ]);
    const { client, mockDAVClient } = createMockedRecurringClient(ical);

    // Title-only change should not trigger the recurring guard
    await client.updateCalendarEvent('recur@fm', { title: 'New Title' });

    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SUMMARY:New Title'));
    assert.ok(updatedData.includes('Exception Week 2')); // Exception preserved
  });
});

describe('CalDAVCalendarClient.createCalendarEvent with participants', () => {
  function createMockedCreateClient() {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      createCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('includes ATTENDEE and ORGANIZER lines with mailto URI', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [
        { email: 'alice@example.com', name: 'Alice' },
        { email: 'bob@example.com', name: 'Bob' },
      ],
    });

    const ical = mockDAVClient.createCalendarObject.mock.calls[0].arguments[0].iCalString;
    assert.ok(ical.includes(':mailto:me@fastmail.com'), 'ORGANIZER should have mailto URI');
    assert.ok(/ORGANIZER;CN=.+:mailto:me@fastmail.com/.test(ical), 'ORGANIZER should have CN parameter');
    assert.ok(ical.includes('ATTENDEE;CN=Alice:mailto:alice@example.com'));
    assert.ok(ical.includes('ATTENDEE;CN=Bob:mailto:bob@example.com'));
  });

  it('does not include ATTENDEE/ORGANIZER when no participants', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Solo Event',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
    });

    const ical = mockDAVClient.createCalendarObject.mock.calls[0].arguments[0].iCalString;
    assert.ok(!ical.includes('ATTENDEE'));
    assert.ok(!ical.includes('ORGANIZER'));
  });

  it('does not emit RSVP=TRUE', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [{ email: 'alice@example.com' }],
    });

    const ical = mockDAVClient.createCalendarObject.mock.calls[0].arguments[0].iCalString;
    assert.ok(!ical.includes('RSVP'));
  });

  it('rejects email with injection attempt', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'Personal',
        title: 'Meeting',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
        participants: [{ email: 'a@b.com\r\nX-INJECT:true' }],
      }),
      /illegal/i
    );
  });

  it('uses DQUOTE quoting for CN, not backslash escaping', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [{ email: 'alice@example.com', name: 'Doe, Alice' }],
    });

    const ical = mockDAVClient.createCalendarObject.mock.calls[0].arguments[0].iCalString;
    // Should use DQUOTE quoting: CN="Doe, Alice"
    assert.ok(ical.includes('CN="Doe, Alice"'));
    // Should NOT use backslash escaping: CN=Doe\, Alice
    assert.ok(!ical.includes('CN=Doe\\, Alice'));
  });

  it('handles date-only start/end (all-day event)', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'All Day',
      start: '2026-04-07',
      end: '2026-04-08',
    });

    const ical = mockDAVClient.createCalendarObject.mock.calls[0].arguments[0].iCalString;
    assert.ok(ical.includes('DTSTART;VALUE=DATE:20260407'));
    assert.ok(ical.includes('DTEND;VALUE=DATE:20260408'));
  });

  it('rejects date-only start and end with same value', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'Personal',
        title: 'Bad',
        start: '2026-04-07',
        end: '2026-04-07',
      }),
      /DTEND is exclusive/
    );
  });

  it('ends with trailing CRLF', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Test',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
    });

    const ical = mockDAVClient.createCalendarObject.mock.calls[0].arguments[0].iCalString;
    assert.ok(ical.endsWith('\r\n'));
  });
});

describe('CRLF vs LF line ending preservation', () => {
  it('replaceICalProperty preserves LF-only endings', () => {
    const event = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nSUMMARY:Old\nEND:VEVENT\nEND:VCALENDAR';
    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:New');
    assert.ok(!result.includes('\r\n'));
    assert.ok(result.includes('SUMMARY:New'));
  });

  it('replaceICalProperty preserves CRLF endings', () => {
    const event = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nSUMMARY:Old\r\nEND:VEVENT\r\nEND:VCALENDAR';
    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:New');
    assert.ok(result.includes('\r\n'));
    assert.ok(result.includes('SUMMARY:New'));
  });

  it('no mixed line endings when patching LF-only input', () => {
    const event = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nSUMMARY:Old\nEND:VEVENT\nEND:VCALENDAR';
    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:New');
    assert.ok(!result.includes('\r\n'), 'Should not contain CRLF in LF-only input');
  });

  it('foldICalLine with wrong lineEnding produces consistent output', () => {
    // Simulate a caller using CRLF fold on what will be inserted into LF document
    const folded = foldICalLine('DESCRIPTION:' + 'x'.repeat(80), '\r\n');
    // The fold itself should use CRLF consistently
    assert.ok(folded.includes('\r\n'));
    // When replaceICalProperty re-splits and re-joins with LF, CRLF folds are preserved
    // inside the replacement line — this is the caller's responsibility to match
    const lfEvent = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nSUMMARY:Old\nEND:VEVENT\nEND:VCALENDAR';
    const result = replaceICalProperty(lfEvent, 'SUMMARY', foldICalLine('SUMMARY:Short', '\n'));
    assert.ok(!result.includes('\r\n'), 'Caller using correct lineEnding prevents mixing');
  });
});

describe('VEVENT extraction consistency', () => {
  it('extractVEvent and replaceICalProperty agree on VEVENT boundaries', () => {
    const multiVevent = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'SUMMARY:Master',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'RECURRENCE-ID:20260408T100000Z',
      'SUMMARY:Exception',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const extracted = extractVEvent(multiVevent)!;
    assert.ok(extracted.includes('SUMMARY:Master'));
    assert.ok(!extracted.includes('SUMMARY:Exception'));

    // replaceICalProperty should operate on the same first VEVENT
    const patched = replaceICalProperty(multiVevent, 'SUMMARY', 'SUMMARY:Updated');
    assert.ok(patched.includes('SUMMARY:Updated'));
    assert.ok(patched.includes('SUMMARY:Exception')); // Second VEVENT untouched
  });
});

describe('Additional plan-required updateCalendarEvent tests', () => {
  function makeRichIcal(uid: string, extra: string[] = []): string {
    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Google Inc//Google Calendar//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTAMP:20260401T000000Z',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'DTEND;TZID=Europe/Rome:20260401T110000',
      'SUMMARY:Original Title',
      'DESCRIPTION:Original description',
      'LOCATION:Room A',
      'ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED:mailto:alice@example.com',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'ACTION:DISPLAY',
      'DESCRIPTION:Reminder',
      'END:VALARM',
      'SEQUENCE:2',
      ...extra,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedClient(calendarObjects: Array<{ data: string; url: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async () => calendarObjects),
      updateCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('increments SEQUENCE for start change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('seqstart@fm');
    const { client, mockDAVClient } = createMockedClient([{ data: ical, url: '/cal/seqstart.ics' }]);

    await client.updateCalendarEvent('seqstart@fm', { start: '2026-04-02T10:00:00' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3'));
  });

  it('increments SEQUENCE for end change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('seqend@fm');
    const { client, mockDAVClient } = createMockedClient([{ data: ical, url: '/cal/seqend.ics' }]);

    await client.updateCalendarEvent('seqend@fm', { end: '2026-04-01T12:00:00' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3'));
  });

  it('increments SEQUENCE for participants change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('seqpart@fm');
    const { client, mockDAVClient } = createMockedClient([{ data: ical, url: '/cal/seqpart.ics' }]);

    await client.updateCalendarEvent('seqpart@fm', {
      participants: [{ email: 'new@example.com', name: 'New' }],
    });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3'));
  });

  it('floating end time falls back to DTSTART TZID when DTEND was DURATION-computed', async () => {
    const durationIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:durtz@fm',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedClient([{ data: durationIcal, url: '/cal/durtz.ics' }]);

    await client.updateCalendarEvent('durtz@fm', { end: '2026-04-01T13:00:00' });

    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    // Should fall back to DTSTART's TZID since there was no DTEND
    assert.ok(updatedData.includes('DTEND;TZID=Europe/Rome:20260401T130000'));
  });

  it('rejects date-only start and end with same value on update', async () => {
    const allDayIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday2@fm',
      'DTSTART;VALUE=DATE:20260401',
      'DTEND;VALUE=DATE:20260402',
      'SUMMARY:All Day',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client } = createMockedClient([{ data: allDayIcal, url: '/cal/allday2.ics' }]);

    await assert.rejects(
      () => client.updateCalendarEvent('allday2@fm', { start: '2026-04-05', end: '2026-04-05' }),
      /DTEND is exclusive/
    );
  });
});

describe('Recurring event: no orphans proceeds without confirmRecurring', () => {
  it('time-only shift that keeps all exceptions valid proceeds without confirmRecurring', async () => {
    // Weekly on Monday at 10:00. Exception on April 13 at 10:00 (still a Monday).
    // Shift time to 11:00 but same day — April 13 at 11:00 is still week 2.
    // The exception's RECURRENCE-ID (April 13T10:00Z) should still match
    // because rrule expansion produces occurrences on the same dates.
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noorphan@fm',
      'DTSTAMP:20260401T000000Z',
      'DTSTART:20260406T100000Z',
      'DTEND:20260406T110000Z',
      'RRULE:FREQ=WEEKLY;COUNT=10',
      'SUMMARY:Weekly',
      'SEQUENCE:0',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:noorphan@fm',
      'RECURRENCE-ID:20260413T100000Z',
      'DTSTAMP:20260401T000000Z',
      'DTSTART:20260413T100000Z',
      'DTEND:20260413T110000Z',
      'SUMMARY:Modified Week 2',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const objects = [{ data: ical, url: '/cal/noorphan.ics' }];
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async () => objects),
      updateCalendarObject: mock.fn(async () => ({})),
    };
    (client as any).client = mockDAVClient;

    // Change only end time (not start) — RRULE anchor stays the same, no orphans
    await client.updateCalendarEvent('noorphan@fm', { end: '2026-04-06T12:00:00Z' });

    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedData = mockDAVClient.updateCalendarObject.mock.calls[0].arguments[0].calendarObject.data;
    assert.ok(updatedData.includes('Modified Week 2')); // Exception preserved
    assert.ok(updatedData.includes('DTEND:20260406T120000Z')); // End updated
  });
});
