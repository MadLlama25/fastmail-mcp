import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import {
  extractVEvent,
  parseICalValue,
  formatICalDate,
  parseCalendarObject,
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
