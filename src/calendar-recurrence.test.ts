import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { fetchCalendarObjects } from 'tsdav';
import { CalDAVCalendarClient } from './caldav-client.js';

const start = '2026-10-01T00:00:00Z';
const end = '2026-11-01T00:00:00Z';
const vevent = (...lines: string[]) => ['BEGIN:VEVENT', 'UID:synthetic-recurring-series', ...lines, 'END:VEVENT'].join('\r\n');
const calendar = (...events: string[]) => ['BEGIN:VCALENDAR', 'VERSION:2.0', ...events, 'END:VCALENDAR'].join('\r\n');
const occurrence = (day: string, extra: string[] = []) => vevent(
  `RECURRENCE-ID:202610${day}T080000Z`, `DTSTART:202610${day}T080000Z`,
  `DTEND:202610${day}T090000Z`, 'SUMMARY:Weekly lesson', ...extra,
);
function clientFor(data: string) {
  const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
  const dav = {
    fetchCalendars: mock.fn(async () => [{ displayName: 'Test', url: 'https://example.test/cal/' }]),
    fetchCalendarObjects: mock.fn(async (_options: any) => [{ data, url: 'https://example.test/cal/series.ics' }]),
  };
  (client as any).client = dav;
  return { client, dav };
}

describe('bounded recurrence listing', () => {
  it('requests server expansion and returns all dated occurrences, including a moved exception', async () => {
    // CalDAV expansion has already applied RRULE, RDATE, EXDATE and overrides.
    // October 13 is excluded; October 20 is moved to October 21 at a new time.
    const { client, dav } = clientFor(calendar(
      occurrence('06'),
      vevent('RECURRENCE-ID:20261020T080000Z', 'DTSTART:20261021T120000Z',
        'DTEND:20261021T130000Z', 'SUMMARY:Moved lesson'),
      occurrence('27'),
    ));
    const events = await client.getCalendarEvents(undefined, 50, start, end);
    assert.equal(dav.fetchCalendarObjects.mock.calls[0].arguments[0].expand, true);
    assert.deepEqual(events.map(e => e.start), ['2026-10-06T08:00:00Z', '2026-10-21T12:00:00Z', '2026-10-27T08:00:00Z']);
    assert.equal(events[1].recurrenceId, '2026-10-20T08:00:00Z');
    assert.equal(events[1].title, 'Moved lesson');
    assert.ok(events.every(e => e.id === 'synthetic-recurring-series'));
    assert.ok(events.every(e => e.url === 'https://example.test/cal/series.ics'));
  });

  it('omits cancelled instances and handles an empty expansion without a phantom event', async () => {
    const { client } = clientFor(calendar(occurrence('06', ['STATUS:CANCELLED']), occurrence('13')));
    assert.equal((await client.getCalendarEvents(undefined, 50, start, end)).length, 1);
    const empty = clientFor(calendar()).client;
    assert.deepEqual(await empty.getCalendarEvents(undefined, 50, start, end), []);
  });

  it('preserves all-day dates, duration and UTC instants across a DST transition', async () => {
    // Paris 09:00 is 07:00Z before and 08:00Z after October's DST transition.
    const { client } = clientFor(calendar(
      vevent('RECURRENCE-ID:20261020T070000Z', 'DTSTART:20261020T070000Z', 'DURATION:PT1H'),
      occurrence('27'),
      vevent('RECURRENCE-ID;VALUE=DATE:20261028', 'DTSTART;VALUE=DATE:20261028', 'DTEND;VALUE=DATE:20261029'),
    ));
    const events = await client.getCalendarEvents(undefined, 50, start, end);
    assert.equal(events[0].start, '2026-10-20T07:00:00Z');
    assert.equal(events[0].end, '2026-10-20T08:00:00Z');
    assert.equal(events[1].start, '2026-10-27T08:00:00Z');
    assert.equal(events[2].start, '2026-10-28');
    assert.equal(events[2].end, '2026-10-29');
    assert.equal(events[2].recurrenceId, '2026-10-28');
  });

  it('fails explicitly if a server ignores expand, rather than returning an old master as today', async () => {
    for (const recurrence of ['RRULE:FREQ=WEEKLY', 'RDATE:20261013T080000Z']) {
      const { client } = clientFor(calendar(vevent('DTSTART:20200101T080000Z', recurrence)));
      await assert.rejects(client.getCalendarEvents(undefined, 50, start, end), /unexpanded recurrence/);
    }
  });

  it('preserves series-level unbounded and one-sided queries without expanding decades', async () => {
    for (const bounds of [[], [start], [undefined, end]] as Array<Array<string | undefined>>) {
      const { client, dav } = clientFor(calendar(vevent('DTSTART:20200101T080000Z', 'RRULE:FREQ=WEEKLY'), occurrence('06')));
      const events = await client.getCalendarEvents(undefined, 50, bounds[0], bounds[1]);
      assert.equal(events.length, 1);
      assert.equal(events[0].start, '2020-01-01T08:00:00Z');
      assert.equal(dav.fetchCalendarObjects.mock.calls[0].arguments[0].expand, undefined);
    }
  });

  it('sorts across all calendars before applying the occurrence limit', async () => {
    const { client, dav } = clientFor('');
    dav.fetchCalendars = mock.fn(async () => [
      { displayName: 'Late', url: 'https://example.test/late/' },
      { displayName: 'Early', url: 'https://example.test/early/' },
    ]);
    dav.fetchCalendarObjects = mock.fn(async (options: any) => [{
      data: calendar(occurrence(options.calendar.displayName === 'Late' ? '27' : '06')),
      url: options.calendar.url + 'event.ics',
    }]);
    const events = await client.getCalendarEvents(undefined, 1, start, end);
    assert.equal(events.length, 1);
    assert.equal(events[0].start, '2026-10-06T08:00:00Z');
    assert.equal(dav.fetchCalendarObjects.mock.calls.length, 2);
  });

  it('uses the installed tsdav transport to request c:expand and retain every VEVENT', async () => {
    const data = calendar(occurrence('06'), occurrence('13'));
    const requests: string[] = [];
    const objects = await fetchCalendarObjects({
      calendar: { url: 'https://example.test/cal/' }, timeRange: { start, end }, expand: true,
      fetch: async (_url: any, options: any) => {
        requests.push(options.body);
        assert.equal(options.method, 'REPORT');
        return new Response(`<?xml version="1.0"?><d:multistatus xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:response><d:href>/cal/series.ics</d:href><d:propstat><d:prop><d:getetag>"v1"</d:getetag><c:calendar-data><![CDATA[${data}]]></c:calendar-data></d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response></d:multistatus>`, { status: 207, headers: { 'Content-Type': 'application/xml' } });
      },
    });
    assert.equal(requests.length, 1);
    assert.match(requests[0], /c:expand[^>]*start="20261001T000000Z"[^>]*end="20261101T000000Z"/);
    const { client, dav } = clientFor('');
    dav.fetchCalendarObjects = mock.fn(async () => objects as any);
    assert.equal((await client.getCalendarEvents(undefined, 50, start, end)).length, 2);
  });
});
