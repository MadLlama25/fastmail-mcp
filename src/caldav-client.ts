import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';
import rruleLib from 'rrule';
const { rrulestr } = rruleLib;

export interface CalDAVConfig {
  username: string;
  password: string;
  serverUrl?: string;
}

export interface CalendarInfo {
  id: string;
  displayName: string;
  url: string;
  description?: string;
  color?: string;
}

export interface Participant {
  email: string;
  name?: string;
  role?: string;       // REQ-PARTICIPANT, OPT-PARTICIPANT, CHAIR
  status?: string;     // PARTSTAT: ACCEPTED, DECLINED, TENTATIVE, NEEDS-ACTION
  cutype?: string;     // CUTYPE: INDIVIDUAL, ROOM, RESOURCE, GROUP, UNKNOWN
  rsvp?: boolean;      // RSVP: TRUE/FALSE
}

export interface CalendarEvent {
  id: string;
  url: string;
  title: string;
  description?: string;
  start?: string;
  end?: string;
  location?: string;
  organizer?: Participant;
  participants?: Participant[];
}

/**
 * Extract the VEVENT block from iCalendar data.
 * This avoids matching properties from VTIMEZONE or other components.
 */
export function extractVEvent(data: string): string | null {
  const match = data.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/);
  return match ? match[0] : null;
}

/**
 * Find the index of the first colon that separates iCal property parameters
 * from the property value. Colons inside quoted parameter values (e.g.
 * DELEGATED-FROM="mailto:boss@example.com") are skipped.
 * Also correctly handles properties like DESCRIPTION;ALTREP="http://...":text
 */
export function findValueBoundary(line: string): number {
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inQuote = !inQuote;
    } else if (ch === ':' && !inQuote) {
      return i;
    }
  }
  return -1;
}

/**
 * Parse an iCalendar property value from within a VEVENT block.
 * Handles simple (KEY:value), parameterized (KEY;TZID=...:value),
 * and VALUE=DATE (KEY;VALUE=DATE:20260319) forms.
 * Also handles line folding (continuation lines starting with space/tab).
 */
export function parseICalValue(vevent: string, key: string): string | undefined {
  // Match KEY followed by either ; (params) or : (value), capturing the rest
  const regex = new RegExp(`^(${key}[;:].*)$`, 'm');
  const match = vevent.match(regex);
  if (!match) return undefined;

  // Strip trailing \r that multiline regex captures on CRLF input
  let fullLine = match[1].replace(/\r$/, '');
  const lines = vevent.split(/\r?\n/);
  const matchIdx = lines.findIndex(l => l === fullLine);
  if (matchIdx >= 0) {
    for (let i = matchIdx + 1; i < lines.length; i++) {
      if (lines[i].startsWith(' ') || lines[i].startsWith('\t')) {
        fullLine += lines[i].substring(1);
      } else {
        break;
      }
    }
  }

  // Use quote-aware colon detection for the parameter/value boundary
  const colonIdx = findValueBoundary(fullLine);
  if (colonIdx === -1) return undefined;
  return fullLine.substring(colonIdx + 1).trim();
}

/**
 * Return all occurrences of a property key as full unfolded raw lines.
 * Needed because ATTENDEE/EXDATE etc. can appear multiple times.
 */
export function parseAllICalProperties(vevent: string, key: string): string[] {
  const lines = vevent.split(/\r?\n/);
  const regex = new RegExp(`^${key}[;:]`);
  const results: string[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/\r$/, '');
    if (!regex.test(line)) continue;

    // Unfold continuation lines
    let fullLine = line;
    for (let j = i + 1; j < lines.length; j++) {
      const next = lines[j];
      if (next.startsWith(' ') || next.startsWith('\t')) {
        fullLine += next.substring(1);
        i = j; // skip continuation lines in outer loop
      } else {
        break;
      }
    }
    results.push(fullLine);
  }

  return results;
}

/**
 * Parse a raw ATTENDEE or ORGANIZER line into a Participant.
 * Uses quote-aware scanning for parameter/value boundary detection.
 */
export function parseAttendee(rawLine: string): Participant {
  // Find the parameter/value boundary (first colon outside quotes)
  const boundaryIdx = findValueBoundary(rawLine);
  const paramPart = boundaryIdx >= 0 ? rawLine.substring(0, boundaryIdx) : rawLine;
  const valuePart = boundaryIdx >= 0 ? rawLine.substring(boundaryIdx + 1) : '';

  // Extract email from cal-address value
  const email = valuePart.replace(/^mailto:/i, '');

  // Split parameters on semicolons, respecting quotes
  const params: string[] = [];
  let current = '';
  let inQuote = false;
  // Skip the property name (ATTENDEE or ORGANIZER) — start after first ;
  const firstSemi = paramPart.indexOf(';');
  const paramStr = firstSemi >= 0 ? paramPart.substring(firstSemi + 1) : '';

  for (let i = 0; i < paramStr.length; i++) {
    const ch = paramStr[i];
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
    } else if (ch === ';' && !inQuote) {
      if (current) params.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  if (current) params.push(current);

  // Extract known parameters
  const result: Participant = { email };

  for (const param of params) {
    const eqIdx = param.indexOf('=');
    if (eqIdx === -1) continue;
    const pName = param.substring(0, eqIdx).toUpperCase();
    let pValue = param.substring(eqIdx + 1);
    // Strip surrounding quotes
    if (pValue.startsWith('"') && pValue.endsWith('"')) {
      pValue = pValue.slice(1, -1);
    }

    switch (pName) {
      case 'CN':
        if (pValue) result.name = pValue;
        break;
      case 'PARTSTAT':
        result.status = pValue;
        break;
      case 'ROLE':
        result.role = pValue;
        break;
      case 'CUTYPE':
        result.cutype = pValue;
        break;
      case 'RSVP':
        result.rsvp = pValue.toUpperCase() === 'TRUE';
        break;
    }
  }

  return result;
}

/**
 * Format an iCalendar date/datetime string to ISO 8601.
 * Input formats: 20260320T083000, 20260320T083000Z, 20260324
 * Output: 2026-03-20T08:30:00, 2026-03-20T08:30:00Z, 2026-03-24
 */
export function formatICalDate(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const cleaned = raw.replace(/\r/g, '');

  // All-day date: 20260324 (8 digits)
  if (/^\d{8}$/.test(cleaned)) {
    return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}`;
  }

  // DateTime: 20260320T083000 or 20260320T083000Z
  const dtMatch = cleaned.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
  if (dtMatch) {
    const [, y, m, d, hh, mm, ss, z] = dtMatch;
    return `${y}-${m}-${d}T${hh}:${mm}:${ss}${z}`;
  }

  return cleaned;
}

/**
 * Convert an ISO 8601 datetime string to iCalendar UTC format.
 * Handles timezone offsets by converting to UTC via Date.
 * Preserves floating times (no offset, no Z) as-is.
 * e.g. "2026-04-07T18:45:00+10:00" → "20260407T084500Z"
 */
export function toICalUTC(isoString: string): string {
  // Guard: date-only input must be handled by caller, not passed here
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoString)) {
    throw new Error('date-only input must be handled by caller, not passed to toICalUTC');
  }
  // Floating time (no offset, no Z) — preserve as local iCal datetime
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(isoString)) {
    return isoString.replace(/[-:]/g, '');
  }
  const d = new Date(isoString);
  if (isNaN(d.getTime())) throw new Error(`Invalid date: ${isoString}`);
  return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
}

/**
 * Fold an iCalendar content line at 75 octets per RFC 5545 §3.1.
 * @param lineEnding Line ending to use for fold breaks (default '\r\n')
 */
export function foldICalLine(line: string, lineEnding: string = '\r\n'): string {
  const parts: string[] = [];
  while (Buffer.byteLength(line, 'utf8') > 75) {
    // Find the largest character count that fits in 75 bytes
    let cut = 75;
    while (cut > 0 && Buffer.byteLength(line.slice(0, cut), 'utf8') > 75) {
      cut--;
    }
    // Don't split a surrogate pair (characters outside BMP like emoji)
    if (cut > 0 && cut < line.length) {
      const code = line.charCodeAt(cut);
      if (code >= 0xDC00 && code <= 0xDFFF) cut--;
    }
    parts.push(line.slice(0, cut));
    line = ' ' + line.slice(cut);
  }
  parts.push(line);
  return parts.join(lineEnding);
}

/**
 * Detect line ending style from iCal data.
 */
export function detectLineEnding(data: string): string {
  return data.includes('\r\n') ? '\r\n' : '\n';
}

/**
 * Replace or insert an iCal property within the first VEVENT block.
 * Operates on lines only within the first BEGIN:VEVENT/END:VEVENT pair,
 * skipping nested sub-components (VALARM etc.).
 * @param newLine Pre-folded replacement line, or null to remove the property.
 */
export function replaceICalProperty(icalData: string, key: string, newLine: string | null): string {
  if (!icalData) throw new Error('replaceICalProperty: empty input');

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  const veventStart = lines.findIndex(l => l.trim() === 'BEGIN:VEVENT');
  if (veventStart === -1) throw new Error('replaceICalProperty: BEGIN:VEVENT not found');

  let veventEnd = -1;
  let depth = 0;
  for (let i = veventStart; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith('BEGIN:')) depth++;
    if (trimmed.startsWith('END:')) {
      depth--;
      if (depth === 0) {
        veventEnd = i;
        break;
      }
    }
  }
  if (veventEnd === -1) throw new Error('replaceICalProperty: END:VEVENT not found');

  const propRegex = new RegExp(`^${key}[;:]`);
  let foundIdx = -1;
  let foundEndIdx = -1;
  let nestDepth = 0;

  for (let i = veventStart + 1; i < veventEnd; i++) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith('BEGIN:')) { nestDepth++; continue; }
    if (trimmed.startsWith('END:')) { nestDepth--; continue; }
    if (nestDepth > 0) continue;

    if (propRegex.test(lines[i])) {
      foundIdx = i;
      // Find end of this property (including continuation lines)
      foundEndIdx = i + 1;
      while (foundEndIdx < veventEnd && (lines[foundEndIdx].startsWith(' ') || lines[foundEndIdx].startsWith('\t'))) {
        foundEndIdx++;
      }
      break;
    }
  }

  if (foundIdx >= 0) {
    // Replace or remove existing property
    const newLines = newLine !== null ? newLine.split(/\r?\n/) : [];
    lines.splice(foundIdx, foundEndIdx - foundIdx, ...newLines);
  } else if (newLine !== null) {
    // Insert before END:VEVENT
    const newLines = newLine.split(/\r?\n/);
    lines.splice(veventEnd, 0, ...newLines);
  }

  return lines.join(lineEnding);
}

/**
 * Remove ALL occurrences of a property within the first VEVENT block.
 * Skips nested sub-components.
 */
export function removeAllICalProperties(icalData: string, key: string): string {
  if (!icalData) throw new Error('removeAllICalProperties: empty input');

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  const veventStart = lines.findIndex(l => l.trim() === 'BEGIN:VEVENT');
  if (veventStart === -1) throw new Error('removeAllICalProperties: BEGIN:VEVENT not found');

  let veventEnd = -1;
  let depth = 0;
  for (let i = veventStart; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith('BEGIN:')) depth++;
    if (trimmed.startsWith('END:')) {
      depth--;
      if (depth === 0) {
        veventEnd = i;
        break;
      }
    }
  }
  if (veventEnd === -1) throw new Error('removeAllICalProperties: END:VEVENT not found');

  const propRegex = new RegExp(`^${key}[;:]`);
  // Collect indices to remove (in reverse order to avoid index shifting)
  const toRemove: Array<[number, number]> = [];
  let nestDepth = 0;

  for (let i = veventStart + 1; i < veventEnd; i++) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith('BEGIN:')) { nestDepth++; continue; }
    if (trimmed.startsWith('END:')) { nestDepth--; continue; }
    if (nestDepth > 0) continue;

    if (propRegex.test(lines[i])) {
      const startIdx = i;
      let endIdx = i + 1;
      while (endIdx < veventEnd && (lines[endIdx].startsWith(' ') || lines[endIdx].startsWith('\t'))) {
        endIdx++;
      }
      toRemove.push([startIdx, endIdx - startIdx]);
      i = endIdx - 1; // skip past continuation lines
    }
  }

  // Remove in reverse to preserve indices
  for (let r = toRemove.length - 1; r >= 0; r--) {
    lines.splice(toRemove[r][0], toRemove[r][1]);
  }

  return lines.join(lineEnding);
}

/**
 * Insert a property line into the first VEVENT block, before any sub-components
 * (VALARM etc.) per RFC 5545 ABNF (eventprop before alarmc).
 * Falls back to before END:VEVENT if no sub-components exist.
 */
export function insertBeforeEndVEvent(icalData: string, newLine: string): string {
  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  const veventStart = lines.findIndex(l => l.trim() === 'BEGIN:VEVENT');
  if (veventStart === -1) throw new Error('insertBeforeEndVEvent: BEGIN:VEVENT not found');

  let veventEnd = -1;
  let firstSubComponent = -1;
  let depth = 0;
  for (let i = veventStart; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith('BEGIN:')) {
      depth++;
      // Track first nested sub-component (depth 2 = inside VEVENT)
      if (depth === 2 && firstSubComponent === -1) {
        firstSubComponent = i;
      }
    }
    if (trimmed.startsWith('END:')) {
      depth--;
      if (depth === 0) { veventEnd = i; break; }
    }
  }
  if (veventEnd === -1) throw new Error('insertBeforeEndVEvent: END:VEVENT not found');

  // Insert before first sub-component (VALARM etc.) or before END:VEVENT
  const insertIdx = firstSubComponent !== -1 ? firstSubComponent : veventEnd;
  const newLines = newLine.split(/\r?\n/);
  lines.splice(insertIdx, 0, ...newLines);
  return lines.join(lineEnding);
}

/**
 * Remove orphaned VTIMEZONE blocks whose TZID has no remaining references
 * in the file (outside VTIMEZONE blocks themselves).
 */
export function removeOrphanedVTimezones(icalData: string): string {
  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  // Find all VTIMEZONE blocks and their TZIDs
  const tzBlocks: Array<{ tzid: string; start: number; end: number }> = [];
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim() === 'BEGIN:VTIMEZONE') {
      const blockStart = i;
      let blockEnd = -1;
      for (let j = i + 1; j < lines.length; j++) {
        if (lines[j].trim() === 'END:VTIMEZONE') {
          blockEnd = j;
          break;
        }
      }
      if (blockEnd === -1) { i = lines.length; break; }
      // Use parseICalValue for proper unfolding support
      const tzBlock = lines.slice(blockStart, blockEnd + 1).join('\n');
      const tzid = parseICalValue(tzBlock, 'TZID') || '';
      tzBlocks.push({ tzid, start: blockStart, end: blockEnd });
      i = blockEnd;
    }
  }

  if (tzBlocks.length === 0) return icalData;

  // Build content outside VTIMEZONE blocks for reference scanning
  const nonTzLines: string[] = [];
  let inTz = false;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim() === 'BEGIN:VTIMEZONE') { inTz = true; continue; }
    if (lines[i].trim() === 'END:VTIMEZONE') { inTz = false; continue; }
    if (!inTz) nonTzLines.push(lines[i]);
  }
  const nonTzContent = nonTzLines.join('\n');

  // Check each VTIMEZONE for references
  const orphaned = tzBlocks.filter(tz => {
    if (!tz.tzid) return false;
    // Check for ;TZID=<tzid> references in any property
    return !nonTzContent.includes(`;TZID=${tz.tzid}`);
  });

  // Remove orphaned blocks in reverse order
  for (let i = orphaned.length - 1; i >= 0; i--) {
    lines.splice(orphaned[i].start, orphaned[i].end - orphaned[i].start + 1);
  }

  return lines.join(lineEnding);
}

/**
 * Remove exception VEVENT blocks whose RECURRENCE-ID matches one of the orphaned dates.
 * Operates on the full iCal string. Never touches the master VEVENT (no RECURRENCE-ID).
 */
export function removeExceptionVEvents(icalData: string, orphanedRecurrenceIds: Date[]): string {
  if (orphanedRecurrenceIds.length === 0) return icalData;

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  // Find all VEVENT blocks
  const veventBlocks: Array<{ start: number; end: number; recurrenceId?: string }> = [];
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim() === 'BEGIN:VEVENT') {
      const blockStart = i;
      for (let j = i + 1; j < lines.length; j++) {
        if (lines[j].trim() === 'END:VEVENT') {
          // Extract RECURRENCE-ID using parseICalValue for consistency
          // with the orphan detection code path (handles unfolding)
          const veventText = lines.slice(blockStart, j + 1).join('\n');
          const recId = parseICalValue(veventText, 'RECURRENCE-ID');
          veventBlocks.push({ start: blockStart, end: j, recurrenceId: recId });
          i = j;
          break;
        }
      }
    }
  }

  // Only remove exception VEVENTs (those with RECURRENCE-ID) that are orphaned.
  // Compare on ISO date strings (YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS) to avoid
  // timezone interpretation issues (floating vs UTC) with millisecond comparison.
  const orphanedDateStrings = orphanedRecurrenceIds.map(d => {
    // Normalize to ISO date string for comparison
    return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
  });
  const toRemove = veventBlocks.filter(block => {
    if (!block.recurrenceId) return false; // master VEVENT — never remove
    const recIdFormatted = formatICalDate(block.recurrenceId);
    if (!recIdFormatted) return false;
    // Compare as UTC ISO string — handles both date-only and datetime
    const recDate = new Date(recIdFormatted);
    const recDateStr = recDate.toISOString().replace(/\.\d{3}Z$/, 'Z');
    return orphanedDateStrings.includes(recDateStr);
  });

  // Remove in reverse order
  for (let i = toRemove.length - 1; i >= 0; i--) {
    lines.splice(toRemove[i].start, toRemove[i].end - toRemove[i].start + 1);
  }

  return lines.join(lineEnding);
}

/**
 * Parse an iCalendar DURATION value and compute end datetime.
 * RFC 5545 §3.3.6: [+/-]P[nW | nDTnHnMnS]
 * Returns ISO 8601 end datetime, or undefined for malformed input.
 */
export function parseICalDuration(duration: string, start: string): string | undefined {
  const m = duration.match(/^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return undefined;

  const [, sign, weeks, days, hours, minutes, seconds] = m;

  // At least one component must be present (reject bare "P")
  if (!weeks && !days && !hours && !minutes && !seconds) return undefined;
  // If T is present in input, at least one time component must exist (reject "P1DT")
  if (duration.includes('T') && !hours && !minutes && !seconds) return undefined;

  const ms =
    (parseInt(weeks || '0', 10) * 7 * 86400000) +
    (parseInt(days || '0', 10) * 86400000) +
    (parseInt(hours || '0', 10) * 3600000) +
    (parseInt(minutes || '0', 10) * 60000) +
    (parseInt(seconds || '0', 10) * 1000);

  const startDate = new Date(start);
  if (isNaN(startDate.getTime())) return undefined;

  const endMs = sign === '-' ? startDate.getTime() - ms : startDate.getTime() + ms;
  const endDate = new Date(endMs);

  // Return in same format as input start
  if (/^\d{4}-\d{2}-\d{2}$/.test(start)) {
    // Date-only: return date-only
    return endDate.toISOString().slice(0, 10);
  }

  // Floating time (no Z, no offset): return floating to match start format.
  // new Date() interprets floating as local, so we add the duration in ms
  // and format back as floating by doing manual arithmetic instead of toISOString().
  const isFloating = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(start);
  if (isFloating) {
    // Parse start components directly to avoid local-time interpretation
    const [datePart, timePart] = start.split('T');
    const [y, mo, d] = datePart.split('-').map(Number);
    const [h, mi, s] = timePart.split(':').map(Number);
    const utcStart = Date.UTC(y, mo - 1, d, h, mi, s);
    const utcEnd = sign === '-' ? utcStart - ms : utcStart + ms;
    const e = new Date(utcEnd);
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${e.getUTCFullYear()}-${pad(e.getUTCMonth() + 1)}-${pad(e.getUTCDate())}T${pad(e.getUTCHours())}:${pad(e.getUTCMinutes())}:${pad(e.getUTCSeconds())}`;
  }

  return endDate.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function parseCalendarObject(obj: DAVCalendarObject, options?: { includeParticipants?: boolean }): CalendarEvent {
  const vevent = extractVEvent(obj.data || '');
  if (!vevent) {
    // No VEVENT found — return minimal event
    return {
      id: obj.url || '',
      url: obj.url || '',
      title: 'Untitled',
    };
  }

  const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
  const description = parseICalValue(vevent, 'DESCRIPTION');
  const rawStart = parseICalValue(vevent, 'DTSTART');
  let rawEnd = parseICalValue(vevent, 'DTEND');
  const location = parseICalValue(vevent, 'LOCATION');
  const uid = parseICalValue(vevent, 'UID') || obj.url || '';

  // DURATION parsing: compute end from start + duration if DTEND absent
  if (!rawEnd && rawStart) {
    const rawDuration = parseICalValue(vevent, 'DURATION');
    if (rawDuration) {
      const startIso = formatICalDate(rawStart);
      if (startIso) {
        const computedEnd = parseICalDuration(rawDuration, startIso);
        if (computedEnd) {
          // computedEnd is already ISO format, return it directly
          const event: CalendarEvent = {
            id: uid,
            url: obj.url || '',
            title: unescapeICalText(title),
            description: description ? unescapeICalText(description) : undefined,
            start: formatICalDate(rawStart),
            end: computedEnd,
            location: location ? unescapeICalText(location) : undefined,
          };
          if (options?.includeParticipants) {
            addParticipantsToEvent(event, vevent);
          }
          return event;
        }
      }
    }
  }

  const event: CalendarEvent = {
    id: uid,
    url: obj.url || '',
    title: unescapeICalText(title),
    description: description ? unescapeICalText(description) : undefined,
    start: formatICalDate(rawStart),
    end: formatICalDate(rawEnd),
    location: location ? unescapeICalText(location) : undefined,
  };

  if (options?.includeParticipants) {
    addParticipantsToEvent(event, vevent);
  }

  return event;
}

function addParticipantsToEvent(event: CalendarEvent, vevent: string): void {
  const attendeeLines = parseAllICalProperties(vevent, 'ATTENDEE');
  if (attendeeLines.length > 0) {
    event.participants = attendeeLines.map(parseAttendee);
  }
  const organizerLines = parseAllICalProperties(vevent, 'ORGANIZER');
  if (organizerLines.length > 0) {
    event.organizer = parseAttendee(organizerLines[0]);
  }
}

/**
 * Unescape an iCalendar text value (RFC 5545 §3.3.11).
 * Reverses escaping of newlines, semicolons, commas, and backslashes.
 */
export function unescapeICalText(value: string): string {
  return value
    .replace(/\\n/gi, '\n')
    .replace(/\\;/g, ';')
    .replace(/\\,/g, ',')
    .replace(/\\\\/g, '\\');
}

/**
 * Escape a text value for use in an iCalendar property (RFC 5545 §3.3.11).
 * Backslashes, newlines, commas, and semicolons must be escaped.
 */
export function escapeICalText(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\r?\n/g, '\\n');
}

/**
 * Validate an email address for use in ATTENDEE lines.
 * Prevents iCal property injection via malicious email values.
 */
export function validateAttendeeEmail(email: string): void {
  if (!email || typeof email !== 'string') {
    throw new Error('Participant email is required');
  }
  if (!/^[^@]+@[^@]+$/.test(email)) {
    throw new Error(`Invalid participant email: ${email}`);
  }
  if (/[\r\n:;"\\]|\s/.test(email)) {
    throw new Error(`Invalid participant email (contains illegal characters): ${email}`);
  }
}

/**
 * Quote a CN parameter value per RFC 5545 §3.2.
 * Uses DQUOTE quoting (NOT escapeICalText backslash escaping).
 * Literal DQUOTEs in the value are replaced with single quotes since
 * RFC 5545 has no escape mechanism for DQUOTE inside quoted parameter values.
 * RFC 6868 caret encoding (^') exists but is poorly adopted;
 * single-quote replacement matches Python icalendar/Outlook behavior.
 */
export function quoteParamValue(value: string): string {
  // Strip newlines to prevent iCal property injection via CN values
  let cleaned = value.replace(/[\r\n]+/g, ' ');
  // Replace literal double quotes with single quotes
  cleaned = cleaned.replace(/"/g, "'");
  // Quote if contains comma, semicolon, colon, or if the original had double quotes
  if (/[,;:]/.test(cleaned) || value.includes('"')) {
    return `"${cleaned}"`;
  }
  return cleaned;
}

/**
 * Format a start/end input value into the correct iCal property line.
 * Handles three cases:
 * 1. Date-only (2026-04-01) → DTXXX;VALUE=DATE:20260401
 * 2. Floating time (2026-03-20T09:30:00) → preserve original TZID
 * 3. UTC/offset (2026-03-20T09:30:00Z) → DTXXX:20260320T093000Z
 */
function formatDateTimeProperty(
  propName: string,
  value: string,
  originalVevent: string | null,
  lineEnding: string
): { line: string; isDateOnly: boolean } {
  // Date-only
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const d = new Date(value);
    if (isNaN(d.getTime()) || !d.toISOString().startsWith(value)) {
      throw new Error(`Invalid date: ${value}`);
    }
    const icalDate = value.replace(/-/g, '');
    return { line: foldICalLine(`${propName};VALUE=DATE:${icalDate}`, lineEnding), isDateOnly: true };
  }

  // Floating time (no offset, no Z)
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(value)) {
    const icalTime = value.replace(/[-:]/g, '');
    // Try to preserve original TZID
    if (originalVevent) {
      const rawLines = parseAllICalProperties(originalVevent, propName);
      if (rawLines.length > 0) {
        const tzMatch = rawLines[0].match(/;TZID=([^;:]+)/);
        if (tzMatch) {
          return { line: foldICalLine(`${propName};TZID=${tzMatch[1]}:${icalTime}`, lineEnding), isDateOnly: false };
        }
      }
      // If propName is DTEND and no TZID found (DURATION-based), fall back to DTSTART's TZID
      if (propName === 'DTEND') {
        const startLines = parseAllICalProperties(originalVevent, 'DTSTART');
        if (startLines.length > 0) {
          const tzMatch = startLines[0].match(/;TZID=([^;:]+)/);
          if (tzMatch) {
            return { line: foldICalLine(`${propName};TZID=${tzMatch[1]}:${icalTime}`, lineEnding), isDateOnly: false };
          }
        }
      }
    }
    // No TZID to preserve — emit as floating
    return { line: foldICalLine(`${propName}:${icalTime}`, lineEnding), isDateOnly: false };
  }

  // UTC or offset — convert to UTC
  return { line: foldICalLine(`${propName}:${toICalUTC(value)}`, lineEnding), isDateOnly: false };
}

/**
 * Check if a raw iCal property line represents a date-only value (VALUE=DATE).
 */
function isDateOnlyProperty(rawLine: string): boolean {
  return /;VALUE=DATE[;:]/.test(rawLine) || /;VALUE=DATE$/.test(rawLine);
}

/**
 * Validate that DTSTART and DTEND have consistent value types and DTEND > DTSTART.
 */
function validateDateConsistency(
  startIsDateOnly: boolean | null,
  endIsDateOnly: boolean | null,
  startValue?: string,
  endValue?: string
): void {
  if (startIsDateOnly !== null && endIsDateOnly !== null) {
    if (startIsDateOnly !== endIsDateOnly) {
      throw new Error('DTSTART and DTEND must have the same value type (both date-only or both datetime) per RFC 5545 §3.6.1');
    }
  }

  // Validate DTEND > DTSTART when both are date-only
  if (startValue && endValue && startIsDateOnly && endIsDateOnly) {
    if (startValue >= endValue) {
      throw new Error(
        `DTEND is exclusive per RFC 5545 — for a one-day event on ${startValue}, ` +
        `pass end: '${nextDay(startValue)}'`
      );
    }
  }
}

function nextDay(dateStr: string): string {
  const d = new Date(dateStr);
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

export class CalDAVCalendarClient {
  private config: CalDAVConfig;
  private client: DAVClient | null = null;
  private calendars: DAVCalendar[] | null = null;

  constructor(config: CalDAVConfig) {
    this.config = config;
  }

  private async getClient(): Promise<DAVClient> {
    if (this.client) return this.client;

    this.client = new DAVClient({
      serverUrl: this.config.serverUrl || 'https://caldav.fastmail.com',
      credentials: {
        username: this.config.username,
        password: this.config.password,
      },
      authMethod: 'Basic',
      defaultAccountType: 'caldav',
    });

    await this.client.login();
    return this.client;
  }

  async getCalendars(): Promise<CalendarInfo[]> {
    const client = await this.getClient();
    const calendars = await client.fetchCalendars();
    this.calendars = calendars;

    return calendars
      .filter(c => c.displayName !== 'DEFAULT_TASK_CALENDAR_NAME')
      .map(c => ({
        id: c.url || '',
        displayName: String(c.displayName || 'Unnamed'),
        url: c.url || '',
        description: c.description || undefined,
        color: (c as any).calendarColor || undefined,
      }));
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50, startDate?: string, endDate?: string): Promise<CalendarEvent[]> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    let targetCalendars = this.calendars.filter(
      c => c.displayName !== 'DEFAULT_TASK_CALENDAR_NAME'
    );
    if (calendarId) {
      targetCalendars = targetCalendars.filter(
        c => c.url === calendarId || c.displayName === calendarId
      );
    }

    const fetchOptions: any = {};
    if (startDate || endDate) {
      fetchOptions.timeRange = {
        start: startDate || '1970-01-01T00:00:00Z',
        end: endDate || '2099-12-31T23:59:59Z',
      };
    }

    const allEvents: CalendarEvent[] = [];
    for (const cal of targetCalendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal, ...fetchOptions });
      for (const obj of objects) {
        allEvents.push(parseCalendarObject(obj));
      }
      if (allEvents.length >= limit) break;
    }

    // Sort by start date ascending
    allEvents.sort((a, b) => (a.start || '').localeCompare(b.start || ''));

    return allEvents.slice(0, limit);
  }

  private async findCalendarObjectByUID(eventId: string): Promise<DAVCalendarObject | null> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    for (const cal of this.calendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal });
      for (const obj of objects) {
        const vevent = extractVEvent(obj.data || '');
        if (!vevent) continue;
        const uid = parseICalValue(vevent, 'UID');
        if (uid === eventId || obj.url === eventId) {
          return obj;
        }
      }
    }

    return null;
  }

  async getCalendarEventById(eventId: string): Promise<CalendarEvent | null> {
    const obj = await this.findCalendarObjectByUID(eventId);
    return obj ? parseCalendarObject(obj, { includeParticipants: true }) : null;
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string;
    end: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
  }): Promise<string> {
    const client = await this.getClient();

    if (!this.calendars) {
      this.calendars = await client.fetchCalendars();
    }

    const targetCal = this.calendars.find(
      c => c.url === event.calendarId || c.displayName === event.calendarId
    );
    if (!targetCal) {
      throw new Error(`Calendar not found: ${event.calendarId}`);
    }

    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@fastmail-mcp`;
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

    // Format start/end with all-day event support
    const startResult = formatDateTimeProperty('DTSTART', event.start, null, '\r\n');
    const endResult = formatDateTimeProperty('DTEND', event.end, null, '\r\n');

    // Value type consistency check
    validateDateConsistency(
      startResult.isDateOnly,
      endResult.isDateOnly,
      event.start,
      event.end
    );

    const icalLines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `LAST-MODIFIED:${now}`,
      startResult.line,
      endResult.line,
      foldICalLine(`SUMMARY:${escapeICalText(event.title)}`),
    ];

    if (event.description) {
      icalLines.push(foldICalLine(`DESCRIPTION:${escapeICalText(event.description)}`));
    }
    if (event.location) {
      icalLines.push(foldICalLine(`LOCATION:${escapeICalText(event.location)}`));
    }

    // Participant support
    if (event.participants && event.participants.length > 0) {
      // Validate all emails first
      for (const p of event.participants) {
        validateAttendeeEmail(p.email);
      }

      // ORGANIZER required when ATTENDEEs present
      const caldavUsername = this.config.username;
      if (!caldavUsername.includes('@')) {
        throw new Error('Cannot add participants: CalDAV username is not an email address, required for ORGANIZER');
      }
      const displayName = process.env.FASTMAIL_CALDAV_DISPLAY_NAME || caldavUsername;
      const cnPart = `;CN=${quoteParamValue(displayName)}`;
      icalLines.push(foldICalLine(`ORGANIZER${cnPart}:mailto:${caldavUsername}`));

      // ATTENDEE lines — do NOT emit RSVP=TRUE by default (RFC 5545 §3.2.17 defaults to FALSE)
      for (const p of event.participants) {
        const cnParam = p.name ? `;CN=${quoteParamValue(p.name)}` : '';
        icalLines.push(foldICalLine(`ATTENDEE${cnParam}:mailto:${p.email}`));
      }
    }

    icalLines.push('END:VEVENT');
    icalLines.push('END:VCALENDAR');

    // Trailing CRLF per RFC 5545 §3.1
    const ical = icalLines.join('\r\n') + '\r\n';

    await client.createCalendarObject({
      calendar: targetCal,
      filename: `${uid}.ics`,
      iCalString: ical,
    });

    return uid;
  }

  async updateCalendarEvent(eventId: string, fields: {
    title?: string;
    description?: string;
    start?: string;
    end?: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    confirmRecurring?: boolean;
  }): Promise<string> {
    const client = await this.getClient();
    const obj = await this.findCalendarObjectByUID(eventId);
    if (!obj) {
      throw new Error(`Calendar event not found: ${eventId}`);
    }

    if (!obj.data || !obj.data.includes('BEGIN:VEVENT')) {
      throw new Error('Cannot update event: no iCal data found');
    }

    // Validate date inputs early before any processing
    const datePattern = /^\d{4}-\d{2}-\d{2}$/;
    const dateTimePattern = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;
    if (fields.start !== undefined && !datePattern.test(fields.start) && !dateTimePattern.test(fields.start)) {
      throw new Error(`Invalid start date format: ${fields.start}. Expected ISO 8601 (e.g. 2026-04-07T14:00:00Z or 2026-04-07)`);
    }
    if (fields.end !== undefined && !datePattern.test(fields.end) && !dateTimePattern.test(fields.end)) {
      throw new Error(`Invalid end date format: ${fields.end}. Expected ISO 8601 (e.g. 2026-04-07T14:00:00Z or 2026-04-07)`);
    }

    const lineEnding = detectLineEnding(obj.data);
    const fold = (line: string) => foldICalLine(line, lineEnding);

    // Capture original VEVENT before any patching for reads
    const originalVevent = extractVEvent(obj.data);
    if (!originalVevent) {
      throw new Error('Cannot update event: no VEVENT block found');
    }

    const existingUid = parseICalValue(originalVevent, 'UID') || eventId;
    let data = obj.data;

    // --- Recurring event guard ---
    const hasRRule = /^RRULE[;:]/m.test(originalVevent);
    const isTimeChange = fields.start !== undefined || fields.end !== undefined;

    if (hasRRule && isTimeChange) {
      // Find exception VEVENTs (same UID, have RECURRENCE-ID)
      const allVevents = data.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];
      const exceptions = allVevents.filter((v: string) => /^RECURRENCE-ID[;:]/m.test(v));

      if (exceptions.length > 0) {
        // Check which exceptions would be orphaned
        const rruleLine = parseICalValue(originalVevent, 'RRULE');
        const existingStart = parseICalValue(originalVevent, 'DTSTART');
        const newStartRaw = fields.start || (existingStart ? formatICalDate(existingStart) : undefined);

        if (rruleLine && newStartRaw) {
          try {
            // Convert offset-bearing timestamps to UTC before passing to rrule
            let dtStartForRrule: string;
            if (/[+-]\d{2}:\d{2}$/.test(newStartRaw)) {
              dtStartForRrule = toICalUTC(newStartRaw).replace(/Z$/, '');
            } else {
              dtStartForRrule = newStartRaw.replace(/[-:]/g, '').replace(/Z$/, '');
            }
            const rruleString = `RRULE:${rruleLine}\nDTSTART:${dtStartForRrule}`;
            const rule = rrulestr(rruleString, { forceset: false });

            const orphanedDates: Date[] = [];
            const validDates: Date[] = [];
            for (const excVevent of exceptions) {
              const recIdRaw = parseICalValue(excVevent, 'RECURRENCE-ID');
              if (!recIdRaw) continue;
              const recIdFormatted = formatICalDate(recIdRaw);
              if (!recIdFormatted) continue;
              const recDate = new Date(recIdFormatted);
              // Check if this recurrence-id still matches an occurrence
              const matches = rule.between(
                new Date(recDate.getTime() - 1000),
                new Date(recDate.getTime() + 1000),
                true
              );
              if (matches.length === 0) {
                orphanedDates.push(recDate);
              } else {
                validDates.push(recDate);
              }
            }

            if (orphanedDates.length > 0 && !fields.confirmRecurring) {
              // List the orphaned exceptions
              const dateList = orphanedDates.map(d => d.toISOString().slice(0, 10)).join(', ');
              throw new Error(
                `This recurring event has ${exceptions.length} exception(s). ` +
                `Changing start/end will orphan ${orphanedDates.length} of them (${dateList}). ` +
                `These will be removed to prevent server errors. Pass confirmRecurring: true to proceed.`
              );
            }

            // If confirmRecurring, remove orphaned exceptions after patching
            if (orphanedDates.length > 0 && fields.confirmRecurring) {
              data = removeExceptionVEvents(data, orphanedDates);
            }
          } catch (e) {
            if (e instanceof Error && e.message.includes('confirmRecurring')) throw e;
            // If RRULE parsing fails, proceed without pruning (best effort)
          }
        }
      }
    }

    // --- Patch fields ---
    let startIsDateOnly: boolean | null = null;
    let endIsDateOnly: boolean | null = null;
    let timeChanged = false;

    if (fields.title !== undefined) {
      data = replaceICalProperty(data, 'SUMMARY', fold(`SUMMARY:${escapeICalText(fields.title)}`));
    }

    if (fields.description !== undefined) {
      data = replaceICalProperty(data, 'DESCRIPTION', fold(`DESCRIPTION:${escapeICalText(fields.description)}`));
    }

    if (fields.start !== undefined) {
      const result = formatDateTimeProperty('DTSTART', fields.start, originalVevent, lineEnding);
      data = replaceICalProperty(data, 'DTSTART', result.line);
      startIsDateOnly = result.isDateOnly;
      timeChanged = true;
    }

    if (fields.end !== undefined) {
      const result = formatDateTimeProperty('DTEND', fields.end, originalVevent, lineEnding);
      data = replaceICalProperty(data, 'DTEND', result.line);
      endIsDateOnly = result.isDateOnly;
      // Remove DURATION — DTEND and DURATION are mutually exclusive (RFC 5545 §3.6.1)
      data = removeAllICalProperties(data, 'DURATION');
      timeChanged = true;
    }

    // Value type consistency: check against existing properties when only one is provided
    if (fields.start !== undefined && fields.end === undefined) {
      const existingEndLines = parseAllICalProperties(originalVevent, 'DTEND');
      if (existingEndLines.length > 0) {
        endIsDateOnly = isDateOnlyProperty(existingEndLines[0]);
        validateDateConsistency(startIsDateOnly, endIsDateOnly);
      }
    } else if (fields.end !== undefined && fields.start === undefined) {
      const existingStartLines = parseAllICalProperties(originalVevent, 'DTSTART');
      if (existingStartLines.length > 0) {
        startIsDateOnly = isDateOnlyProperty(existingStartLines[0]);
        validateDateConsistency(startIsDateOnly, endIsDateOnly);
      }
    } else if (fields.start !== undefined && fields.end !== undefined) {
      validateDateConsistency(startIsDateOnly, endIsDateOnly, fields.start, fields.end);
    }

    if (fields.location !== undefined) {
      data = replaceICalProperty(data, 'LOCATION', fold(`LOCATION:${escapeICalText(fields.location)}`));
    }

    if (fields.participants !== undefined) {
      // Validate emails
      for (const p of fields.participants) {
        validateAttendeeEmail(p.email);
      }
      // Remove all existing ATTENDEE lines (ORGANIZER is preserved)
      data = removeAllICalProperties(data, 'ATTENDEE');
      // Build and insert all ATTENDEE lines in one pass
      if (fields.participants.length > 0) {
        const attendeeLines = fields.participants.map(p => {
          const cnParam = p.name ? `;CN=${quoteParamValue(p.name)}` : '';
          return fold(`ATTENDEE${cnParam}:mailto:${p.email}`);
        }).join(lineEnding);
        data = insertBeforeEndVEvent(data, attendeeLines);
      }
      // Add ORGANIZER if absent and participants are being added (RFC 5545 §3.8.4.1)
      if (fields.participants.length > 0 && !/^ORGANIZER[;:]/m.test(extractVEvent(data) || '')) {
        const caldavUsername = this.config.username;
        if (!caldavUsername.includes('@')) {
          throw new Error('Cannot add participants: CalDAV username is not an email address, required for ORGANIZER');
        }
        const displayName = process.env.FASTMAIL_CALDAV_DISPLAY_NAME || caldavUsername;
        const cnPart = displayName ? `;CN=${quoteParamValue(displayName)}` : '';
        data = replaceICalProperty(data, 'ORGANIZER', fold(`ORGANIZER${cnPart}:mailto:${caldavUsername}`));
      }
    }

    // --- SEQUENCE increment ---
    const hasAttendees = /^ATTENDEE[;:]/m.test(originalVevent);
    const schedulingSignificant = fields.start !== undefined || fields.end !== undefined ||
      fields.participants !== undefined || fields.location !== undefined;

    if (hasAttendees && schedulingSignificant) {
      const existingSeq = parseInt(parseICalValue(originalVevent, 'SEQUENCE') || '0', 10) || 0;
      data = replaceICalProperty(data, 'SEQUENCE', `SEQUENCE:${existingSeq + 1}`);
    }

    // --- Update DTSTAMP and LAST-MODIFIED ---
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    data = replaceICalProperty(data, 'DTSTAMP', `DTSTAMP:${now}`);
    data = replaceICalProperty(data, 'LAST-MODIFIED', `LAST-MODIFIED:${now}`);

    // --- Orphaned VTIMEZONE cleanup (LAST — after all modifications) ---
    if (timeChanged) {
      data = removeOrphanedVTimezones(data);
    }

    obj.data = data;
    await client.updateCalendarObject({ calendarObject: obj });

    return existingUid;
  }

  async deleteCalendarEvent(eventId: string): Promise<void> {
    const client = await this.getClient();
    const obj = await this.findCalendarObjectByUID(eventId);
    if (!obj) {
      throw new Error(`Calendar event not found: ${eventId}`);
    }

    await client.deleteCalendarObject({ calendarObject: obj });
  }
}
