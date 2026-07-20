/**
 * Represents a CalDAV calendar object (VTODO)
 */
export interface CalendarObject {
  data: string;
  etag?: string;
  url: string;
}

import { CommonTask } from '../sync/types';
import { extractInlineTags, stripInlineTags } from '../utils/inlineTags';

/** Fields returned by vtodoToTask — everything except uid, which is extracted separately */
type VTODOTaskFields = Omit<CommonTask, 'uid'>;

/**
 * Maps between CommonTask fields and CalDAV VTODO iCalendar format.
 */
export class VTODOMapper {
  /**
   * Convert a CommonTask to VTODO iCalendar string.
   *
   * When `existingData` is supplied (update/complete paths), only the
   * properties this plugin manages are replaced; all other properties
   * (RELATED-TO, VALARM blocks, X-* extension lines, PERCENT-COMPLETE
   * when not completing, etc.) are carried through unchanged. This
   * preserves data written by other CalDAV clients such as jtx Board.
   *
   * When `existingData` is absent (create path), a fresh VCALENDAR is built.
   */
  taskToVTODO(task: Omit<CommonTask, 'uid'>, uid: string, existingData?: string): string {
    return existingData
      ? this.mergeIntoVTODO(task, existingData)
      : this.buildFreshVTODO(task, uid);
  }

  private buildFreshVTODO(task: Omit<CommonTask, 'uid'>, uid: string): string {
    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Obsidian//Tasks CalDAV Sync//EN',
      'BEGIN:VTODO',
      `UID:${uid}`,
      ...this.buildManagedLines(task),
      'END:VTODO',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  /**
   * The VTODO properties owned by this plugin. Written on every create/update.
   * Any property NOT in this list is foreign and must pass through untouched
   * in merge mode.
   *
   * When `mergeContext` is provided we are on the update/complete path:
   *  - STATUS is omitted for TODO tasks so the server's IN-PROCESS (or
   *    NEEDS-ACTION) passes through untouched from the merge loop.
   *  - DUE/DTSTART carry the existing timezone and time-of-day when the
   *    server stored a datetime value, so jtx Board's time precision survives
   *    an Obsidian title or date change.
   */
  private buildManagedLines(
    task: Omit<CommonTask, 'uid'>,
    mergeContext?: { existingDue?: string | null; existingDtstart?: string | null },
  ): string[] {
    const lines: string[] = [];
    lines.push(`DTSTAMP:${this.formatDateTimeUTC(new Date())}`);
    lines.push(`LAST-MODIFIED:${this.formatDateTimeUTC(new Date())}`);
    lines.push(`SUMMARY:${this.escapeText(task.title)}`);

    // Description (body text), with optional obsidian link prepended
    const description = this.buildDescription(task.body, task.obsidianUrl);
    if (description) {
      lines.push(`DESCRIPTION:${this.escapeText(description)}`);
    }

    // Obsidian vault link. When set, this plugin owns the URL property —
    // any value previously set by another CalDAV client will be overwritten.
    if (task.obsidianUrl) {
      lines.push(`URL:${task.obsidianUrl}`);
    }

    // Status mapping. In merge mode, STATUS for open (TODO) tasks is preserved
    // from the existing VTODO by the merge loop — emitting it here would
    // overwrite IN-PROCESS.
    if (!mergeContext || task.status !== 'TODO') {
      lines.push(`STATUS:${this.mapStatusToVTODO(task.status)}`);
    }

    // Due date
    if (task.dueDate) {
      lines.push(
        mergeContext?.existingDue
          ? this.mergeDatetime('DUE', task.dueDate, mergeContext.existingDue)
          : `DUE;VALUE=DATE:${this.formatDate(task.dueDate)}`,
      );
    }

    // DTSTART carries the scheduled date (⏳) — the field CalDAV clients plan
    // by. The start date (🛫) has no CalDAV counterpart and never syncs.
    if (task.scheduledDate) {
      lines.push(
        mergeContext?.existingDtstart
          ? this.mergeDatetime('DTSTART', task.scheduledDate, mergeContext.existingDtstart)
          : `DTSTART;VALUE=DATE:${this.formatDate(task.scheduledDate)}`,
      );
    }

    // Completed date
    if (task.completedDate) {
      lines.push(`COMPLETED:${this.formatDateTimeUTC(this.toCompletedInstant(task.completedDate))}`);
      lines.push('PERCENT-COMPLETE:100');
    }

    // Priority mapping (Obsidian: lowest/low/none/medium/high/highest -> VTODO: 0-9)
    lines.push(`PRIORITY:${this.mapPriorityToVTODO(task.priority)}`);

    // Recurrence rule
    if (task.recurrenceRule) {
      lines.push(`RRULE:${task.recurrenceRule}`);
    }

    // Tags as categories
    if (task.tags.length > 0) {
      lines.push(`CATEGORIES:${task.tags.map(t => this.escapeText(t)).join(',')}`);
    }

    return lines;
  }

  /**
   * Patch an existing iCalendar string in place. Strips only the properties
   * this plugin owns, then injects fresh values before END:VTODO. Sub-components
   * (VALARM, etc.) and unrecognised properties pass through verbatim.
   *
   * PERCENT-COMPLETE is treated specially: it is only stripped (and re-emitted
   * as 100) when the task is being completed. Otherwise the server's value
   * (e.g. jtx Board's in-progress percentage) is preserved.
   */
  private mergeIntoVTODO(task: Omit<CommonTask, 'uid'>, existingData: string): string {
    const ALWAYS_MANAGED = new Set([
      'SUMMARY', 'DESCRIPTION', 'DUE', 'DTSTART',
      'PRIORITY', 'RRULE', 'CATEGORIES', 'DTSTAMP', 'LAST-MODIFIED', 'COMPLETED', 'URL',
      // STATUS and PERCENT-COMPLETE are handled inline below
    ]);
    const isCompleting = !!task.completedDate;

    const unfolded = this.unfold(existingData);
    const mergeContext = {
      existingDue: this.extractRawDatetimeLine(unfolded, 'DUE'),
      existingDtstart: this.extractRawDatetimeLine(unfolded, 'DTSTART'),
    };

    const lines = unfolded.split(/\r?\n/).filter(l => l.length > 0);
    const out: string[] = [];
    let inVTODO = false;
    let depth = 0;

    for (const line of lines) {
      if (!inVTODO) {
        out.push(line);
        if (line === 'BEGIN:VTODO') inVTODO = true;
        continue;
      }

      if (line === 'END:VTODO') {
        out.push(...this.buildManagedLines(task, mergeContext));
        out.push(line);
        inVTODO = false;
        continue;
      }

      if (line.startsWith('BEGIN:')) { depth++; out.push(line); continue; }
      if (line.startsWith('END:')) { depth--; out.push(line); continue; }
      if (depth > 0) { out.push(line); continue; }

      const propName = line.split(/[;:]/)[0].toUpperCase();
      if (ALWAYS_MANAGED.has(propName)) continue;
      if (propName === 'PERCENT-COMPLETE' && isCompleting) continue;

      // STATUS: preserve the server's value (IN-PROCESS or NEEDS-ACTION) when
      // Obsidian has no opinion — i.e. the task is open (TODO). Terminal states
      // (DONE, CANCELLED) are stripped here and re-emitted by buildManagedLines.
      if (propName === 'STATUS') {
        if (task.status === 'TODO') out.push(line);
        continue;
      }

      out.push(line);
    }

    return out.join('\r\n');
  }

  /**
   * Convert VTODO iCalendar object to CommonTask fields (minus uid).
   * @param vtodo The CalDAV calendar object containing VTODO
   */
  vtodoToTask(vtodo: CalendarObject): VTODOTaskFields {
    const unfolded = this.unfold(vtodo.data);
    // Extract only the VTODO section to avoid matching properties from VTIMEZONE or other components,
    // then strip sub-components (VALARM etc.) so their properties don't bleed into task fields
    // — a sub-component's DESCRIPTION is not the task's description.
    const vtodoMatch = unfolded.match(/BEGIN:VTODO[\s\S]*?END:VTODO/);
    const data = (vtodoMatch ? vtodoMatch[0] : unfolded)
      .replace(/BEGIN:(?!VTODO\b)\w+[\s\S]*?END:\w+(\r?\n|$)/g, '');

    // Inline #tags in SUMMARY (written by older plugin versions or other
    // clients) move into tags[], so corrupted tasks heal instead of gaining
    // a duplicate tag on every sync — issue #114.
    const summary = this.extractRawProperty(data, 'SUMMARY') || '';

    return {
      title: stripInlineTags(summary) || 'Untitled Task',
      status: this.mapStatusFromVTODO(this.extractProperty(data, 'STATUS') || 'NEEDS-ACTION') as CommonTask['status'],
      dueDate: this.extractDateProperty(data, 'DUE'),
      scheduledDate: this.extractDateProperty(data, 'DTSTART'),
      startDate: null,
      completedDate: this.extractDateTimeProperty(data, 'COMPLETED'),
      priority: this.mapPriorityFromVTODO(this.extractProperty(data, 'PRIORITY') || '0') as CommonTask['priority'],
      recurrenceRule: this.extractProperty(data, 'RRULE') || '',
      tags: this.dedupeTags([...this.extractCategories(data), ...extractInlineTags(summary)]),
      body: this.stripObsidianLinks(this.extractRawProperty(data, 'DESCRIPTION') || ''),
    };
  }

  /**
   * Extract UID from VTODO data
   */
  extractUID(data: string): string {
    const unfolded = this.unfold(data);
    const match = unfolded.match(/^UID:(.+)$/m);
    return match ? match[1].trim() : '';
  }

  /**
   * Extract LAST-MODIFIED timestamp from VTODO data
   * Returns ISO 8601 string or null if not present
   */
  extractLastModified(data: string): string | null {
    const match = this.unfold(data).match(/^LAST-MODIFIED:(.+)$/m);
    if (!match) return null;

    const timestamp = match[1].trim();
    // Parse iCalendar datetime format (YYYYMMDDTHHMMSSZ)
    const year = timestamp.substring(0, 4);
    const month = timestamp.substring(4, 6);
    const day = timestamp.substring(6, 8);
    const hour = timestamp.substring(9, 11);
    const minute = timestamp.substring(11, 13);
    const second = timestamp.substring(13, 15);

    return `${year}-${month}-${day}T${hour}:${minute}:${second}Z`;
  }

  /**
   * Map Obsidian task status to VTODO status
   */
  private mapStatusToVTODO(status: string): string {
    switch (status) {
      case 'TODO':
        return 'NEEDS-ACTION';
      case 'IN_PROGRESS':
        return 'IN-PROCESS';
      case 'DONE':
        return 'COMPLETED';
      case 'CANCELLED':
        return 'CANCELLED';
      default:
        return 'NEEDS-ACTION';
    }
  }

  /**
   * Map VTODO status to Obsidian task status
   */
  private mapStatusFromVTODO(status: string): string {
    switch (status) {
      case 'NEEDS-ACTION':
        return 'TODO';
      case 'IN-PROCESS':
        // Obsidian-tasks has no in-progress checkbox; treat as open so the diff
        // sees TODO on both sides and doesn't generate a spurious update that
        // would overwrite IN-PROCESS with NEEDS-ACTION on every sync.
        return 'TODO';
      case 'COMPLETED':
        return 'DONE';
      case 'CANCELLED':
        return 'CANCELLED';
      default:
        return 'TODO';
    }
  }

  /**
   * Map Obsidian priority to VTODO priority (0-9, where 1 is highest)
   */
  private mapPriorityToVTODO(priority: string): number {
    switch (priority) {
      case 'highest':
        return 1;
      case 'high':
        return 3;
      case 'medium':
        return 5;
      case 'low':
        return 7;
      case 'lowest':
        return 9;
      default:
        return 0; // undefined
    }
  }

  /**
   * Map VTODO priority to Obsidian priority
   */
  private mapPriorityFromVTODO(priorityStr: string): string {
    const priority = parseInt(priorityStr);

    if (priority === 0) return 'none';
    if (priority <= 2) return 'highest';
    if (priority <= 4) return 'high';
    if (priority <= 6) return 'medium';
    if (priority <= 8) return 'low';
    return 'lowest';
  }

  /**
   * Extract the raw params+value segment of a date/datetime property line so
   * the merge path can preserve time precision and timezone.
   * Returns e.g. `;TZID=Europe/Berlin:20240115T140000` or `;VALUE=DATE:20240115`.
   */
  private extractRawDatetimeLine(unfolded: string, property: string): string | null {
    const regex = new RegExp(`^${property}(;[^:]+)?:(.+)$`, 'm');
    const match = unfolded.match(regex);
    if (!match) return null;
    return `${match[1] ?? ''}:${match[2].trim()}`;
  }

  /**
   * Emit a DUE or DTSTART line that preserves the server's timezone and
   * time-of-day when it stored a datetime value, updating only the date digits.
   * Falls back to VALUE=DATE when the existing value was date-only.
   */
  private mergeDatetime(prop: string, newDate: string, existingParamsAndValue: string): string {
    const colonIdx = existingParamsAndValue.indexOf(':');
    const params = colonIdx >= 0 ? existingParamsAndValue.substring(0, colonIdx) : '';
    const value = colonIdx >= 0 ? existingParamsAndValue.substring(colonIdx + 1) : existingParamsAndValue;
    if (value.length > 8 && value.includes('T')) {
      // Datetime: update the 8-digit date prefix, carry the time suffix (T…Z or T…)
      return `${prop}${params}:${this.formatDate(newDate)}${value.substring(8)}`;
    }
    return `${prop};VALUE=DATE:${this.formatDate(newDate)}`;
  }

  /**
   * RFC 5545 Section 3.1: Unfold long content lines.
   * Lines folded with CRLF+space/tab continuation are joined.
   */
  private unfold(data: string): string {
    return data.replace(/\r?\n[ \t]/g, '');
  }

  /**
   * Extract a simple property value from iCalendar data
   */
  private extractProperty(data: string, property: string): string | null {
    const regex = new RegExp(`^${property}[;:](.+)$`, 'm');
    const match = data.match(regex);

    if (match) {
      // Extract value after last colon (handles parameters like DUE;VALUE=DATE:20250105)
      const fullValue = match[1];
      const colonIndex = fullValue.lastIndexOf(':');
      const value = colonIndex >= 0 ? fullValue.substring(colonIndex + 1).trim() : fullValue.trim();

      // Unescape iCalendar special characters
      return this.unescapeText(value);
    }

    return null;
  }

  /**
   * Extract a property value without splitting on colons within the value.
   * Used for DESCRIPTION and other text properties where colons are valid content.
   */
  private extractRawProperty(data: string, property: string): string | null {
    const regex = new RegExp(`^${property}:(.+)$`, 'm');
    const match = data.match(regex);
    if (!match) return null;
    return this.unescapeText(match[1].trim());
  }

  /**
   * Extract date property (VALUE=DATE format)
   */
  private extractDateProperty(data: string, property: string): string | null {
    const value = this.extractProperty(data, property);
    if (!value) return null;

    // Parse YYYYMMDD format (VALUE=DATE)
    if (value.length === 8 && /^\d{8}$/.test(value)) {
      const year = value.substring(0, 4);
      const month = value.substring(4, 6);
      const day = value.substring(6, 8);
      return `${year}-${month}-${day}`;
    }

    // Parse YYYYMMDDTHHMMSS format (TZID parameter or datetime without Z)
    if (value.length >= 15 && value.includes('T')) {
      const year = value.substring(0, 4);
      const month = value.substring(4, 6);
      const day = value.substring(6, 8);
      return `${year}-${month}-${day}`;
    }

    return null;
  }

  /**
   * Extract datetime property
   */
  private extractDateTimeProperty(data: string, property: string): string | null {
    const value = this.extractProperty(data, property);
    if (!value) return null;

    // Parse YYYYMMDDTHHMMSSZ format
    if (value.length >= 15 && value.includes('T')) {
      const year = value.substring(0, 4);
      const month = value.substring(4, 6);
      const day = value.substring(6, 8);
      const hour = value.substring(9, 11);
      const minute = value.substring(11, 13);
      const second = value.substring(13, 15);
      return `${year}-${month}-${day}T${hour}:${minute}:${second}Z`;
    }

    return null;
  }

  /**
   * Extract categories (tags)
   * Handles both comma-separated (CATEGORIES:a,b,c) and multiple lines
   * (CATEGORIES:a\nCATEGORIES:b) as servers use both formats.
   */
  private extractCategories(data: string): string[] {
    const regex = /^CATEGORIES[;:](.+)$/gm;
    const categories: string[] = [];
    let match;

    while ((match = regex.exec(data)) !== null) {
      // Extract value after last colon (handles parameters)
      const fullValue = match[1];
      const colonIndex = fullValue.lastIndexOf(':');
      const value = colonIndex >= 0 ? fullValue.substring(colonIndex + 1).trim() : fullValue.trim();

      // Split by unescaped commas: split on commas that aren't preceded by backslash
      const parts = value.split(/(?<!\\),/);
      for (const part of parts) {
        categories.push(this.unescapeText(part.trim()));
      }
    }

    return categories;
  }

  /** Case-insensitive, order-preserving dedupe — Obsidian treats tags case-insensitively. */
  private dedupeTags(tags: string[]): string[] {
    const seen = new Set<string>();
    return tags.filter((tag) => {
      const key = tag.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  /**
   * Format date as YYYYMMDD
   * For date-only strings (YYYY-MM-DD), parses without timezone conversion
   */
  private formatDate(dateInput: Date | string): string {
    // If it's already a YYYY-MM-DD string, parse it directly without timezone issues
    if (typeof dateInput === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateInput)) {
      const [year, month, day] = dateInput.split('-');
      return `${year}${month}${day}`;
    }

    // Otherwise treat as Date object (use local time)
    const date = dateInput instanceof Date ? dateInput : new Date(dateInput);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}${month}${day}`;
  }

  /**
   * Resolve a CommonTask completedDate to the instant written as COMPLETED.
   * Obsidian completion is date-only (✅ YYYY-MM-DD) with no time; anchor it
   * at local noon so the UTC timestamp maps back to the same local calendar
   * day in any timezone (round-trip safe). A full datetime is already an
   * instant and is preserved as-is. See issue #43.
   */
  private toCompletedInstant(completedDate: string): Date {
    if (/^\d{4}-\d{2}-\d{2}$/.test(completedDate)) {
      const [year, month, day] = completedDate.split('-').map(Number);
      return new Date(year, month - 1, day, 12, 0, 0);
    }
    return new Date(completedDate);
  }

  /**
   * Format datetime as YYYYMMDDTHHMMSSZ (UTC)
   */
  private formatDateTimeUTC(date: Date): string {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const day = String(date.getUTCDate()).padStart(2, '0');
    const hour = String(date.getUTCHours()).padStart(2, '0');
    const minute = String(date.getUTCMinutes()).padStart(2, '0');
    const second = String(date.getUTCSeconds()).padStart(2, '0');
    return `${year}${month}${day}T${hour}${minute}${second}Z`;
  }

  /**
   * Escape special characters in iCalendar text
   */
  private escapeText(text: string): string {
    return text
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\n/g, '\\n');
  }

  private buildDescription(body: string, obsidianUrl?: string): string {
    if (!obsidianUrl && !body) return '';
    if (!obsidianUrl) return body;
    if (!body) return obsidianUrl;
    return `${obsidianUrl}\n\n${body}`;
  }

  private stripObsidianLinks(body: string): string {
    const lines = body.split('\n');
    const filtered = lines.filter(line => !line.match(/^obsidian:\/\/open\?vault=/));
    return filtered.join('\n').replace(/^\n+/, '');
  }

  /**
   * Unescape special characters from iCalendar text
   */
  private unescapeText(text: string): string {
    return text
      .replace(/\\n/g, '\n')
      .replace(/\\,/g, ',')
      .replace(/\\;/g, ';')
      .replace(/\\\\/g, '\\');
  }
}
