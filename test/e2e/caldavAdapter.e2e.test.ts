import { CalDAVClientDirect } from '../../src/caldav/calDAVClientDirect';
import { CalDAVAdapter } from '../../src/sync/caldavAdapter';
import { CommonTask } from '../../src/sync/types';
import { IdMapping } from '../../src/types';
import { FetchHttpClient } from '../helpers/fetchHttpClient';
import { RADICALE, createIsolatedCalendar } from '../helpers/radicaleSetup';

const emptyIdMapping: IdMapping = { taskIdToCaldavUid: {}, caldavUidToTaskId: {} };

const httpClient = new FetchHttpClient();

let calendarName: string;
let clean: () => Promise<void>;
let cleanup: () => Promise<void>;

function makeClient(): CalDAVClientDirect {
  return new CalDAVClientDirect(
    {
      serverUrl: RADICALE.baseUrl,
      username: RADICALE.username,
      password: RADICALE.password,
      calendarName,
    },
    httpClient,
  );
}

function buildVTODO(uid: string, summary: string, extra: string[] = []): string {
  const hasStatus = extra.some(l => l.startsWith('STATUS:'));
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//E2E Test//EN',
    'BEGIN:VTODO',
    `UID:${uid}`,
    'DTSTAMP:20250101T000000Z',
    `SUMMARY:${summary}`,
    ...(hasStatus ? [] : ['STATUS:NEEDS-ACTION']),
    ...extra,
    'END:VTODO',
    'END:VCALENDAR',
  ].join('\r\n');
}

beforeAll(async () => {
  const cal = await createIsolatedCalendar();
  calendarName = cal.calendarName;
  clean = cal.clean;
  cleanup = cal.cleanup;
});

beforeEach(async () => {
  await clean();
});

afterAll(async () => {
  await cleanup();
});

describe('CalDAVAdapter E2E', () => {

  describe('normalize round-trip', () => {
    it('should normalize VTODOs from a real server into CommonTasks', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const uid = `e2e-adapt-${Date.now()}`;
      const vtodo = buildVTODO(uid, 'Buy milk', [
        'DUE;VALUE=DATE:20250615',
        'PRIORITY:3',
        'CATEGORIES:sync,groceries',
      ]);

      await client.createVTODO(vtodo, uid);

      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].uid).toBe(uid);
      expect(tasks[0].title).toBe('Buy milk');
      expect(tasks[0].status).toBe('TODO');
      expect(tasks[0].dueDate).toBe('2025-06-15');
      expect(tasks[0].priority).toBe('high');
      expect(tasks[0].tags).toEqual(['sync', 'groceries']);
    });

    it('should use mapped obsidian ID when available', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const uid = `e2e-mapped-${Date.now()}`;
      await client.createVTODO(buildVTODO(uid, 'Mapped task'), uid);

      const vtodos = await client.fetchVTODOs();
      const idMapping: IdMapping = {
        taskIdToCaldavUid: { 'obsidian-task-id-123': uid },
        caldavUidToTaskId: { [uid]: 'obsidian-task-id-123' },
      };
      const tasks = adapter.normalize(vtodos, idMapping);

      expect(tasks[0].uid).toBe('obsidian-task-id-123');
    });
  });

  describe('fromCommonTask round-trip', () => {
    it('should create a VTODO from CommonTask and read it back', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const task: CommonTask = {
        uid: 'round-trip-id',
        title: 'Round trip test',
        status: 'TODO',
        dueDate: '2025-07-01',
        startDate: null,
        scheduledDate: '2025-06-28',
        completedDate: null,
        priority: 'high',
        tags: ['sync', 'test'],
        recurrenceRule: '',
        body: '',
      };

      const caldavUID = `e2e-roundtrip-${Date.now()}`;
      const vtodoData = adapter.fromCommonTask(task, caldavUID);
      await client.createVTODO(vtodoData, caldavUID);

      // Fetch back and normalize
      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].title).toBe('Round trip test');
      expect(tasks[0].status).toBe('TODO');
      expect(tasks[0].dueDate).toBe('2025-07-01');
      expect(tasks[0].scheduledDate).toBe('2025-06-28');
      expect(tasks[0].priority).toBe('high');
      expect(tasks[0].tags).toEqual(['sync', 'test']);
    });

    it('should round-trip a completed task', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const task: CommonTask = {
        uid: 'done-id',
        title: 'Completed task',
        status: 'DONE',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: '2025-06-10',
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: '',
      };

      const caldavUID = `e2e-done-${Date.now()}`;
      const vtodoData = adapter.fromCommonTask(task, caldavUID);
      await client.createVTODO(vtodoData, caldavUID);

      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks[0].status).toBe('DONE');
      expect(tasks[0].completedDate).toBe('2025-06-10');
    });
  });

  describe('DESCRIPTION round-trip', () => {
    it('should round-trip DESCRIPTION through the server', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const task: CommonTask = {
        uid: 'desc-rt-id',
        title: 'Task with notes',
        status: 'TODO',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: 'Remember to check the farmers market\nAlso need cleaning supplies',
      };

      const caldavUID = `e2e-desc-rt-${Date.now()}`;
      const vtodoData = adapter.fromCommonTask(task, caldavUID);
      await client.createVTODO(vtodoData, caldavUID);

      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].body).toBe('Remember to check the farmers market\nAlso need cleaning supplies');
    });

    it('should handle DESCRIPTION with special characters', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const task: CommonTask = {
        uid: 'desc-special-id',
        title: 'Special chars test',
        status: 'TODO',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: 'Commas, semicolons; colons: and backslashes\\',
      };

      const caldavUID = `e2e-desc-special-${Date.now()}`;
      const vtodoData = adapter.fromCommonTask(task, caldavUID);
      await client.createVTODO(vtodoData, caldavUID);

      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].body).toBe('Commas, semicolons; colons: and backslashes\\');
    });

    it('should return empty body when VTODO has no DESCRIPTION', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const caldavUID = `e2e-no-desc-${Date.now()}`;
      const vtodo = buildVTODO(caldavUID, 'No description task');
      await client.createVTODO(vtodo, caldavUID);

      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].body).toBe('');
    });

    it('should update DESCRIPTION on existing VTODO', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      // Create initial VTODO without description
      const caldavUID = `e2e-desc-update-${Date.now()}`;
      const vtodo = buildVTODO(caldavUID, 'Task to update');
      await client.createVTODO(vtodo, caldavUID);

      // Update with description
      const task: CommonTask = {
        uid: 'update-desc-id',
        title: 'Task to update',
        status: 'TODO',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: 'New description added',
      };

      const idMapping: IdMapping = {
        taskIdToCaldavUid: { 'update-desc-id': caldavUID },
        caldavUidToTaskId: { [caldavUID]: 'update-desc-id' },
      };
      await adapter.applyChanges(
        [{ type: 'update', task }],
        idMapping,
      );

      const vtodos = await client.fetchVTODOs();
      const tasks = adapter.normalize(vtodos, emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].body).toBe('New description added');
    });
  });

  describe('applyChanges', () => {
    it('should create, update, and delete VTODOs', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      // Create a task to later update and delete
      const existingUID = `e2e-existing-${Date.now()}`;
      await client.createVTODO(buildVTODO(existingUID, 'Existing task'), existingUID);

      const toDeleteUID = `e2e-delete-${Date.now()}`;
      await client.createVTODO(buildVTODO(toDeleteUID, 'To delete'), toDeleteUID);

      let vtodos = await client.fetchVTODOs();
      expect(vtodos.length).toBe(2);

      const idMapping: IdMapping = {
        taskIdToCaldavUid: { 'obs-existing': existingUID, 'obs-delete': toDeleteUID },
        caldavUidToTaskId: { [existingUID]: 'obs-existing', [toDeleteUID]: 'obs-delete' },
      };

      // Apply: create new + update existing + delete one
      const newTask: CommonTask = {
        uid: 'obs-new',
        title: 'Brand new task',
        status: 'TODO',
        dueDate: '2025-08-01',
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'medium',
        tags: ['sync'],
        recurrenceRule: '',
        body: '',
      };

      const updatedTask: CommonTask = {
        uid: 'obs-existing',
        title: 'Updated existing task',
        status: 'DONE',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: '2025-07-15',
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: '',
      };

      const deletedTask: CommonTask = {
        uid: 'obs-delete',
        title: 'To delete',
        status: 'TODO',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: '',
      };

      await adapter.applyChanges(
        [
          { type: 'create', task: newTask },
          { type: 'update', task: updatedTask },
          { type: 'delete', task: deletedTask },
        ],
        idMapping,
      );

      // Verify final state
      vtodos = await client.fetchVTODOs();
      expect(vtodos.length).toBe(2); // 2 original - 1 deleted + 1 created = 2

      const tasks = adapter.normalize(vtodos, emptyIdMapping);
      const descriptions = tasks.map(t => t.title).sort();
      expect(descriptions).toEqual(['Brand new task', 'Updated existing task']);

      const updated = tasks.find(t => t.title === 'Updated existing task');
      expect(updated?.status).toBe('DONE');
    });
  });

  describe('foreign property preservation (jtx Board / RFC 5545)', () => {
    it('preserves RELATED-TO, VALARM, X- extension, and PERCENT-COMPLETE through an Obsidian update', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const uid = `e2e-jtx-${Date.now()}`;
      const jtxVTODO = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//jtx Board//EN',
        'BEGIN:VTODO',
        `UID:${uid}`,
        'DTSTAMP:20250101T000000Z',
        'SUMMARY:Plan trip',
        'STATUS:IN-PROCESS',
        'PERCENT-COMPLETE:40',
        'RELATED-TO;RELTYPE=PARENT:parent-uid-123',
        'X-JTX-FOOBAR:custom-metadata',
        'BEGIN:VALARM',
        'ACTION:DISPLAY',
        'DESCRIPTION:Reminder',
        'TRIGGER:-PT15M',
        'END:VALARM',
        'END:VTODO',
        'END:VCALENDAR',
      ].join('\r\n');

      await client.createVTODO(jtxVTODO, uid);

      const updatedTask: CommonTask = {
        uid: 'obs-jtx',
        title: 'Plan trip (updated)',
        status: 'TODO',
        dueDate: '2026-12-01',
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'high',
        tags: [],
        recurrenceRule: '',
        body: '',
      };

      const idMapping: IdMapping = {
        taskIdToCaldavUid: { 'obs-jtx': uid },
        caldavUidToTaskId: { [uid]: 'obs-jtx' },
      };
      await adapter.applyChanges([{ type: 'update', task: updatedTask }], idMapping);

      const vtodos = await client.fetchVTODOs();
      expect(vtodos).toHaveLength(1);
      const rawData = vtodos[0].data;

      expect(rawData).toContain('RELATED-TO');
      expect(rawData).toContain('parent-uid-123');
      expect(rawData).toContain('X-JTX-FOOBAR:custom-metadata');
      expect(rawData).toContain('BEGIN:VALARM');
      expect(rawData).toContain('END:VALARM');
      expect(rawData).toContain('PERCENT-COMPLETE:40');
      expect(rawData).toContain('STATUS:IN-PROCESS'); // must survive the Obsidian update

      const tasks = adapter.normalize(vtodos, emptyIdMapping);
      expect(tasks[0].title).toBe('Plan trip (updated)');
      expect(tasks[0].dueDate).toBe('2026-12-01');
      expect(tasks[0].priority).toBe('high');
    });

    it('strips PERCENT-COMPLETE:40 and sets 100 on completion, preserves X- properties', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const uid = `e2e-jtx-complete-${Date.now()}`;
      const jtxVTODO = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//jtx Board//EN',
        'BEGIN:VTODO',
        `UID:${uid}`,
        'DTSTAMP:20250101T000000Z',
        'SUMMARY:In progress task',
        'STATUS:IN-PROCESS',
        'PERCENT-COMPLETE:60',
        'X-JTX-FOOBAR:preserved',
        'END:VTODO',
        'END:VCALENDAR',
      ].join('\r\n');

      await client.createVTODO(jtxVTODO, uid);

      const completedTask: CommonTask = {
        uid: 'obs-jtx-done',
        title: 'In progress task',
        status: 'DONE',
        dueDate: null,
        startDate: null,
        scheduledDate: null,
        completedDate: '2026-07-18',
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: '',
      };

      const idMapping: IdMapping = {
        taskIdToCaldavUid: { 'obs-jtx-done': uid },
        caldavUidToTaskId: { [uid]: 'obs-jtx-done' },
      };
      await adapter.applyChanges([{ type: 'complete', task: completedTask }], idMapping);

      const vtodos = await client.fetchVTODOs();
      const rawData = vtodos[0].data;

      expect(rawData).toContain('PERCENT-COMPLETE:100');
      expect(rawData).not.toContain('PERCENT-COMPLETE:60');
      expect(rawData).toContain('X-JTX-FOOBAR:preserved');
      expect(rawData).toContain('STATUS:COMPLETED');
    });

    it('preserves TZID and time-of-day on DUE through an Obsidian date change', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client);
      await client.connect();

      const uid = `e2e-jtx-datetime-${Date.now()}`;
      const jtxVTODO = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//jtx Board//EN',
        'BEGIN:VTODO',
        `UID:${uid}`,
        'DTSTAMP:20250101T000000Z',
        'SUMMARY:Meeting prep',
        'STATUS:IN-PROCESS',
        'DUE;TZID=Europe/Berlin:20240115T140000',
        'END:VTODO',
        'END:VCALENDAR',
      ].join('\r\n');

      await client.createVTODO(jtxVTODO, uid);

      const updatedTask: CommonTask = {
        uid: 'obs-dt',
        title: 'Meeting prep',
        status: 'TODO',
        dueDate: '2026-09-01',
        startDate: null,
        scheduledDate: null,
        completedDate: null,
        priority: 'none',
        tags: [],
        recurrenceRule: '',
        body: '',
      };

      const idMapping: IdMapping = {
        taskIdToCaldavUid: { 'obs-dt': uid },
        caldavUidToTaskId: { [uid]: 'obs-dt' },
      };
      await adapter.applyChanges([{ type: 'update', task: updatedTask }], idMapping);

      const vtodos = await client.fetchVTODOs();
      const rawData = vtodos[0].data;

      // Date updated; timezone and time-of-day must be preserved
      expect(rawData).toMatch(/DUE;TZID=Europe\/Berlin:20260901T140000/);
      expect(rawData).not.toMatch(/DUE;VALUE=DATE/);
      // STATUS:IN-PROCESS must also survive
      expect(rawData).toContain('STATUS:IN-PROCESS');
    });
  });

  describe('fetchTasks category filter', () => {
    it('pulls every server task when the category is empty (iOS Reminders case, issue #94)', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client, '');
      await client.connect();

      const uidTagged = `e2e-cat-tagged-${Date.now()}`;
      const uidBare = `e2e-cat-bare-${Date.now()}`;
      await client.createVTODO(
        buildVTODO(uidTagged, 'Tagged task', ['CATEGORIES:sync']),
        uidTagged,
      );
      await client.createVTODO(
        buildVTODO(uidBare, 'iOS task with no CATEGORIES'),
        uidBare,
      );

      const tasks = await adapter.fetchTasks(emptyIdMapping);

      const titles = tasks.map((t: CommonTask) => t.title).sort();
      expect(titles).toEqual(['Tagged task', 'iOS task with no CATEGORIES'].sort());
    });

    it('pulls only matching tasks when a category is set', async () => {
      const client = makeClient();
      const adapter = new CalDAVAdapter(client, 'work');
      await client.connect();

      const uidWork = `e2e-cat-work-${Date.now()}`;
      const uidPersonal = `e2e-cat-personal-${Date.now()}`;
      await client.createVTODO(
        buildVTODO(uidWork, 'Work task', ['CATEGORIES:work']),
        uidWork,
      );
      await client.createVTODO(
        buildVTODO(uidPersonal, 'Personal task', ['CATEGORIES:personal']),
        uidPersonal,
      );

      const tasks = await adapter.fetchTasks(emptyIdMapping);

      expect(tasks).toHaveLength(1);
      expect(tasks[0].title).toBe('Work task');
      expect(tasks[0].tags).toEqual([]);
    });
  });
});
