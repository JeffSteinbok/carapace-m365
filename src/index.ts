/**
 * Microsoft 365 plugin — unified mail, calendar, tasks, and OneDrive tools via Microsoft Graph.
 */

import {
  definePlugin,
  type PluginApi,
  type PluginEntry,
} from "carapace-plugin-sdk";
import { Type } from "@sinclair/typebox";
import {
  DEFAULT_TOKEN_BROKER_URL,
  isM365FeatureEnabled,
  parseM365Features,
  type M365Feature,
} from "@carapace/m365-graph-auth";
import {
  fetchCalendar,
  createEvent,
  updateEvent,
  deleteEvent,
  createMeeting,
  queryEvents,
  getInbox,
  searchMail,
  readMessage,
  saveAttachments,
  saveDraft,
  sendMessage,
  replyToMessage,
  forwardMessage,
  moveMessage,
  flagMessage,
  listTaskListsHandler,
  listTasks,
  createTask,
  updateTask,
  completeTask,
  deleteTask,
  type OutlookCalendarConfig,
} from "./handlers.js";
import {
  createOneDriveFolder,
  deleteOneDriveItem,
  downloadOneDriveFile,
  getOneDriveMetadata,
  listOneDrive,
  moveOneDriveItem,
  searchOneDrive,
  uploadOneDriveFile,
} from "./onedrive.js";
import { DEFAULT_DIRECT_TOKEN_STATE_PATH } from "./direct-token-state.js";

/**
 * Published public-client app registration (PKCE, personal Microsoft accounts).
 * The client ID of a public client is not a secret, so it ships as the default:
 * users can run `npm run login` and consent with their own account without
 * registering their own Azure app. Override via config.clientId or
 * M365_CLIENT_ID to point at your own registration.
 */
export const DEFAULT_CLIENT_ID = "0c3df71b-4dc2-49a7-b6e7-e5c3c48bf501";

const createBaseEntry = definePlugin({
  id: "m365",
  name: "Microsoft 365",
  description: "Outlook mail, calendar, tasks, and OneDrive tools via Microsoft Graph",

  configSchema: Type.Object({
    clientId: Type.Optional(Type.String({ description: "Microsoft OAuth client ID" })),
    clientSecret: Type.Optional(Type.String({ description: "Microsoft OAuth client secret" })),
    refreshToken: Type.Optional(Type.String({ description: "Microsoft OAuth refresh token" })),
    directTokenStatePath: Type.Optional(Type.String({ description: "Direct-mode refresh-token state path" })),
    tenant: Type.Optional(Type.String({ description: "Microsoft OAuth tenant (default: consumers)" })),
    tokenBrokerUrl: Type.Optional(Type.String({ description: "Authoritative Microsoft 365 token broker URL" })),
    tokenBrokerSecret: Type.Optional(Type.String({ description: "Bearer secret used to authenticate to the token broker" })),
    features: Type.Optional(
      Type.Array(Type.String(), {
        description: "Microsoft 365 features to register. OneDrive requires explicit enablement.",
      }),
    ),
    personalCalendarNames: Type.Optional(
      Type.Array(Type.String(), { description: "Additional personal calendar names to match." }),
    ),
    familyCalendarNames: Type.Optional(
      Type.Array(Type.String(), { description: "Additional family calendar names to match." }),
    ),
  }),

  tools: (tool) => [

    // -------------------------------------------------------------------------
    // Mail tools
    // -------------------------------------------------------------------------

    tool({
      name: "outlook_inbox",
      label: "Outlook Inbox",
      description: "List recent messages from the Outlook inbox, or any other mail folder.",
      parameters: Type.Object({
        folder: Type.Optional(Type.String({ description: "Mail folder to read (default: inbox). Well-known folder names: inbox, junkemail, deleteditems, sentitems, drafts, outbox, archive." })),
        limit: Type.Optional(Type.Integer({ description: "Maximum number of messages to return (default 10)." })),
        unread: Type.Optional(Type.Boolean({ description: "Only show unread messages." })),
      }),
      async execute({ folder, limit, unread }, config) {
        try {
          return await getInbox(resolveConfig(config), { folder, limit, unread });
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_search",
      label: "Outlook Search",
      description: "Search Outlook messages by query text, sender, subject, or date range.",
      parameters: Type.Object({
        query: Type.Optional(Type.String({ description: "Full-text search across subject and body." })),
        from: Type.Optional(Type.String({ description: "Filter by sender email address." })),
        subject: Type.Optional(Type.String({ description: "Filter by subject (substring match)." })),
        since: Type.Optional(Type.String({ description: "Only messages received on or after this date (YYYY-MM-DD)." })),
        before: Type.Optional(Type.String({ description: "Only messages received on or before this date (YYYY-MM-DD)." })),
        limit: Type.Optional(Type.Integer({ description: "Maximum number of results (default 10)." })),
      }),
      async execute({ query, from, subject, since, before, limit }, config) {
        try {
          return await searchMail(resolveConfig(config), { query, from, subject, since, before, limit });
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_read",
      label: "Outlook Read",
      description: "Read a specific Outlook message by its ID, including full body content.",
      parameters: Type.Object({
        message_id: Type.String({ description: "The Microsoft Graph message ID to retrieve." }),
      }),
      async execute({ message_id }, config) {
        try {
          return await readMessage(resolveConfig(config), { message_id });
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_save_attachments",
      label: "Outlook Save Attachments",
      description: "Download attachments from an Outlook message to a local directory.",
      parameters: Type.Object({
        message_id: Type.String({ description: "The Microsoft Graph message ID." }),
        output_dir: Type.String({ description: "Local directory path to save attachments to (created if needed)." }),
        content_types: Type.Optional(Type.Array(Type.String(), { description: "Content type filters (e.g. ['image/*']). Defaults to ['image/*']." })),
      }),
      async execute({ message_id, output_dir, content_types }, config) {
        try {
          return await saveAttachments(resolveConfig(config), { message_id, output_dir, content_types });
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_save_draft",
      label: "Outlook Save Draft",
      description: "Save a draft email to the Outlook Drafts folder. All fields are optional — useful for saving a partial draft before sending.",
      parameters: Type.Object({
        to: Type.Optional(Type.Union([Type.String(), Type.Array(Type.String())], { description: "Recipient email address(es)." })),
        subject: Type.Optional(Type.String({ description: "Email subject line." })),
        body: Type.Optional(Type.String({ description: "Plain-text email body." })),
        cc: Type.Optional(Type.Array(Type.String(), { description: "CC recipient email address(es)." })),
        attachment: Type.Optional(Type.Array(Type.String(), { description: "File path(s) to attach." })),
        in_reply_to: Type.Optional(Type.String({ description: "Message-ID of the email being replied to (for threading)." })),
        signature: Type.Optional(Type.String({ description: "Signature block appended after body." })),
      }),
      async execute(params, config) {
        try {
          return await saveDraft(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_send",
      label: "Outlook Send",
      description: "Send a plain-text email via Outlook, with optional file attachments.",
      parameters: Type.Object({
        to: Type.Union([Type.String(), Type.Array(Type.String())], { description: "Recipient email address(es)." }),
        subject: Type.String({ description: "Email subject line." }),
        body: Type.String({ description: "Plain-text email body." }),
        cc: Type.Optional(Type.Array(Type.String(), { description: "CC recipient email address(es)." })),
        attachment: Type.Optional(Type.Array(Type.String(), { description: "File path(s) to attach." })),
        in_reply_to: Type.Optional(Type.String({ description: "Message-ID of the email being replied to (enables threading). Include angle brackets, e.g. <abc@mail.example.com>." })),
        references: Type.Optional(Type.String({ description: "Space-separated list of Message-IDs for the full thread References header." })),
        signature: Type.Optional(Type.String({ description: "Signature block appended after body." })),
      }),
      async execute(params, config) {
        try {
          return await sendMessage(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_reply",
      label: "Outlook Reply",
      description: "Reply to an existing Outlook message with proper threading. Handles In-Reply-To and References headers automatically via Microsoft Graph.",
      parameters: Type.Object({
        message_id: Type.String({ description: "The Microsoft Graph message ID to reply to." }),
        body: Type.String({ description: "Reply body text." }),
        reply_all: Type.Optional(Type.Boolean({ description: "Reply to all recipients (default: false)." })),
        signature: Type.Optional(Type.String({ description: "Signature block appended after body." })),
      }),
      async execute(params, config) {
        try {
          return await replyToMessage(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_forward",
      label: "Outlook Forward",
      description: "Forward an existing Outlook message to new recipients.",
      parameters: Type.Object({
        message_id: Type.String({ description: "The Microsoft Graph message ID to forward." }),
        to: Type.Union([Type.String(), Type.Array(Type.String())], { description: "Recipient email address(es) to forward to." }),
        comment: Type.Optional(Type.String({ description: "Optional note to prepend to the forwarded message." })),
      }),
      async execute(params, config) {
        try {
          return await forwardMessage(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_move",
      label: "Outlook Move",
      description: "Move an Outlook message to a different mail folder.",
      parameters: Type.Object({
        message_id: Type.String({ description: "The Microsoft Graph message ID to move." }),
        destination_folder: Type.String({ description: "Target folder name or well-known folder name (inbox, archive, deleteditems, junkemail, sentitems, drafts)." }),
      }),
      async execute(params, config) {
        try {
          return await moveMessage(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_flag",
      label: "Outlook Flag",
      description: "Flag, complete, or unflag an Outlook message.",
      parameters: Type.Object({
        message_id: Type.String({ description: "The Microsoft Graph message ID to flag." }),
        flag_status: Type.Union([Type.Literal("flagged"), Type.Literal("complete"), Type.Literal("notFlagged")], { description: "Flag status: 'flagged', 'complete', or 'notFlagged'." }),
      }),
      async execute(params, config) {
        try {
          return await flagMessage(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    // -------------------------------------------------------------------------
    // Tasks tools
    // -------------------------------------------------------------------------

    tool({
      name: "outlook_task_lists",
      label: "Outlook Task Lists",
      description: "List Microsoft To Do / Outlook task lists available on the account.",
      parameters: Type.Object({}),
      async execute(_, config) {
        try {
          return await listTaskListsHandler(resolveConfig(config));
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_tasks",
      label: "Outlook Tasks",
      description: "List tasks from a Microsoft To Do list.",
      parameters: Type.Object({
        task_list: Type.Optional(Type.String({ description: "Task list name or list ID. Defaults to the Tasks/To Do list if found." })),
        include_completed: Type.Optional(Type.Boolean({ description: "Include completed tasks (default: false)." })),
        limit: Type.Optional(Type.Integer({ description: "Maximum number of tasks to return (default 20)." })),
      }),
      async execute({ task_list, include_completed, limit }, config) {
        try {
          return await listTasks(resolveConfig(config), { task_list, include_completed, limit });
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_create_task",
      label: "Outlook Create Task",
      description: "Create a Microsoft To Do task, optionally with due/reminder dates.",
      parameters: Type.Object({
        title: Type.String({ description: "Task title." }),
        task_list: Type.Optional(Type.String({ description: "Task list name or list ID. Defaults to the Tasks/To Do list if found." })),
        due: Type.Optional(Type.String({ description: "Due datetime in ISO format." })),
        reminder: Type.Optional(Type.String({ description: "Reminder datetime in ISO format." })),
        notes: Type.Optional(Type.String({ description: "Task notes or body text." })),
        importance: Type.Optional(Type.Union([Type.Literal("low"), Type.Literal("normal"), Type.Literal("high")], { description: "Task importance." })),
      }),
      async execute(params, config) {
        try {
          return await createTask(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_update_task",
      label: "Outlook Update Task",
      description: "Update an existing Microsoft To Do task.",
      parameters: Type.Object({
        task_id: Type.String({ description: "Task ID to update." }),
        task_list: Type.Optional(Type.String({ description: "Task list name or list ID." })),
        title: Type.Optional(Type.String({ description: "New task title." })),
        due: Type.Optional(Type.String({ description: "New due datetime in ISO format. Set empty string to clear." })),
        reminder: Type.Optional(Type.String({ description: "New reminder datetime in ISO format. Set empty string to clear." })),
        notes: Type.Optional(Type.String({ description: "New notes/body text." })),
        importance: Type.Optional(Type.Union([Type.Literal("low"), Type.Literal("normal"), Type.Literal("high")], { description: "Task importance." })),
        status: Type.Optional(Type.Union([Type.Literal("notStarted"), Type.Literal("inProgress"), Type.Literal("completed")], { description: "Task status." })),
      }),
      async execute(params, config) {
        try {
          return await updateTask(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_complete_task",
      label: "Outlook Complete Task",
      description: "Mark a Microsoft To Do task as completed.",
      parameters: Type.Object({
        task_id: Type.String({ description: "Task ID to complete." }),
        task_list: Type.Optional(Type.String({ description: "Task list name or list ID." })),
      }),
      async execute(params, config) {
        try {
          return await completeTask(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_delete_task",
      label: "Outlook Delete Task",
      description: "Delete a Microsoft To Do task.",
      parameters: Type.Object({
        task_id: Type.String({ description: "Task ID to delete." }),
        task_list: Type.Optional(Type.String({ description: "Task list name or list ID." })),
      }),
      async execute(params, config) {
        try {
          return await deleteTask(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    // -------------------------------------------------------------------------
    // Calendar tools
    // -------------------------------------------------------------------------

    tool({
      name: "outlook_calendar_fetch",
      label: "Outlook Calendar",
      description: "Fetch upcoming events from Outlook personal, family, or combined calendars.",
      parameters: Type.Object({
        calendar: Type.Optional(Type.Union([Type.Literal("personal"), Type.Literal("family"), Type.Literal("all")], { description: "Which calendar to fetch: personal, family, or all (default: all)." })),
        days: Type.Optional(Type.Integer({ description: "Number of days ahead to fetch events for (default: 7)." })),
      }),
      async execute({ calendar, days }, config) {
        try {
          return await fetchCalendar(resolveConfig(config), { calendar, days });
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_create_event",
      label: "Outlook Create Event",
      description: "Create a new event on an Outlook personal or family calendar.",
      parameters: Type.Object({
        subject: Type.String({ description: "Event title." }),
        start: Type.String({ description: "Start datetime in ISO format (e.g. 2026-03-15T14:00)." }),
        duration: Type.Optional(Type.String({ description: "Duration: '1h', '30m', '1.5h' (default: 1h). Ignored when end is supplied." })),
        end: Type.Optional(Type.String({ description: "End datetime in ISO format. Overrides duration." })),
        timezone: Type.Optional(Type.String({ description: "IANA timezone (default: America/Los_Angeles)." })),
        location: Type.Optional(Type.String({ description: "Meeting location." })),
        description: Type.Optional(Type.String({ description: "Event description / agenda." })),
        attendees: Type.Optional(Type.Array(Type.String(), { description: "Attendee email addresses." })),
        calendar: Type.Optional(Type.Union([Type.Literal("personal"), Type.Literal("family")], { description: "Target calendar (default: personal)." })),
      }),
      async execute(params, config) {
        try {
          return await createEvent(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_update_event",
      label: "Outlook Update Event",
      description: "Update an existing Outlook calendar event by event ID. Get the event ID from outlook_calendar_fetch.",
      parameters: Type.Object({
        event_id: Type.String({ description: "Graph event ID from outlook_calendar_fetch." }),
        subject: Type.Optional(Type.String({ description: "New event title." })),
        start: Type.Optional(Type.String({ description: "New start datetime in ISO format." })),
        end: Type.Optional(Type.String({ description: "New end datetime in ISO format." })),
        duration: Type.Optional(Type.String({ description: "New duration (e.g. '1h', '30m'). Requires start." })),
        timezone: Type.Optional(Type.String({ description: "IANA timezone for new start/end (default: America/Los_Angeles)." })),
        location: Type.Optional(Type.String({ description: "New location." })),
        description: Type.Optional(Type.String({ description: "New description." })),
        add_attendees: Type.Optional(Type.Array(Type.String(), { description: "Email addresses to add as attendees." })),
        remove_attendees: Type.Optional(Type.Array(Type.String(), { description: "Email addresses to remove from attendees." })),
        status: Type.Optional(Type.Union([Type.Literal("confirmed"), Type.Literal("tentative"), Type.Literal("cancelled")], { description: "Update event status." })),
      }),
      async execute(params, config) {
        try {
          return await updateEvent(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_delete_event",
      label: "Outlook Delete Event",
      description: "Delete (cancel) an Outlook calendar event by event ID.",
      parameters: Type.Object({
        event_id: Type.String({ description: "Graph event ID to delete. Get this from outlook_calendar_fetch." }),
      }),
      async execute(params, config) {
        try {
          return await deleteEvent(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_meeting",
      label: "Outlook Create Meeting",
      description: "Create a calendar meeting invite via Microsoft Graph and send invitations to attendees.",
      parameters: Type.Object({
        to: Type.Union([Type.String(), Type.Array(Type.String())], { description: "Attendee email address(es)." }),
        cc: Type.Optional(Type.Array(Type.String(), { description: "Optional attendees (will be marked as optional)." })),
        subject: Type.String({ description: "Meeting title." }),
        start: Type.String({ description: "Start datetime in ISO format (e.g. 2026-03-15T14:00)." }),
        duration: Type.Optional(Type.String({ description: "Duration: '1h', '30m', '1.5h' (default: 1h)." })),
        end: Type.Optional(Type.String({ description: "End datetime in ISO format. Overrides duration." })),
        timezone: Type.Optional(Type.String({ description: "IANA timezone (default: America/Los_Angeles)." })),
        location: Type.Optional(Type.String({ description: "Meeting location." })),
        description: Type.Optional(Type.String({ description: "Meeting description / agenda." })),
        signature: Type.Optional(Type.String({ description: "Signature block for the invite email." })),
      }),
      async execute(params, config) {
        try {
          return await createMeeting(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "outlook_query_events",
      label: "Outlook Query Events",
      description: "Query calendar events by date range, text, attendee email, or iCalUId.",
      parameters: Type.Object({
        after: Type.Optional(Type.String({ description: "Only events starting at or after this date (ISO, e.g. 2026-03-01)." })),
        before: Type.Optional(Type.String({ description: "Only events starting before this date (ISO, e.g. 2026-04-01)." })),
        text: Type.Optional(Type.String({ description: "Filter by text match on title." })),
        attendee: Type.Optional(Type.String({ description: "Filter to events including this attendee email." })),
        uid: Type.Optional(Type.String({ description: "Return the single event with this exact iCalUId." })),
      }),
      async execute(params, config) {
        try {
          return await queryEvents(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    // -------------------------------------------------------------------------
    // Mail store tool
    // -------------------------------------------------------------------------

    tool({
      name: "outlook_search_store",
      label: "Outlook Search Store",
      description: "Query the local SQLite mail cache for fast message lookup without hitting the Graph API. Returns graph_id (for reply/forward), internet_message_id (for threading), sender, subject, and date. Only available when M365_MAIL_STORE_PATH is configured.",
      parameters: Type.Object({
        graph_id: Type.Optional(Type.String({ description: "Look up a single message by Microsoft Graph ID (takes priority over search filters, returns one row or null)." })),
        sender_email: Type.Optional(Type.String({ description: "Filter by sender email address." })),
        subject_contains: Type.Optional(Type.String({ description: "Substring match on subject." })),
        since: Type.Optional(Type.String({ description: "ISO datetime lower bound for received_at." })),
        before: Type.Optional(Type.String({ description: "ISO datetime upper bound for received_at." })),
        limit: Type.Optional(Type.Integer({ description: "Maximum number of results to return (default 20, max 50)." })),
      }),
      async execute({ graph_id, sender_email, subject_contains, since, before, limit }) {
        const storePath = process.env.M365_MAIL_STORE_PATH?.trim();
        if (!storePath) {
          return { error: "Mail store is not configured (M365_MAIL_STORE_PATH not set)" };
        }
        // Dynamic import avoids bundling node:sqlite into the plugin bundle;
        // node:sqlite is a Node built-in (v22+) that must stay external.
        const { MailStore } = await import("@carapace/m365-mail-store");
        const store = new MailStore(storePath);
        try {
          if (graph_id) {
            const row = store.getByGraphId(graph_id);
            return { result: row };
          }
          const rows = store.search({
            sender_email,
            subject_contains,
            since,
            before,
            limit: Math.min(limit ?? 20, 50),
          });
          return { results: rows, count: rows.length };
        } finally {
          store.close();
        }
      },
    }),

    // -------------------------------------------------------------------------
    // OneDrive tools
    // -------------------------------------------------------------------------

    tool({
      name: "onedrive_list",
      label: "OneDrive List",
      description: "List files and folders in the OneDrive root or a folder addressed by item ID or path.",
      parameters: Type.Object({
        item_id: Type.Optional(Type.String({ description: "OneDrive folder item ID. Mutually exclusive with path." })),
        path: Type.Optional(Type.String({ description: "OneDrive folder path relative to the drive root. Omit for root." })),
        limit: Type.Optional(Type.Integer({ description: "Maximum items to return, 1-200 (default 50)." })),
      }),
      async execute(params, config) {
        try {
          return await listOneDrive(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_search",
      label: "OneDrive Search",
      description: "Search OneDrive files and folders by name or content.",
      parameters: Type.Object({
        query: Type.String({ description: "Search text." }),
        limit: Type.Optional(Type.Integer({ description: "Maximum items to return, 1-200 (default 25)." })),
      }),
      async execute(params, config) {
        try {
          return await searchOneDrive(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_metadata",
      label: "OneDrive Metadata",
      description: "Get metadata for the OneDrive root or an item addressed by ID or path.",
      parameters: Type.Object({
        item_id: Type.Optional(Type.String({ description: "OneDrive item ID. Mutually exclusive with path." })),
        path: Type.Optional(Type.String({ description: "OneDrive path relative to the drive root. Omit for root." })),
      }),
      async execute(params, config) {
        try {
          return await getOneDriveMetadata(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_download",
      label: "OneDrive Download",
      description: "Download a OneDrive file to a local path without overwriting unless explicitly requested.",
      parameters: Type.Object({
        item_id: Type.Optional(Type.String({ description: "OneDrive file item ID. Mutually exclusive with path." })),
        path: Type.Optional(Type.String({ description: "OneDrive file path relative to the drive root." })),
        output_path: Type.String({ description: "Local destination file path." }),
        overwrite: Type.Optional(Type.Boolean({ description: "Replace an existing local file (default false)." })),
      }),
      async execute(params, config) {
        try {
          return await downloadOneDriveFile(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_upload",
      label: "OneDrive Upload",
      description: "Upload a local file up to 4 MiB to OneDrive.",
      parameters: Type.Object({
        local_path: Type.String({ description: "Local file path to upload." }),
        name: Type.Optional(Type.String({ description: "Remote file name. Defaults to the local basename." })),
        parent_id: Type.Optional(Type.String({ description: "Destination folder item ID. Mutually exclusive with parent_path." })),
        parent_path: Type.Optional(Type.String({ description: "Destination folder path relative to root. Omit for root." })),
        overwrite: Type.Optional(Type.Boolean({ description: "Replace an existing remote file (default false)." })),
      }),
      async execute(params, config) {
        try {
          return await uploadOneDriveFile(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_create_folder",
      label: "OneDrive Create Folder",
      description: "Create a OneDrive folder without replacing an existing item.",
      parameters: Type.Object({
        name: Type.String({ description: "New folder name." }),
        parent_id: Type.Optional(Type.String({ description: "Parent folder item ID. Mutually exclusive with parent_path." })),
        parent_path: Type.Optional(Type.String({ description: "Parent folder path relative to root. Omit for root." })),
      }),
      async execute(params, config) {
        try {
          return await createOneDriveFolder(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_move",
      label: "OneDrive Move or Rename",
      description: "Move or rename a OneDrive item addressed by item ID or path.",
      parameters: Type.Object({
        item_id: Type.Optional(Type.String({ description: "Source item ID. Mutually exclusive with path." })),
        path: Type.Optional(Type.String({ description: "Source item path relative to root." })),
        new_name: Type.Optional(Type.String({ description: "New file or folder name." })),
        parent_id: Type.Optional(Type.String({ description: "Destination folder item ID. Mutually exclusive with parent_path." })),
        parent_path: Type.Optional(Type.String({ description: "Destination folder path relative to root." })),
      }),
      async execute(params, config) {
        try {
          return await moveOneDriveItem(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

    tool({
      name: "onedrive_delete",
      label: "OneDrive Delete",
      description: "Delete a OneDrive item addressed by item ID or path.",
      parameters: Type.Object({
        item_id: Type.Optional(Type.String({ description: "OneDrive item ID. Mutually exclusive with path." })),
        path: Type.Optional(Type.String({ description: "OneDrive item path relative to root." })),
      }),
      async execute(params, config) {
        try {
          return await deleteOneDriveItem(resolveConfig(config), params);
        } catch (e) {
          return { error: (e as Error).message };
        }
      },
    }),

  ],
});

export const TOOL_REQUIRED_FEATURE = {
  outlook_inbox: "mail-read",
  outlook_search: "mail-read",
  outlook_read: "mail-read",
  outlook_save_attachments: "mail-read",
  outlook_save_draft: "mail-write",
  outlook_send: "mail-send",
  outlook_reply: "mail-send",
  outlook_forward: "mail-send",
  outlook_move: "mail-write",
  outlook_flag: "mail-write",
  outlook_task_lists: "tasks-read",
  outlook_tasks: "tasks-read",
  outlook_create_task: "tasks-write",
  outlook_update_task: "tasks-write",
  outlook_complete_task: "tasks-write",
  outlook_delete_task: "tasks-write",
  outlook_calendar_fetch: "calendar-read",
  outlook_create_event: "calendar-write",
  outlook_update_event: "calendar-write",
  outlook_delete_event: "calendar-write",
  outlook_meeting: "calendar-write",
  outlook_query_events: "calendar-read",
  onedrive_list: "onedrive-read",
  onedrive_search: "onedrive-read",
  onedrive_metadata: "onedrive-read",
  onedrive_download: "onedrive-read",
  onedrive_upload: "onedrive-write",
  onedrive_create_folder: "onedrive-write",
  onedrive_move: "onedrive-write",
  onedrive_delete: "onedrive-write",
  outlook_search_store: "mail-read",
} as const satisfies Record<string, M365Feature>;

function configuredFeatures(config: Record<string, unknown> | undefined): M365Feature[] {
  if (config && Object.prototype.hasOwnProperty.call(config, "features")) {
    return parseM365Features(config.features as readonly unknown[]);
  }
  const envFeatures = process.env.M365_FEATURES?.trim();
  return parseM365Features(envFeatures || undefined);
}

export function createEntry(): PluginEntry {
  const entry = createBaseEntry();
  const registerAll = entry.register.bind(entry);
  return {
    ...entry,
    register(api: PluginApi) {
      const features = configuredFeatures(api.pluginConfig);
      registerAll({
        ...api,
        registerTool(tool: unknown) {
          const name = (tool as { name?: unknown })?.name;
          if (typeof name !== "string" || !(name in TOOL_REQUIRED_FEATURE)) {
            throw new Error(`Microsoft 365 tool is missing a feature policy: ${String(name)}`);
          }
          const required = TOOL_REQUIRED_FEATURE[
            name as keyof typeof TOOL_REQUIRED_FEATURE
          ];
          if (isM365FeatureEnabled(features, required)) {
            api.registerTool(tool);
          }
        },
      });
    },
  };
}

// ---------------------------------------------------------------------------
// Config resolver
// ---------------------------------------------------------------------------

export function resolveConfig(config: {
  clientId?: string;
  clientSecret?: string;
  refreshToken?: string;
  directTokenStatePath?: string;
  tenant?: string;
  tokenBrokerUrl?: string;
  tokenBrokerSecret?: string;
  features?: string[];
  personalCalendarNames?: string[];
  familyCalendarNames?: string[];
}): OutlookCalendarConfig {
  const env = (name: string): string => process.env[name]?.trim() || "";
  const parseNames = (values: string[] | undefined, key: string): string[] => {
    if (Array.isArray(values)) return values.map(n => n.trim().toLowerCase()).filter(Boolean);
    return env(key).split(",").map(n => n.trim().toLowerCase()).filter(Boolean);
  };
  const brokerSecret = config.tokenBrokerSecret?.trim()
    || env("M365_TOKEN_BROKER_SECRET");
  const brokerUrl = config.tokenBrokerUrl?.trim()
    || env("M365_TOKEN_BROKER_URL")
    || (brokerSecret ? DEFAULT_TOKEN_BROKER_URL : "");
  return {
    clientId: config.clientId?.trim() || env("M365_CLIENT_ID") || DEFAULT_CLIENT_ID,
    clientSecret: config.clientSecret?.trim() || env("M365_CLIENT_SECRET"),
    refreshToken: config.refreshToken?.trim() || env("M365_REFRESH_TOKEN"),
    directTokenStatePath: config.directTokenStatePath?.trim()
      || env("M365_DIRECT_TOKEN_STATE_PATH")
      || DEFAULT_DIRECT_TOKEN_STATE_PATH,
    tenant: config.tenant?.trim() || env("M365_TENANT") || "consumers",
    tokenBrokerUrl: brokerUrl,
    tokenBrokerSecret: brokerSecret,
    features: configuredFeatures(config),
    personalCalendarNames: parseNames(
      config.personalCalendarNames,
      "M365_PERSONAL_CALENDAR_NAMES",
    ),
    familyCalendarNames: parseNames(
      config.familyCalendarNames,
      "M365_FAMILY_CALENDAR_NAMES",
    ),
  };
}
