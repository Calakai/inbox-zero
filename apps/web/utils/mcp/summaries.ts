// One sentence per write tool, built only from the counts and names the tool
// actually got back, so a model that repeats it cannot claim more than happened.

export function emails(n: number) {
  return count(n, "email");
}

export function archiveFolderSummary({
  path,
  archived,
  failed,
  moreLeft,
}: {
  path: string;
  archived: number;
  failed: number;
  moreLeft: boolean;
}) {
  return folderMoveSummary({
    verb: "Archived",
    destination: "",
    tool: "archive_folder",
    path,
    done: archived,
    failed,
    moreLeft,
  });
}

export function moveFolderToInboxSummary({
  path,
  moved,
  failed,
  moreLeft,
}: {
  path: string;
  moved: number;
  failed: number;
  moreLeft: boolean;
}) {
  return folderMoveSummary({
    verb: "Moved",
    destination: " to the Inbox",
    tool: "move_folder_to_inbox",
    path,
    done: moved,
    failed,
    moreLeft,
  });
}

export function markReadSummary({
  path,
  sender,
  read,
  marked,
  failed,
  moreLeft,
}: {
  path?: string;
  sender?: string;
  read: boolean;
  marked: number;
  failed: number;
  moreLeft: boolean;
}) {
  const scope = [sender && `from ${sender}`, path && `in ${path}`]
    .filter(Boolean)
    .join(" ");
  const state = read ? "read" : "unread";
  const before = read ? "unread" : "read";
  const total = marked + failed;
  const more = moreLeft ? "; more are left, so call mark_read again" : "";
  const stillUnchanged = `could not be changed and ${isOrAre(failed)} still ${before}`;

  if (total === 0) {
    if (moreLeft) return `Marked nothing ${scope}${more}.`;
    return `Marked nothing: no ${before} emails ${scope}.`;
  }
  if (marked === 0)
    return `Marked nothing: ${emails(failed)} ${scope} ${stillUnchanged}${more}.`;
  if (failed === 0) return `Marked ${emails(marked)} ${scope} ${state}${more}.`;
  return `Marked ${marked} of ${emails(total)} ${scope} ${state}; ${failed} ${stillUnchanged}${more}.`;
}

export function deleteFolderSummary(path: string) {
  return `Deleted the empty folder ${path}.`;
}

export function updateThreadSummary({
  action,
  trashFolder,
}: {
  action: "archive" | "unarchive" | "trash" | "mark_read" | "mark_unread";
  trashFolder: "Deleted Items" | "Trash";
}) {
  switch (action) {
    case "archive":
      return "Archived 1 conversation; unarchive moves it back.";
    case "unarchive":
      // Outlook moves only what sits in Archive, and a thread with nothing
      // there is left alone, so this states the end state rather than a move.
      return "The conversation has nothing left in Archive; anything that was there is back in the Inbox.";
    case "trash":
      return `Moved 1 conversation to ${trashFolder}; it can be restored.`;
    case "mark_read":
      return "Marked 1 conversation read.";
    case "mark_unread":
      return "Marked 1 conversation unread.";
  }
}

export function labelThreadSummary({
  labelName,
  removed,
  messages = 0,
}: {
  labelName: string;
  removed: boolean;
  messages?: number;
}) {
  if (removed)
    return `The conversation no longer has the label "${labelName}".`;
  if (messages === 0)
    return `Added nothing: the conversation has no messages, so the label "${labelName}" was not added.`;
  return `Added the label "${labelName}" to ${count(messages, "message")} in 1 conversation.`;
}

export function createDraftSummary({
  to,
  reply,
}: {
  to: string;
  reply: boolean;
}) {
  if (reply)
    return `Saved a reply draft to ${to} in the same conversation; it was not sent.`;
  return `Saved a draft to ${to}; it was not sent.`;
}

export function runRulesSummary({
  preview,
  results,
}: {
  preview: boolean;
  results: RuleResult[];
}) {
  const handled = results.filter((result) => result.alreadyProcessed);
  if (handled.length > 0) return handled.map(handledRuleSentence).join(" ");

  const matched = results.filter((result) => result.ruleName);
  if (preview) {
    if (matched.length === 0)
      return "Preview only, nothing changed: no rule matches this email.";
    return `Preview only, nothing changed: ${matched
      .map(
        (result) =>
          `the rule "${result.ruleName}" would match (${describeActions(result.actions)})`,
      )
      .join("; ")}.`;
  }

  if (matched.length === 0)
    return "No rule matched this email; nothing changed.";
  return matched.map(appliedRuleSentence).join(" ");
}

export function createRuleSummary(name: string) {
  return `Created the rule "${name}".`;
}

export function updateRuleSummary(name: string) {
  return `Replaced the rule "${name}" with the conditions and actions given.`;
}

export function deleteRuleSummary(name: string) {
  return `Deleted the rule "${name}".`;
}

type RuleResult = {
  ruleName: string | null;
  status: string;
  alreadyProcessed: boolean;
  actions: Array<{
    type: string;
    label: string | null;
    folderName: string | null;
    delayInMinutes?: number | null;
  }>;
};

function folderMoveSummary({
  verb,
  destination,
  tool,
  path,
  done,
  failed,
  moreLeft,
}: {
  verb: string;
  destination: string;
  tool: string;
  path: string;
  done: number;
  failed: number;
  moreLeft: boolean;
}) {
  const total = done + failed;
  const more = moreLeft ? `; more are left, so call ${tool} again` : "";
  const stillThere = `could not be moved and ${isOrAre(failed)} still there`;

  if (total === 0) {
    if (moreLeft) return `${verb} nothing from ${path}${more}.`;
    return `${verb} nothing: ${path} was already empty.`;
  }
  if (done === 0)
    return `${verb} nothing: ${emails(failed)} in ${path} ${stillThere}${more}.`;
  if (failed === 0)
    return `${verb} ${emails(done)} from ${path}${destination}${moreLeft ? more : "; it is now empty"}.`;
  return `${verb} ${done} of ${emails(total)} from ${path}${destination}; ${failed} ${stillThere}${more}.`;
}

function appliedRuleSentence(result: RuleResult) {
  const actions = describeActions(result.actions);
  switch (result.status) {
    case "ERROR":
      return `The rule "${result.ruleName}" matched but failed while running (${actions}); some of it may have gone through.`;
    case "SKIPPED":
      return `The rule "${result.ruleName}" matched but was skipped; nothing changed.`;
  }

  const delayed = result.actions.filter(isDelayed);
  if (delayed.length > 0) {
    const ran = result.actions.filter((action) => !isDelayed(action));
    const later = delayed.map(describeAction).join(", ");
    if (ran.length === 0)
      return `Started the rule "${result.ruleName}"; scheduled for later: ${later}.`;
    return `Ran the rule "${result.ruleName}" (${ran.map(describeAction).join(", ")}); scheduled for later: ${later}.`;
  }
  if (result.status === "APPLYING")
    return `Started the rule "${result.ruleName}" (${actions}); some of its actions are scheduled for later.`;
  return `Ran the rule "${result.ruleName}" (${actions}).`;
}

function handledRuleSentence(result: RuleResult) {
  if (!result.ruleName)
    return "Already handled earlier: no rule matched this email; nothing ran again.";
  const actions = describeActions(result.actions);
  switch (result.status) {
    case "ERROR":
      return `The rule "${result.ruleName}" already ran on this email earlier and failed (${actions}); some of it may have gone through, and nothing ran again.`;
    case "APPLYING":
      return `The rule "${result.ruleName}" already started on this email earlier (${actions}); some of its actions are scheduled for later, and nothing ran again.`;
    case "SKIPPED":
      return `The rule "${result.ruleName}" matched this email earlier but was skipped; nothing ran again.`;
    default:
      return `Already handled earlier by the rule "${result.ruleName}" (${actions}); nothing ran again.`;
  }
}

// Delayed actions are only scheduled when a rule runs, so they are marked
// "later" wherever actions are listed, and never counted as done.
function describeActions(actions: RuleResult["actions"]) {
  if (actions.length === 0) return "no actions";
  return actions
    .map((action) =>
      isDelayed(action)
        ? `${describeAction(action)} later`
        : describeAction(action),
    )
    .join(", ");
}

function describeAction({ type, label, folderName }: RuleResult["actions"][0]) {
  if (type === "LABEL" && label) return `label "${label}"`;
  if (type === "MOVE_FOLDER" && folderName)
    return `move to the folder ${folderName}`;
  return ACTION_PHRASES[type] ?? type.toLowerCase().replaceAll("_", " ");
}

function isDelayed(action: RuleResult["actions"][0]) {
  return (action.delayInMinutes ?? 0) > 0;
}

const ACTION_PHRASES: Record<string, string> = {
  ARCHIVE: "archive",
  LABEL: "label",
  REPLY: "reply",
  SEND_EMAIL: "send an email",
  FORWARD: "forward",
  DRAFT_EMAIL: "draft a reply",
  DRAFT_MESSAGING_CHANNEL: "draft in a messaging channel",
  NOTIFY_MESSAGING_CHANNEL: "notify a messaging channel",
  MARK_SPAM: "mark as spam",
  CALL_WEBHOOK: "call a webhook",
  MARK_READ: "mark read",
  STAR: "star",
  DELETE: "delete",
  DIGEST: "add to the digest",
  MOVE_FOLDER: "move to a folder",
  NOTIFY_SENDER: "notify the sender",
  INTEGRATION: "run an integration",
};

function count(n: number, noun: string) {
  return `${n} ${noun}${n === 1 ? "" : "s"}`;
}

function isOrAre(n: number) {
  return n === 1 ? "is" : "are";
}
