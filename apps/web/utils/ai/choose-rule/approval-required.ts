import { ActionType } from "@/generated/prisma/enums";

// Actions that lose mail or send mail or its content out of the mailbox. A rule never
// runs these on its own; they wait for the user's approval, whatever the rule's
// automation says. Drafts stay automatic because nothing leaves until the user sends.
const APPROVAL_REQUIRED_ACTION_TYPES: ReadonlySet<ActionType> = new Set([
  ActionType.DELETE,
  ActionType.FORWARD,
  ActionType.SEND_EMAIL,
  ActionType.REPLY,
  ActionType.CALL_WEBHOOK,
  ActionType.INTEGRATION,
]);

export function requiresApproval(type: ActionType) {
  return APPROVAL_REQUIRED_ACTION_TYPES.has(type);
}
