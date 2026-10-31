import { ActionType } from "@/generated/prisma/enums";

// Actions that lose mail or send it out of the mailbox. A rule never runs these on
// its own; they wait for the user's approval, whatever the rule's automation says.
const APPROVAL_REQUIRED_ACTION_TYPES: ReadonlySet<ActionType> = new Set([
  ActionType.DELETE,
  ActionType.FORWARD,
]);

export function requiresApproval(type: ActionType) {
  return APPROVAL_REQUIRED_ACTION_TYPES.has(type);
}
